// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

import * as msRest from "ms-rest-js";
import Constants from "./util/constants";
import PollingState from "./pollingState";
import { DefaultAzureHttpPipelineOptions, createDefaultAzureHttpPipeline } from "./azureHttpPipeline";
import { HttpMethod } from "ms-rest-js";
const LroStates = Constants.LongRunningOperationStates;

export class AzureServiceClient extends msRest.ServiceClient {
  longRunningOperationRetryTimeout = 30;

  /**
   * Initializes a new instance of the AzureServiceClient class.
   * @param {msRest.HttpPipeline | DefaultAzureHttpPipelineOptions} httpPipeline - The HttpPipeline
   * that this AzureServiceClient will use to send HttpRequests, or the
   * DefaultAzureHttpPipelineOptions that will be used to create the default HttpPipeline for
   * AzureServiceClients.
   */
  constructor(httpPipeline?: msRest.HttpPipeline | DefaultAzureHttpPipelineOptions) {
    super(httpPipeline instanceof msRest.HttpPipeline ? httpPipeline : createDefaultAzureHttpPipeline(httpPipeline));
  }

  /**
   * Provides a mechanism to make a request that will poll and provide the final result.
   * @param {msRest.RequestPrepareOptions|msRest.WebResource} request - The request object
   * @param {msRest.RequestOptionsBase} [options] Additional options to be sent while making the request
   * @returns {Promise<msRest.HttpOperationResponse>} The HttpOperationResponse containing the final polling request, response and the responseBody.
   */
  async sendLongRunningRequest(request: msRest.HttpRequest): Promise<msRest.HttpResponse> {
    const initialResponse: msRest.HttpResponse = await this.sendRequest(request);
    return await this.getLongRunningOperationResult(initialResponse);
  }

  /**
   * Poll Azure long running PUT, PATCH, POST or DELETE operations.
   * @param {msRest.HttpResponse} initialResponse - result/response of the initial request which is
   * a part of the asynchronous polling operation.
   * @returns {Promise<msRest.HttpResponse>} result - The final response after polling is complete.
   */
  async getLongRunningOperationResult(initialResponse: msRest.HttpResponse): Promise<msRest.HttpResponse> {
    const initialRequestMethod: msRest.HttpMethod = initialResponse.request.method as msRest.HttpMethod;

    if (this.checkResponseStatusCodeFailed(initialResponse)) {
      throw new Error(`Unexpected polling status code from long running operation "${initialResponse.statusCode}" for method "${initialRequestMethod}".`);
    }

    const pollingState = new PollingState(initialResponse, this.longRunningOperationRetryTimeout);

    const resourceUrl: string = initialResponse.request.url;
    while (![LroStates.Succeeded, LroStates.Failed, LroStates.Canceled].some((e) => { return e === pollingState.status; })) {
      await msRest.delay(pollingState.getTimeoutInMilliseconds());
      if (pollingState.azureAsyncOperationHeaderLink) {
        await this.updateStateFromAzureAsyncOperationHeader(pollingState, true);
      } else if (pollingState.locationHeaderLink) {
        await this.updateStateFromLocationHeader(initialRequestMethod, pollingState);
      } else if (initialRequestMethod === "PUT") {
        await this.updateStateFromGetResourceOperation(resourceUrl, pollingState);
      } else {
        throw new Error("Location header is missing from long running operation.");
      }
    }

    if (pollingState.status === LroStates.Succeeded) {
      if ((pollingState.azureAsyncOperationHeaderLink || !pollingState.resource) &&
        (initialRequestMethod === "PUT" || initialRequestMethod === "PATCH")) {
        await this.updateStateFromGetResourceOperation(resourceUrl, pollingState);
      }
    } else {
      throw pollingState.getRestError();
    }

    return pollingState.latestResponse;
  }

  /**
   * Verified whether an unexpected polling status code for long running operation was received for the response of the initial request.
   * @param {msRest.HttpResponse} initialResponse - Response to the initial request that was sent as a part of the asynchronous operation.
   */
  private checkResponseStatusCodeFailed(initialResponse: msRest.HttpResponse): boolean {
    const statusCode: number = initialResponse.statusCode;
    const method: msRest.HttpMethod = initialResponse.request.method as msRest.HttpMethod;
    if (statusCode === 200 || statusCode === 202 ||
      (statusCode === 201 && method === msRest.HttpMethod.PUT) ||
      (statusCode === 204 && (method === msRest.HttpMethod.DELETE || method === msRest.HttpMethod.POST))) {
      return false;
    } else {
      return true;
    }
  }

  /**
   * Retrieve operation status by polling from "azure-asyncoperation" header.
   * @param {PollingState} pollingState - The object to persist current operation state.
   * @param {boolean} inPostOrDelete - Invoked by Post Or Delete operation.
   */
  private async updateStateFromAzureAsyncOperationHeader(pollingState: PollingState, inPostOrDelete = false): Promise<void> {
    const statusResponse: msRest.HttpResponse = await this.getStatus(pollingState.azureAsyncOperationHeaderLink as string);

    const parsedResponse: { [propertyName: string]: any } = await statusResponse.deserializedBody();

    if (!parsedResponse) {
      throw new Error("The response from long running operation does not contain a body.");
    } else if (parsedResponse && !parsedResponse.status) {
      throw new Error(`The response "${JSON.stringify(parsedResponse)}" from long running operation does not contain the status property.`);
    }
    pollingState.status = parsedResponse.status;
    pollingState.error = parsedResponse.error;
    pollingState.updateResponse(statusResponse);
    pollingState.resource = undefined;
    if (inPostOrDelete) {
      pollingState.resource = parsedResponse;
    }
  }

  /**
   * Retrieve PUT operation status by polling from "location" header.
   * @param {string} method - The HTTP method.
   * @param {PollingState} pollingState - The object to persist current operation state.
   */
  private async updateStateFromLocationHeader(method: string, pollingState: PollingState): Promise<void> {
    const statusResponse: msRest.HttpResponse = await this.getStatus(pollingState.locationHeaderLink as string);

    const parsedResponse: { [propertyName: string]: any } = await statusResponse.deserializedBody();

    pollingState.updateResponse(statusResponse);
    const statusCode: number = statusResponse.statusCode;
    if (statusCode === 202) {
      pollingState.status = LroStates.InProgress;
    } else if (statusCode === 200 ||
      (statusCode === 201 && (method === msRest.HttpMethod.PUT || method === msRest.HttpMethod.PATCH)) ||
      (statusCode === 204 && (method === msRest.HttpMethod.DELETE || method === msRest.HttpMethod.POST))) {
      pollingState.status = LroStates.Succeeded;
      pollingState.resource = parsedResponse;
      // we might not throw an error, but initialize here just in case.
      pollingState.error = new msRest.RestError(`Long running operation failed with status "${pollingState.status}".`, {
        code: pollingState.status
      });
    } else {
      throw new Error(`The response with status code ${statusCode} from polling for long running operation url "${pollingState.locationHeaderLink}" is not valid.`);
    }
  }

  /**
   * Polling for resource status.
   * @param {string} resourceUrl - The url of resource.
   * @param {PollingState} pollingState - The object to persist current operation state.
   */
  private async updateStateFromGetResourceOperation(resourceUrl: string, pollingState: PollingState): Promise<void> {
    const statusResponse: msRest.HttpResponse = await this.getStatus(resourceUrl);

    const deserializedBody: { [propertyName: string]: any } = await statusResponse.deserializedBody();
    if (!deserializedBody) {
      throw new Error("The response from long running operation does not contain a body.");
    }

    pollingState.status = LroStates.Succeeded;
    if (deserializedBody && deserializedBody.properties && deserializedBody.properties.provisioningState) {
      pollingState.status = deserializedBody.properties.provisioningState;
    }
    pollingState.updateResponse(statusResponse);
    pollingState.resource = deserializedBody;
    // we might not throw an error, but initialize here just in case.
    pollingState.error = new msRest.RestError(`Long running operation failed with status "${pollingState.status}".`, {
      code: pollingState.status
    });
  }

  /**
   * Retrieves operation status by querying the operation URL.
   * @param {string} operationUrl - URL used to poll operation result.
   */
  private async getStatus(operationUrl: string): Promise<msRest.HttpResponse> {
    // Construct URL
    const requestUrl: string = operationUrl.replace(" ", "%20");

    // Create HTTP request object
    const httpRequest = new msRest.HttpRequest({
      method: HttpMethod.GET,
      url: requestUrl
    });

    const statusResponse: msRest.HttpResponse = await this.sendRequest(httpRequest);
    const statusCode: number = statusResponse.statusCode;
    const responseBody: { [propertyName: string]: any } = await statusResponse.deserializedBody();
    if (statusCode !== 200 && statusCode !== 201 && statusCode !== 202 && statusCode !== 204) {
      throw new msRest.RestError(`Invalid status code with response body "${JSON.stringify(responseBody)}" occurred when polling for operation status.`, {
        statusCode: statusCode,
        request: statusResponse.request,
        response: statusResponse,
        body: responseBody
      });
    }

    return statusResponse;
  }
}