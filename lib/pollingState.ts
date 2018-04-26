// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

import Constants from "./util/constants";
import * as msRest from "ms-rest-js";
const LroStates = Constants.LongRunningOperationStates;

/**
 * @class
 * Initializes a new instance of the PollingState class.
 */
export default class PollingState {
  /**
   * @param {msRest.WebResource} [request] - provides information about the request made for polling.
   */
  request: msRest.WebResource;
  /**
   * @param {Response} [response] - The response object to extract longrunning operation status.
   */
  private _latestResponse: msRest.HttpResponse;
  /**
   * @param {any} [resource] - Provides information about the response body received in the polling request. Particularly useful when polling via provisioningState.
   */
  resource: any;
  /**
   * @param {string} [azureAsyncOperationHeaderLink] - The url that is present in "azure-asyncoperation" response header.
   */
  azureAsyncOperationHeaderLink?: string;
  /**
   * @param {string} [locationHeaderLink] - The url that is present in "Location" response header.
   */
  locationHeaderLink?: string;
  /**
   * @param {string} [status] - The status of polling. "Succeeded, Failed, Cancelled, Updating, Creating, etc."
   */
  status?: string;
  /**
   * @param {msRest.RestError} [error] - Provides information about the error that happened while polling.
   */
  error?: msRest.RestError;

  /**
   * Create a new PollingState object.
   * @param {msRest.HttpResponse} _initialResponse - Response of the initial request that was made as a part of the asynchronous operation.
   * @param {number} _retryTimeoutInSeconds - The timeout in seconds to retry on intermediate operation results. Default Value is 30.
   */
  constructor(private readonly _initialResponse: msRest.HttpResponse, private _retryTimeoutInSeconds: number) {
    this.updateResponse(this._initialResponse);

    // Parse response.body & assign it as the resource.
    try {
      if (this._initialResponse.bodyAsText && this._initialResponse.bodyAsText.length > 0) {
        this.resource = JSON.parse(this._initialResponse.bodyAsText);
      } else {
        this.resource = this._initialResponse.parsedBody;
      }
    } catch (error) {
      const deserializationError = new msRest.RestError(`Error "${error}" occurred in parsing the responseBody " +
        "while creating the PollingState for Long Running Operation- "${this._initialResponse.bodyAsText}"`);
      deserializationError.request = this._initialResponse.request;
      deserializationError.response = this._initialResponse.response;
      throw deserializationError;
    }
    switch (this.response.status) {
      case 202:
        this.status = LroStates.InProgress;
        break;

      case 204:
        this.status = LroStates.Succeeded;
        break;

      case 201:
        if (this.resource && this.resource.properties && this.resource.properties.provisioningState) {
          this.status = this.resource.properties.provisioningState;
        } else {
          this.status = LroStates.InProgress;
        }
        break;

      case 200:
        if (this.resource && this.resource.properties && this.resource.properties.provisioningState) {
          this.status = this.resource.properties.provisioningState;
        } else {
          this.status = LroStates.Succeeded;
        }
        break;

      default:
        this.status = LroStates.Failed;
        break;
    }
  }

  /**
   * Update cached data using the provided response object
   * @param {Response} [response] - provider response object.
   */
  updateResponse(response: msRest.HttpResponse) {
    this._latestResponse = response;
    if (response && response.headers) {
      const asyncOperationHeader: string | undefined = response.headers.get("azure-asyncoperation");
      if (asyncOperationHeader) {
        this.azureAsyncOperationHeaderLink = asyncOperationHeader;
      }

      const locationHeader: string | undefined = response.headers.get("location");
      if (locationHeader) {
        this.locationHeaderLink = locationHeader;
      }

      const retryAfterHeader: string | undefined = response.headers.get("retry-after");
      if (retryAfterHeader) {
        this._retryTimeoutInSeconds = parseInt(retryAfterHeader);
      }
    }
  }

  /**
   * Gets timeout in seconds.
   * @returns {number} timeout
   */
  getTimeoutInSeconds() {
    return this._retryTimeoutInSeconds;
  }

  /**
   * Gets timeout in millisecondsseconds.
   * @returns {number} timeout
   */
  getTimeoutInMilliseconds() {
    return this.getTimeoutInSeconds() * 1000;
  }

  /**
   * Returns long running operation result.
   * @returns {msRest.HttpOperationResponse} HttpOperationResponse
   */
  getOperationResponse(): msRest.HttpOperationResponse {
    const result = new msRest.HttpOperationResponse(this.request, this.response);
    if (this.resource && typeof this.resource.valueOf() === "string") {
      result.bodyAsText = this.resource;
      result.parsedBody = JSON.parse(this.resource);
    } else {
      result.parsedBody = this.resource;
      result.bodyAsText = JSON.stringify(this.resource);
    }
    return result;
  }

  /**
   * Returns an Error on operation failure.
   * @param {Error} err - The error object.
   * @returns {msRest.RestError} The RestError defined in the runtime.
   */
  getRestError(err?: Error): msRest.RestError {
    let errMsg: string;
    let errCode: string | undefined = undefined;

    const error = new msRest.RestError("");
    error.request = msRest.stripRequest(this.request);
    error.response = this.response;
    const parsedResponse = this.resource as { [key: string]: any };

    if (err && err.message) {
      errMsg = `Long running operation failed with error: "${err.message}".`;
    } else {
      errMsg = `Long running operation failed with status: "${this.status}".`;
    }

    if (parsedResponse) {
      if (parsedResponse.error && parsedResponse.error.message) {
        errMsg = `Long running operation failed with error: "${parsedResponse.error.message}".`;
      }
      if (parsedResponse.error && parsedResponse.error.code) {
        errCode = parsedResponse.error.code as string;
      }
    }

    error.message = errMsg;
    if (errCode) error.code = errCode;
    error.body = parsedResponse;
    return error;
  }
}