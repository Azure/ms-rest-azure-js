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
   * @param {msRest.HttpRequest} [request] - provides information about the request made for polling.
   */
  public request: msRest.HttpRequest;
  /**
   * @param {Response} [response] - The response object to extract longrunning operation status.
   */
  public latestResponse: msRest.HttpResponse;
  /**
   * @param {any} [resource] - Provides information about the response body received in the polling request. Particularly useful when polling via provisioningState.
   */
  public resource: any;
  /**
   * @param {string} [azureAsyncOperationHeaderLink] - The url that is present in "azure-asyncoperation" response header.
   */
  public azureAsyncOperationHeaderLink?: string;
  /**
   * @param {string} [locationHeaderLink] - The url that is present in "Location" response header.
   */
  public locationHeaderLink?: string;
  /**
   * @param {string} [status] - The status of polling. "Succeeded, Failed, Cancelled, Updating, Creating, etc."
   */
  public status?: string;
  /**
   * @param {msRest.RestError} [error] - Provides information about the error that happened while polling.
   */
  public error?: msRest.RestError;

  /**
   * Create a new PollingState object.
   * @param {msRest.HttpResponse} _initialResponse - Response of the initial request that was made as a part of the asynchronous operation.
   * @param {number} _retryTimeoutInSeconds - The timeout in seconds to retry on intermediate operation results. Default Value is 30.
   */
  constructor(private readonly _initialResponse: msRest.HttpResponse, private _retryTimeoutInSeconds: number) {
    this.request = this._initialResponse.request;
    this.latestResponse = this._initialResponse;

    this.updateResponse(this._initialResponse);

    // Parse response.body & assign it as the resource.
    try {
      this.resource = this._initialResponse.deserializedBody();
    } catch (error) {
      throw new msRest.RestError(`Error "${error}" occurred in parsing the responseBody while creating the PollingState for Long Running Operation`, {
        request: this._initialResponse.request,
        response: this._initialResponse
      });
    }

    switch (this.latestResponse.statusCode) {
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
    this.latestResponse = response;
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
   * Returns an Error on operation failure.
   * @param {Error} err - The error object.
   * @returns {msRest.RestError} The RestError defined in the runtime.
   */
  getRestError(err?: Error): msRest.RestError {
    let errorMessage: string;
    if (err && err.message) {
      errorMessage = `Long running operation failed with error: "${err.message}".`;
    } else {
      errorMessage = `Long running operation failed with status: "${this.status}".`;
    }
    
    let errorCode: string | undefined = undefined;
    if (this.resource) {
      if (this.resource.error && this.resource.error.message) {
        errorMessage = `Long running operation failed with error: "${this.resource.error.message}".`;
      }
      if (this.resource.error && this.resource.error.code) {
        errorCode = this.resource.error.code as string;
      }
    }

    return new msRest.RestError(errorMessage, {
      request: this.request,
      response: this.latestResponse,
      code: errorCode,
      body: this.resource
    });
  }
}