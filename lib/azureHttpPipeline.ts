// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
import * as msRest from "ms-rest-js";
import Constants from "./util/constants";

/**
 * Options that can be used to configure the default HttpPipeline configuration.
 */
export interface DefaultAzureHttpPipelineOptions extends msRest.DefaultHttpPipelineOptions {
  /**
   * @property {number} [options.longRunningOperationRetryTimeout] - Gets or sets the retry timeout in seconds for
   * Long Running Operations. Default value is 30.
   */
  longRunningOperationRetryTimeout?: number;
}

/**
 * Get the default HttpPipeline.
 */
export function createDefaultAzureHttpPipeline(options?: DefaultAzureHttpPipelineOptions): msRest.HttpPipeline {
  if (!options) {
    options = {};
  }

  const requestPolicyFactories: msRest.RequestPolicyFactory[] = [];

  if (options.credentials) {
    requestPolicyFactories.push(msRest.signingPolicy(options.credentials));
  }

  if (options.generateClientRequestId) {
    requestPolicyFactories.push(msRest.generateClientRequestIdPolicy());
  }

  if (msRest.isNode) {
    if (!options.nodeJsUserAgentPackage) {
      options.nodeJsUserAgentPackage = `ms-rest-azure-js/${Constants.msRestAzureVersion}`;
    }
    requestPolicyFactories.push(msRest.msRestNodeJsUserAgentPolicy([options.nodeJsUserAgentPackage]));
  }

  requestPolicyFactories.push(msRest.serializationPolicy(options.serializationOptions));

  requestPolicyFactories.push(msRest.redirectPolicy());
  requestPolicyFactories.push(msRest.rpRegistrationPolicy(options.rpRegistrationRetryTimeoutInSeconds));

  if (options.addRetryPolicies) {
    requestPolicyFactories.push(msRest.exponentialRetryPolicy());
    requestPolicyFactories.push(msRest.systemErrorRetryPolicy());
  }

  const httpPipelineOptions: msRest.HttpPipelineOptions = {
    httpClient: options.httpClient || msRest.getDefaultHttpClient(),
    pipelineLogger: options.logger
  };

  return new msRest.HttpPipeline(requestPolicyFactories, httpPipelineOptions);
}