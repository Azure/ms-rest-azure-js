// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.

export type Authenticator = (challenge: object, callback: (error: Error, authorizationValue: string) => void) => void;
