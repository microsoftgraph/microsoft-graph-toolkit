// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import jwt, { SigningKeyCallback, JwtHeader } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import * as msal from '@azure/msal-node';

const keyClient = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/common/discovery/keys'
});

/**
 * Parses the JWT header and retrieves the appropriate public key
 * @param {JwtHeader} header - The JWT header
 * @param {SigningKeyCallback} callback - Callback function
 */
function getSigningKey(header: JwtHeader, callback: SigningKeyCallback): void {
  if (header) {
    keyClient.getSigningKey(header.kid!, (err, key) => {
      if (err) {
        callback(err, undefined);
      } else {
        callback(null, key.getPublicKey());
      }
    });
  }
}

/**
 * Validates a JWT
 * @param {string} authHeader - The Authorization header value containing a JWT bearer token
 * @returns {Promise<string | null>} - Returns the token if valid, returns null if invalid
 */
async function validateJwt(authHeader: string): Promise<string | null> {
  return new Promise((resolve, reject) => {
    const token = authHeader.split(' ')[1];

    const validationOptions = {
      audience: process.env.PROXY_APP_ID,
      issuer: `https://login.microsoftonline.com/${process.env.PROXY_APP_TENANT_ID}/v2.0`
    };

    jwt.verify(token, getSigningKey, validationOptions, (err, payload) => {
      if (err) {
        resolve(null);
      }

      resolve(token);
    });
  });
}

/**
 * Gets an access token for the user using the on-behalf-of flow
 * @param authHeader - The Authorization header value containing a JWT bearer token
 * @returns {Promise<string | null>} - Returns the access token if successful, null if not
 */
export default async function getAccessTokenOnBehalfOf(authHeader: string): Promise<string | null> {
  // Validate the token
  const token = await validateJwt(authHeader);

  if (token) {
    // Create an MSAL client
    const msalClient = new msal.ConfidentialClientApplication({
      auth: {
        clientId: process.env.PROXY_APP_ID!,
        clientSecret: process.env.PROXY_APP_SECRET
      }
    });

    try {
      // Make the on-behalf-of request
      // This exchanges the incoming token (which is scoped for the proxy service)
      // for a new token that is scoped for Microsoft Graph
      const result = await msalClient.acquireTokenOnBehalfOf({
        oboAssertion: token,
        skipCache: true,
        scopes: ['https://graph.microsoft.com/.default']
      });

      return result.accessToken;
    } catch (error) {
      console.log(`Token error: ${error}`);
      return null;
    }

  } else {
    return null;
  }
}
