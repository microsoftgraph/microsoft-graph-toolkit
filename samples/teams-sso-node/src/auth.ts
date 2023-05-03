// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as msal from '@azure/msal-node';
import { NextFunction, Request, Response } from 'express';
import jwt, { JwtHeader, SigningKeyCallback } from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';
import jwt_decode from 'jwt-decode';

/**
 * Validates a JWT
 *
 * @param {Request} req - The incoming request
 * @param {Response} res - The outgoing response
 * @returns {Promise<string | null>} - Returns the token if valid, returns null if invalid
 */
export const validateJwt = (req: Request, res: Response, next: NextFunction): void => {
  // eslint-disable-next-line @typescript-eslint/no-non-null-assertion, @typescript-eslint/no-unnecessary-type-assertion
  const authHeader = req.headers.authorization!;
  const ssoToken = req.query.ssoToken ? req.query.ssoToken?.toString() : authHeader.split(' ')[1];
  if (ssoToken) {
    const validationOptions = {
      audience: process.env.CLIENT_ID
    };
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    jwt.verify(ssoToken, getSigningKey, validationOptions, (err, _payload) => {
      if (err) {
        return res.sendStatus(403);
      }
      next();
    });
  } else {
    res.sendStatus(401);
  }
};

/**
 * Parses the JWT header and retrieves the appropriate public key
 *
 * @param {JwtHeader} header - The JWT header
 * @param {SigningKeyCallback} callback - Callback function
 */
const getSigningKey = (header: JwtHeader, callback: SigningKeyCallback): void => {
  const client = jwksClient({
    jwksUri: 'https://login.microsoftonline.com/common/discovery/keys'
  });
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-assertion, @typescript-eslint/no-non-null-assertion
  client.getSigningKey(header.kid!, (err, key) => {
    if (err) {
      callback(err, undefined);
    } else {
      callback(null, key.getPublicKey());
    }
  });
};

/**
 * Gets an access token for the user using the on-behalf-of flow
 *
 * @param authHeader - The Authorization header value containing a JWT bearer token
 * @returns {Promise<string | null>} - Returns the access token if successful, null if not
 */
export const getAccessTokenOnBehalfOfPost = async (req: Request, res: Response): Promise<void> => {
  // The token has already been validated, just grab it
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-assertion, @typescript-eslint/no-non-null-assertion
  const authHeader = req.headers.authorization!;
  const ssoToken = authHeader.split(' ')[1];

  // Create an MSAL client
  const msalClient = new msal.ConfidentialClientApplication({
    auth: {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
      clientId: req.body.clientid,
      clientSecret: process.env.APP_SECRET
    }
  });

  try {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    const tid = jwt_decode<any>(ssoToken).tid as string;
    const result = await msalClient.acquireTokenOnBehalfOf({
      authority: `https://login.microsoftonline.com/${tid}`,
      oboAssertion: ssoToken,
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
      scopes: req.body.scopes as string[],
      skipCache: true
    });
    res.json({ access_token: result?.accessToken });
  } catch (error: any) {
    if (error instanceof msal.AuthError) {
      if (error.errorCode === 'invalid_grant' || error.errorCode === 'interaction_required') {
        // This is expected if it's the user's first time running the app ( user must consent ) or the admin requires MFA
        res.status(403).json({ error: 'consent_required' }); // This error triggers the consent flow in the client.
      } else {
        // Unknown error
        res.status(500).json({ error: error.message });
      }
    }
  }
};

/**
 * Gets an access token for the user using the on-behalf-of flow
 *
 * @returns {Promise<string | null>} - Returns the access token if successful, null if not
 */
export const getAccessTokenOnBehalfOf = async (req: Request, res: Response): Promise<void> => {
  // The token has already been validated.
  const ssoToken: string = req.query.ssoToken?.toString() || '';
  const clientId: string = req.query.clientId?.toString() || '';
  const graphScopes: string[] = req.query.scopes?.toString().split(',') || [];

  // Get tenantId from the SSO Token
  // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
  const tenantId: string = jwt_decode<any>(ssoToken).tid as string;
  const clientSecret: string = process.env.APP_SECRET || '';

  // Create an MSAL client
  const msalClient = new msal.ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret
    }
  });

  try {
    const result = await msalClient.acquireTokenOnBehalfOf({
      authority: `https://login.microsoftonline.com/${tenantId}`,
      oboAssertion: ssoToken,
      scopes: graphScopes,
      skipCache: true
    });
    res.json({ access_token: result?.accessToken });
  } catch (error: any) {
    if (error instanceof msal.AuthError) {
      if (error.errorCode === 'invalid_grant' || error.errorCode === 'interaction_required') {
        // This is expected if it's the user's first time running the app ( user must consent ) or the admin requires MFA
        res.status(403).json({ error: 'consent_required' }); // This error triggers the consent flow in the client.
      } else {
        // Unknown error
        res.status(500).json({ error: error.message });
      }
    }
  }
};
