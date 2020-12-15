// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import dotenv from 'dotenv';
import express from 'express';
import path from 'path';
import { getAccessTokenOnBehalfOf, validateJwt } from './auth';

// Load .env file
dotenv.config();

const app = express();
const PORT = process.env.port || process.env.PORT || 5000;

// Support JSON payloads
app.use(express.json());
app.use(express.static(path.join(__dirname, 'client')));

// Validate the Jwt token using middleware
app.get('/api/token', validateJwt, async (req, res) => {
  await getAccessTokenOnBehalfOf(req, res);
});

app.listen(PORT, () => {
  console.log(`⚡️[server]: Server is running at http://localhost:${PORT}`);
});
