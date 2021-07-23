// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import dotenv from 'dotenv';
import express from 'express';
import path from 'path';
import { getAccessTokenOnBehalfOf, getAccessTokenOnBehalfOfPost, validateJwt } from './auth';

var cors = require('cors');

// Load .env file
dotenv.config();

const app = express();
const PORT = process.env.port || process.env.PORT || 5000;

// Support JSON payloads
app.use(express.json());
app.use(express.static(path.join(__dirname, 'client')));

// enable cors for localhost for this sample
app.use(cors());

// An example for using GET and with token validation using middleware
app.get('/api/token', validateJwt, async (req, res) => {
  await getAccessTokenOnBehalfOf(req, res);
});

// An example for using POST and with token validation using middleware
app.post('/api/token', validateJwt, async (req, res) => {
  await getAccessTokenOnBehalfOfPost(req, res);
});

app.listen(PORT, () => {
  console.log(`⚡️[server]: Server is running at http://localhost:${PORT}`);
});
