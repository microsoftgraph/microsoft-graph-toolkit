// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import express from 'express';
import Router from 'express-promise-router';
import dotenv from 'dotenv';
import path from 'path';

import proxyRequest from './proxy';

// Load .env file
dotenv.config();

const app = express();
const PORT = 8000;
const router = Router();

// Support JSON payloads
app.use(express.json());
app.use(express.static(path.join(__dirname, 'client')));

router.get('/*', async (req, res) => await proxyRequest(req, res));
router.post('/*', async (req, res) => await proxyRequest(req, res));
router.put('/*', async (req, res) => await proxyRequest(req, res));
router.patch('/*', async (req, res) => await proxyRequest(req, res));
router.delete('/*', async (req, res) => await proxyRequest(req, res));

app.use('/apiproxy', router);

app.listen(PORT, () => {
  console.log(`⚡️[server]: Server is running at http://localhost:${PORT}`);
  console.log(`Auth mode: ${process.env.AUTH_PASS_THROUGH == 'true' ? 'pass-through' : 'on-behalf-of'}`);
});
