'use strict';

// Header: MCP HTTP server for Google Sheets and Drive tools

const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));

// Utilities
function getBearerToken(req) {
  const header = req.get('Authorization') || req.get('authorization');
  if (!header) throw new Error('Missing Authorization header');
  const match = header.match(/^Bearer\s+(.+)$/i);
  if (!match) throw new Error('Invalid Authorization header format');
  return match[1];
}

function getGoogleClients(req) {
  const token = getBearerToken(req);
  const oauth2 = new google.auth.OAuth2();
  oauth2.setCredentials({ access_token: token });
  const sheets = google.sheets({ version: 'v4', auth: oauth2 });
  const drive = google.drive({ version: 'v3', auth: oauth2 });
  return { oauth2, sheets, drive };
}

async function getSheetIdByName(sheetsApi, spreadsheetId, sheetName) {
  const meta = await sheetsApi.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(sheetId,title))',
  });
  const sheet = (meta.data.sheets || []).find(
    (s) => s.properties && s.properties.title === sheetName
  );
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);
  return sheet.properties.sheetId;
}

function columnLettersToIndex(letters) {
  let index = 0;
  for (let i = 0; i < letters.length; i++) {
    index *= 26;
    index += letters.charCodeAt(i) - 64; // 'A' => 1
  }
  return index - 1; // zero-based
}

async function parseA1ToGridRange(sheetsApi, spreadsheetId, a1Range) {
  // Expecting "SheetName!A1:B2"
  const [sheetPart, rangePartRaw] = a1Range.split('!');
  if (!rangePartRaw) {
    throw new Error('range must include sheet name, e.g., "Sheet1!A1:B2"');
  }
  const sheetName = sheetPart.replace(/^'(.+)'$/, '$1');
  const sheetId = await getSheetIdByName(sheetsApi, spreadsheetId, sheetName);

  const rangePart = rangePartRaw.toUpperCase();
  // Support A1:B2 only (no open-ended ranges)
  const m = rangePart.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
  if (!m) throw new Error('range must be like A1:B2');
  const [, c1, r1, c2, r2] = m;
  const startColumnIndex = columnLettersToIndex(c1);
  const endColumnIndex = columnLettersToIndex(c2) + 1;
  const startRowIndex = parseInt(r1, 10) - 1;
  const endRowIndex = parseInt(r2, 10);

  return { sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex };
}

// Tool implementations
const toolHandlers = {
  // Google Sheets tools
  'sheets.create_spreadsheet': async ({ title }, { sheets }) => {
    const res = await sheets.spreadsheets.create({ requestBody: { properties: { title } } });
    return res.data;
  },

  'sheets.get_values': async ({ spreadsheetId, range }, { sheets }) => {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    return res.data;
  },

  'sheets.update_values': async ({ spreadsheetId, range, values }, { sheets }) => {
    const res = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      requestBody: { values },
    });
    return res.data;
  },

  'sheets.append_values': async ({ spreadsheetId, range, values }, { sheets }) => {
    const res = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values },
    });
    return res.data;
  },

  'sheets.clear_values': async ({ spreadsheetId, range }, { sheets }) => {
    const res = await sheets.spreadsheets.values.clear({ spreadsheetId, range, requestBody: {} });
    return res.data;
  },

  'sheets.create_sheet': async ({ spreadsheetId, sheetName }, { sheets }) => {
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetName,
              },
            },
          },
        ],
      },
    });
    return res.data;
  },

  'sheets.delete_sheet': async ({ spreadsheetId, sheetName }, { sheets }) => {
    const sheetId = await getSheetIdByName(sheets, spreadsheetId, sheetName);
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            deleteSheet: { sheetId },
          },
        ],
      },
    });
    return res.data;
  },

  'sheets.get_info': async ({ spreadsheetId }, { sheets }) => {
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    return res.data;
  },

  'sheets.format_cells': async ({ spreadsheetId, range, format }, { sheets }) => {
    const gridRange = await parseA1ToGridRange(sheets, spreadsheetId, range);
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: gridRange,
              cell: { userEnteredFormat: format },
              fields: 'userEnteredFormat',
            },
          },
        ],
      },
    });
    return res.data;
  },

  'sheets.batch_update': async ({ spreadsheetId, requests }, { sheets }) => {
    const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
    return res.data;
  },

  // Google Drive tools
  'drive_list_files': async ({ query, pageSize }, { drive }) => {
    const res = await drive.files.list({
      q: query || undefined,
      pageSize: pageSize || 100,
      fields: 'files(id,name,mimeType,parents,owners,modifiedTime,size),nextPageToken',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    return res.data;
  },

  'drive_create_folder': async ({ name, parentId }, { drive }) => {
    const res = await drive.files.create({
      requestBody: {
        name,
        mimeType: 'application/vnd.google-apps.folder',
        parents: parentId ? [parentId] : undefined,
      },
      fields: 'id,name,parents',
      supportsAllDrives: true,
    });
    return res.data;
  },

  'drive_upload_file': async ({ name, mimeType, data, parentId }, { drive }) => {
    const buffer = Buffer.from(data, 'base64');
    if (buffer.length > 10 * 1024 * 1024) {
      throw new Error('File too large: max 10MB base64 payload supported');
    }
    const res = await drive.files.create({
      requestBody: {
        name,
        mimeType,
        parents: parentId ? [parentId] : undefined,
      },
      media: {
        mimeType,
        body: buffer,
      },
      fields: 'id,name,mimeType,parents',
      supportsAllDrives: true,
    });
    return res.data;
  },

  'drive_delete_file': async ({ fileId }, { drive }) => {
    const res = await drive.files.delete({ fileId, supportsAllDrives: true });
    return { success: true, status: res.status };
  },

  'drive_get_file_info': async ({ fileId }, { drive }) => {
    const res = await drive.files.get({
      fileId,
      fields: 'id,name,mimeType,parents,owners,modifiedTime,size',
      supportsAllDrives: true,
    });
    return res.data;
  },

  'drive_download_file': async ({ fileId }, { drive }) => {
    const meta = await drive.files.get({ fileId, fields: 'id,name,mimeType', supportsAllDrives: true });
    const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
    const base64 = Buffer.from(res.data).toString('base64');
    return { id: meta.data.id, name: meta.data.name, mimeType: meta.data.mimeType, data: base64 };
  },

  'drive_share_file': async ({ fileId, email, role }, { drive }) => {
    // role: reader | commenter | writer | organizer | fileOrganizer | owner
    const effectiveRole = role || 'reader';
    const res = await drive.permissions.create({
      fileId,
      requestBody: {
        type: 'user',
        role: effectiveRole,
        emailAddress: email,
      },
      sendNotificationEmail: false,
      supportsAllDrives: true,
      fields: 'id,role',
    });
    return res.data;
  },

  'drive_move_file': async ({ fileId, folderId }, { drive }) => {
    const file = await drive.files.get({ fileId, fields: 'parents', supportsAllDrives: true });
    const previousParents = (file.data.parents || []).join(',');
    const res = await drive.files.update({
      fileId,
      addParents: folderId,
      removeParents: previousParents || undefined,
      fields: 'id,parents',
      supportsAllDrives: true,
    });
    return res.data;
  },

  'drive_rename_file': async ({ fileId, newName }, { drive }) => {
    const res = await drive.files.update({
      fileId,
      requestBody: { name: newName },
      fields: 'id,name',
      supportsAllDrives: true,
    });
    return res.data;
  },
};

// Manifest
const manifestTools = [
  {
    name: 'sheets.create_spreadsheet',
    description: 'Create a new Google Spreadsheet with a title.',
    input_schema: {
      type: 'object',
      properties: { title: { type: 'string' } },
      required: ['title'],
    },
  },
  {
    name: 'sheets.get_values',
    description: 'Get values from a range in A1 notation.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        range: { type: 'string' },
      },
      required: ['spreadsheetId', 'range'],
    },
  },
  {
    name: 'sheets.update_values',
    description: 'Update values in a range (RAW) using a 2D array.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        range: { type: 'string' },
        values: { type: 'array', items: { type: 'array' } },
      },
      required: ['spreadsheetId', 'range', 'values'],
    },
  },
  {
    name: 'sheets.append_values',
    description: 'Append rows to a range (RAW) using a 2D array.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        range: { type: 'string' },
        values: { type: 'array', items: { type: 'array' } },
      },
      required: ['spreadsheetId', 'range', 'values'],
    },
  },
  {
    name: 'sheets.clear_values',
    description: 'Clear values in a range.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        range: { type: 'string' },
      },
      required: ['spreadsheetId', 'range'],
    },
  },
  {
    name: 'sheets.create_sheet',
    description: 'Add a new sheet (tab) to a spreadsheet.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        sheetName: { type: 'string' },
      },
      required: ['spreadsheetId', 'sheetName'],
    },
  },
  {
    name: 'sheets.delete_sheet',
    description: 'Delete a sheet (tab) by name.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        sheetName: { type: 'string' },
      },
      required: ['spreadsheetId', 'sheetName'],
    },
  },
  {
    name: 'sheets.get_info',
    description: 'Get spreadsheet metadata and sheets.',
    input_schema: {
      type: 'object',
      properties: { spreadsheetId: { type: 'string' } },
      required: ['spreadsheetId'],
    },
  },
  {
    name: 'sheets.format_cells',
    description: "Apply userEnteredFormat to cells in a range. Range must include sheet name, e.g., 'Sheet1!A1:B2'.",
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        range: { type: 'string' },
        format: { type: 'object', description: 'Google Sheets CellFormat object' },
      },
      required: ['spreadsheetId', 'range', 'format'],
    },
  },
  {
    name: 'sheets.batch_update',
    description: 'Send raw batchUpdate requests to the Sheets API.',
    input_schema: {
      type: 'object',
      properties: {
        spreadsheetId: { type: 'string' },
        requests: { type: 'array', items: { type: 'object' } },
      },
      required: ['spreadsheetId', 'requests'],
    },
  },
  {
    name: 'drive_list_files',
    description: 'List Drive files by query.',
    input_schema: {
      type: 'object',
      properties: {
        query: { type: ['string', 'null'] },
        pageSize: { type: ['number', 'null'] },
      },
      required: [],
    },
  },
  {
    name: 'drive_create_folder',
    description: 'Create a Drive folder optionally under a parent.',
    input_schema: {
      type: 'object',
      properties: {
        name: { type: 'string' },
        parentId: { type: ['string', 'null'] },
      },
      required: ['name'],
    },
  },
  {
    name: 'drive_upload_file',
    description: 'Upload a file with base64-encoded data.',
    input_schema: {
      type: 'object',
      properties: {
        name: { type: 'string' },
        mimeType: { type: 'string' },
        data: { type: 'string', description: 'base64-encoded content' },
        parentId: { type: ['string', 'null'] },
      },
      required: ['name', 'mimeType', 'data'],
    },
  },
  {
    name: 'drive_delete_file',
    description: 'Delete a file by ID.',
    input_schema: {
      type: 'object',
      properties: { fileId: { type: 'string' } },
      required: ['fileId'],
    },
  },
  {
    name: 'drive_get_file_info',
    description: 'Get file metadata by ID.',
    input_schema: {
      type: 'object',
      properties: { fileId: { type: 'string' } },
      required: ['fileId'],
    },
  },
  {
    name: 'drive_download_file',
    description: 'Download a file as base64 content.',
    input_schema: {
      type: 'object',
      properties: { fileId: { type: 'string' } },
      required: ['fileId'],
    },
  },
  {
    name: 'drive_share_file',
    description: 'Share a file with a user and role.',
    input_schema: {
      type: 'object',
      properties: {
        fileId: { type: 'string' },
        email: { type: 'string' },
        role: { type: 'string', enum: ['reader', 'commenter', 'writer', 'organizer', 'fileOrganizer', 'owner'] },
      },
      required: ['fileId', 'email', 'role'],
    },
  },
  {
    name: 'drive_move_file',
    description: 'Move a file to a folder.',
    input_schema: {
      type: 'object',
      properties: {
        fileId: { type: 'string' },
        folderId: { type: 'string' },
      },
      required: ['fileId', 'folderId'],
    },
  },
  {
    name: 'drive_rename_file',
    description: 'Rename a file.',
    input_schema: {
      type: 'object',
      properties: {
        fileId: { type: 'string' },
        newName: { type: 'string' },
      },
      required: ['fileId', 'newName'],
    },
  },
];

// OAuth Configuration
const oauthConfig = {
  client_id: process.env.GOOGLE_CLIENT_ID || 'your-google-client-id',
  client_secret: process.env.GOOGLE_CLIENT_SECRET || 'your-google-client-secret',
  auth_uri: 'https://accounts.google.com/o/oauth2/auth',
  token_uri: 'https://oauth2.googleapis.com/token',
  redirect_uris: [process.env.REDIRECT_URI || 'http://localhost:3000/oauth/callback'],
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
  ]
};

// Endpoints
app.get('/', (_req, res) => {
  res.status(200).send('MCP OK');
});

// OAuth configuration endpoint
app.get('/oauth/config', (_req, res) => {
  console.log('OAuth config requested');
  res.json({
    client_id: oauthConfig.client_id,
    auth_uri: oauthConfig.auth_uri,
    token_uri: oauthConfig.token_uri,
    redirect_uris: oauthConfig.redirect_uris,
    scopes: oauthConfig.scopes,
    response_type: 'code',
    access_type: 'offline',
    prompt: 'consent'
  });
});

// Alternative OAuth config endpoint that ChatGPT might expect
app.get('/.well-known/oauth-configuration', (_req, res) => {
  console.log('Well-known OAuth config requested');
  res.json({
    client_id: oauthConfig.client_id,
    auth_uri: oauthConfig.auth_uri,
    token_uri: oauthConfig.token_uri,
    redirect_uris: oauthConfig.redirect_uris,
    scopes: oauthConfig.scopes,
    response_type: 'code',
    access_type: 'offline',
    prompt: 'consent'
  });
});

// OAuth callback endpoint
app.get('/oauth/callback', (req, res) => {
  const { code, state } = req.query;
  if (!code) {
    return res.status(400).json({ error: 'Authorization code not provided' });
  }
  
  // Exchange code for tokens
  const { OAuth2Client } = require('google-auth-library');
  const oauth2Client = new OAuth2Client(
    oauthConfig.client_id,
    oauthConfig.client_secret,
    oauthConfig.redirect_uris[0]
  );
  
  oauth2Client.getToken(code)
    .then(({ tokens }) => {
      res.json({
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token,
        expires_in: tokens.expiry_date ? Math.floor((tokens.expiry_date - Date.now()) / 1000) : 3600
      });
    })
    .catch((error) => {
      console.error('Error exchanging code for tokens:', error);
      res.status(400).json({ error: 'Failed to exchange authorization code' });
    });
});

app.get('/mcp/manifest', (_req, res) => {
  res.json({
    name: 'mcp-google-tools',
    version: '1.0.0',
    oauth: {
      client_id: oauthConfig.client_id,
      auth_uri: oauthConfig.auth_uri,
      token_uri: oauthConfig.token_uri,
      redirect_uris: oauthConfig.redirect_uris,
      scopes: oauthConfig.scopes,
      response_type: 'code',
      access_type: 'offline',
      prompt: 'consent'
    },
    tools: manifestTools.map((t) => {
      if (t.name === 'drive_upload_file') {
        // Document the 10MB limit in the schema description
        return {
          ...t,
          input_schema: {
            ...t.input_schema,
            properties: {
              ...t.input_schema.properties,
              data: { ...t.input_schema.properties.data, description: 'base64-encoded content (<= 10MB)' },
            },
          },
        };
      }
      if (t.name === 'drive_share_file') {
        return {
          ...t,
          input_schema: {
            ...t.input_schema,
            required: ['fileId', 'email'],
            properties: {
              ...t.input_schema.properties,
              role: { type: 'string', enum: ['reader', 'commenter', 'writer', 'organizer', 'fileOrganizer', 'owner'], default: 'reader' },
            },
          },
        };
      }
      return t;
    }),
  });
});

app.post('/mcp/tool', async (req, res) => {
  try {
    const { name, args } = req.body || {};
    if (!name) return res.status(400).json({ error: 'Missing tool name' });
    const handler = toolHandlers[name];
    if (!handler) return res.status(404).json({ error: `Unknown tool: ${name}` });
    const clients = getGoogleClients(req);
    const result = await handler(args || {}, clients);
    return res.json(result);
  } catch (err) {
    const message = err && err.message ? err.message : 'Unknown error';
    return res.status(400).json({ error: message });
  }
});

app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`MCP server listening on port ${PORT}`);
});


