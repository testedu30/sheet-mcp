import express from 'express';
import { google } from 'googleapis';
import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';

const REQUIRED_ENV = ['GCP_SERVICE_ACCOUNT_EMAIL', 'GCP_PROJECT_ID', 'DRIVE_FOLDER_ID'];
REQUIRED_ENV.forEach(key => {
  if (!process.env[key]) {
    console.warn(`[sheets-mcp] Missing ${key}. Set it for full functionality.`);
  }
});

const DRIVE_FOLDER_ID = process.env.DRIVE_FOLDER_ID;
const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/spreadsheets'
];

let googleClientsPromise;

// async function getGoogleClients() {
//   if (!googleClientsPromise) {
//     googleClientsPromise = (async () => {
//       const auth = await google.auth.getClient({
//         projectId: process.env.GCP_PROJECT_ID,
//         scopes: SCOPES
//       });

//       return {
//         sheets: google.sheets({ version: 'v4', auth }),
//         drive: google.drive({ version: 'v3', auth })
//       };
//     })();
//   }

//   return googleClientsPromise;
// }


async function getGoogleClients() {
  if (!googleClientsPromise) {
    googleClientsPromise = (async () => {

      const auth = new google.auth.GoogleAuth({
        credentials: {
          client_email: process.env.GCP_SERVICE_ACCOUNT_EMAIL,
          private_key: process.env.GCP_PRIVATE_KEY.replace(/\\n/g, '\n')
        },
        scopes: SCOPES
      });

      return {
        sheets: google.sheets({ version: 'v4', auth }),
        drive: google.drive({ version: 'v3', auth })
      };
    })();
  }

  return googleClientsPromise;
}


function respondWith(content) {
  const text = typeof content === 'string' ? content : JSON.stringify(content, null, 2);
  return {
    content: [{ type: 'text', text }],
    structuredContent: content
  };
}

const server = new McpServer({
  name: 'sheets-mcp-node',
  version: '0.1.0'
});

server.registerTool(
  'list_spreadsheets',
  {
    title: 'List spreadsheets',
    description: 'Lists spreadsheets in the configured Drive folder.',
    inputSchema: {
      pageSize: z.number().int().min(1).max(100).optional()
    },
    outputSchema: {
      spreadsheets: z.array(
        z.object({
          id: z.string(),
          name: z.string()
        })
      )
    }
  },
  async ({ pageSize = 25 }) => {
    if (!DRIVE_FOLDER_ID) {
      throw new Error('DRIVE_FOLDER_ID is required to list spreadsheets.');
    }

    const { drive } = await getGoogleClients();
    const response = await drive.files.list({
      q: `'${DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
      pageSize,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'files(id,name)'
    });

    const spreadsheets = response.data.files?.map(file => ({
      id: file.id,
      name: file.name
    })) ?? [];

    return respondWith({ spreadsheets });
  }
);

server.registerTool(
  'create_spreadsheet',
  {
    title: 'Create spreadsheet',
    description: 'Creates a new Google Sheet and stores it in the configured folder.',
    inputSchema: {
      title: z.string().min(1)
    },
    outputSchema: {
      spreadsheetId: z.string(),
      title: z.string()
    }
  },
  async ({ title }) => {
    const { sheets, drive } = await getGoogleClients();

    const creation = await sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title
        }
      },
      fields: 'spreadsheetId,properties/title'
    });

    const spreadsheetId = creation.data.spreadsheetId;

    if (!spreadsheetId) {
      throw new Error('Failed to create spreadsheet.');
    }

    if (DRIVE_FOLDER_ID) {
      await drive.files.update({
        fileId: spreadsheetId,
        addParents: DRIVE_FOLDER_ID,
        supportsAllDrives: true,
        fields: 'id, parents'
      });
    }

    return respondWith({
      spreadsheetId,
      title: creation.data.properties?.title ?? title
    });
  }
);

server.registerTool(
  'append_rows',
  {
    title: 'Append rows',
    description: 'Appends rows to a sheet using A1 notation.',
    inputSchema: {
      spreadsheetId: z.string(),
      range: z.string().describe('Range like Sheet1!A1'),
      values: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean()])))
        .min(1)
    },
    outputSchema: {
      updatedRange: z.string(),
      updatedRows: z.number().int()
    }
  },
  async ({ spreadsheetId, range, values }) => {
    const { sheets } = await getGoogleClients();
    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values
      }
    });

    const updates = response.data.updates;
    return respondWith({
      updatedRange: updates?.updatedRange ?? '',
      updatedRows: updates?.updatedRows ?? 0
    });
  }
);


server.registerTool(
  'get_sheet_info',
  {
    title: 'Get Sheet Info',
    description: 'Get metadata about spreadsheet (tabs, rows, columns)',
    inputSchema: {
      url_or_id: z.string()
    },
    outputSchema: {
      sheets: z.array(
        z.object({
          title: z.string(),
          rows: z.number(),
          cols: z.number()
        })
      )
    }
  },
  async ({ url_or_id }) => {
    const { sheets } = await getGoogleClients();

    // extract ID from URL if needed
    const spreadsheetId = url_or_id.includes('docs.google.com')
      ? url_or_id.split('/d/')[1].split('/')[0]
      : url_or_id;

    const res = await sheets.spreadsheets.get({ spreadsheetId });

    const info = res.data.sheets.map((s) => ({
      title: s.properties.title,
      rows: s.properties.gridProperties.rowCount,
      cols: s.properties.gridProperties.columnCount,
    }));

    return respondWith({ sheets: info });
  }
);


server.registerTool(
  'read_range',
  {
    title: 'Read Range',
    description: 'Read data from a Google Sheet using A1 notation',
    inputSchema: {
      url_or_id: z.string().describe('Google Sheet URL or ID'),
      range: z.string().describe('Range like Sheet1!A1:C10')
    },
    outputSchema: {
      values: z.array(z.array(z.string()))
    }
  },
  async ({ url_or_id, range }) => {
    const { sheets } = await getGoogleClients();

    // URL se ID extract
    const spreadsheetId = url_or_id.includes('docs.google.com')
      ? url_or_id.split('/d/')[1].split('/')[0]
      : url_or_id;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    });

    const values = response.data.values || [];

    return respondWith({ values });
  }
);

server.registerTool(
  'write_sheet',
  {
    title: 'Write Range',
    description: 'Write or update values in a Google Sheet (overwrite existing data)',
    inputSchema: {
      url_or_id: z.string().describe('Google Sheet URL or ID'),
      range: z.string().describe('Range like Sheet1!A1:C10'),
      values: z.array(
        z.array(z.union([z.string(), z.number(), z.boolean()]))
      ).min(1)
    },
    outputSchema: {
      updatedRange: z.string(),
      updatedRows: z.number()
    }
  },
  async ({ url_or_id, range, values }) => {
    const { sheets } = await getGoogleClients();

    // URL → ID extract
    const spreadsheetId = url_or_id.includes('docs.google.com')
      ? url_or_id.split('/d/')[1].split('/')[0]
      : url_or_id;

    const response = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values
      }
    });

    return respondWith({
      updatedRange: response.data.updatedRange || '',
      updatedRows: response.data.updatedRows || 0
    });
  }
);

server.registerTool(
  'clear_range',
  {
    title: 'Clear Range',
    description: 'Clear data from a specific range in a Google Sheet',
    inputSchema: {
      url_or_id: z.string().describe('Google Sheet URL or ID'),
      sheet: z.string().describe('Sheet name (e.g. Sheet1)'),
      range: z.string().describe('Range like A1:C10')
    },
    outputSchema: {
      clearedRange: z.string()
    }
  },
  async ({ url_or_id, sheet, range }) => {
    const { sheets } = await getGoogleClients();

    // URL → ID extract
    const spreadsheetId = url_or_id.includes('docs.google.com')
      ? url_or_id.split('/d/')[1].split('/')[0]
      : url_or_id;

    const fullRange = `${sheet}!${range}`;

    const response = await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: fullRange
    });

    return respondWith({
      clearedRange: response.data.clearedRange || fullRange
    });
  }
);



const app = express();
app.use(express.json());

app.post('/mcp', async (req, res) => {
  try {
    const transport = new StreamableHTTPServerTransport({
      enableJsonResponse: true
    });

    res.on('close', () => {
      transport.close();
    });

    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error('[sheets-mcp] Error handling /mcp request', error);
    if (!res.headersSent) {
      res.status(500).json({
        jsonrpc: '2.0',
        error: {
          code: -32603,
          message: error?.message ?? 'Internal server error'
        },
        id: null
      });
    }
  }
});

const sseTransports = new Map();

app.get('/sse', async (req, res) => {
  const transport = new SSEServerTransport('/messages', res);
  sseTransports.set(transport.sessionId, transport);

  res.on('close', () => {
    sseTransports.delete(transport.sessionId);
  });

  await server.connect(transport);
});

app.post('/messages', async (req, res) => {
  const sessionId = req.query.sessionId;
  if (!sessionId || !sseTransports.has(sessionId)) {
    res.status(400).json({ error: 'Unknown session' });
    return;
  }

  const transport = sseTransports.get(sessionId);
  await transport.handlePostMessage(req, res, req.body);
});

const port = parseInt(process.env.PORT ?? '8080', 10);

app.listen(port, () => {
  console.log(`[sheets-mcp] Server listening on port ${port}`);
  console.log(`[sheets-mcp] Streamable HTTP endpoint: POST /mcp`);
  console.log(`[sheets-mcp] SSE endpoint: GET /sse`);
}).on('error', error => {
  console.error('[sheets-mcp] Failed to start server', error);
  process.exit(1);
});
