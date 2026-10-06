/**
 * Configuration management for Outlook Calendar MCP Server
 */

import dotenv from 'dotenv';

// Load environment variables
dotenv.config();

export const config = {
  server: {
    port: parseInt(process.env.PORT || '3003', 10),
    nodeEnv: process.env.NODE_ENV || 'development',
    name: process.env.MCP_SERVER_NAME || 'outlook-calendar-mcp-server',
    version: process.env.MCP_SERVER_VERSION || '1.0.0',
  },
  microsoft: {
    clientId: process.env.MICROSOFT_CLIENT_ID || '',
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET || '',
    tenantId: process.env.MICROSOFT_TENANT_ID || 'common',
    redirectUri: process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3003/oauth/callback',
  },
  frontend: {
    url: process.env.FRONTEND_URL || 'http://localhost:5173',
  },
  graph: {
    scopes: [
      'https://graph.microsoft.com/Calendars.Read',
      'https://graph.microsoft.com/User.Read',
    ],
    endpoint: 'https://graph.microsoft.com/v1.0',
  },
};

// Validate required configuration
export function validateConfig(): void {
  const required = [
    { key: 'MICROSOFT_CLIENT_ID', value: config.microsoft.clientId },
    { key: 'MICROSOFT_CLIENT_SECRET', value: config.microsoft.clientSecret },
  ];

  const missing = required.filter((item) => !item.value);

  if (missing.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missing.map((item) => item.key).join(', ')}`
    );
  }
}







