/**
 * Cloudflare Workers Entry Point for Outlook Calendar MCP Server
 * Implements MCP protocol with SSE support for ChatGPT integration
 */

import 'isomorphic-fetch';
import { outlookCalendarTools } from './mcp/tools.js';

// Environment interface for Cloudflare Workers
interface Env {
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_CLIENT_SECRET: string;
  MICROSOFT_TENANT_ID: string;
  MICROSOFT_REDIRECT_URI: string;
  FRONTEND_URL: string;
  BACKEND_API_URL?: string; // Optional: Backend API URL for token fallback
  MCP_SERVER_NAME: string;
  MCP_SERVER_VERSION: string;
  NODE_ENV: string;
  SESSIONS: KVNamespace;
}

// Session data structure
interface SessionData {
  userId: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt?: string;
  createdAt: string;
}

// OAuth token structure
interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  scope?: string;
  id_token?: string;
}

// Microsoft Graph scopes
const GRAPH_SCOPES = [
  'https://graph.microsoft.com/Calendars.Read',
  'https://graph.microsoft.com/User.Read',
];

/**
 * Outlook Calendar Client for Cloudflare Workers
 */
class OutlookCalendarClientWorker {
  private accessToken: string;
  private graphEndpoint = 'https://graph.microsoft.com/v1.0';

  constructor(accessToken: string) {
    this.accessToken = accessToken;
  }

  private async makeGraphRequest(
    endpoint: string,
    method: string = 'GET',
    body?: any
  ): Promise<any> {
    const url = `${this.graphEndpoint}${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${this.accessToken}`,
      'Content-Type': 'application/json',
    };

    const options: RequestInit = {
      method,
      headers,
    };

    if (body) {
      options.body = JSON.stringify(body);
    }

    const response = await fetch(url, options);

    if (!response.ok) {
      const error = await response.text();
      console.error('Graph API request failed:', { endpoint, status: response.status, error });
      throw new Error(`Microsoft Graph API error: ${response.status}`);
    }

    return response.json();
  }

  private simplifyEvent(event: any) {
    return {
      id: event.id,
      subject: event.subject,
      bodyPreview: event.bodyPreview,
      startDateTime: event.start.dateTime,
      endDateTime: event.end.dateTime,
      timeZone: event.start.timeZone,
      location: event.location?.displayName,
      organizer: event.organizer?.emailAddress.address,
      attendees: event.attendees?.map((attendee: any) => ({
        email: attendee.emailAddress.address,
        name: attendee.emailAddress.name,
        response: attendee.status?.response,
      })),
      isAllDay: event.isAllDay,
      webLink: event.webLink,
      onlineMeetingUrl: event.onlineMeetingUrl || event.onlineMeeting?.joinUrl,
      categories: event.categories,
    };
  }

  async searchEvents(options: any = {}) {
    const params = new URLSearchParams();

    // Build filter query
    const filters: string[] = [];
    if (options.startDateTime && options.endDateTime) {
      filters.push(
        `start/dateTime ge '${options.startDateTime}' and end/dateTime le '${options.endDateTime}'`
      );
    } else if (options.startDateTime) {
      filters.push(`start/dateTime ge '${options.startDateTime}'`);
    } else if (options.endDateTime) {
      filters.push(`end/dateTime le '${options.endDateTime}'`);
    }

    if (options.subject) {
      filters.push(`contains(subject,'${options.subject}')`);
    }

    if (options.location) {
      filters.push(`contains(location/displayName,'${options.location}')`);
    }

    if (options.category) {
      filters.push(`categories/any(c:c eq '${options.category}')`);
    }

    if (filters.length > 0) {
      params.append('$filter', filters.join(' and '));
    }

    if (options.top) {
      params.append('$top', options.top.toString());
    }
    if (options.skip) {
      params.append('$skip', options.skip.toString());
    }
    if (options.orderBy) {
      params.append('$orderby', options.orderBy);
    } else {
      params.append('$orderby', 'start/dateTime');
    }

    const endpoint = `/me/events${params.toString() ? '?' + params.toString() : ''}`;
    const data = await this.makeGraphRequest(endpoint);

    const events = data.value || [];
    return events.map((event: any) => this.simplifyEvent(event));
  }

  async fetchEvent(eventId: string) {
    const endpoint = `/me/events/${eventId}`;
    const event = await this.makeGraphRequest(endpoint);
    return this.simplifyEvent(event);
  }

  async fetchEventsBatch(eventIds: string[]) {
    const batchRequests = eventIds.map((id, index) => ({
      id: index.toString(),
      method: 'GET',
      url: `/me/events/${id}`,
    }));

    const batchRequestBody = {
      requests: batchRequests,
    };

    const endpoint = '/$batch';
    const response = await this.makeGraphRequest(endpoint, 'POST', batchRequestBody);

    const events: any[] = [];
    for (const batchResponse of response.responses || []) {
      if (batchResponse.status === 200 && batchResponse.body) {
        events.push(this.simplifyEvent(batchResponse.body));
      }
    }

    return events;
  }

  async listEvents(startDateTime: string, endDateTime: string, options: any = {}) {
    return await this.searchEvents({
      startDateTime,
      endDateTime,
      top: options.top || 50,
      skip: options.skip,
      orderBy: 'start/dateTime',
    });
  }

  async getProfile() {
    const endpoint = '/me';
    const profile = await this.makeGraphRequest(endpoint);
    return {
      id: profile.id,
      displayName: profile.displayName,
      givenName: profile.givenName,
      surname: profile.surname,
      mail: profile.mail,
      userPrincipalName: profile.userPrincipalName,
      jobTitle: profile.jobTitle,
      officeLocation: profile.officeLocation,
      mobilePhone: profile.mobilePhone,
      businessPhones: profile.businessPhones,
    };
  }
}

/**
 * OAuth Manager for Cloudflare Workers
 */
class CloudflareOAuthManager {
  private env: Env;

  constructor(env: Env) {
    this.env = env;
  }

  getAuthorizationUrl(state: string): string {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      response_type: 'code',
      redirect_uri: this.env.MICROSOFT_REDIRECT_URI,
      response_mode: 'query',
      scope: GRAPH_SCOPES.join(' '),
      state,
      prompt: 'consent',
    });

    return `https://login.microsoftonline.com/${this.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
  }

  async exchangeCodeForTokens(code: string): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      client_secret: this.env.MICROSOFT_CLIENT_SECRET,
      code,
      redirect_uri: this.env.MICROSOFT_REDIRECT_URI,
      grant_type: 'authorization_code',
    });

    const response = await fetch(
      `https://login.microsoftonline.com/${this.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      console.error('Token exchange failed:', error);
      throw new Error('Failed to exchange authorization code for tokens');
    }

    const data = await response.json() as {
      access_token: string;
      refresh_token?: string;
      expires_in?: number;
      token_type?: string;
      scope?: string;
      id_token?: string;
    };

    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token || undefined,
      expires_in: data.expires_in || 3600,
      token_type: data.token_type || 'Bearer',
      scope: data.scope || undefined,
      id_token: data.id_token || undefined,
    };
  }

  async refreshAccessToken(refreshToken: string): Promise<OAuthTokens> {
    const params = new URLSearchParams({
      client_id: this.env.MICROSOFT_CLIENT_ID,
      client_secret: this.env.MICROSOFT_CLIENT_SECRET,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
    });

    const response = await fetch(
      `https://login.microsoftonline.com/${this.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      }
    );

    if (!response.ok) {
      const error = await response.text();
      console.error('Token refresh failed:', error);
      throw new Error('Failed to refresh access token');
    }

    const data = await response.json() as {
      access_token: string;
      refresh_token?: string;
      expires_in?: number;
      token_type?: string;
      scope?: string;
      id_token?: string;
    };

    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token || refreshToken,
      expires_in: data.expires_in || 3600,
      token_type: data.token_type || 'Bearer',
      scope: data.scope || undefined,
      id_token: data.id_token || undefined,
    };
  }

  async storeSession(userId: string, tokens: OAuthTokens): Promise<void> {
    const expiresAt = tokens.expires_in
      ? new Date(Date.now() + tokens.expires_in * 1000).toISOString()
      : undefined;

    const sessionData: SessionData = {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt,
      createdAt: new Date().toISOString(),
    };

    await this.env.SESSIONS.put(
      `session:${userId}`,
      JSON.stringify(sessionData),
      { expirationTtl: 60 * 60 * 24 * 30 }
    );
  }

  async getSession(userId: string): Promise<SessionData | null> {
    const data = await this.env.SESSIONS.get(`session:${userId}`);
    return data ? JSON.parse(data) : null;
  }

  async isSessionValid(userId: string): Promise<boolean> {
    const session = await this.getSession(userId);
    if (!session) return false;
    if (!session.expiresAt) return true;
    return new Date(session.expiresAt) > new Date();
  }

  async getValidAccessToken(userId: string): Promise<string | null> {
    // First, try to get from KV store (fast path)
    const session = await this.getSession(userId);
    if (session && (await this.isSessionValid(userId))) {
      return session.accessToken;
    }

    // If session exists but expired, try to refresh
    if (session && session.refreshToken) {
      try {
        const newTokens = await this.refreshAccessToken(session.refreshToken);
        await this.storeSession(userId, newTokens);
        return newTokens.access_token;
      } catch (error) {
        console.error('Failed to refresh token from KV:', error);
      }
    }

    // Fallback: Query backend API for tokens from profile.settings.oauth_tokens
    try {
      const backendUrl =
        this.env.BACKEND_API_URL ||
        (this.env.FRONTEND_URL ? this.env.FRONTEND_URL.replace(/\/$/, '') : null) ||
        'https://api.zerotwo.app';

      const response = await fetch(
        `${backendUrl}/api/ai/tools/outlook/token?userId=${userId}`,
        {
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
          },
        }
      );

      if (response.ok) {
        const data = (await response.json()) as {
          accessToken?: string;
          refreshToken?: string;
          expiresIn?: number;
        };
        if (data.accessToken) {
          await this.storeSession(userId, {
            access_token: data.accessToken,
            refresh_token: data.refreshToken,
            expires_in: data.expiresIn,
            token_type: 'Bearer',
          });
          return data.accessToken;
        }
      }
    } catch (error) {
      console.error('Failed to fetch token from backend API (fallback):', error);
    }

    return null;
  }

  async removeSession(userId: string): Promise<void> {
    await this.env.SESSIONS.delete(`session:${userId}`);
  }
}

/**
 * CORS headers
 */
function getCorsHeaders(origin?: string): Record<string, string> {
  const allowedOrigins = [
    'http://localhost:5173',
    'http://localhost:3000',
    'https://zerotwo.app',
  ];

  const requestOrigin = origin || '';
  const allowOrigin = allowedOrigins.includes(requestOrigin)
    ? requestOrigin
    : allowedOrigins[0];

  return {
    'Access-Control-Allow-Origin': allowOrigin,
    'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, mcp-session-id, Authorization',
    'Access-Control-Expose-Headers': 'Mcp-Session-Id',
    'Access-Control-Allow-Credentials': 'true',
  };
}

/**
 * Handle OPTIONS requests
 */
function handleOptions(request: Request): Response {
  return new Response(null, {
    status: 204,
    headers: getCorsHeaders(request.headers.get('Origin') || undefined),
  });
}

/**
 * Execute MCP tool
 */
async function executeTool(
  toolName: string,
  args: any,
  oauthManager: CloudflareOAuthManager,
  requestHeaders?: Headers
): Promise<any> {
  // Get userId from args or from X-User-Id header
  let userId = args?.userId;
  if (!userId && requestHeaders) {
    userId = requestHeaders.get('X-User-Id');
  }

  if (!userId) {
    return {
      content: [
        {
          type: 'text',
          text: 'User ID is required for all Outlook Calendar operations. Please provide userId in tool arguments or X-User-Id header.',
        },
      ],
      isError: true,
    };
  }

  // First, try to get access token from X-Access-Token header (passed directly from backend)
  let accessToken = requestHeaders?.get('X-Access-Token') || null;
  
  // If not in header, fall back to KV storage lookup
  if (!accessToken) {
    accessToken = await oauthManager.getValidAccessToken(userId);
  }
  
  if (!accessToken) {
    return {
      content: [
        {
          type: 'text',
          text: 'Authentication required. Please authenticate with Microsoft Outlook first.',
        },
      ],
      isError: true,
    };
  }

  const client = new OutlookCalendarClientWorker(accessToken);

  try {
    switch (toolName) {
      case 'outlook_calendar_search_events': {
        const {
          startDateTime,
          endDateTime,
          subject,
          location,
          category,
          top,
          skip,
          orderBy,
        } = args;
        const events = await client.searchEvents({
          startDateTime,
          endDateTime,
          subject,
          location,
          category,
          top,
          skip,
          orderBy,
        });

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ events, count: events.length }, null, 2),
            },
          ],
        };
      }

      case 'outlook_calendar_fetch_event': {
        const { eventId } = args;
        const event = await client.fetchEvent(eventId);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ event }, null, 2),
            },
          ],
        };
      }

      case 'outlook_calendar_fetch_events_batch': {
        const { eventIds } = args;
        if (!eventIds || !Array.isArray(eventIds) || eventIds.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: 'Error: eventIds must be a non-empty array',
              },
            ],
            isError: true,
          };
        }

        const events = await client.fetchEventsBatch(eventIds);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ events, count: events.length }, null, 2),
            },
          ],
        };
      }

      case 'outlook_calendar_list_events': {
        const { startDateTime, endDateTime, top, skip } = args;
        const events = await client.listEvents(startDateTime, endDateTime, { top, skip });

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ events, count: events.length }, null, 2),
            },
          ],
        };
      }

      case 'outlook_calendar_get_profile': {
        const profile = await client.getProfile();

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ profile }, null, 2),
            },
          ],
        };
      }

      default:
        return {
          content: [
            {
              type: 'text',
              text: `Unknown tool: ${toolName}`,
            },
          ],
          isError: true,
        };
    }
  } catch (error: any) {
    console.error(`Error executing tool ${toolName}:`, error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${error.message}`,
        },
      ],
      isError: true,
    };
  }
}

/**
 * Handle MCP protocol requests
 */
async function handleMcpRequest(
  request: Request,
  env: Env,
  oauthManager: CloudflareOAuthManager
): Promise<Response> {
  try {
    const body = (await request.json()) as {
      method?: string;
      id?: string | number;
      params?: {
        name?: string;
        arguments?: any;
      };
    };

    // Handle initialize request
    if (body.method === 'initialize') {
      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            protocolVersion: '2024-11-05',
            capabilities: {
              tools: {},
            },
            serverInfo: {
              name: env.MCP_SERVER_NAME || 'outlook-calendar-mcp-server',
              version: env.MCP_SERVER_VERSION || '1.0.0',
            },
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    // Handle tools/list request
    if (body.method === 'tools/list') {
      const tools = outlookCalendarTools.map((tool) => {
        // Extract schema properties from Zod schema
        const schema = tool.definition.inputSchema as any;
        const properties: Record<string, any> = {};

        if (schema && typeof schema === 'object') {
          if (schema.shape) {
            Object.entries(schema.shape).forEach(([key, value]: [string, any]) => {
              properties[key] = {
                type:
                  value._def?.typeName === 'ZodString'
                    ? 'string'
                    : value._def?.typeName === 'ZodNumber'
                    ? 'number'
                    : value._def?.typeName === 'ZodBoolean'
                    ? 'boolean'
                    : value._def?.typeName === 'ZodArray'
                    ? 'array'
                    : 'string',
                description: value.description || value._def?.description || '',
              };
            });
          }
        }

        return {
          name: tool.name,
          description: tool.definition.description || tool.definition.title || '',
          inputSchema: {
            type: 'object',
            properties,
            required: [],
          },
        };
      });

      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            tools,
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    // Handle tools/call request
    if (body.method === 'tools/call') {
      if (!body.params) {
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            id: body.id,
            error: {
              code: -32602,
              message: 'Invalid params',
            },
          }),
          {
            status: 400,
            headers: {
              'Content-Type': 'application/json',
              ...getCorsHeaders(),
            },
          }
        );
      }

      const { name, arguments: args } = body.params;
      if (!name) {
        return new Response(
          JSON.stringify({
            jsonrpc: '2.0',
            id: body.id,
            error: {
              code: -32602,
              message: 'Tool name is required',
            },
          }),
          {
            status: 400,
            headers: {
              'Content-Type': 'application/json',
              ...getCorsHeaders(),
            },
          }
        );
      }

      const result = await executeTool(name, args || {}, oauthManager, request.headers);

      return new Response(
        JSON.stringify({
          jsonrpc: '2.0',
          id: body.id,
          result: {
            content: result.content,
            isError: result.isError || false,
          },
        }),
        {
          headers: {
            'Content-Type': 'application/json',
            ...getCorsHeaders(request.headers.get('Origin') || undefined),
          },
        }
      );
    }

    // Handle notifications (no id field, no response expected)
    // Notifications like notifications/initialized don't have an id and don't require a JSON-RPC response
    // Use 204 No Content for proper HTTP semantics
    if (body.method?.startsWith('notifications/') || (!body.id && body.method)) {
      // For notifications, just acknowledge with 204 No Content
      return new Response(null, {
        status: 204,
        headers: getCorsHeaders(request.headers.get('Origin') || undefined),
      });
    }

    // Unknown method (only for requests with id)
    return new Response(
      JSON.stringify({
        jsonrpc: '2.0',
        id: body.id,
        error: {
          code: -32601,
          message: 'Method not found',
        },
      }),
      {
        status: 400,
        headers: {
          'Content-Type': 'application/json',
          ...getCorsHeaders(request.headers.get('Origin') || undefined),
        },
      }
    );
  } catch (error: any) {
    console.error('MCP request error:', error);
    return new Response(
      JSON.stringify({
        jsonrpc: '2.0',
        id: null,
        error: {
          code: -32603,
          message: 'Internal error',
          data: error.message,
        },
      }),
      {
        status: 500,
        headers: {
          'Content-Type': 'application/json',
          ...getCorsHeaders(request.headers.get('Origin') || undefined),
        },
      }
    );
  }
}

/**
 * SSE endpoint for MCP protocol
 */
function createSseResponse(env: Env): Response {
  const encoder = new TextEncoder();
  let keepAliveInterval: ReturnType<typeof setInterval> | null = null;

  const stream = new ReadableStream({
    start(controller) {
      const initMessage = {
        jsonrpc: '2.0',
        method: 'notifications/initialized',
        params: {},
      };
      controller.enqueue(encoder.encode(`data: ${JSON.stringify(initMessage)}\n\n`));

      keepAliveInterval = setInterval(() => {
        try {
          controller.enqueue(encoder.encode(': keepalive\n\n'));
        } catch (e) {
          if (keepAliveInterval) {
            clearInterval(keepAliveInterval);
            keepAliveInterval = null;
          }
        }
      }, 30000);
    },
    cancel() {
      if (keepAliveInterval) {
        clearInterval(keepAliveInterval);
        keepAliveInterval = null;
      }
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
      ...getCorsHeaders(),
    },
  });
}

/**
 * Main Worker fetch handler
 */
export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin');
    const corsHeaders = getCorsHeaders(origin || undefined);

    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      return handleOptions(request);
    }

    try {
      const oauthManager = new CloudflareOAuthManager(env);

      // Health check
      if (url.pathname === '/health' || url.pathname === '/') {
        return new Response(
          JSON.stringify({
            status: 'healthy',
            server: env.MCP_SERVER_NAME,
            version: env.MCP_SERVER_VERSION,
            timestamp: new Date().toISOString(),
          }),
          {
            headers: {
              'Content-Type': 'application/json',
              ...corsHeaders,
            },
          }
        );
      }

      // OAuth authorization endpoint
      if (url.pathname === '/oauth/authorize' && request.method === 'GET') {
        const userId = url.searchParams.get('userId');
        const state = url.searchParams.get('state');

        if (!userId || !state) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and state' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const authUrl = oauthManager.getAuthorizationUrl(state);

        return new Response(JSON.stringify({ authorizationUrl: authUrl, state }), {
          headers: { 'Content-Type': 'application/json', ...corsHeaders },
        });
      }

      // OAuth callback endpoint
      if (url.pathname === '/oauth/callback' && request.method === 'GET') {
        const code = url.searchParams.get('code');
        const state = url.searchParams.get('state');
        const error = url.searchParams.get('error');

        if (error) {
          return Response.redirect(
            `${env.FRONTEND_URL}/settings?oauth_error=${encodeURIComponent(error)}`,
            302
          );
        }

        if (!code || !state) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: code and state' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const tokens = await oauthManager.exchangeCodeForTokens(code);
        const userId = state;

        await oauthManager.storeSession(userId, tokens);

        return Response.redirect(
          `${env.FRONTEND_URL}/settings?oauth_success=true&provider=outlook_calendar`,
          302
        );
      }

      // OAuth refresh endpoint
      if (url.pathname === '/oauth/refresh' && request.method === 'POST') {
        const body = (await request.json()) as { userId?: string; refreshToken?: string };
        const { userId, refreshToken } = body;

        if (!userId || !refreshToken) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and refreshToken' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        const tokens = await oauthManager.refreshAccessToken(refreshToken);
        await oauthManager.storeSession(userId, tokens);

        return new Response(
          JSON.stringify({
            accessToken: tokens.access_token,
            refreshToken: tokens.refresh_token,
            expiresIn: tokens.expires_in,
            timestamp: new Date().toISOString(),
          }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      // OAuth disconnect endpoint
      if (url.pathname === '/oauth/disconnect' && request.method === 'POST') {
        const body = (await request.json()) as { userId?: string };
        const { userId } = body;

        if (!userId) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameter: userId' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        await oauthManager.removeSession(userId);

        return new Response(
          JSON.stringify({ success: true, message: 'Successfully disconnected' }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      // OAuth token sync endpoint
      if (url.pathname === '/oauth/sync' && request.method === 'POST') {
        const body = (await request.json()) as {
          userId?: string;
          accessToken?: string;
          refreshToken?: string;
          expiresIn?: number;
        };
        const { userId, accessToken, refreshToken, expiresIn } = body;

        if (!userId || !accessToken) {
          return new Response(
            JSON.stringify({ error: 'Missing required parameters: userId and accessToken' }),
            {
              status: 400,
              headers: { 'Content-Type': 'application/json', ...corsHeaders },
            }
          );
        }

        await oauthManager.storeSession(userId, {
          access_token: accessToken,
          refresh_token: refreshToken,
          expires_in: expiresIn,
          token_type: 'Bearer',
        });

        return new Response(
          JSON.stringify({ success: true, message: 'Tokens synced successfully' }),
          {
            headers: { 'Content-Type': 'application/json', ...corsHeaders },
          }
        );
      }

      // MCP endpoint - POST for JSON-RPC
      if (url.pathname === '/mcp' && request.method === 'POST') {
        return handleMcpRequest(request, env, oauthManager);
      }

      // SSE endpoint for MCP protocol
      if (url.pathname === '/sse' && request.method === 'GET') {
        return createSseResponse(env);
      }

      // 404 - Not Found
      return new Response(
        JSON.stringify({ error: 'Not Found', message: 'The requested endpoint does not exist' }),
        {
          status: 404,
          headers: { 'Content-Type': 'application/json', ...corsHeaders },
        }
      );
    } catch (error: any) {
      console.error('Worker error:', error);
      return new Response(
        JSON.stringify({
          error: 'Internal Server Error',
          message: error.message,
        }),
        {
          status: 500,
          headers: { 'Content-Type': 'application/json', ...corsHeaders },
        }
      );
    }
  },
};







