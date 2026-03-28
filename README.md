# Outlook Calendar MCP Server

A Model Context Protocol (MCP) server for Microsoft Outlook Calendar integration with OAuth 2.0 authentication via Microsoft Graph API.

## Features

- **OAuth 2.0 Authentication**: Secure Microsoft OAuth flow with automatic token refresh
- **MCP Protocol Support**: Full implementation of the Model Context Protocol
- **Outlook Calendar Operations**:
  - Search events with advanced filters
  - Fetch single event details
  - Fetch multiple events in batch
  - List events within date ranges
  - Get user profile information
- **Session Management**: Persistent sessions with automatic token refresh
- **TypeScript**: Fully typed with TypeScript for better development experience
- **Cloudflare Workers**: Can be deployed as a Cloudflare Worker for serverless operation

## Prerequisites

- Node.js 18+
- Microsoft Azure App Registration with Microsoft Graph API enabled
- OAuth 2.0 credentials (Client ID and Client Secret)

## Setup

### 1. Microsoft Azure Configuration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure your app:
   - Name: `Outlook Calendar MCP Server`
   - Supported account types: Choose based on your needs
   - Redirect URI: `http://localhost:3003/oauth/callback` (Web)
5. After creation, note your **Application (client) ID** and **Directory (tenant) ID**
6. Go to **Certificates & secrets** → **New client secret**
   - Add a description and expiration
   - Copy the secret value immediately (you won't see it again)
7. Go to **API permissions** → **Add a permission**
   - Select **Microsoft Graph** → **Delegated permissions**
   - Add these permissions:
     - `Calendars.Read`
     - `User.Read`
   - Click **Grant admin consent** (if you have admin rights)

### 2. Installation

```bash
# Clone or navigate to the project directory
cd outlook-calendar-mcp-server

# Install dependencies
npm install
```

### 3. Configuration

Create a `.env` file in the root directory:

```bash
cp env.example .env
```

Edit `.env` with your configuration:

```env
# Server Configuration
PORT=3003
NODE_ENV=development

# Microsoft OAuth Configuration
MICROSOFT_CLIENT_ID=your_microsoft_client_id_here
MICROSOFT_CLIENT_SECRET=your_microsoft_client_secret_here
MICROSOFT_TENANT_ID=common
MICROSOFT_REDIRECT_URI=http://localhost:3003/oauth/callback

# Frontend URL (where to redirect after OAuth)
FRONTEND_URL=http://localhost:5173

# MCP Server Configuration
MCP_SERVER_NAME=outlook-calendar-mcp-server
MCP_SERVER_VERSION=1.0.0
```

**Note:** Use `common` for multi-tenant apps, or your specific tenant ID for single-tenant apps.

## Usage

### Development Mode

```bash
npm run dev
```

### Production Mode

```bash
# Build
npm run build

# Start
npm start
```

### Cloudflare Workers Deployment

```bash
# Create KV namespace for sessions
npm run kv:create

# Update wrangler.toml with the KV namespace ID

# Deploy
npm run deploy
```

## API Endpoints

### Health Check
```
GET /health
```

### OAuth Flow

#### 1. Initiate Authorization
```
GET /oauth/authorize?userId={userId}&state={state}
```

Response:
```json
{
  "authorizationUrl": "https://login.microsoftonline.com/...",
  "state": "your-state-value"
}
```

#### 2. OAuth Callback (handled automatically)
```
GET /oauth/callback?code={code}&state={state}
```

#### 3. Refresh Token
```
POST /oauth/refresh
Content-Type: application/json

{
  "userId": "user123",
  "refreshToken": "refresh_token_here"
}
```

#### 4. Disconnect
```
POST /oauth/disconnect
Content-Type: application/json

{
  "userId": "user123"
}
```

### MCP Endpoint

```
POST /mcp
GET /mcp
DELETE /mcp
```

The MCP endpoint follows the Model Context Protocol specification for tool execution.

## MCP Tools

### 1. outlook_calendar_search_events

Search Outlook Calendar events with date filters and other criteria.

**Input:**
```json
{
  "userId": "user123",
  "startDateTime": "2024-01-01T00:00:00",
  "endDateTime": "2024-12-31T23:59:59",
  "subject": "Meeting",
  "location": "Conference Room",
  "category": "Work",
  "top": 20,
  "skip": 0,
  "orderBy": "start/dateTime"
}
```

**Output:**
```json
{
  "events": [
    {
      "id": "event123",
      "subject": "Team Meeting",
      "startDateTime": "2024-01-15T10:00:00",
      "endDateTime": "2024-01-15T11:00:00",
      "timeZone": "Pacific Standard Time",
      "location": "Conference Room A",
      "organizer": "organizer@example.com",
      "attendees": [
        {
          "email": "attendee@example.com",
          "name": "John Doe",
          "response": "accepted"
        }
      ],
      "isAllDay": false,
      "webLink": "https://outlook.office365.com/...",
      "onlineMeetingUrl": "https://teams.microsoft.com/..."
    }
  ],
  "count": 1
}
```

**Scopes Required:** `Calendars.Read`

### 2. outlook_calendar_fetch_event

Retrieve details for a single calendar event by ID.

**Input:**
```json
{
  "userId": "user123",
  "eventId": "event123"
}
```

**Output:**
```json
{
  "event": {
    "id": "event123",
    "subject": "Team Meeting",
    "bodyPreview": "Discuss Q4 goals...",
    "startDateTime": "2024-01-15T10:00:00",
    "endDateTime": "2024-01-15T11:00:00",
    "location": "Conference Room A",
    "organizer": "organizer@example.com"
  }
}
```

**Scopes Required:** `Calendars.Read`

### 3. outlook_calendar_fetch_events_batch

Retrieve multiple events in one call by their IDs.

**Input:**
```json
{
  "userId": "user123",
  "eventIds": ["event123", "event456", "event789"]
}
```

**Output:**
```json
{
  "events": [
    {
      "id": "event123",
      "subject": "Team Meeting",
      "startDateTime": "2024-01-15T10:00:00",
      "endDateTime": "2024-01-15T11:00:00"
    },
    {
      "id": "event456",
      "subject": "Client Call",
      "startDateTime": "2024-01-16T14:00:00",
      "endDateTime": "2024-01-16T15:00:00"
    }
  ],
  "count": 2
}
```

**Scopes Required:** `Calendars.Read`

### 4. outlook_calendar_list_events

List calendar events within a specific date range.

**Input:**
```json
{
  "userId": "user123",
  "startDateTime": "2024-01-01T00:00:00",
  "endDateTime": "2024-01-31T23:59:59",
  "top": 50,
  "skip": 0
}
```

**Output:**
```json
{
  "events": [
    {
      "id": "event123",
      "subject": "Team Meeting",
      "startDateTime": "2024-01-15T10:00:00",
      "endDateTime": "2024-01-15T11:00:00",
      "location": "Conference Room A"
    }
  ],
  "count": 1
}
```

**Scopes Required:** `Calendars.Read`

### 5. outlook_calendar_get_profile

Retrieve the current user's Microsoft profile information.

**Input:**
```json
{
  "userId": "user123"
}
```

**Output:**
```json
{
  "profile": {
    "id": "user-uuid",
    "displayName": "John Doe",
    "givenName": "John",
    "surname": "Doe",
    "mail": "john.doe@example.com",
    "userPrincipalName": "john.doe@company.com",
    "jobTitle": "Software Engineer",
    "officeLocation": "Building 1",
    "mobilePhone": "+1234567890",
    "businessPhones": ["+0987654321"]
  }
}
```

**Scopes Required:** `User.Read`

## Architecture

```
┌─────────────────┐
│   Frontend      │
│   (ZeroTwo)     │
└────────┬────────┘
         │
         │ OAuth Flow
         ▼
┌─────────────────┐      ┌──────────────────┐
│  Outlook        │◄────►│  Microsoft OAuth │
│  Calendar MCP   │      │  & Graph API     │
│  Server         │      └──────────────────┘
└────────┬────────┘
         │
         │ MCP Protocol
         ▼
┌─────────────────┐
│   Backend       │
│   (ZeroTwoApi)  │
└─────────────────┘
```

## Development

### Project Structure

```
outlook-calendar-mcp-server/
├── src/
│   ├── auth/
│   │   └── oauth-manager.ts    # OAuth token management
│   ├── config/
│   │   └── index.ts            # Configuration management
│   ├── outlook/
│   │   └── client.ts           # Microsoft Graph API wrapper
│   ├── mcp/
│   │   ├── server.ts           # MCP server implementation
│   │   └── tools.ts            # MCP tool definitions
│   ├── types/
│   │   └── index.ts            # TypeScript type definitions
│   ├── utils/
│   │   └── logger.ts           # Logging utility
│   ├── index.ts                # Entry point
│   └── worker.ts               # Cloudflare Worker entry point
├── package.json
├── tsconfig.json
├── wrangler.toml
├── env.example
└── README.md
```

### Scripts

- `npm run dev` - Start development server with hot reload
- `npm run dev:worker` - Start Cloudflare Workers development mode
- `npm run build` - Build TypeScript to JavaScript
- `npm run start` - Start production server
- `npm run deploy` - Deploy to Cloudflare Workers
- `npm run lint` - Run ESLint
- `npm run format` - Format code with Prettier

## Microsoft Graph API Reference

This server uses the following Microsoft Graph API endpoints:

- **Events**: `GET /me/events`
- **Single Event**: `GET /me/events/{eventId}`
- **Batch Request**: `POST /$batch`
- **User Profile**: `GET /me`

For more information, see [Microsoft Graph Calendar API Documentation](https://learn.microsoft.com/en-us/graph/api/resources/calendar).

## Security Considerations

1. **OAuth Tokens**: Tokens are stored in memory (Node.js) or KV (Cloudflare Workers). For production, consider encrypted storage.
2. **CORS**: Configure CORS origins appropriately for production.
3. **HTTPS**: Always use HTTPS in production for secure communication.
4. **Environment Variables**: Never commit `.env` files. Use secure secret management in production.
5. **Rate Limiting**: Implement rate limiting for production deployments.
6. **Tenant ID**: Use specific tenant ID instead of `common` for single-tenant applications.

## Troubleshooting

### OAuth Errors

- **"Missing required environment variables"**: Check your `.env` file has all required variables
- **"Failed to exchange authorization code"**: Verify your redirect URI matches exactly in Azure Portal
- **"Token expired"**: The server automatically refreshes tokens if a refresh token is available
- **"AADSTS50011: The reply URL specified in the request does not match"**: Ensure the redirect URI in Azure matches your configuration

### MCP Connection Issues

- **"No valid session ID"**: Ensure the MCP client sends proper session headers
- **"Authentication required"**: User needs to complete OAuth flow first

### Microsoft Graph API Errors

- **"401 Unauthorized"**: Token may be expired or invalid. Try re-authenticating.
- **"403 Forbidden"**: Check that all required permissions are granted and admin consent is provided.
- **"404 Not Found"**: Verify the event ID or resource exists.

## Microsoft Graph API Scopes

This server requires the following Microsoft Graph scopes:

| Scope | Type | Description | Admin Consent Required |
|-------|------|-------------|----------------------|
| `Calendars.Read` | Delegated | Read user calendars | No |
| `User.Read` | Delegated | Read user profile | No |

## Integration with ZeroTwo

To integrate this Outlook Calendar MCP server with your ZeroTwo application:

### 1. Add OAuth Provider Configuration

Create a new OAuth provider configuration for Outlook Calendar:

```javascript
// In your OAuth providers configuration
outlook_calendar: {
  provider: "outlook_calendar",
  clientId: import.meta.env.VITE_MICROSOFT_CLIENT_ID,
  flowType: OAuthFlowType.REDIRECT,
  scopes: [
    "https://graph.microsoft.com/Calendars.Read",
    "https://graph.microsoft.com/User.Read",
  ],
  authEndpoint: "http://localhost:3003/oauth/authorize",
  backendExchangeEndpoint: "/api/auth/outlook/callback",
  backendRefreshEndpoint: "/api/auth/outlook/refresh",
  mcpEnabled: true,
  mcpEndpoint: "http://localhost:3003/mcp",
}
```

### 2. Update MCP Client

Add Outlook Calendar support to your MCP client configuration.

### 3. Backend Integration

Create backend endpoints to handle OAuth callback and token management, storing tokens in user profiles.

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions:
- Check the [Troubleshooting](#troubleshooting) section
- Review [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/overview)
- Open an issue on GitHub








---

## Powered by ZeroTwo

This Outlook Calendar MCP connector is part of the [ZeroTwo AI platform](https://zerotwo.ai) — the all-in-one AI workspace that lets you schedule meetings, manage events, and automate your Microsoft calendar through GPT-5, Claude, and Gemini.

| | |
|---|---|
| 🌐 **[ZeroTwo — All AI Models in One App](https://zerotwo.ai)** | Manage your Outlook calendar with GPT-5, Claude, and Gemini — all in one AI workspace. |
| ✨ **[ZeroTwo Features](https://zerotwo.ai/features)** | AI scheduling, email drafting, web search, and MCP-powered Microsoft 365 tools. |
| 🤖 **[AI Models — GPT-5, Claude & Gemini](https://zerotwo.ai/zerotwo-models)** | Use the world's best AI to book meetings, find availability, and manage your time. |
| 🔌 **[ZeroTwo Connectors & Integrations](https://zerotwo.ai/connectors)** | Connect Outlook Calendar, Gmail, Teams, SharePoint, and more to your AI workflow. |
| 💰 **[ZeroTwo Pricing](https://zerotwo.ai/pricing)** | One subscription that replaces ChatGPT Plus, Claude Pro, and Gemini Advanced. |
| 📝 **[ZeroTwo Blog](https://zerotwo.ai/blog)** | AI productivity tips, Microsoft 365 guides, and ZeroTwo product updates. |
| 🚀 **[Try ZeroTwo Free](https://app.zerotwo.ai/auth/login)** | Let AI manage your Outlook schedule — get started free today. |

> **Built for ZeroTwo** — Use this Outlook Calendar MCP server with [ZeroTwo's AI connector system](https://zerotwo.ai/connectors) to create events, check your schedule, and manage meetings through natural language in your AI assistant.
