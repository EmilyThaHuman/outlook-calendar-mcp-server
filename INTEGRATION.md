# ZeroTwo Integration Guide - Outlook Calendar MCP Server

This guide provides step-by-step instructions for integrating the Outlook Calendar MCP Server with your ZeroTwo application.

## Overview

The Outlook Calendar MCP Server enables your ZeroTwo AI assistant to:
- Search and retrieve calendar events
- Access user profile information
- Filter events by date, location, subject, and categories
- Fetch multiple events efficiently in batch operations

## Deployment Information

- **Server URL:** `https://outlook-calendar-mcp-server.reed-b9b.workers.dev`
- **Client ID:** `4beaa1c6-5219-4626-ad6e-203fb72e45b3`
- **Provider:** `outlook_calendar`
- **Scopes:** `Calendars.Read`, `User.Read`

## Step 1: Update Azure App Registration

Before integration, ensure your Azure App Registration is configured correctly:

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Select the app with Client ID: `4beaa1c6-5219-4626-ad6e-203fb72e45b3`
4. Under **Authentication**, add redirect URIs:
   - `https://outlook-calendar-mcp-server.reed-b9b.workers.dev/oauth/callback`
   - `https://zerotwo.ai/auth/callback` (if using frontend callback)
5. Under **API permissions**, verify:
   - `Calendars.Read` (Microsoft Graph, Delegated)
   - `User.Read` (Microsoft Graph, Delegated)
6. Grant admin consent if required

## Step 2: Frontend Configuration

### Add OAuth Provider

In your OAuth providers configuration (e.g., `src/services/oauthProviders.ts`):

```typescript
import { OAuthFlowType } from './types';

export const oauthProviders = {
  // ... existing providers
  
  outlook_calendar: {
    provider: "outlook_calendar",
    displayName: "Outlook Calendar",
    icon: "calendar", // or appropriate icon
    clientId: "4beaa1c6-5219-4626-ad6e-203fb72e45b3",
    flowType: OAuthFlowType.REDIRECT,
    scopes: [
      "https://graph.microsoft.com/Calendars.Read",
      "https://graph.microsoft.com/User.Read",
    ],
    authEndpoint: "https://outlook-calendar-mcp-server.reed-b9b.workers.dev/oauth/authorize",
    tokenEndpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    backendExchangeEndpoint: "/api/auth/outlook_calendar/callback",
    backendRefreshEndpoint: "/api/auth/outlook_calendar/refresh",
    backendDisconnectEndpoint: "/api/auth/outlook_calendar/disconnect",
    mcpEnabled: true,
    mcpEndpoint: "https://outlook-calendar-mcp-server.reed-b9b.workers.dev/mcp",
    description: "Connect your Outlook Calendar to search and manage events",
  },
};
```

### Add to Settings Page

In your settings/integrations page:

```tsx
<IntegrationCard
  provider="outlook_calendar"
  title="Outlook Calendar"
  description="Access your Microsoft Outlook calendar events"
  icon={<CalendarIcon />}
  scopes={[
    { name: "Calendars.Read", description: "Read your calendar events" },
    { name: "User.Read", description: "Read your profile information" },
  ]}
  onConnect={handleConnect}
  onDisconnect={handleDisconnect}
  isConnected={isOutlookConnected}
/>
```

## Step 3: Backend API Implementation

### Create Backend Routes

Create routes to handle OAuth callbacks and token management:

```typescript
// routes/auth/outlook-calendar.ts

import { Router } from 'express';
import { outlookCalendarService } from '../../services/outlookCalendarService';

const router = Router();

// OAuth callback handler
router.get('/callback', async (req, res) => {
  try {
    const { code, state, error } = req.query;
    
    if (error) {
      return res.redirect(`${process.env.FRONTEND_URL}/settings?oauth_error=${error}`);
    }
    
    if (!code || !state) {
      return res.status(400).json({ error: 'Missing code or state' });
    }
    
    const userId = state as string; // Validate/decode state properly
    
    // Exchange code for tokens via MCP server
    const tokens = await outlookCalendarService.exchangeCode(code);
    
    // Store tokens in database
    await outlookCalendarService.storeUserTokens(userId, tokens);
    
    // Sync tokens to MCP server KV
    await outlookCalendarService.syncTokensToWorker(userId, tokens);
    
    res.redirect(`${process.env.FRONTEND_URL}/settings?oauth_success=true&provider=outlook_calendar`);
  } catch (error) {
    console.error('Outlook OAuth callback error:', error);
    res.redirect(`${process.env.FRONTEND_URL}/settings?oauth_error=callback_failed`);
  }
});

// Token refresh handler
router.post('/refresh', async (req, res) => {
  try {
    const { userId } = req.body;
    
    if (!userId) {
      return res.status(400).json({ error: 'Missing userId' });
    }
    
    const tokens = await outlookCalendarService.refreshUserTokens(userId);
    
    res.json({
      success: true,
      expiresIn: tokens.expires_in,
    });
  } catch (error) {
    console.error('Token refresh error:', error);
    res.status(500).json({ error: 'Failed to refresh token' });
  }
});

// Disconnect handler
router.post('/disconnect', async (req, res) => {
  try {
    const { userId } = req.body;
    
    if (!userId) {
      return res.status(400).json({ error: 'Missing userId' });
    }
    
    await outlookCalendarService.disconnectUser(userId);
    
    res.json({ success: true });
  } catch (error) {
    console.error('Disconnect error:', error);
    res.status(500).json({ error: 'Failed to disconnect' });
  }
});

// Token retrieval for MCP worker (fallback)
router.get('/token', async (req, res) => {
  try {
    const { userId } = req.query;
    
    if (!userId) {
      return res.status(400).json({ error: 'Missing userId' });
    }
    
    const tokens = await outlookCalendarService.getUserTokens(userId as string);
    
    if (!tokens) {
      return res.status(404).json({ error: 'No tokens found' });
    }
    
    res.json({
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresIn: tokens.expires_in,
    });
  } catch (error) {
    console.error('Token retrieval error:', error);
    res.status(500).json({ error: 'Failed to retrieve tokens' });
  }
});

export default router;
```

### Create Service Layer

```typescript
// services/outlookCalendarService.ts

import axios from 'axios';
import { supabase } from '../lib/supabase';

const MCP_SERVER_URL = 'https://outlook-calendar-mcp-server.reed-b9b.workers.dev';

export const outlookCalendarService = {
  async exchangeCode(code: string): Promise<any> {
    // This would typically be done by the MCP server
    // But you can also handle it here if needed
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: process.env.MICROSOFT_CLIENT_ID!,
        client_secret: process.env.MICROSOFT_CLIENT_SECRET!,
        code,
        redirect_uri: `${MCP_SERVER_URL}/oauth/callback`,
        grant_type: 'authorization_code',
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }
    );
    
    return response.data;
  },

  async storeUserTokens(userId: string, tokens: any): Promise<void> {
    const { error } = await supabase
      .from('profiles')
      .update({
        settings: {
          oauth_tokens: {
            outlook_calendar: {
              access_token: tokens.access_token,
              refresh_token: tokens.refresh_token,
              expires_at: new Date(Date.now() + tokens.expires_in * 1000).toISOString(),
              scope: tokens.scope,
              updated_at: new Date().toISOString(),
            },
          },
        },
      })
      .eq('id', userId);
    
    if (error) {
      throw new Error(`Failed to store tokens: ${error.message}`);
    }
  },

  async syncTokensToWorker(userId: string, tokens: any): Promise<void> {
    await axios.post(`${MCP_SERVER_URL}/oauth/sync`, {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresIn: tokens.expires_in,
    });
  },

  async getUserTokens(userId: string): Promise<any> {
    const { data, error } = await supabase
      .from('profiles')
      .select('settings')
      .eq('id', userId)
      .single();
    
    if (error) {
      throw new Error(`Failed to retrieve tokens: ${error.message}`);
    }
    
    return data?.settings?.oauth_tokens?.outlook_calendar;
  },

  async refreshUserTokens(userId: string): Promise<any> {
    const tokens = await this.getUserTokens(userId);
    
    if (!tokens?.refresh_token) {
      throw new Error('No refresh token available');
    }
    
    const response = await axios.post(`${MCP_SERVER_URL}/oauth/refresh`, {
      userId,
      refreshToken: tokens.refresh_token,
    });
    
    await this.storeUserTokens(userId, response.data);
    
    return response.data;
  },

  async disconnectUser(userId: string): Promise<void> {
    // Remove from MCP server KV
    await axios.post(`${MCP_SERVER_URL}/oauth/disconnect`, { userId });
    
    // Remove from database
    const { error } = await supabase
      .from('profiles')
      .update({
        settings: {
          oauth_tokens: {
            outlook_calendar: null,
          },
        },
      })
      .eq('id', userId);
    
    if (error) {
      throw new Error(`Failed to disconnect: ${error.message}`);
    }
  },
};
```

## Step 4: MCP Client Integration

### Update MCP Client Configuration

In your MCP client setup:

```typescript
// lib/mcp-client.ts

export async function createMcpClient(integration: any) {
  let endpointUrl = integration.endpoint;
  let transportOptions: any = {};

  // Configure for Outlook Calendar MCP
  if (integration.provider === 'outlook_calendar') {
    endpointUrl = 'https://outlook-calendar-mcp-server.reed-b9b.workers.dev/mcp';
    
    // Use OAuth token from profile if available
    if (integration.oauth_token) {
      transportOptions.headers = {
        'X-User-Id': integration.user_id,
        ...integration.headers,
      };
    }
  }

  // ... rest of client setup
}
```

## Step 5: AI Assistant Tool Usage

### Register MCP Tools

The following tools will be automatically available to your AI assistant:

1. **search_events** - Search calendar events
2. **fetch_event** - Get specific event details
3. **fetch_events_batch** - Get multiple events at once
4. **list_events** - List events in a date range
5. **get_profile** - Get user profile

### Example Tool Calls

The AI assistant can now respond to queries like:

```
User: "What meetings do I have tomorrow?"
Assistant: [Calls search_events with tomorrow's date range]

User: "Show me all my meetings with John this week"
Assistant: [Calls search_events with date range and subject filter]

User: "Get details for event abc123"
Assistant: [Calls fetch_event with event ID]
```

## Step 6: Testing the Integration

### Test OAuth Flow

1. Navigate to your settings page
2. Click "Connect Outlook Calendar"
3. Complete Microsoft authentication
4. Verify success redirect to settings page
5. Check that integration shows as "Connected"

### Test MCP Tools

In your AI chat interface:

```
User: "What's on my calendar today?"
Expected: AI calls list_events tool and displays events

User: "Show my profile information"
Expected: AI calls get_profile tool and displays user details
```

### Debug Mode

Enable debug logging to see MCP tool calls:

```typescript
// In your AI chat component
const debug = true;

if (debug) {
  console.log('MCP Tool Call:', {
    tool: toolName,
    args: toolArgs,
    response: toolResponse,
  });
}
```

## Step 7: Error Handling

### Common Errors and Solutions

#### "Authentication required"
- User hasn't completed OAuth flow
- Token expired without refresh token
- Session not found in KV or database

**Solution:** Prompt user to reconnect Outlook Calendar

#### "Token refresh failed"
- Refresh token expired or revoked
- Client credentials invalid

**Solution:** Force re-authentication

#### "CORS error"
- Frontend domain not in allowed origins

**Solution:** Update CORS configuration in worker

## Security Considerations

1. **Token Storage**
   - Store tokens encrypted in database
   - Use secure session management
   - Implement token rotation

2. **Scope Management**
   - Only request necessary scopes
   - Inform users about data access
   - Implement scope upgrade flow if needed

3. **Rate Limiting**
   - Implement rate limiting on backend endpoints
   - Cache frequently accessed data
   - Monitor Microsoft Graph API quotas

## Monitoring

### Key Metrics to Track

1. OAuth conversion rate
2. Token refresh success rate
3. MCP tool call frequency
4. Error rates by type
5. API response times

### Logging

Ensure logging for:
- OAuth flow steps
- Token refresh attempts
- MCP tool calls
- API errors

## Support Resources

- **Microsoft Graph API Docs:** https://learn.microsoft.com/en-us/graph/
- **Azure AD Documentation:** https://learn.microsoft.com/en-us/azure/active-directory/
- **MCP Protocol Spec:** https://modelcontextprotocol.io/

## Troubleshooting Checklist

- [ ] Azure App Registration redirect URI matches
- [ ] Required Microsoft Graph permissions granted
- [ ] Secrets configured in Cloudflare Worker
- [ ] Frontend OAuth provider configured
- [ ] Backend routes implemented
- [ ] Database schema supports token storage
- [ ] CORS domains whitelisted
- [ ] MCP client can connect to worker
- [ ] Error handling implemented
- [ ] Monitoring and logging enabled

## Next Steps

1. Complete Azure App Registration configuration
2. Implement backend routes
3. Add frontend integration components
4. Test OAuth flow end-to-end
5. Test MCP tool calls through AI assistant
6. Monitor logs for errors
7. Gather user feedback
8. Optimize based on usage patterns

---

**Integration Guide Version:** 1.0  
**Last Updated:** November 15, 2025  
**MCP Server:** https://outlook-calendar-mcp-server.reed-b9b.workers.dev







