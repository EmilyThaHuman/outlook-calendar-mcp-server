/**
 * MCP Tool definitions for Outlook Calendar operations
 */

import { z } from 'zod';
import { outlookCalendarClient } from '../outlook/client.js';
import { oauthManager } from '../auth/oauth-manager.js';
import { logger } from '../utils/logger.js';

/**
 * Search Events Tool
 */
export const searchEventsTool = {
  name: 'outlook_calendar_search_events',
  definition: {
    title: 'Search Outlook Calendar Events',
    description: 'Use this to find Outlook Calendar events by time range and optional subject, location, or category filters. It returns event IDs that should be passed to `outlook_calendar_fetch_event` or `outlook_calendar_fetch_events_batch` when full details are needed.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      startDateTime: z
        .string()
        .optional()
        .describe('Start date/time in ISO format (e.g., 2024-01-01T00:00:00)'),
      endDateTime: z
        .string()
        .optional()
        .describe('End date/time in ISO format (e.g., 2024-12-31T23:59:59)'),
      subject: z.string().optional().describe('Search by event subject/title'),
      location: z.string().optional().describe('Search by location'),
      category: z.string().optional().describe('Filter by category'),
      top: z.number().min(1).max(100).default(20).describe('Maximum number of events to return'),
      skip: z.number().min(0).optional().describe('Number of events to skip for pagination'),
      orderBy: z.string().optional().describe('Field to order by (e.g., start/dateTime)'),
    },
    outputSchema: {
      events: z.array(
        z.object({
          id: z.string(),
          subject: z.string(),
          startDateTime: z.string(),
          endDateTime: z.string(),
          location: z.string().optional(),
          organizer: z.string().optional(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const {
        userId,
        startDateTime,
        endDateTime,
        subject,
        location,
        category,
        top,
        skip,
        orderBy,
      } = args;

      logger.info('[Tool:outlook_calendar_search_events] Executing', { userId, startDateTime, endDateTime });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Search events
      const events = await outlookCalendarClient.searchEvents(accessToken, {
        startDateTime,
        endDateTime,
        subject,
        location,
        category,
        top,
        skip,
        orderBy,
      });

      const output = {
        events,
        count: events.length,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:outlook_calendar_search_events] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error searching events: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Fetch Event Tool
 */
export const fetchEventTool = {
  name: 'outlook_calendar_fetch_event',
  definition: {
    title: 'Retrieve details for a single event',
    description: 'Use this to read one specific Outlook Calendar event when you already have its `eventId`, usually from `outlook_calendar_search_events` or `outlook_calendar_list_events`.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      eventId: z.string().describe('Outlook Calendar event ID returned by `outlook_calendar_search_events` or `outlook_calendar_list_events`.'),
    },
    outputSchema: {
      event: z.object({
        id: z.string(),
        subject: z.string(),
        startDateTime: z.string(),
        endDateTime: z.string(),
        location: z.string().optional(),
        organizer: z.string().optional(),
      }),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, eventId } = args;

      logger.info('[Tool:outlook_calendar_fetch_event] Executing', { userId, eventId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Fetch event
      const event = await outlookCalendarClient.fetchEvent(accessToken, eventId);

      const output = { event };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:outlook_calendar_fetch_event] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching event: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Fetch Events Batch Tool
 */
export const fetchEventsBatchTool = {
  name: 'outlook_calendar_fetch_events_batch',
  definition: {
    title: 'Retrieve multiple events in one call',
    description: 'Use this to read several known Outlook Calendar events in one request after a search or list step. Pass the event IDs returned by the Outlook Calendar discovery tools.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      eventIds: z.array(z.string()).describe('Array of Outlook Calendar event IDs returned by `outlook_calendar_search_events` or `outlook_calendar_list_events`.'),
    },
    outputSchema: {
      events: z.array(
        z.object({
          id: z.string(),
          subject: z.string(),
          startDateTime: z.string(),
          endDateTime: z.string(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, eventIds } = args;

      logger.info('[Tool:outlook_calendar_fetch_events_batch] Executing', { userId, count: eventIds?.length });

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

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Fetch events in batch
      const events = await outlookCalendarClient.fetchEventsBatch(accessToken, eventIds);

      const output = {
        events,
        count: events.length,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:outlook_calendar_fetch_events_batch] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching events batch: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * List Events Tool
 */
export const listEventsTool = {
  name: 'outlook_calendar_list_events',
  definition: {
    title: 'List calendar events within a date range',
    description: 'Use this to browse Outlook Calendar events inside a required date window when you do not need text search. It returns event IDs for follow-up reads with `outlook_calendar_fetch_event`.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      startDateTime: z
        .string()
        .describe('Start date/time in ISO format (e.g., 2024-01-01T00:00:00)'),
      endDateTime: z
        .string()
        .describe('End date/time in ISO format (e.g., 2024-12-31T23:59:59)'),
      top: z.number().min(1).max(100).default(50).describe('Maximum number of events to return'),
      skip: z.number().min(0).optional().describe('Number of events to skip for pagination'),
    },
    outputSchema: {
      events: z.array(
        z.object({
          id: z.string(),
          subject: z.string(),
          startDateTime: z.string(),
          endDateTime: z.string(),
          location: z.string().optional(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, startDateTime, endDateTime, top, skip } = args;

      logger.info('[Tool:outlook_calendar_list_events] Executing', { userId, startDateTime, endDateTime });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // List events
      const events = await outlookCalendarClient.listEvents(
        accessToken,
        startDateTime,
        endDateTime,
        { top, skip }
      );

      const output = {
        events,
        count: events.length,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:outlook_calendar_list_events] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error listing events: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Get Profile Tool
 */
export const getProfileTool = {
  name: 'outlook_calendar_get_profile',
  definition: {
    title: "Retrieve the current user's profile",
    description: 'Use this when you need the authenticated Outlook/Microsoft 365 user profile for calendar context. It does not return events.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
    },
    outputSchema: {
      profile: z.object({
        id: z.string(),
        displayName: z.string(),
        mail: z.string().optional(),
        userPrincipalName: z.string(),
      }),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId } = args;

      logger.info('[Tool:outlook_calendar_get_profile] Executing', { userId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
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

      // Get profile
      const profile = await outlookCalendarClient.getProfile(accessToken);

      const output = { profile };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:outlook_calendar_get_profile] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error getting profile: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Export all tools
 */
export const outlookCalendarTools = [
  searchEventsTool,
  fetchEventTool,
  fetchEventsBatchTool,
  listEventsTool,
  getProfileTool,
];






