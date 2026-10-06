/**
 * Outlook Calendar API Client wrapper using Microsoft Graph
 */

import { config } from '../config/index.js';
import { logger } from '../utils/logger.js';
import {
  OutlookEvent,
  SimplifiedEvent,
  OutlookProfile,
  EventFilterOptions,
} from '../types/index.js';

export class OutlookCalendarClient {
  /**
   * Make authenticated request to Microsoft Graph API
   */
  private async makeGraphRequest(
    accessToken: string,
    endpoint: string,
    method: string = 'GET',
    body?: any
  ): Promise<any> {
    const url = `${config.graph.endpoint}${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
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
      logger.error('[OutlookCalendarClient] Graph API request failed:', {
        endpoint,
        status: response.status,
        error,
      });
      throw new Error(`Microsoft Graph API error: ${response.status} ${response.statusText}`);
    }

    return response.json();
  }

  /**
   * Simplify event for output
   */
  private simplifyEvent(event: OutlookEvent): SimplifiedEvent {
    return {
      id: event.id,
      subject: event.subject,
      bodyPreview: event.bodyPreview,
      startDateTime: event.start.dateTime,
      endDateTime: event.end.dateTime,
      timeZone: event.start.timeZone,
      location: event.location?.displayName,
      organizer: event.organizer?.emailAddress.address,
      attendees: event.attendees?.map((attendee) => ({
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

  /**
   * Build filter query string
   */
  private buildFilterQuery(options: EventFilterOptions): string {
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

    return filters.length > 0 ? filters.join(' and ') : '';
  }

  /**
   * Search calendar events with filters
   */
  async searchEvents(
    accessToken: string,
    options: EventFilterOptions = {}
  ): Promise<SimplifiedEvent[]> {
    try {
      logger.info('[OutlookCalendarClient] Searching events', options);

      const params = new URLSearchParams();

      // Build filter query
      const filterQuery = this.buildFilterQuery(options);
      if (filterQuery) {
        params.append('$filter', filterQuery);
      }

      // Add search if subject is provided
      if (options.subject && !filterQuery.includes('subject')) {
        params.append('$search', `"${options.subject}"`);
      }

      // Add pagination
      if (options.top) {
        params.append('$top', options.top.toString());
      }
      if (options.skip) {
        params.append('$skip', options.skip.toString());
      }

      // Add ordering
      if (options.orderBy) {
        params.append('$orderby', options.orderBy);
      } else {
        params.append('$orderby', 'start/dateTime');
      }

      const endpoint = `/me/events${params.toString() ? '?' + params.toString() : ''}`;
      const data = await this.makeGraphRequest(accessToken, endpoint);

      const events: OutlookEvent[] = data.value || [];
      const simplified = events.map((event) => this.simplifyEvent(event));

      logger.info('[OutlookCalendarClient] Successfully searched events', {
        count: simplified.length,
      });

      return simplified;
    } catch (error) {
      logger.error('[OutlookCalendarClient] Error searching events:', error);
      throw new Error('Failed to search calendar events');
    }
  }

  /**
   * Fetch a single event by ID
   */
  async fetchEvent(accessToken: string, eventId: string): Promise<SimplifiedEvent> {
    try {
      logger.info('[OutlookCalendarClient] Fetching event', { eventId });

      const endpoint = `/me/events/${eventId}`;
      const event: OutlookEvent = await this.makeGraphRequest(accessToken, endpoint);

      logger.info('[OutlookCalendarClient] Successfully fetched event', { eventId });

      return this.simplifyEvent(event);
    } catch (error) {
      logger.error('[OutlookCalendarClient] Error fetching event:', error);
      throw new Error('Failed to fetch calendar event');
    }
  }

  /**
   * Fetch multiple events in batch
   */
  async fetchEventsBatch(
    accessToken: string,
    eventIds: string[]
  ): Promise<SimplifiedEvent[]> {
    try {
      logger.info('[OutlookCalendarClient] Fetching events batch', {
        count: eventIds.length,
      });

      // Microsoft Graph batch request
      const batchRequests = eventIds.map((id, index) => ({
        id: index.toString(),
        method: 'GET',
        url: `/me/events/${id}`,
      }));

      const batchRequestBody = {
        requests: batchRequests,
      };

      const endpoint = '/$batch';
      const response = await this.makeGraphRequest(
        accessToken,
        endpoint,
        'POST',
        batchRequestBody
      );

      const events: SimplifiedEvent[] = [];

      for (const batchResponse of response.responses || []) {
        if (batchResponse.status === 200 && batchResponse.body) {
          events.push(this.simplifyEvent(batchResponse.body as OutlookEvent));
        } else {
          logger.warn('[OutlookCalendarClient] Failed to fetch event in batch', {
            id: batchResponse.id,
            status: batchResponse.status,
          });
        }
      }

      logger.info('[OutlookCalendarClient] Successfully fetched events batch', {
        requested: eventIds.length,
        fetched: events.length,
      });

      return events;
    } catch (error) {
      logger.error('[OutlookCalendarClient] Error fetching events batch:', error);
      throw new Error('Failed to fetch calendar events in batch');
    }
  }

  /**
   * List calendar events within a date range
   */
  async listEvents(
    accessToken: string,
    startDateTime: string,
    endDateTime: string,
    options: { top?: number; skip?: number } = {}
  ): Promise<SimplifiedEvent[]> {
    try {
      logger.info('[OutlookCalendarClient] Listing events', {
        startDateTime,
        endDateTime,
        ...options,
      });

      return await this.searchEvents(accessToken, {
        startDateTime,
        endDateTime,
        top: options.top || 50,
        skip: options.skip,
        orderBy: 'start/dateTime',
      });
    } catch (error) {
      logger.error('[OutlookCalendarClient] Error listing events:', error);
      throw new Error('Failed to list calendar events');
    }
  }

  /**
   * Get user profile
   */
  async getProfile(accessToken: string): Promise<OutlookProfile> {
    try {
      logger.info('[OutlookCalendarClient] Getting user profile');

      const endpoint = '/me';
      const profile: OutlookProfile = await this.makeGraphRequest(accessToken, endpoint);

      logger.info('[OutlookCalendarClient] Successfully retrieved user profile', {
        userId: profile.id,
      });

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
    } catch (error) {
      logger.error('[OutlookCalendarClient] Error getting user profile:', error);
      throw new Error('Failed to get user profile');
    }
  }
}

// Singleton instance
export const outlookCalendarClient = new OutlookCalendarClient();







