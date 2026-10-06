/**
 * Type definitions for Outlook Calendar MCP Server
 */

// Microsoft Graph Calendar Event
export interface OutlookEvent {
  id: string;
  subject: string;
  body?: {
    contentType: string;
    content: string;
  };
  bodyPreview?: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
    locationType?: string;
    uniqueId?: string;
    uniqueIdType?: string;
  };
  locations?: Array<{
    displayName: string;
    locationType?: string;
  }>;
  attendees?: Array<{
    emailAddress: {
      address: string;
      name?: string;
    };
    status?: {
      response: string;
      time: string;
    };
    type: string;
  }>;
  organizer?: {
    emailAddress: {
      address: string;
      name?: string;
    };
  };
  isAllDay: boolean;
  isCancelled: boolean;
  isOrganizer: boolean;
  responseRequested: boolean;
  responseStatus?: {
    response: string;
    time: string;
  };
  importance: string;
  sensitivity: string;
  showAs: string;
  recurrence?: {
    pattern: {
      type: string;
      interval: number;
      daysOfWeek?: string[];
      dayOfMonth?: number;
      month?: number;
    };
    range: {
      type: string;
      startDate: string;
      endDate?: string;
      numberOfOccurrences?: number;
    };
  };
  webLink?: string;
  onlineMeetingUrl?: string;
  isOnlineMeeting?: boolean;
  onlineMeeting?: {
    joinUrl: string;
    conferenceId?: string;
    tollNumber?: string;
  };
  categories?: string[];
  createdDateTime: string;
  lastModifiedDateTime: string;
}

// Simplified event for output
export interface SimplifiedEvent {
  id: string;
  subject: string;
  bodyPreview?: string;
  startDateTime: string;
  endDateTime: string;
  timeZone: string;
  location?: string;
  organizer?: string;
  attendees?: Array<{
    email: string;
    name?: string;
    response?: string;
  }>;
  isAllDay: boolean;
  webLink?: string;
  onlineMeetingUrl?: string;
  categories?: string[];
}

// Microsoft Graph User Profile
export interface OutlookProfile {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
}

// OAuth token structure
export interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  scope?: string;
  id_token?: string;
}

// Session data structure
export interface SessionData {
  userId: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt?: Date;
  createdAt: Date;
}

// Event search/filter options
export interface EventFilterOptions {
  startDateTime?: string;
  endDateTime?: string;
  subject?: string;
  location?: string;
  attendee?: string;
  category?: string;
  top?: number;
  skip?: number;
  orderBy?: string;
}

// Batch fetch request
export interface BatchEventRequest {
  eventIds: string[];
}







