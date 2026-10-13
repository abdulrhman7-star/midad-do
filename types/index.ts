export interface GuestSession {
  id: string;
  isGuest: boolean;
  email?: string;
  name?: string;
  credits: number;
  maxCredits: number;
  lastRefreshed: number;
  createdAt: number;
}

export interface SavedDocument {
  id: string;
  title: string;
  markdown: string;
  html: string;
  updatedAt: number;
  isPublic?: boolean;
}

export type SupportedLanguage = 'en' | 'ar';
