export type TaskType = 'translation' | 'analysis' | 'summarization' | 'custom';

export interface TokenUsage {
  input: number;
  output: number;
}

export interface ApiResponse {
  success: true;
  task: TaskType;
  sourceText: string;
  result: string;
  model: string;
  tokens?: TokenUsage;
  processingTime: number; // in seconds
  timestamp: string;
  metadata?: {
    confidence?: number;
    language?: string;
    [key: string]: unknown;
  };
}

export interface ApiErrorResponse {
  success: false;
  error: 'SERVICE_UNAVAILABLE' | 'TIMEOUT' | 'INVALID_INPUT' | 'AUTH_ERROR' | 'UNKNOWN_ERROR';
  message: string; // Vietnamese error description
  suggestion?: string; // Actionable suggestion in Vietnamese
  requestId: string;
}

export interface AppConfig {
  ollamaUrl: string;
  customApiUrl: string;
  customApiKey: string; // encrypted
  activeService: 'ollama' | 'custom';
  fallbackEnabled: boolean;
  streamingEnabled: boolean;
  models: {
    translation: string;
    analysis: string;
    summarization: string;
    custom: string;
  };
}

export interface HistoryItem {
  id: string;
  task: TaskType;
  timestamp: string;
  sourceText: string;
  result: string;
  model: string;
  processingTime: number;
  serviceUsed: 'ollama' | 'custom';
  status: 'success' | 'error';
  errorMessage?: string;
}

export interface GlossaryItem {
  japanese: string;
  vietnamese: string;
  note?: string;
}
