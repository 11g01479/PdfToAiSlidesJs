
export interface Slide {
  title: string;
  notes: string;
  pageIndex: number;
  imageUrl?: string;
}

export interface AnalysisResult {
  presentationTitle: string;
  summary: string;
  slides: Slide[];
}

export interface AppState {
  status: 'idle' | 'rendering' | 'analyzing' | 'reviewing' | 'completed' | 'error';
  progress: number;
  error?: string;
}

export enum ModelName {
  TEXT = 'gemini-3-flash-preview'
}
