export interface SearchResult {
  id: string;
  title: string;
  description: string;
  branch: string;
  lastUpdateDate: string;
  state: DocumentState;
  totalTokens: number;
  totalSnippets: number;
  totalPages: number;
  stars?: number;
  trustScore?: number;
  versions?: string[];
}

export interface SearchResponse {
  error?: string;
  results: SearchResult[];
}

// Version state is still needed for validating search results
export type DocumentState = "initial" | "finalized" | "error" | "delete";

// VBA Support Types
export type OfficeApplication = 
  | "Excel" 
  | "Word" 
  | "Access" 
  | "PowerPoint" 
  | "Outlook" 
  | "Project" 
  | "Publisher";

export type VBACategory = 
  | "Workbook" 
  | "Worksheet" 
  | "Range" 
  | "Chart" 
  | "PivotTable" 
  | "Document" 
  | "Table" 
  | "Form" 
  | "Query" 
  | "Slide" 
  | "Shape" 
  | "Email" 
  | "Calendar";

export type VBADifficulty = 
  | "Beginner" 
  | "Intermediate" 
  | "Advanced";

export interface VBAExample {
  id: string;                   // Уникальный ID примера
  title: string;               // Заголовок
  description: string;         // Описание
  code: string;               // Код VBA
  category: VBACategory;      // Категория
  difficulty: VBADifficulty;  // Сложность
  tags: string[];             // Теги
  author?: string;            // Автор
  createdDate: Date;          // Дата создания
  rating: number;             // Рейтинг (1-5)
}

export interface VBALibrary {
  id: string;                    // Уникальный идентификатор
  name: string;                  // Название библиотеки
  description: string;           // Описание
  officeApp: OfficeApplication;  // Приложение Office
  apiVersion: string;           // Версия API
  examples: VBAExample[];       // Примеры кода
  documentation: string;        // Документация
  lastUpdated: Date;           // Дата последнего обновления
  trustScore: number;          // Оценка доверия (0-10)
  usageCount: number;          // Количество использований
}

export interface VBASearchResponse {
  error?: string;              // Ошибка, если есть
  results: VBALibrary[];       // Результаты поиска
  totalCount: number;          // Общее количество
  searchTime: number;          // Время поиска в мс
  suggestions?: string[];      // Предложения для поиска
}

export interface VBADocumentationResponse {
  libraryId: string;           // ID библиотеки
  content: string;             // Содержимое документации
  codeExamples: VBAExample[];  // Примеры кода
  relatedLibraries: string[];  // Связанные библиотеки
  lastUpdated: Date;          // Дата обновления
  source: string;             // Источник документации
}
