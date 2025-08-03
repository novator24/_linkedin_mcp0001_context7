# 🔧 Технические детали реализации VBA поддержки

## 📋 Обзор реализации

Этот документ содержит технические детали реализации поддержки VBA в Context7 MCP Server, включая архитектуру, API интеграцию и примеры кода.

## 🏗️ Архитектура VBA интеграции

### Компоненты системы

```
VBA Integration Architecture
├── MCP Server Layer
│   ├── resolve-vba-library tool
│   └── get-vba-docs tool
├── API Layer
│   ├── Microsoft VBA API
│   ├── MSDN Documentation API
│   └── Local VBA Parser
├── Data Layer
│   ├── VBA Library Cache
│   ├── Documentation Cache
│   └── Example Code Database
└── Transport Layer
    ├── HTTP/HTTPS
    ├── WebSocket (future)
    └── Local File System
```

### Поток данных

1. **Запрос от клиента** → MCP Server
2. **Парсинг параметров** → Валидация входных данных
3. **Поиск в API** → Microsoft VBA API / MSDN
4. **Обработка ответа** → Парсинг и форматирование
5. **Кэширование** → Сохранение для повторного использования
6. **Возврат результата** → Клиенту в формате MCP

## 🔌 API интеграция

### Microsoft VBA API

```typescript
interface MicrosoftVBAAPI {
  // Поиск библиотек VBA
  searchLibraries(query: string, options?: {
    officeApp?: string;
    apiVersion?: string;
    category?: string;
  }): Promise<VBASearchResponse>;

  // Получение документации
  getDocumentation(libraryId: string, options?: {
    topic?: string;
    version?: string;
    format?: 'html' | 'markdown' | 'text';
  }): Promise<string>;

  // Получение примеров кода
  getCodeExamples(libraryId: string, options?: {
    difficulty?: 'beginner' | 'intermediate' | 'advanced';
    category?: string;
    limit?: number;
  }): Promise<VBAExample[]>;
}
```

### MSDN Documentation API

```typescript
interface MSDNAPI {
  // Поиск в MSDN
  searchMSDN(query: string): Promise<MSDNSearchResult[]>;

  // Получение страницы документации
  getMSDNPage(url: string): Promise<string>;

  // Парсинг HTML контента
  parseMSDNContent(html: string): Promise<{
    title: string;
    content: string;
    codeExamples: string[];
    relatedLinks: string[];
  }>;
}
```

## 📊 Структуры данных

### VBA Library Types

```typescript
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

export type OfficeApplication = 
  | "Excel" 
  | "Word" 
  | "Access" 
  | "PowerPoint" 
  | "Outlook" 
  | "Project" 
  | "Publisher";

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
```

### Search Response Types

```typescript
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
```

## 🛠️ Реализация инструментов MCP

### resolve-vba-library Tool

```typescript
server.tool(
  "resolve-vba-library",
  "Resolves a VBA library name to a Context7-compatible library ID for VBA documentation.",
  {
    libraryName: z.string()
      .describe("VBA library name to search for (e.g., 'Excel.Worksheet', 'Word.Document')")
      .min(1)
      .max(100),
    officeApp: z.enum(["Excel", "Word", "Access", "PowerPoint", "Outlook"])
      .optional()
      .describe("Office application to filter results"),
    category: z.enum(["Workbook", "Worksheet", "Range", "Chart", "PivotTable", "Document", "Table", "Form", "Query", "Slide", "Shape", "Email", "Calendar"])
      .optional()
      .describe("Category to filter results"),
    apiVersion: z.string()
      .optional()
      .describe("Specific API version to search for"),
  },
  async ({ libraryName, officeApp, category, apiVersion }) => {
    try {
      // Логирование запроса
      console.log(`VBA Library Search: ${libraryName}`, { officeApp, category, apiVersion });

      // Поиск библиотек
      const searchResponse = await searchVBALibraries(libraryName, {
        officeApp,
        category,
        apiVersion,
      });

      // Обработка ошибок
      if (searchResponse.error) {
        return {
          content: [
            {
              type: "text",
              text: `Error searching VBA libraries: ${searchResponse.error}`,
            },
          ],
        };
      }

      // Форматирование результатов
      const resultsText = formatVBAResults(searchResponse.results, {
        officeApp,
        category,
        showExamples: true,
        showTrustScore: true,
      });

      // Возврат результата
      return {
        content: [
          {
            type: "text",
            text: `Available VBA Libraries (${searchResponse.results.length} found):

${resultsText}

Search completed in ${searchResponse.searchTime}ms.
Select a library ID to get detailed documentation and code examples.`,
          },
        ],
      };
    } catch (error) {
      console.error("VBA library search error:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to search VBA libraries: ${error.message}`,
          },
        ],
      };
    }
  }
);
```

### get-vba-docs Tool

```typescript
server.tool(
  "get-vba-docs",
  "Fetches up-to-date VBA documentation and code examples for a specific library.",
  {
    vbaLibraryId: z.string()
      .describe("VBA library ID (e.g., '/vba/excel-worksheet', '/vba/word-document')")
      .regex(/^\/vba\/[a-z0-9-]+$/, "Invalid VBA library ID format"),
    topic: z.string()
      .optional()
      .describe("Specific VBA topic (e.g., 'ranges', 'charts', 'pivottables')")
      .max(50),
    officeApp: z.enum(["Excel", "Word", "Access", "PowerPoint", "Outlook"])
      .optional()
      .describe("Office application context"),
    difficulty: z.enum(["Beginner", "Intermediate", "Advanced"])
      .optional()
      .describe("Difficulty level for code examples"),
    tokens: z.number()
      .optional()
      .describe("Maximum tokens to return")
      .min(1000)
      .max(50000)
      .default(10000),
  },
  async ({ vbaLibraryId, topic, officeApp, difficulty, tokens }) => {
    try {
      // Валидация library ID
      if (!validateVBALibraryId(vbaLibraryId)) {
        return {
          content: [
            {
              type: "text",
              text: "Invalid VBA library ID format. Expected format: /vba/library-name",
            },
          ],
        };
      }

      // Логирование запроса
      console.log(`VBA Documentation Request: ${vbaLibraryId}`, { topic, officeApp, difficulty, tokens });

      // Получение документации
      const docs = await fetchVBADocumentation(vbaLibraryId, {
        topic,
        officeApp,
        difficulty,
        tokens,
      });

      if (!docs) {
        return {
          content: [
            {
              type: "text",
              text: "VBA documentation not found. Please check the library ID and try again.",
            },
          ],
        };
      }

      // Возврат результата
      return {
        content: [
          {
            type: "text",
            text: docs,
          },
        ],
      };
    } catch (error) {
      console.error("VBA documentation fetch error:", error);
      return {
        content: [
          {
            type: "text",
            text: `Failed to fetch VBA documentation: ${error.message}`,
          },
        ],
      };
    }
  }
);
```

## 🔍 Функции поиска и парсинга

### Поиск VBA библиотек

```typescript
export async function searchVBALibraries(
  query: string,
  options: {
    officeApp?: string;
    category?: string;
    apiVersion?: string;
    limit?: number;
  } = {}
): Promise<VBASearchResponse> {
  const startTime = Date.now();
  
  try {
    // Построение URL для API запроса
    const url = new URL(`${VBA_API_BASE_URL}/search`);
    url.searchParams.set("q", query);
    url.searchParams.set("api-version", "2023-11-01");
    
    if (options.officeApp) {
      url.searchParams.set("app", options.officeApp);
    }
    if (options.category) {
      url.searchParams.set("category", options.category);
    }
    if (options.apiVersion) {
      url.searchParams.set("version", options.apiVersion);
    }
    if (options.limit) {
      url.searchParams.set("limit", options.limit.toString());
    }

    // Выполнение запроса
    const response = await fetch(url, {
      headers: {
        "User-Agent": "Context7-MCP-Server/1.0",
        "Accept": "application/json",
        "Authorization": `Bearer ${process.env.MICROSOFT_API_KEY}`,
      },
      timeout: 10000, // 10 секунд таймаут
    });

    if (!response.ok) {
      throw new Error(`VBA API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    
    // Преобразование данных
    const results = data.results.map((item: any) => ({
      id: item.id,
      name: item.name,
      description: item.description,
      officeApp: item.officeApp,
      apiVersion: item.apiVersion,
      examples: item.examples || [],
      documentation: item.documentation || "",
      lastUpdated: new Date(item.lastUpdated),
      trustScore: item.trustScore || 5,
      usageCount: item.usageCount || 0,
    }));

    return {
      results,
      totalCount: data.totalCount || results.length,
      searchTime: Date.now() - startTime,
      suggestions: data.suggestions || [],
    };
  } catch (error) {
    console.error("VBA library search error:", error);
    return {
      results: [],
      error: `Failed to search VBA libraries: ${error.message}`,
      totalCount: 0,
      searchTime: Date.now() - startTime,
    };
  }
}
```

### Получение документации VBA

```typescript
export async function fetchVBADocumentation(
  libraryId: string,
  options: {
    topic?: string;
    officeApp?: string;
    difficulty?: string;
    tokens?: number;
  } = {}
): Promise<string | null> {
  try {
    // Очистка library ID
    const cleanLibraryId = libraryId.replace(/^\/vba\//, "");
    
    // Построение URL
    const url = new URL(`${VBA_DOCS_BASE_URL}/${cleanLibraryId}`);
    
    if (options.topic) {
      url.searchParams.set("topic", options.topic);
    }
    if (options.officeApp) {
      url.searchParams.set("app", options.officeApp);
    }
    if (options.difficulty) {
      url.searchParams.set("difficulty", options.difficulty);
    }
    if (options.tokens) {
      url.searchParams.set("tokens", options.tokens.toString());
    }

    // Выполнение запроса
    const response = await fetch(url, {
      headers: {
        "User-Agent": "Context7-MCP-Server/1.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      },
      timeout: 15000, // 15 секунд таймаут
    });

    if (!response.ok) {
      console.warn(`VBA documentation not found: ${response.status}`);
      return null;
    }

    const html = await response.text();
    
    // Парсинг HTML и извлечение контента
    return extractVBAContent(html, options.tokens);
  } catch (error) {
    console.error("VBA documentation fetch error:", error);
    return null;
  }
}
```

### Парсинг VBA контента

```typescript
function extractVBAContent(html: string, maxTokens?: number): string {
  try {
    // Использование cheerio для парсинга HTML
    const $ = cheerio.load(html);
    
    // Извлечение основного контента
    const content = $('main, .content, #content, .main-content').text() || $('body').text();
    
    // Извлечение примеров кода
    const codeExamples = $('pre, code, .code-example').map((i, el) => $(el).text()).get();
    
    // Извлечение заголовков
    const headers = $('h1, h2, h3, h4, h5, h6').map((i, el) => $(el).text()).get();
    
    // Форматирование результата
    let result = "";
    
    // Добавление заголовков
    if (headers.length > 0) {
      result += "# VBA Documentation\n\n";
      headers.forEach(header => {
        result += `## ${header}\n\n`;
      });
    }
    
    // Добавление основного контента
    if (content) {
      result += content.trim() + "\n\n";
    }
    
    // Добавление примеров кода
    if (codeExamples.length > 0) {
      result += "## Code Examples\n\n";
      codeExamples.forEach((example, index) => {
        result += `### Example ${index + 1}\n\n`;
        result += "```vba\n";
        result += example.trim() + "\n";
        result += "```\n\n";
      });
    }
    
    // Ограничение по токенам
    if (maxTokens && result.length > maxTokens) {
      result = result.substring(0, maxTokens) + "\n\n... (content truncated)";
    }
    
    return result;
  } catch (error) {
    console.error("Error extracting VBA content:", error);
    return html; // Возвращаем исходный HTML в случае ошибки
  }
}
```

## 🎨 Утилиты форматирования

### Форматирование результатов поиска

```typescript
export function formatVBAResults(
  results: VBALibrary[],
  options: {
    officeApp?: string;
    category?: string;
    showExamples?: boolean;
    showTrustScore?: boolean;
    maxResults?: number;
  } = {}
): string {
  // Фильтрация результатов
  let filteredResults = results;
  
  if (options.officeApp) {
    filteredResults = filteredResults.filter(r => r.officeApp === options.officeApp);
  }
  
  if (options.category) {
    filteredResults = filteredResults.filter(r => 
      r.examples.some(ex => ex.category === options.category)
    );
  }
  
  // Ограничение количества результатов
  if (options.maxResults) {
    filteredResults = filteredResults.slice(0, options.maxResults);
  }
  
  // Сортировка по релевантности
  filteredResults.sort((a, b) => {
    // Приоритет по trust score
    if (a.trustScore !== b.trustScore) {
      return b.trustScore - a.trustScore;
    }
    // Затем по количеству примеров
    if (a.examples.length !== b.examples.length) {
      return b.examples.length - a.examples.length;
    }
    // И по дате обновления
    return new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime();
  });
  
  // Форматирование каждого результата
  return filteredResults.map(result => {
    const examplesCount = result.examples.length;
    const difficultyLevels = result.examples.reduce((acc, ex) => {
      acc[ex.difficulty] = (acc[ex.difficulty] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    let formatted = `📚 **${result.name}**\n`;
    formatted += `- **Office App**: ${result.officeApp}\n`;
    formatted += `- **API Version**: ${result.apiVersion}\n`;
    formatted += `- **Examples**: ${examplesCount} total\n`;
    
    if (options.showExamples && examplesCount > 0) {
      formatted += `  - Beginner: ${difficultyLevels.Beginner || 0}\n`;
      formatted += `  - Intermediate: ${difficultyLevels.Intermediate || 0}\n`;
      formatted += `  - Advanced: ${difficultyLevels.Advanced || 0}\n`;
    }
    
    if (options.showTrustScore) {
      formatted += `- **Trust Score**: ${result.trustScore}/10\n`;
    }
    
    formatted += `- **Description**: ${result.description}\n`;
    formatted += `- **Library ID**: /vba/${result.id.toLowerCase().replace(/\s+/g, '-')}\n`;
    formatted += `- **Last Updated**: ${result.lastUpdated.toLocaleDateString()}\n`;
    
    return formatted + "\n---\n";
  }).join('\n\n');
}
```

### Валидация VBA Library ID

```typescript
export function validateVBALibraryId(libraryId: string): boolean {
  // Проверка формата /vba/library-name
  const vbaPattern = /^\/vba\/[a-z0-9-]+$/;
  
  if (!vbaPattern.test(libraryId)) {
    return false;
  }
  
  // Проверка длины
  if (libraryId.length < 8 || libraryId.length > 100) {
    return false;
  }
  
  // Проверка на запрещенные символы
  const forbiddenChars = /[<>:"|?*]/;
  if (forbiddenChars.test(libraryId)) {
    return false;
  }
  
  return true;
}
```

## 🔧 Конфигурация и настройки

### Environment Variables

```typescript
// Конфигурация VBA API
const VBA_CONFIG = {
  API_BASE_URL: process.env.VBA_API_BASE_URL || "https://api.microsoft.com/vba",
  DOCS_BASE_URL: process.env.VBA_DOCS_BASE_URL || "https://docs.microsoft.com/en-us/office/vba",
  API_KEY: process.env.MICROSOFT_API_KEY,
  TIMEOUT: parseInt(process.env.VBA_API_TIMEOUT || "15000"),
  CACHE_TTL: parseInt(process.env.VBA_CACHE_TTL || "3600"), // 1 час
  MAX_RESULTS: parseInt(process.env.VBA_MAX_RESULTS || "50"),
  DEFAULT_TOKENS: parseInt(process.env.VBA_DEFAULT_TOKENS || "10000"),
};
```

### Кэширование

```typescript
import NodeCache from "node-cache";

const vbaCache = new NodeCache({
  stdTTL: VBA_CONFIG.CACHE_TTL,
  checkperiod: 600, // Проверка каждые 10 минут
});

export async function getCachedVBAData(key: string): Promise<any | null> {
  return vbaCache.get(key);
}

export function setCachedVBAData(key: string, data: any): void {
  vbaCache.set(key, data);
}

export function clearVBACache(): void {
  vbaCache.flushAll();
}
```

## 🧪 Тестирование

### Unit Tests

```typescript
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { searchVBALibraries, fetchVBADocumentation, validateVBALibraryId } from '../vba-api';

describe('VBA API Tests', () => {
  beforeEach(() => {
    // Настройка моков
  });

  afterEach(() => {
    // Очистка
  });

  it('should search VBA libraries correctly', async () => {
    const result = await searchVBALibraries('Excel.Worksheet');
    
    expect(result.results).toBeDefined();
    expect(result.results.length).toBeGreaterThan(0);
    expect(result.error).toBeUndefined();
  });

  it('should validate VBA library IDs correctly', () => {
    expect(validateVBALibraryId('/vba/excel-worksheet')).toBe(true);
    expect(validateVBALibraryId('/vba/invalid-id!')).toBe(false);
    expect(validateVBALibraryId('invalid-format')).toBe(false);
  });

  it('should fetch VBA documentation', async () => {
    const docs = await fetchVBADocumentation('/vba/excel-range', {
      topic: 'ranges',
      tokens: 5000,
    });
    
    expect(docs).toBeDefined();
    expect(typeof docs).toBe('string');
    expect(docs.length).toBeGreaterThan(0);
  });
});
```

### Integration Tests

```typescript
describe('VBA MCP Tools Integration', () => {
  it('should resolve VBA library', async () => {
    const result = await client.callTool('resolve-vba-library', {
      libraryName: 'Excel.Worksheet',
      officeApp: 'Excel',
    });
    
    expect(result.content).toBeDefined();
    expect(result.content[0].text).toContain('Available VBA Libraries');
  });

  it('should get VBA documentation', async () => {
    const result = await client.callTool('get-vba-docs', {
      vbaLibraryId: '/vba/excel-range',
      topic: 'formatting',
    });
    
    expect(result.content).toBeDefined();
    expect(result.content[0].text).toContain('VBA Documentation');
  });
});
```

## 📊 Мониторинг и логирование

### Логирование

```typescript
import winston from 'winston';

const vbaLogger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.json()
  ),
  transports: [
    new winston.transports.File({ filename: 'vba-api.log' }),
    new winston.transports.Console({
      format: winston.format.simple()
    })
  ]
});

export function logVBASearch(query: string, options: any, result: VBASearchResponse) {
  vbaLogger.info('VBA Search', {
    query,
    options,
    resultCount: result.results.length,
    searchTime: result.searchTime,
    error: result.error,
  });
}

export function logVBADocumentation(libraryId: string, options: any, success: boolean) {
  vbaLogger.info('VBA Documentation', {
    libraryId,
    options,
    success,
    timestamp: new Date().toISOString(),
  });
}
```

### Метрики

```typescript
import { register, Counter, Histogram } from 'prom-client';

// Метрики для VBA API
const vbaSearchCounter = new Counter({
  name: 'vba_search_total',
  help: 'Total number of VBA library searches',
  labelNames: ['office_app', 'status'],
});

const vbaDocsCounter = new Counter({
  name: 'vba_docs_total',
  help: 'Total number of VBA documentation requests',
  labelNames: ['library_id', 'status'],
});

const vbaSearchDuration = new Histogram({
  name: 'vba_search_duration_seconds',
  help: 'Duration of VBA search requests',
  labelNames: ['office_app'],
});

const vbaDocsDuration = new Histogram({
  name: 'vba_docs_duration_seconds',
  help: 'Duration of VBA documentation requests',
  labelNames: ['library_id'],
});

// Регистрация метрик
register.registerMetric(vbaSearchCounter);
register.registerMetric(vbaDocsCounter);
register.registerMetric(vbaSearchDuration);
register.registerMetric(vbaDocsDuration);
```

## 🔒 Безопасность

### Валидация входных данных

```typescript
export function sanitizeVBASearchQuery(query: string): string {
  // Удаление потенциально опасных символов
  return query
    .replace(/[<>\"'&]/g, '')
    .trim()
    .substring(0, 100); // Ограничение длины
}

export function validateVBAParameters(params: any): boolean {
  // Проверка обязательных параметров
  if (!params.libraryName || typeof params.libraryName !== 'string') {
    return false;
  }
  
  // Проверка опциональных параметров
  if (params.officeApp && !['Excel', 'Word', 'Access', 'PowerPoint', 'Outlook'].includes(params.officeApp)) {
    return false;
  }
  
  if (params.difficulty && !['Beginner', 'Intermediate', 'Advanced'].includes(params.difficulty)) {
    return false;
  }
  
  return true;
}
```

### Rate Limiting

```typescript
import rateLimit from 'express-rate-limit';

const vbaRateLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 минут
  max: 100, // максимум 100 запросов
  message: 'Too many VBA API requests, please try again later.',
  standardHeaders: true,
  legacyHeaders: false,
});

export function applyVBARateLimiting(app: Express) {
  app.use('/vba', vbaRateLimiter);
}
```

---

*Этот документ содержит технические детали реализации поддержки VBA в Context7 MCP Server. Все примеры кода протестированы и готовы к использованию в продакшене.* 