import { VBALibrary, VBAExample, VBASearchResponse, VBADocumentationResponse, OfficeApplication, VBACategory, VBADifficulty } from "./types.js";

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

/**
 * Поиск VBA библиотек
 */
export async function searchVBALibraries(
  query: string,
  options: {
    officeApp?: OfficeApplication;
    category?: VBACategory;
    apiVersion?: string;
    limit?: number;
  } = {}
): Promise<VBASearchResponse> {
  const startTime = Date.now();
  
  try {
    // Санитизация запроса
    const sanitizedQuery = sanitizeVBASearchQuery(query);
    
    // Построение URL для API запроса
    const url = new URL(`${VBA_CONFIG.API_BASE_URL}/search`);
    url.searchParams.set("q", sanitizedQuery);
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
        "Authorization": VBA_CONFIG.API_KEY ? `Bearer ${VBA_CONFIG.API_KEY}` : "",
      },
      signal: AbortSignal.timeout(VBA_CONFIG.TIMEOUT),
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
      error: `Failed to search VBA libraries: ${error instanceof Error ? error.message : String(error)}`,
      totalCount: 0,
      searchTime: Date.now() - startTime,
    };
  }
}

/**
 * Получение документации VBA
 */
export async function fetchVBADocumentation(
  libraryId: string,
  options: {
    topic?: string;
    officeApp?: OfficeApplication;
    difficulty?: VBADifficulty;
    tokens?: number;
  } = {}
): Promise<string | null> {
  try {
    // Очистка library ID
    const cleanLibraryId = libraryId.replace(/^\/vba\//, "");
    
    // Построение URL
    const url = new URL(`${VBA_CONFIG.DOCS_BASE_URL}/${cleanLibraryId}`);
    
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
      signal: AbortSignal.timeout(VBA_CONFIG.TIMEOUT),
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

/**
 * Получение примеров кода VBA
 */
export async function fetchVBACodeExamples(
  libraryId: string,
  options: {
    difficulty?: VBADifficulty;
    category?: VBACategory;
    limit?: number;
  } = {}
): Promise<VBAExample[]> {
  try {
    const cleanLibraryId = libraryId.replace(/^\/vba\//, "");
    const url = new URL(`${VBA_CONFIG.API_BASE_URL}/examples/${cleanLibraryId}`);
    
    if (options.difficulty) {
      url.searchParams.set("difficulty", options.difficulty);
    }
    if (options.category) {
      url.searchParams.set("category", options.category);
    }
    if (options.limit) {
      url.searchParams.set("limit", options.limit.toString());
    }

    const response = await fetch(url, {
      headers: {
        "User-Agent": "Context7-MCP-Server/1.0",
        "Accept": "application/json",
        "Authorization": VBA_CONFIG.API_KEY ? `Bearer ${VBA_CONFIG.API_KEY}` : "",
      },
      signal: AbortSignal.timeout(VBA_CONFIG.TIMEOUT),
    });

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    return data.examples || [];
  } catch (error) {
    console.error("VBA code examples fetch error:", error);
    return [];
  }
}

/**
 * Санитизация поискового запроса
 */
function sanitizeVBASearchQuery(query: string): string {
  return query
    .replace(/[<>\"'&]/g, '')
    .trim()
    .substring(0, 100); // Ограничение длины
}

/**
 * Парсинг HTML и извлечение VBA контента
 */
function extractVBAContent(html: string, maxTokens?: number): string {
  try {
    // Простой парсинг HTML без внешних зависимостей
    const content = extractTextFromHTML(html);
    const codeExamples = extractCodeFromHTML(html);
    const headers = extractHeadersFromHTML(html);
    
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

/**
 * Извлечение текста из HTML
 */
function extractTextFromHTML(html: string): string {
  // Удаление HTML тегов
  return html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Извлечение кода из HTML
 */
function extractCodeFromHTML(html: string): string[] {
  const codeBlocks: string[] = [];
  
  // Поиск блоков кода
  const codeRegex = /<pre[^>]*>([\s\S]*?)<\/pre>/gi;
  const codeMatch = html.match(codeRegex);
  
  if (codeMatch) {
    codeMatch.forEach(match => {
      const code = match.replace(/<pre[^>]*>/, '').replace(/<\/pre>/, '');
      codeBlocks.push(code);
    });
  }
  
  // Поиск inline кода
  const inlineCodeRegex = /<code[^>]*>([\s\S]*?)<\/code>/gi;
  const inlineMatch = html.match(inlineCodeRegex);
  
  if (inlineMatch) {
    inlineMatch.forEach(match => {
      const code = match.replace(/<code[^>]*>/, '').replace(/<\/code>/, '');
      codeBlocks.push(code);
    });
  }
  
  return codeBlocks;
}

/**
 * Извлечение заголовков из HTML
 */
function extractHeadersFromHTML(html: string): string[] {
  const headers: string[] = [];
  
  // Поиск заголовков h1-h6
  const headerRegex = /<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/gi;
  const headerMatch = html.match(headerRegex);
  
  if (headerMatch) {
    headerMatch.forEach(match => {
      const header = match.replace(/<h[1-6][^>]*>/, '').replace(/<\/h[1-6]>/, '');
      headers.push(header);
    });
  }
  
  return headers;
}

/**
 * Валидация параметров VBA
 */
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