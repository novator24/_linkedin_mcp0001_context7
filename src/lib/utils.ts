import { SearchResponse, SearchResult, VBALibrary, VBASearchResponse, OfficeApplication, VBACategory } from "./types.js";

/**
 * Formats a search result into a human-readable string representation.
 * Only shows code snippet count and GitHub stars when available (not equal to -1).
 *
 * @param result The SearchResult object to format
 * @returns A formatted string with library information
 */
export function formatSearchResult(result: SearchResult): string {
  // Always include these basic details
  const formattedResult = [
    `- Title: ${result.title}`,
    `- Context7-compatible library ID: ${result.id}`,
    `- Description: ${result.description}`,
  ];

  // Only add code snippets count if it's a valid value
  if (result.totalSnippets !== -1 && result.totalSnippets !== undefined) {
    formattedResult.push(`- Code Snippets: ${result.totalSnippets}`);
  }

  // Only add trust score if it's a valid value
  if (result.trustScore !== -1 && result.trustScore !== undefined) {
    formattedResult.push(`- Trust Score: ${result.trustScore}`);
  }

  // Only add versions if it's a valid value
  if (result.versions !== undefined && result.versions.length > 0) {
    formattedResult.push(`- Versions: ${result.versions.join(", ")}`);
  }

  // Join all parts with newlines
  return formattedResult.join("\n");
}

/**
 * Formats a search response into a human-readable string representation.
 * Each result is formatted using formatSearchResult.
 *
 * @param searchResponse The SearchResponse object to format
 * @returns A formatted string with search results
 */
export function formatSearchResults(searchResponse: SearchResponse): string {
  if (!searchResponse.results || searchResponse.results.length === 0) {
    return "No documentation libraries found matching your query.";
  }

  const formattedResults = searchResponse.results.map(formatSearchResult);
  return formattedResults.join("\n----------\n");
}

/**
 * Форматирование результатов поиска VBA библиотек
 */
export function formatVBAResults(
  results: VBALibrary[],
  options: {
    officeApp?: OfficeApplication;
    category?: VBACategory;
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

/**
 * Валидация VBA Library ID
 */
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

/**
 * Форматирование VBA результатов поиска
 */
export function formatVBASearchResults(searchResponse: VBASearchResponse, options: {
  officeApp?: OfficeApplication;
  category?: VBACategory;
  showExamples?: boolean;
  showTrustScore?: boolean;
  maxResults?: number;
} = {}): string {
  if (!searchResponse.results || searchResponse.results.length === 0) {
    return searchResponse.error || "No VBA libraries found matching your query.";
  }

  const formattedResults = formatVBAResults(searchResponse.results, options);
  
  let result = `Available VBA Libraries (${searchResponse.results.length} found):\n\n`;
  result += formattedResults;
  
  if (searchResponse.searchTime) {
    result += `\nSearch completed in ${searchResponse.searchTime}ms.`;
  }
  
  if (searchResponse.suggestions && searchResponse.suggestions.length > 0) {
    result += `\n\nSuggestions: ${searchResponse.suggestions.join(', ')}`;
  }
  
  return result;
}
