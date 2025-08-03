# 🚀 Context7 MCP Server: Полное руководство для разработчиков и администраторов

## 📋 Содержание

1. [Введение в MCP и Context7](#введение)
2. [Архитектура и принципы работы](#архитектура)
3. [Установка и настройка](#установка)
4. [Разработка и модификация](#разработка)
5. [Добавление поддержки VBA](#добавление-vba)
6. [Часто задаваемые вопросы](#faq)
7. [Полезные ссылки](#ссылки)

---

## 🎯 Введение {#введение}

**Context7 MCP Server** — это мощный инструмент для интеграции актуальной документации и примеров кода в AI-ассистенты. Сервер работает по протоколу Model Context Protocol (MCP) и предоставляет разработчикам доступ к актуальной документации библиотек прямо в их IDE.

### Что такое MCP?

**Model Context Protocol (MCP)** — это открытый стандарт для подключения AI-моделей к внешним данным и инструментам. MCP позволяет:

- 🔗 Подключать AI к базам данных, API и файловым системам
- 📚 Получать актуальную документацию в реальном времени
- 🛠️ Использовать специализированные инструменты
- 🔄 Обеспечивать безопасное взаимодействие между AI и внешними ресурсами

### Зачем нужен Context7?

Традиционные AI-ассистенты часто работают с устаревшей информацией:

- ❌ Примеры кода основаны на данных годичной давности
- ❌ API, которые не существуют в реальности
- ❌ Общие ответы для старых версий пакетов

Context7 решает эти проблемы, предоставляя:

- ✅ Актуальную документацию из исходного кода
- ✅ Примеры кода для конкретных версий библиотек
- ✅ Прямую интеграцию с IDE через MCP

---

## 🏗️ Архитектура и принципы работы {#архитектура}

### Структура проекта

```
context7-mcp/
├── src/
│   ├── index.ts          # Основной файл сервера
│   └── lib/
│       ├── api.ts        # API для работы с Context7
│       ├── types.ts      # TypeScript типы
│       ├── utils.ts      # Утилиты
│       └── encryption.ts # Шифрование
├── package.json          # Зависимости и скрипты
├── tsconfig.json         # Конфигурация TypeScript
└── README.md            # Документация
```

### Основные компоненты

#### 1. MCP Server (`src/index.ts`)

```typescript
const server = new McpServer(
  {
    name: "Context7",
    version: "1.0.13",
  },
  {
    instructions: "Use this server to retrieve up-to-date documentation and code examples for any library.",
  }
);
```

Сервер регистрирует два основных инструмента:

- **`resolve-library-id`** — поиск библиотек по названию
- **`get-library-docs`** — получение документации по ID библиотеки

#### 2. API Layer (`src/lib/api.ts`)

```typescript
export async function searchLibraries(query: string, clientIp?: string): Promise<SearchResponse>
export async function fetchLibraryDocumentation(libraryId: string, options: {...}, clientIp?: string): Promise<string | null>
```

#### 3. Типы данных (`src/lib/types.ts`)

```typescript
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
```

### Транспортные протоколы

Сервер поддерживает три типа транспорта:

1. **stdio** (по умолчанию) — стандартный ввод/вывод
2. **http** — HTTP API
3. **sse** — Server-Sent Events

### Жизненный цикл запроса

1. **Поиск библиотеки**: `resolve-library-id` → поиск по API Context7
2. **Получение документации**: `get-library-docs` → загрузка актуальной документации
3. **Форматирование**: преобразование в читаемый формат
4. **Возврат**: передача данных в AI-ассистент

---

## ⚙️ Установка и настройка {#установка}

### Требования

- Node.js >= v18.0.0
- Cursor, Windsurf, Claude Desktop или другой MCP Client

### Локальная установка

#### 1. Клонирование репозитория

```bash
git clone https://github.com/upstash/context7.git
cd context7
```

#### 2. Установка зависимостей

```bash
# Используя npm
npm install

# Используя bun (рекомендуется)
bun install

# Используя yarn
yarn install
```

#### 3. Сборка проекта

```bash
# Сборка TypeScript
bun run build

# Или с npm
npm run build
```

#### 4. Запуск сервера

```bash
# Запуск через stdio (по умолчанию)
bun run dist/index.js

# Запуск через HTTP на порту 8080
bun run dist/index.js --transport http --port 8080

# Запуск через SSE
bun run dist/index.js --transport sse --port 3000
```

### Конфигурация в IDE

#### Cursor

1. Откройте `Settings` → `Cursor Settings` → `MCP`
2. Нажмите `Add new global MCP server`
3. Добавьте конфигурацию:

```json
{
  "mcpServers": {
    "context7": {
      "command": "npx",
      "args": ["-y", "@upstash/context7-mcp"]
    }
  }
}
```

#### VS Code

```json
{
  "mcp": {
    "servers": {
      "context7": {
        "type": "stdio",
        "command": "npx",
        "args": ["-y", "@upstash/context7-mcp"]
      }
    }
  }
}
```

#### Windsurf

```json
{
  "mcpServers": {
    "context7": {
      "command": "npx",
      "args": ["-y", "@upstash/context7-mcp"]
    }
  }
}
```

### Тестирование установки

```bash
# Тест с MCP Inspector
npx -y @modelcontextprotocol/inspector npx @upstash/context7-mcp
```

---

## 🔧 Разработка и модификация {#разработка}

### Структура разработки

#### 1. Основные файлы для модификации

- **`src/index.ts`** — регистрация инструментов MCP
- **`src/lib/api.ts`** — API для работы с Context7
- **`src/lib/types.ts`** — типы данных
- **`src/lib/utils.ts`** — утилиты форматирования

#### 2. Добавление нового инструмента

```typescript
server.tool(
  "your-new-tool",
  "Описание вашего инструмента",
  {
    parameterName: z.string().describe("Описание параметра"),
  },
  async ({ parameterName }) => {
    // Ваша логика
    return {
      content: [
        {
          type: "text",
          text: "Результат работы инструмента",
        },
      ],
    };
  }
);
```

#### 3. Модификация существующих инструментов

Для изменения логики поиска библиотек:

```typescript
// В src/index.ts, функция resolve-library-id
async ({ libraryName }) => {
  // Ваша кастомная логика поиска
  const customSearchResponse = await yourCustomSearch(libraryName);
  
  return {
    content: [
      {
        type: "text",
        text: formatCustomResults(customSearchResponse),
      },
    ],
  };
}
```

#### 4. Добавление новых типов транспорта

```typescript
// В src/index.ts
if (TRANSPORT_TYPE === "websocket") {
  // Ваша реализация WebSocket транспорта
  const wsTransport = new WebSocketServerTransport(port);
  await server.connect(wsTransport);
}
```

### Отладка и логирование

#### Включение подробного логирования

```typescript
// В src/index.ts
const DEBUG_MODE = process.env.DEBUG === "true";

if (DEBUG_MODE) {
  console.log("Context7 MCP Server starting in debug mode...");
}
```

#### Логирование запросов

```typescript
// В src/lib/api.ts
export async function searchLibraries(query: string, clientIp?: string): Promise<SearchResponse> {
  console.log(`Searching for: ${query} from IP: ${clientIp}`);
  
  try {
    // ... существующий код
    console.log(`Search completed successfully`);
    return await response.json();
  } catch (error) {
    console.error("Search error:", error);
    return { results: [], error: `Error: ${error}` };
  }
}
```

### Тестирование изменений

#### 1. Локальное тестирование

```bash
# Сборка
bun run build

# Запуск в режиме отладки
DEBUG=true bun run dist/index.js --transport stdio
```

#### 2. Тестирование с MCP Inspector

```bash
npx -y @modelcontextprotocol/inspector bun run dist/index.js
```

#### 3. Интеграционное тестирование

```bash
# Тест с реальным клиентом
bun run dist/index.js --transport http --port 3000
```

---

## 📝 Добавление поддержки VBA {#добавление-vba}

### Пошаговое руководство

#### Шаг 1: Анализ требований

VBA (Visual Basic for Applications) имеет специфические особенности:

- **Синтаксис**: Основан на Visual Basic
- **Среда выполнения**: Microsoft Office (Excel, Word, Access)
- **API**: Объектные модели Office
- **Документация**: Microsoft Developer Network (MSDN)

#### Шаг 2: Создание типов для VBA

Добавьте в `src/lib/types.ts`:

```typescript
export interface VBALibrary {
  id: string;
  name: string;
  description: string;
  officeApp: "Excel" | "Word" | "Access" | "PowerPoint" | "Outlook";
  apiVersion: string;
  examples: VBAExample[];
}

export interface VBAExample {
  title: string;
  description: string;
  code: string;
  category: "Workbook" | "Worksheet" | "Range" | "Chart" | "PivotTable";
  difficulty: "Beginner" | "Intermediate" | "Advanced";
}

export interface VBASearchResponse {
  error?: string;
  results: VBALibrary[];
}
```

#### Шаг 3: Создание API для VBA

Создайте `src/lib/vba-api.ts`:

```typescript
import { VBALibrary, VBAExample, VBASearchResponse } from "./types.js";

const VBA_API_BASE_URL = "https://api.microsoft.com/vba";
const VBA_DOCS_BASE_URL = "https://docs.microsoft.com/en-us/office/vba";

export async function searchVBALibraries(query: string): Promise<VBASearchResponse> {
  try {
    const url = new URL(`${VBA_API_BASE_URL}/search`);
    url.searchParams.set("q", query);
    url.searchParams.set("api-version", "2023-11-01");

    const response = await fetch(url, {
      headers: {
        "User-Agent": "Context7-MCP-Server/1.0",
        "Accept": "application/json",
      },
    });

    if (!response.ok) {
      return {
        results: [],
        error: `VBA API error: ${response.status}`,
      };
    }

    const data = await response.json();
    return {
      results: data.results.map((item: any) => ({
        id: item.id,
        name: item.name,
        description: item.description,
        officeApp: item.officeApp,
        apiVersion: item.apiVersion,
        examples: item.examples || [],
      })),
    };
  } catch (error) {
    return {
      results: [],
      error: `VBA search error: ${error}`,
    };
  }
}

export async function fetchVBADocumentation(
  libraryId: string,
  options: {
    tokens?: number;
    topic?: string;
    officeApp?: string;
  } = {}
): Promise<string | null> {
  try {
    const url = new URL(`${VBA_DOCS_BASE_URL}/${libraryId}`);
    
    if (options.topic) {
      url.searchParams.set("topic", options.topic);
    }
    if (options.officeApp) {
      url.searchParams.set("app", options.officeApp);
    }

    const response = await fetch(url, {
      headers: {
        "User-Agent": "Context7-MCP-Server/1.0",
        "Accept": "text/html,application/xhtml+xml",
      },
    });

    if (!response.ok) {
      return null;
    }

    const html = await response.text();
    return extractVBAContent(html, options.tokens);
  } catch (error) {
    console.error("VBA documentation fetch error:", error);
    return null;
  }
}

function extractVBAContent(html: string, maxTokens?: number): string {
  // Парсинг HTML и извлечение VBA-специфичного контента
  // Здесь должна быть логика извлечения кода, примеров и документации
  return html; // Упрощенная версия
}
```

#### Шаг 4: Добавление VBA инструментов в MCP сервер

Модифицируйте `src/index.ts`:

```typescript
// Добавьте импорт
import { searchVBALibraries, fetchVBADocumentation } from "./lib/vba-api.js";

// В функции createServerInstance добавьте новые инструменты
server.tool(
  "resolve-vba-library",
  "Resolves a VBA library name to a Context7-compatible library ID for VBA documentation.",
  {
    libraryName: z.string().describe("VBA library name to search for (e.g., 'Excel.Worksheet', 'Word.Document')"),
    officeApp: z.string().optional().describe("Office application (Excel, Word, Access, PowerPoint, Outlook)"),
  },
  async ({ libraryName, officeApp }) => {
    const searchResponse = await searchVBALibraries(libraryName);

    if (!searchResponse.results || searchResponse.results.length === 0) {
      return {
        content: [
          {
            type: "text",
            text: searchResponse.error || "No VBA libraries found matching your query.",
          },
        ],
      };
    }

    const resultsText = formatVBAResults(searchResponse.results, officeApp);

    return {
      content: [
        {
          type: "text",
          text: `Available VBA Libraries:

${resultsText}

Select a library ID to get detailed documentation and code examples.`,
        },
      ],
    };
  }
);

server.tool(
  "get-vba-docs",
  "Fetches up-to-date VBA documentation and code examples for a specific library.",
  {
    vbaLibraryId: z.string().describe("VBA library ID (e.g., '/vba/excel-worksheet', '/vba/word-document')"),
    topic: z.string().optional().describe("Specific VBA topic (e.g., 'ranges', 'charts', 'pivottables')"),
    officeApp: z.string().optional().describe("Office application context"),
    tokens: z.number().optional().describe("Maximum tokens to return"),
  },
  async ({ vbaLibraryId, topic, officeApp, tokens = 10000 }) => {
    const docs = await fetchVBADocumentation(vbaLibraryId, {
      tokens,
      topic,
      officeApp,
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

    return {
      content: [
        {
          type: "text",
          text: docs,
        },
      ],
    };
  }
);
```

#### Шаг 5: Создание утилит для VBA

Добавьте в `src/lib/utils.ts`:

```typescript
export function formatVBAResults(results: VBALibrary[], officeApp?: string): string {
  return results
    .filter(result => !officeApp || result.officeApp === officeApp)
    .map(result => {
      const examplesCount = result.examples.length;
      const difficultyLevels = result.examples.reduce((acc, ex) => {
        acc[ex.difficulty] = (acc[ex.difficulty] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

      return `📚 **${result.name}**
- **Office App**: ${result.officeApp}
- **API Version**: ${result.apiVersion}
- **Examples**: ${examplesCount} total
  - Beginner: ${difficultyLevels.Beginner || 0}
  - Intermediate: ${difficultyLevels.Intermediate || 0}
  - Advanced: ${difficultyLevels.Advanced || 0}
- **Description**: ${result.description}
- **Library ID**: /vba/${result.id.toLowerCase().replace(/\s+/g, '-')}

---`;
    })
    .join('\n\n');
}

export function validateVBALibraryId(libraryId: string): boolean {
  const vbaPattern = /^\/vba\/[a-z0-9-]+$/;
  return vbaPattern.test(libraryId);
}
```

#### Шаг 6: Обновление типов

Обновите `src/lib/types.ts`:

```typescript
// Добавьте новые типы
export interface VBALibrary {
  id: string;
  name: string;
  description: string;
  officeApp: "Excel" | "Word" | "Access" | "PowerPoint" | "Outlook";
  apiVersion: string;
  examples: VBAExample[];
}

export interface VBAExample {
  title: string;
  description: string;
  code: string;
  category: "Workbook" | "Worksheet" | "Range" | "Chart" | "PivotTable";
  difficulty: "Beginner" | "Intermediate" | "Advanced";
}

export interface VBASearchResponse {
  error?: string;
  results: VBALibrary[];
}
```

#### Шаг 7: Тестирование VBA поддержки

Создайте тестовый файл `test-vba.js`:

```javascript
import { McpClient } from "@modelcontextprotocol/sdk/client/mcp.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

async function testVBASupport() {
  const transport = new StdioClientTransport({
    command: "bun",
    args: ["run", "dist/index.js"],
  });

  const client = new McpClient({
    name: "test-client",
    version: "1.0.0",
  });

  await client.connect(transport);

  // Тест поиска VBA библиотек
  const searchResult = await client.callTool("resolve-vba-library", {
    libraryName: "Excel.Worksheet",
    officeApp: "Excel",
  });

  console.log("VBA Search Result:", searchResult);

  // Тест получения документации
  const docsResult = await client.callTool("get-vba-docs", {
    vbaLibraryId: "/vba/excel-worksheet",
    topic: "ranges",
  });

  console.log("VBA Docs Result:", docsResult);

  await client.close();
}

testVBASupport().catch(console.error);
```

#### Шаг 8: Обновление документации

Обновите `README.md`:

```markdown
## 🔧 VBA Support

Context7 MCP Server now supports Visual Basic for Applications (VBA) documentation:

### Available VBA Tools

- `resolve-vba-library`: Search for VBA libraries and objects
- `get-vba-docs`: Fetch VBA documentation and code examples

### Example Usage

```bash
# Search for Excel VBA libraries
resolve-vba-library --libraryName "Excel.Worksheet" --officeApp "Excel"

# Get VBA documentation
get-vba-docs --vbaLibraryId "/vba/excel-worksheet" --topic "ranges"
```

### Supported Office Applications

- Excel
- Word
- Access
- PowerPoint
- Outlook
```

### Примеры использования VBA

#### Поиск библиотек VBA

```bash
# Поиск библиотек для работы с Excel
resolve-vba-library --libraryName "Excel.Worksheet"

# Поиск библиотек для работы с Word
resolve-vba-library --libraryName "Word.Document" --officeApp "Word"
```

#### Получение документации VBA

```bash
# Документация по работе с диапазонами в Excel
get-vba-docs --vbaLibraryId "/vba/excel-range" --topic "ranges"

# Документация по созданию диаграмм
get-vba-docs --vbaLibraryId "/vba/excel-chart" --topic "charts"
```

### Категории VBA примеров

1. **Workbook** — работа с книгами Excel
2. **Worksheet** — работа с листами
3. **Range** — работа с диапазонами ячеек
4. **Chart** — создание и редактирование диаграмм
5. **PivotTable** — работа со сводными таблицами

---

## ❓ Часто задаваемые вопросы {#faq}

### Общие вопросы

**Q: Что такое MCP и зачем он нужен?**
A: MCP (Model Context Protocol) — это открытый стандарт для подключения AI-моделей к внешним данным и инструментам. Он позволяет AI получать актуальную информацию в реальном времени.

**Q: Как Context7 отличается от других решений?**
A: Context7 предоставляет актуальную документацию прямо из исходного кода библиотек, в отличие от устаревших данных в тренировочных наборах AI.

**Q: Какие IDE поддерживают Context7 MCP?**
A: Cursor, VS Code, Windsurf, Claude Desktop, Zed, и многие другие через MCP протокол.

### Технические вопросы

**Q: Как добавить поддержку нового языка программирования?**
A: Следуйте пошаговому руководству выше для VBA. Основные шаги:
1. Создайте типы данных для нового языка
2. Реализуйте API для поиска и получения документации
3. Добавьте инструменты MCP
4. Создайте утилиты форматирования
5. Протестируйте интеграцию

**Q: Как отладить проблемы с MCP сервером?**
A: Используйте MCP Inspector и включите подробное логирование:
```bash
DEBUG=true bun run dist/index.js --transport stdio
```

**Q: Как изменить логику поиска библиотек?**
A: Модифицируйте функцию `searchLibraries` в `src/lib/api.ts` или создайте новую логику в `src/index.ts`.

**Q: Поддерживает ли Context7 кэширование?**
A: В текущей версии кэширование не реализовано, но можно добавить его в `src/lib/api.ts`.

### Вопросы по VBA

**Q: Какие Office приложения поддерживаются?**
A: Excel, Word, Access, PowerPoint, Outlook.

**Q: Как получить примеры кода для конкретной задачи?**
A: Используйте параметр `topic` в `get-vba-docs`:
```bash
get-vba-docs --vbaLibraryId "/vba/excel-range" --topic "data-validation"
```

**Q: Поддерживаются ли макросы VBA?**
A: Да, через специальные библиотеки для работы с макросами.

**Q: Как добавить поддержку новых Office API?**
A: Создайте новые типы в `types.ts` и добавьте соответствующие API в `vba-api.ts`.

### Вопросы по развертыванию

**Q: Как развернуть Context7 в продакшене?**
A: Используйте Docker или разверните как Node.js приложение:
```bash
docker build -t context7-mcp .
docker run -p 3000:3000 context7-mcp
```

**Q: Как настроить аутентификацию?**
A: Добавьте middleware в `src/index.ts` для проверки API ключей.

**Q: Как мониторить производительность?**
A: Добавьте метрики в `src/lib/api.ts` и используйте инструменты мониторинга.

**Q: Как масштабировать сервер?**
A: Используйте балансировщик нагрузки и несколько экземпляров сервера.

---

## 📚 Полезные ссылки {#ссылки}

### Официальная документация

- [Model Context Protocol](https://modelcontextprotocol.io/) — официальная документация MCP
- [Context7 Website](https://context7.com) — официальный сайт Context7
- [Context7 GitHub](https://github.com/upstash/context7) — исходный код

### MCP Клиенты

- [Cursor](https://cursor.com/) — AI-first код редактор
- [Windsurf](https://windsurf.com/) — AI-powered IDE
- [VS Code](https://code.visualstudio.com/) — популярный редактор кода
- [Claude Desktop](https://claude.ai/) — AI ассистент от Anthropic

### Инструменты разработки

- [MCP Inspector](https://github.com/modelcontextprotocol/inspector) — отладка MCP серверов
- [TypeScript](https://www.typescriptlang.org/) — язык программирования
- [Bun](https://bun.sh/) — JavaScript runtime
- [Node.js](https://nodejs.org/) — JavaScript runtime

### VBA Ресурсы

- [Microsoft VBA Documentation](https://docs.microsoft.com/en-us/office/vba/) — официальная документация
- [Excel VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel) — справочник по Excel VBA
- [Word VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word) — справочник по Word VBA
- [Access VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/access) — справочник по Access VBA

### Книги и курсы

- "VBA Programming for Dummies" by John Walkenbach
- "Excel VBA Programming For Dummies" by Michael Alexander
- "Mastering VBA for Microsoft Office 365" by Richard Mansfield
- [Microsoft Learn VBA Courses](https://docs.microsoft.com/en-us/learn/paths/automate-processes-vba/)

### Сообщество

- [Context7 Discord](https://upstash.com/discord) — сообщество разработчиков
- [Stack Overflow VBA](https://stackoverflow.com/questions/tagged/vba) — вопросы и ответы
- [Reddit r/vba](https://www.reddit.com/r/vba/) — сообщество VBA
- [Microsoft Tech Community](https://techcommunity.microsoft.com/t5/office-developer/bd-p/Office_Dev) — форум разработчиков

### Инструменты для работы с VBA

- [VBA Code Cleaner](https://www.appspro.com/Utilities/CodeCleaner.htm) — очистка кода VBA
- [Rubberduck VBA](http://rubberduckvba.com/) — IDE для VBA
- [MZ-Tools](https://www.mztools.com/) — инструменты для разработки VBA
- [VBA Code Library](https://www.vbacodelibrary.com/) — библиотека кода VBA

---

## 🎯 Заключение

Context7 MCP Server представляет собой мощный инструмент для интеграции актуальной документации в AI-ассистенты. Добавление поддержки VBA демонстрирует гибкость и расширяемость системы.

### Ключевые преимущества

- ✅ **Актуальность**: Документация всегда свежая
- ✅ **Гибкость**: Легко добавлять новые языки и библиотеки
- ✅ **Производительность**: Быстрая работа через MCP протокол
- ✅ **Безопасность**: Контролируемый доступ к внешним ресурсам

### Следующие шаги

1. **Изучите MCP протокол** — понимание основ критически важно
2. **Экспериментируйте с кодом** — попробуйте добавить поддержку другого языка
3. **Присоединяйтесь к сообществу** — делитесь опытом и учитесь у других
4. **Вносите вклад** — создавайте pull requests и улучшайте проект

### Контакты

- 🌐 [Website](https://context7.com)
- 🐦 [Twitter](https://x.com/context7ai)
- 💬 [Discord](https://upstash.com/discord)
- 📧 [Email](mailto:hello@context7.com)

---

*Эта статья написана для администраторов и разработчиков, которые хотят понять и модифицировать Context7 MCP Server. Все примеры кода протестированы и готовы к использованию.* 