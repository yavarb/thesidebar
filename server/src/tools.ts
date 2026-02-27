/**
 * @module tools
 * OpenAI function-calling compatible tool definitions that expose
 * The Sidebar's own API endpoints. Each tool maps to a localhost:3001 API call.
 *
 * Used by the agentic loop to let LLMs interact with the Word document.
 */

/** OpenAI function-calling tool definition */
export interface ToolDefinition {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: {
      type: "object";
      properties: Record<string, any>;
      required?: string[];
    };
  };
}

/** Mapping from tool name to API endpoint info */
export interface ToolEndpoint {
  method: "GET" | "POST" | "PUT";
  path: string;
  /** How to map tool arguments to the request */
  mapArgs?: (args: Record<string, any>) => { path: string; body?: any; query?: Record<string, string> };
}

/**
 * All tool definitions for The Sidebar document operations.
 * These are in OpenAI function-calling format.
 */
export const TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    type: "function",
    function: {
      name: "readDocument",
      description: "Read the full text content of the current Word document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "readParagraphs",
      description: "Read a range of paragraphs from the document. Returns paragraph text with indices.",
      parameters: {
        type: "object",
        properties: {
          from: { type: "number", description: "Starting paragraph index (0-based)" },
          to: { type: "number", description: "Ending paragraph index (exclusive)" },
          compact: { type: "boolean", description: "If true, return compact format (text only)" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "readParagraph",
      description: "Read a single paragraph by its index.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index (0-based)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "replaceParagraph",
      description: "Replace the text of a paragraph at a given index. Use listString to identify the paragraph for safety.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index to replace" },
          text: { type: "string", description: "New text for the paragraph" },
          listString: { type: "string", description: "Expected current text (for verification)" },
        },
        required: ["index", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insertText",
      description: "Insert a new paragraph at a specified position in the document.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Text to insert" },
          position: { type: "string", enum: ["before", "after", "end", "start"], description: "Where to insert relative to the reference index" },
          index: { type: "number", description: "Reference paragraph index" },
          style: { type: "string", description: "Style to apply (e.g., 'Heading 1', 'Normal')" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "findReplace",
      description: "Find and replace text in the document.",
      parameters: {
        type: "object",
        properties: {
          find: { type: "string", description: "Text to find" },
          replace: { type: "string", description: "Replacement text" },
          matchCase: { type: "boolean", description: "Case-sensitive matching" },
          replaceAll: { type: "boolean", description: "Replace all occurrences (default: true)" },
        },
        required: ["find", "replace"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "readFootnotes",
      description: "Read all footnotes in the document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "addFootnote",
      description: "Add a footnote to a paragraph.",
      parameters: {
        type: "object",
        properties: {
          paragraphIndex: { type: "number", description: "Paragraph to attach the footnote to" },
          text: { type: "string", description: "Footnote text" },
        },
        required: ["paragraphIndex", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "formatParagraph",
      description: "Apply formatting to a paragraph (bold, italic, font size, alignment, style).",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index" },
          bold: { type: "boolean", description: "Set bold" },
          italic: { type: "boolean", description: "Set italic" },
          underline: { type: "boolean", description: "Set underline" },
          fontSize: { type: "number", description: "Font size in points" },
          fontName: { type: "string", description: "Font name" },
          alignment: { type: "string", enum: ["left", "center", "right", "justified"], description: "Text alignment" },
          style: { type: "string", description: "Named style to apply" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getStyles",
      description: "List all available styles in the document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getDocumentStats",
      description: "Get document statistics (word count, paragraph count, etc.).",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getStructure",
      description: "Get the document outline/structure (headings tree).",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "readSelection",
      description: "Read the currently selected text in the document.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "navigateTo",
      description: "Navigate to (scroll to and select) a specific paragraph in the document.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index to navigate to" },
        },
        required: ["index"],
      },
    },
  },
];

/**
 * Map from tool name to the API endpoint it calls.
 * The agentic loop uses this to execute tool calls.
 */
export const TOOL_ENDPOINTS: Record<string, ToolEndpoint> = {
  readDocument: { method: "GET", path: "/api/document" },
  readParagraphs: {
    method: "GET",
    path: "/api/document/paragraphs",
    mapArgs: (args) => ({
      path: "/api/document/paragraphs",
      query: {
        ...(args.from !== undefined ? { from: String(args.from) } : {}),
        ...(args.to !== undefined ? { to: String(args.to) } : {}),
        ...(args.compact ? { compact: "true" } : {}),
      },
    }),
  },
  readParagraph: {
    method: "GET",
    path: "/api/paragraph/:index",
    mapArgs: (args) => ({ path: `/api/paragraph/${args.index}` }),
  },
  replaceParagraph: {
    method: "POST",
    path: "/api/paragraph/replace",
    mapArgs: (args) => ({ path: "/api/paragraph/replace", body: args }),
  },
  insertText: {
    method: "POST",
    path: "/api/insert",
    mapArgs: (args) => ({ path: "/api/insert", body: args }),
  },
  findReplace: {
    method: "POST",
    path: "/api/find-replace",
    mapArgs: (args) => ({ path: "/api/find-replace", body: args }),
  },
  readFootnotes: { method: "GET", path: "/api/footnotes" },
  addFootnote: {
    method: "POST",
    path: "/api/footnote",
    mapArgs: (args) => ({ path: "/api/footnote", body: args }),
  },
  formatParagraph: {
    method: "POST",
    path: "/api/format",
    mapArgs: (args) => ({ path: "/api/format", body: args }),
  },
  getStyles: { method: "GET", path: "/api/styles" },
  getDocumentStats: { method: "GET", path: "/api/document/stats" },
  getStructure: { method: "GET", path: "/api/document/structure" },
  readSelection: { method: "GET", path: "/api/selection" },
  navigateTo: {
    method: "POST",
    path: "/api/navigate",
    mapArgs: (args) => ({ path: "/api/navigate", body: { index: args.index } }),
  },
};

/**
 * Get tool definitions, optionally filtered by name.
 * @param names - If provided, only return tools with these names
 */
export function getToolDefinitions(names?: string[]): ToolDefinition[] {
  if (!names) return TOOL_DEFINITIONS;
  return TOOL_DEFINITIONS.filter((t) => names.includes(t.function.name));
}

/**
 * Get the endpoint mapping for a tool by name.
 * @param name - Tool name
 * @returns The endpoint info, or undefined if not found
 */
export function getToolEndpoint(name: string): ToolEndpoint | undefined {
  return TOOL_ENDPOINTS[name];
}
