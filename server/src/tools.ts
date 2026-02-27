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
 */
export const TOOL_DEFINITIONS: ToolDefinition[] = [
  // ═══ Document Reading ═══
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
      name: "getParagraphs",
      description: "Read a range of paragraphs (alias). Returns text with indices, styles, and list info.",
      parameters: {
        type: "object",
        properties: {
          from: { type: "number", description: "Starting paragraph index (0-based)" },
          to: { type: "number", description: "Ending paragraph index (exclusive)" },
          compact: { type: "boolean", description: "If true, return compact format" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getDocumentStats",
      description: "Get document statistics: word count, paragraph count, character count, section count, footnote count.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getStructure",
      description: "Get the document outline/structure as a headings hierarchy tree.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getDocumentProperties",
      description: "Get document metadata (title, author, subject, keywords, creation date, revision number, etc.).",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getToc",
      description: "Get the table of contents entries from the document.",
      parameters: { type: "object", properties: {} },
    },
  },

  // ═══ Selection ═══
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
      name: "editSelection",
      description: "Replace the currently selected text with new text.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "New text to replace the selection with" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "selectParagraph",
      description: "Select a specific paragraph by index (highlights it in the document).",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index to select (0-based)" },
        },
        required: ["index"],
      },
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

  // ═══ Editing ═══
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
      name: "deleteParagraph",
      description: "Delete a paragraph at a specific index.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index to delete (0-based)" },
        },
        required: ["index"],
      },
    },
  },

  // ═══ Find & Replace ═══
  {
    type: "function",
    function: {
      name: "find",
      description: "Find text in the document. Returns match locations and count.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Text to search for" },
          matchCase: { type: "boolean", description: "Case-sensitive matching" },
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

  // ═══ Formatting ═══
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
      name: "setParagraphFormat",
      description: "Set paragraph spacing, indentation, and alignment.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index (0-based)" },
          spaceBefore: { type: "number", description: "Space before paragraph in points" },
          spaceAfter: { type: "number", description: "Space after paragraph in points" },
          lineSpacing: { type: "number", description: "Line spacing in points" },
          leftIndent: { type: "number", description: "Left indent in points" },
          rightIndent: { type: "number", description: "Right indent in points" },
          firstLineIndent: { type: "number", description: "First line indent in points" },
          alignment: { type: "string", enum: ["Left", "Centered", "Right", "Justified"], description: "Paragraph alignment" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "highlightText",
      description: "Highlight all occurrences of text with a color.",
      parameters: {
        type: "object",
        properties: {
          text: { type: "string", description: "Text to highlight" },
          color: { type: "string", description: "Highlight color (e.g., 'yellow', 'green', 'red', 'blue')" },
          matchCase: { type: "boolean", description: "Case-sensitive matching" },
        },
        required: ["text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "setFontColor",
      description: "Set font color for a paragraph or specific text within it.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Paragraph index" },
          color: { type: "string", description: "Font color (e.g., 'red', '#FF0000')" },
          text: { type: "string", description: "Specific text within the paragraph to color (if omitted, colors entire paragraph)" },
        },
        required: ["index", "color"],
      },
    },
  },

  // ═══ Styles ═══
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
      name: "applyStyle",
      description: "Apply a named style to one or more paragraphs.",
      parameters: {
        type: "object",
        properties: {
          fromIndex: { type: "number", description: "Starting paragraph index" },
          toIndex: { type: "number", description: "Ending paragraph index (default: same as fromIndex)" },
          styleName: { type: "string", description: "Style name to apply" },
        },
        required: ["fromIndex", "styleName"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "createStyle",
      description: "Create a new custom paragraph style with font and paragraph settings.",
      parameters: {
        type: "object",
        properties: {
          name: { type: "string", description: "New style name" },
          basedOn: { type: "string", description: "Base style name to inherit from" },
          fontName: { type: "string", description: "Font name" },
          fontSize: { type: "number", description: "Font size in points" },
          bold: { type: "boolean", description: "Bold" },
          italic: { type: "boolean", description: "Italic" },
          color: { type: "string", description: "Font color" },
          spaceBefore: { type: "number", description: "Space before in points" },
          spaceAfter: { type: "number", description: "Space after in points" },
          lineSpacing: { type: "number", description: "Line spacing in points" },
          alignment: { type: "string", description: "Alignment (Left, Centered, Right, Justified)" },
        },
        required: ["name"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "modifyStyle",
      description: "Modify an existing style's font and paragraph properties.",
      parameters: {
        type: "object",
        properties: {
          styleName: { type: "string", description: "Style name to modify" },
          fontName: { type: "string", description: "Font name" },
          fontSize: { type: "number", description: "Font size in points" },
          bold: { type: "boolean", description: "Bold" },
          italic: { type: "boolean", description: "Italic" },
          color: { type: "string", description: "Font color" },
          spaceBefore: { type: "number", description: "Space before in points" },
          spaceAfter: { type: "number", description: "Space after in points" },
          lineSpacing: { type: "number", description: "Line spacing in points" },
          alignment: { type: "string", description: "Alignment" },
        },
        required: ["styleName"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getStyleDetails",
      description: "Get detailed info about a style (font, paragraph format, base style, built-in status).",
      parameters: {
        type: "object",
        properties: {
          styleName: { type: "string", description: "Style name to inspect" },
        },
        required: ["styleName"],
      },
    },
  },

  // ═══ Footnotes ═══
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
      name: "updateFootnote",
      description: "Edit an existing footnote by index.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Footnote index (0-based)" },
          text: { type: "string", description: "New footnote text" },
        },
        required: ["index", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "deleteFootnote",
      description: "Delete a footnote by index.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Footnote index (0-based)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getFootnoteBody",
      description: "Get the full content of a footnote including all its paragraphs and styles.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Footnote index (0-based)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insertFootnoteWithFormat",
      description: "Insert a footnote at a specific anchor text location.",
      parameters: {
        type: "object",
        properties: {
          anchorText: { type: "string", description: "Text in the document to attach the footnote to" },
          footnoteText: { type: "string", description: "Footnote body text" },
          matchCase: { type: "boolean", description: "Case-sensitive anchor search (default: true)" },
        },
        required: ["anchorText", "footnoteText"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "reorderFootnotes",
      description: "List all footnotes with their body text and reference text (for auditing footnote order).",
      parameters: { type: "object", properties: {} },
    },
  },

  // ═══ Comments ═══
  {
    type: "function",
    function: {
      name: "getComments",
      description: "List all comments in the document with content, author, and creation date.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "addComment",
      description: "Add a comment to text in the document.",
      parameters: {
        type: "object",
        properties: {
          paragraphIndex: { type: "number", description: "Paragraph index to attach the comment to" },
          text: { type: "string", description: "Comment text" },
        },
        required: ["paragraphIndex", "text"],
      },
    },
  },

  // ═══ Tables ═══
  {
    type: "function",
    function: {
      name: "getTables",
      description: "List all tables in the document with their row counts and header row info.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "readTable",
      description: "Read a specific table's full contents (all cell values) by index.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Table index (0-based)" },
        },
        required: ["index"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insertTable",
      description: "Insert a new table at the end of the document.",
      parameters: {
        type: "object",
        properties: {
          rows: { type: "number", description: "Number of rows (default: 2)" },
          columns: { type: "number", description: "Number of columns (default: 2)" },
          values: { type: "array", items: { type: "array", items: { type: "string" } }, description: "2D array of cell values" },
          style: { type: "string", description: "Table style name" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "updateTableCell",
      description: "Update the text in a specific table cell.",
      parameters: {
        type: "object",
        properties: {
          tableIndex: { type: "number", description: "Table index (default: 0)" },
          row: { type: "number", description: "Row index (0-based)" },
          column: { type: "number", description: "Column index (0-based)" },
          text: { type: "string", description: "New cell text" },
        },
        required: ["row", "column", "text"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "addTableRow",
      description: "Add a row to an existing table.",
      parameters: {
        type: "object",
        properties: {
          tableIndex: { type: "number", description: "Table index (default: 0)" },
          values: { type: "array", items: { type: "string" }, description: "Cell values for the new row" },
          position: { type: "string", enum: ["start", "end"], description: "Where to add (default: end)" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "addTableColumn",
      description: "Add a column to an existing table.",
      parameters: {
        type: "object",
        properties: {
          tableIndex: { type: "number", description: "Table index (default: 0)" },
          values: { type: "array", items: { type: "string" }, description: "Cell values for the new column" },
          position: { type: "string", enum: ["start", "end"], description: "Where to add (default: end)" },
        },
      },
    },
  },

  // ═══ Headers & Footers ═══
  {
    type: "function",
    function: {
      name: "getHeaderFooter",
      description: "Read the header or footer text from a document section.",
      parameters: {
        type: "object",
        properties: {
          type: { type: "string", enum: ["header", "footer"], description: "Header or footer" },
          sectionIndex: { type: "number", description: "Section index (default: 0)" },
          headerType: { type: "string", enum: ["primary", "firstPage", "evenPages"], description: "Type (default: primary)" },
        },
        required: ["type"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "setHeaderFooter",
      description: "Set the header or footer text for a document section.",
      parameters: {
        type: "object",
        properties: {
          type: { type: "string", enum: ["header", "footer"], description: "Header or footer" },
          text: { type: "string", description: "Text to set" },
          sectionIndex: { type: "number", description: "Section index (default: 0)" },
          headerType: { type: "string", enum: ["primary", "firstPage", "evenPages"], description: "Type (default: primary)" },
        },
        required: ["type", "text"],
      },
    },
  },

  // ═══ Breaks ═══
  {
    type: "function",
    function: {
      name: "insertBreak",
      description: "Insert a page or section break after a paragraph.",
      parameters: {
        type: "object",
        properties: {
          afterParagraph: { type: "number", description: "Paragraph index to insert break after (default: last)" },
          breakType: { type: "string", enum: ["page", "section", "sectionContinuous"], description: "Break type (default: page)" },
        },
      },
    },
  },

  // ═══ Lists ═══
  {
    type: "function",
    function: {
      name: "setListFormat",
      description: "Apply list formatting (bullet, numbered, or none) to a range of paragraphs.",
      parameters: {
        type: "object",
        properties: {
          fromIndex: { type: "number", description: "Starting paragraph index" },
          toIndex: { type: "number", description: "Ending paragraph index (default: same as fromIndex)" },
          type: { type: "string", enum: ["bullet", "numbered", "none"], description: "List type" },
        },
        required: ["fromIndex", "type"],
      },
    },
  },

  // ═══ Bookmarks ═══
  {
    type: "function",
    function: {
      name: "getBookmarks",
      description: "List all bookmarks in the document.",
      parameters: { type: "object", properties: {} },
    },
  },

  // ═══ Tracked Changes ═══
  {
    type: "function",
    function: {
      name: "getTrackedChanges",
      description: "List all tracked changes in the document (requires WordApi 1.6+).",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "acceptTrackedChange",
      description: "Accept a tracked change by index, or accept all.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Index of the tracked change to accept" },
          all: { type: "boolean", description: "Accept all tracked changes" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "rejectTrackedChange",
      description: "Reject a tracked change by index, or reject all.",
      parameters: {
        type: "object",
        properties: {
          index: { type: "number", description: "Index of the tracked change to reject" },
          all: { type: "boolean", description: "Reject all tracked changes" },
        },
      },
    },
  },

  // ═══ Citations / Table of Authorities ═══
  {
    type: "function",
    function: {
      name: "markCitation",
      description: "Mark a citation for Table of Authorities. Inserts a TA field code at the specified text. Categories: 1=Cases, 2=Statutes, 3=Other Authorities, 4=Rules.",
      parameters: {
        type: "object",
        properties: {
          shortCite: { type: "string", description: "Short citation form (e.g., 'Smith v. Jones')" },
          longCite: { type: "string", description: "Full citation (e.g., 'Smith v. Jones, 123 F.3d 456 (2d Cir. 2020)')" },
          category: { type: "number", description: "Category: 1=Cases, 2=Statutes, 3=Other, 4=Rules (default: 1)" },
          searchText: { type: "string", description: "Text to find in document to place the mark (default: shortCite)" },
        },
        required: ["shortCite", "longCite"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "insertTableOfAuthorities",
      description: "Insert a Table of Authorities at a position in the document.",
      parameters: {
        type: "object",
        properties: {
          category: { type: "number", description: "Category to include (0=all, 1=Cases, 2=Statutes, 3=Other, 4=Rules). Default: 0" },
          paragraphIndex: { type: "number", description: "Paragraph index to insert after (default: end of document)" },
        },
      },
    },
  },

  // ═══ Cross-References ═══
  {
    type: "function",
    function: {
      name: "insertCrossReference",
      description: "Insert a cross-reference to a heading, footnote, or bookmark.",
      parameters: {
        type: "object",
        properties: {
          type: { type: "string", enum: ["heading", "footnote", "bookmark"], description: "Reference type" },
          target: { type: "string", description: "Target heading text, footnote number, or bookmark name" },
          text: { type: "string", description: "Display text for the reference (auto-generated if omitted)" },
          paragraphIndex: { type: "number", description: "Paragraph to insert at (default: current selection)" },
        },
        required: ["type", "target"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "validateCrossReferences",
      description: "Scan document for cross-reference patterns (Section X, Article X, see supra/infra, ¶ N, Part X) and validate them against actual headings. Returns issues where referenced sections don't exist.",
      parameters: { type: "object", properties: {} },
    },
  },

  // ═══ Batch Operations ═══
  {
    type: "function",
    function: {
      name: "batch",
      description: "Execute multiple document operations in a single call. Each operation has a 'command' and params.",
      parameters: {
        type: "object",
        properties: {
          operations: {
            type: "array",
            items: {
              type: "object",
              properties: {
                command: { type: "string", description: "Command name" },
              },
              additionalProperties: true,
            },
            description: "Array of operations to execute",
          },
        },
        required: ["operations"],
      },
    },
  },

  // ═══ TOA Page Check ═══
  {
    type: "function",
    function: {
      name: "checkToaPages",
      description: "Export the document to PDF, parse page numbers, read the Table of Authorities, and verify that all citation page numbers are correct. Returns a list of discrepancies.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "getPageSetup",
      description: "Get page margins, gutter, paper size, header/footer distance, and orientation for a document section.",
      parameters: {
        type: "object",
        properties: {
          sectionIndex: { type: "number", description: "Section index (0-based, default 0)" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "setPageSetup",
      description: "Set page margins, gutter, orientation, paper size, header/footer distance. All measurements in points (72 points = 1 inch).",
      parameters: {
        type: "object",
        properties: {
          sectionIndex: { type: "number", description: "Section index (0-based, default 0)" },
          topMargin: { type: "number", description: "Top margin in points" },
          bottomMargin: { type: "number", description: "Bottom margin in points" },
          leftMargin: { type: "number", description: "Left margin in points" },
          rightMargin: { type: "number", description: "Right margin in points" },
          gutter: { type: "number", description: "Gutter in points" },
          orientation: { type: "string", description: "Portrait or Landscape" },
          paperSize: { type: "string", description: "Paper size" },
          headerDistance: { type: "number", description: "Header distance in points" },
          footerDistance: { type: "number", description: "Footer distance in points" },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "getPageNumbers",
      description: "Get section count and section info for the document.",
      parameters: { type: "object", properties: {} },
    },
  },
];


/**
 * Map from tool name to the API endpoint it calls.
 */
export const TOOL_ENDPOINTS: Record<string, ToolEndpoint> = {
  // Session
  searchConversationHistory: { method: "POST", path: "/api/session/search", mapArgs: (args) => ({ path: "/api/session/search", body: { query: args.query } }) },

  // Document reading
  readDocument: { method: "GET", path: "/api/document" },
  readParagraphs: {
    method: "GET", path: "/api/document/paragraphs",
    mapArgs: (args) => ({ path: "/api/document/paragraphs", query: {
      ...(args.from !== undefined ? { from: String(args.from) } : {}),
      ...(args.to !== undefined ? { to: String(args.to) } : {}),
      ...(args.compact ? { compact: "true" } : {}),
    }}),
  },
  readParagraph: { method: "GET", path: "/api/paragraph/:index", mapArgs: (args) => ({ path: `/api/paragraph/${args.index}` }) },
  getParagraphs: {
    method: "GET", path: "/api/document/paragraphs",
    mapArgs: (args) => ({ path: "/api/document/paragraphs", query: {
      ...(args.from !== undefined ? { from: String(args.from) } : {}),
      ...(args.to !== undefined ? { to: String(args.to) } : {}),
      ...(args.compact ? { compact: "true" } : {}),
    }}),
  },
  getDocumentStats: { method: "GET", path: "/api/document/stats" },
  getStructure: { method: "GET", path: "/api/document/structure" },
  getDocumentProperties: { method: "GET", path: "/api/document/properties" },
  getToc: { method: "GET", path: "/api/document/toc" },

  // Selection
  readSelection: { method: "GET", path: "/api/selection" },
  editSelection: { method: "POST", path: "/api/selection/edit", mapArgs: (args) => ({ path: "/api/selection/edit", body: { replacement: args.text } }) },
  selectParagraph: { method: "POST", path: "/api/select", mapArgs: (args) => ({ path: "/api/select", body: { index: args.index } }) },
  navigateTo: { method: "POST", path: "/api/navigate", mapArgs: (args) => ({ path: "/api/navigate", body: { index: args.index } }) },

  // Editing
  replaceParagraph: { method: "POST", path: "/api/paragraph/replace", mapArgs: (args) => ({ path: "/api/paragraph/replace", body: args }) },
  insertText: { method: "POST", path: "/api/insert", mapArgs: (args) => ({ path: "/api/insert", body: args }) },
  deleteParagraph: { method: "POST", path: "/api/paragraph/delete", mapArgs: (args) => ({ path: "/api/paragraph/delete", body: args }) },

  // Find & Replace
  find: { method: "POST", path: "/api/find", mapArgs: (args) => ({ path: "/api/find", body: args }) },
  findReplace: { method: "POST", path: "/api/find-replace", mapArgs: (args) => ({ path: "/api/find-replace", body: args }) },

  // Formatting
  formatParagraph: { method: "POST", path: "/api/format", mapArgs: (args) => ({ path: "/api/format", body: args }) },
  setParagraphFormat: { method: "POST", path: "/api/paragraph/format", mapArgs: (args) => ({ path: "/api/paragraph/format", body: args }) },
  highlightText: { method: "POST", path: "/api/highlight", mapArgs: (args) => ({ path: "/api/highlight", body: args }) },
  setFontColor: { method: "POST", path: "/api/font-color", mapArgs: (args) => ({ path: "/api/font-color", body: args }) },

  // Styles
  getStyles: { method: "GET", path: "/api/styles" },
  applyStyle: { method: "POST", path: "/api/style/apply", mapArgs: (args) => ({ path: "/api/style/apply", body: args }) },
  createStyle: { method: "POST", path: "/api/style/create", mapArgs: (args) => ({ path: "/api/style/create", body: args }) },
  modifyStyle: { method: "POST", path: "/api/style/modify", mapArgs: (args) => ({ path: "/api/style/modify", body: args }) },
  getStyleDetails: { method: "POST", path: "/api/style/details", mapArgs: (args) => ({ path: "/api/style/details", body: args }) },

  // Footnotes
  readFootnotes: { method: "GET", path: "/api/footnotes" },
  addFootnote: { method: "POST", path: "/api/footnote", mapArgs: (args) => ({ path: "/api/footnote", body: args }) },
  updateFootnote: { method: "PUT", path: "/api/footnote/:index", mapArgs: (args) => ({ path: `/api/footnote/${args.index}`, body: { text: args.text } }) },
  deleteFootnote: { method: "POST", path: "/api/footnote/delete", mapArgs: (args) => ({ path: "/api/footnote/delete", body: args }) },
  getFootnoteBody: { method: "GET", path: "/api/footnote/:index/body", mapArgs: (args) => ({ path: `/api/footnote/${args.index}/body` }) },
  insertFootnoteWithFormat: { method: "POST", path: "/api/footnote/insert", mapArgs: (args) => ({ path: "/api/footnote/insert", body: args }) },
  reorderFootnotes: { method: "GET", path: "/api/footnotes/detailed" },

  // Comments
  getComments: { method: "GET", path: "/api/comments" },
  addComment: { method: "POST", path: "/api/comment", mapArgs: (args) => ({ path: "/api/comment", body: args }) },

  // Tables
  getTables: { method: "GET", path: "/api/tables" },
  readTable: { method: "GET", path: "/api/table/:index", mapArgs: (args) => ({ path: `/api/table/${args.index}` }) },
  insertTable: { method: "POST", path: "/api/table/insert", mapArgs: (args) => ({ path: "/api/table/insert", body: args }) },
  updateTableCell: { method: "POST", path: "/api/table/cell", mapArgs: (args) => ({ path: "/api/table/cell", body: args }) },
  addTableRow: { method: "POST", path: "/api/table/row", mapArgs: (args) => ({ path: "/api/table/row", body: args }) },
  addTableColumn: { method: "POST", path: "/api/table/column", mapArgs: (args) => ({ path: "/api/table/column", body: args }) },

  // Headers & Footers
  getHeaderFooter: { method: "POST", path: "/api/header-footer", mapArgs: (args) => ({ path: "/api/header-footer", body: args }) },
  setHeaderFooter: { method: "POST", path: "/api/header-footer/set", mapArgs: (args) => ({ path: "/api/header-footer/set", body: args }) },

  // Breaks
  insertBreak: { method: "POST", path: "/api/break", mapArgs: (args) => ({ path: "/api/break", body: args }) },

  // Lists
  setListFormat: { method: "POST", path: "/api/list-format", mapArgs: (args) => ({ path: "/api/list-format", body: args }) },

  // Bookmarks
  getBookmarks: { method: "GET", path: "/api/bookmarks" },

  // Tracked Changes
  getTrackedChanges: { method: "GET", path: "/api/tracked-changes" },
  acceptTrackedChange: { method: "POST", path: "/api/tracked-changes/accept", mapArgs: (args) => ({ path: "/api/tracked-changes/accept", body: args }) },
  rejectTrackedChange: { method: "POST", path: "/api/tracked-changes/reject", mapArgs: (args) => ({ path: "/api/tracked-changes/reject", body: args }) },

  // Citations / TOA
  markCitation: { method: "POST", path: "/api/citation/mark", mapArgs: (args) => ({ path: "/api/citation/mark", body: args }) },
  insertTableOfAuthorities: { method: "POST", path: "/api/citation/toa", mapArgs: (args) => ({ path: "/api/citation/toa", body: args }) },

  // Cross-References
  insertCrossReference: { method: "POST", path: "/api/cross-reference", mapArgs: (args) => ({ path: "/api/cross-reference", body: args }) },
  validateCrossReferences: { method: "GET", path: "/api/cross-references/validate" },

  // Batch

  // TOA & Page Setup
  checkToaPages: { method: "POST", path: "/api/toa/check" },
  getPageSetup: { method: "GET", path: "/api/page/setup", mapArgs: (args) => ({ path: "/api/page/setup", query: { ...(args.sectionIndex !== undefined ? { sectionIndex: String(args.sectionIndex) } : {}) } }) },
  setPageSetup: { method: "POST", path: "/api/page/setup", mapArgs: (args) => ({ path: "/api/page/setup", body: args }) },
  getPageNumbers: { method: "GET", path: "/api/page/info" },
  batch: { method: "POST", path: "/api/batch", mapArgs: (args) => ({ path: "/api/batch", body: args }) },
};

/**
 * Get tool definitions, optionally filtered by name.
 */
export function getToolDefinitions(names?: string[]): ToolDefinition[] {
  if (!names) return TOOL_DEFINITIONS;
  return TOOL_DEFINITIONS.filter((t) => names.includes(t.function.name));
}

/**
 * Get the endpoint mapping for a tool by name.
 */
export function getToolEndpoint(name: string): ToolEndpoint | undefined {
  return TOOL_ENDPOINTS[name];
}
