/* Office.js integration helpers */
/* eslint-disable @typescript-eslint/no-explicit-any */

declare global {
  interface Window {
    Office?: any;
    Word?: any;
  }
}

let officeLoaded = false;

function loadOfficeScript(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof window === "undefined") { resolve(); return; }
    if (window.Office) { resolve(); return; }
    const script = document.createElement("script");
    script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    script.onload = () => resolve();
    script.onerror = () => {
      console.warn("[OfficeAgent] Failed to load Office.js — running in demo mode.");
      resolve();
    };
    document.head.appendChild(script);
  });
}

export async function initializeOffice(): Promise<boolean> {
  if (typeof window === "undefined") return false;

  // Always try to load Office.js — don't gate on host detection
  try {
    await loadOfficeScript();
    if (window.Office && typeof window.Office.onReady === "function") {
      return new Promise((resolve) => {
        // Add a timeout in case onReady never fires
        const timeout = setTimeout(() => {
          console.warn("[OfficeAgent] Office.onReady timed out — demo mode.");
          resolve(false);
        }, 5000);

        window.Office.onReady((info: { host: string; platform: string }) => {
          clearTimeout(timeout);
          console.log(`[OfficeAgent] Office ready — host: ${info.host}, platform: ${info.platform}`);
          officeLoaded = true;
          resolve(true);
        });
      });
    }
  } catch {
    console.warn("[OfficeAgent] Office.js initialization failed — demo mode.");
  }
  console.log("[OfficeAgent] Running in demo mode (outside Word).");
  return false;
}

/** Check if Office.js is available and ready. Exported so UI can use it. */
export function isOfficeReady(): boolean {
  try {
    return officeLoaded && !!window.Word && !!window.Office?.context;
  } catch {
    return false;
  }
}

/**
 * Safe wrapper around Word.run that handles context errors gracefully.
 * If Office.context is lost (e.g., HMR reload), it attempts re-initialization.
 */
async function safeWordRun<T>(callback: (context: any) => Promise<T>): Promise<T> {
  if (!isOfficeReady()) {
    throw new Error("Not connected to Word (demo mode)");
  }
  try {
    return await window.Word.run(callback);
  } catch (err: any) {
    // If context was lost (the appId error), try re-initializing once
    if (err?.message?.includes("appId") || err?.message?.includes("context")) {
      console.warn("[OfficeAgent] Context lost, attempting re-initialization...");
      officeLoaded = false;
      const reconnected = await initializeOffice();
      if (reconnected && isOfficeReady()) {
        return await window.Word.run(callback);
      }
    }
    throw err;
  }
}

export async function getDocumentText(): Promise<string> {
  if (!isOfficeReady()) {
    return "[Demo mode] Sample document text. In Word, this returns actual content.";
  }
  return safeWordRun(async (context: any) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    return body.text;
  });
}

export async function getDocumentParagraphs(): Promise<Array<{ text: string; style: string; index: number }>> {
  if (!isOfficeReady()) {
    return [
      { text: "[Demo] Sample paragraph 1", style: "Normal", index: 0 },
      { text: "[Demo] Sample heading", style: "Heading1", index: 1 },
      { text: "[Demo] Another paragraph", style: "Normal", index: 2 },
    ];
  }
  return safeWordRun(async (context: any) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(["text", "style"]);
    await context.sync();
    return paragraphs.items.map((p: any, i: number) => ({
      text: p.text,
      style: p.style,
      index: i,
    }));
  });
}

export async function highlightAllText(color: string = "Yellow"): Promise<void> {
  return safeWordRun(async (context: any) => {
    const body = context.document.body;
    body.font.highlightColor = color;
    await context.sync();
  });
}

export async function insertText(text: string): Promise<void> {
  return safeWordRun(async (context: any) => {
    const body = context.document.body;
    body.insertParagraph(text, "End");
    await context.sync();
  });
}

export async function deleteParagraphs(startIndex: number, endIndex: number): Promise<void> {
  return safeWordRun(async (context: any) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    for (let i = Math.min(endIndex, paragraphs.items.length - 1); i >= startIndex; i--) {
      paragraphs.items[i].delete();
    }
    await context.sync();
  });
}

export async function deletePages(startPage: number, endPage: number): Promise<void> {
  return safeWordRun(async (context: any) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    const ranges: any[] = paragraphs.items.map((p: any) => {
      const r = p.getRange("Whole");
      r.load("pageNumber");
      return r;
    });

    let usePageAPI = false;
    try {
      await context.sync();
      if (ranges.length > 0 && typeof ranges[0].pageNumber === "number") {
        usePageAPI = true;
      }
    } catch {
      usePageAPI = false;
    }

    if (usePageAPI) {
      const toDelete = paragraphs.items
        .filter((_: any, i: number) => {
          const pg: number = ranges[i].pageNumber;
          return pg >= startPage && pg <= endPage;
        })
        .reverse();
      for (const para of toDelete) para.delete();
    } else {
      const PARAS_PER_PAGE = 40;
      const start = (startPage - 1) * PARAS_PER_PAGE;
      const end = Math.min(endPage * PARAS_PER_PAGE - 1, paragraphs.items.length - 1);
      for (let i = end; i >= start; i--) {
        paragraphs.items[i].delete();
      }
    }
    await context.sync();
  });
}

export async function keepOnlyParagraphsWithText(searchText: string): Promise<void> {
  return safeWordRun(async (context: any) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    const lower = searchText.toLowerCase();

    const matchIndices = new Set<number>();
    paragraphs.items.forEach((p: any, i: number) => {
      if (p.text.toLowerCase().includes(lower)) {
        matchIndices.add(i);
      }
    });

    console.log(`[OfficeAgent] Found ${matchIndices.size} matching paragraphs out of ${paragraphs.items.length}`);

    if (matchIndices.size === 0) {
      throw new Error(`No paragraphs found containing "${searchText}"`);
    }

    for (let i = paragraphs.items.length - 1; i >= 0; i--) {
      if (!matchIndices.has(i)) {
        paragraphs.items[i].delete();
      }
    }

    await context.sync();
  });
}

export async function formatByStyle(
  targetStyle: string,
  formatting: { bold?: boolean; italic?: boolean; color?: string; size?: number }
): Promise<void> {
  return safeWordRun(async (context: any) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(["text", "style"]);
    await context.sync();
    for (const p of paragraphs.items) {
      if (p.style === targetStyle) {
        if (formatting.bold !== undefined) p.font.bold = formatting.bold;
        if (formatting.italic !== undefined) p.font.italic = formatting.italic;
        if (formatting.color) p.font.color = formatting.color;
        if (formatting.size) p.font.size = formatting.size;
      }
    }
    await context.sync();
  });
}

export async function replaceText(find: string, replace: string): Promise<void> {
  return safeWordRun(async (context: any) => {
    const body = context.document.body;
    const searchResults = body.search(find, { matchCase: false, matchWholeWord: false });
    searchResults.load("text");
    await context.sync();
    for (const result of searchResults.items) {
      result.insertText(replace, "Replace");
    }
    await context.sync();
  });
}

export async function setFont(options: {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}): Promise<void> {
  return safeWordRun(async (context: any) => {
    const body = context.document.body;
    if (options.name) body.font.name = options.name;
    if (options.size) body.font.size = options.size;
    if (options.color) body.font.color = options.color;
    if (options.bold !== undefined) body.font.bold = options.bold;
    if (options.italic !== undefined) body.font.italic = options.italic;
    await context.sync();
  });
}

/* ─── Central Action Dispatcher ─── */
export async function executeAction(action: {
  type: string;
  params: Record<string, any>;
}): Promise<{ success: boolean; message: string; data?: any }> {
  try {
    switch (action.type) {
      case "get_document_text": {
        const text = await getDocumentText();
        return { success: true, message: "Document text retrieved.", data: text };
      }
      case "insert_text": {
        await insertText(action.params.text);
        return { success: true, message: "Text inserted." };
      }
      case "delete_paragraphs": {
        await deleteParagraphs(action.params.startIndex, action.params.endIndex);
        return { success: true, message: `Deleted paragraphs ${action.params.startIndex}–${action.params.endIndex}.` };
      }
      case "delete_pages": {
        await deletePages(action.params.startPage, action.params.endPage);
        return { success: true, message: `Deleted pages ${action.params.startPage}–${action.params.endPage}.` };
      }
      case "keep_only_pages_with_text": {
        await keepOnlyParagraphsWithText(action.params.searchText);
        return { success: true, message: `Kept only content containing "${action.params.searchText}".` };
      }
      case "delete_all_except_name": {
        await keepOnlyParagraphsWithText(action.params.name);
        return { success: true, message: `Kept only content containing "${action.params.name}".` };
      }
      case "highlight_all": {
        await highlightAllText(action.params.color);
        return { success: true, message: `Highlighted all text in ${action.params.color}.` };
      }
      case "format_by_style": {
        await formatByStyle(action.params.targetStyle, action.params.formatting);
        return { success: true, message: `Formatted "${action.params.targetStyle}" paragraphs.` };
      }
      case "replace_text": {
        await replaceText(action.params.find, action.params.replace);
        return { success: true, message: `Replaced "${action.params.find}" with "${action.params.replace}".` };
      }
      case "set_font": {
        await setFont(action.params);
        return { success: true, message: "Font updated." };
      }
      default:
        return { success: false, message: `Unknown action type: ${action.type}` };
    }
  } catch (err: any) {
    return { success: false, message: `Error executing ${action.type}: ${err.message}` };
  }
}