"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  initializeOffice,
  isOfficeReady,
  getDocumentText,
  highlightAllText,
  insertText,
  deleteParagraphs,
  deletePages,
  keepOnlyParagraphsWithText,
  formatByStyle,
  replaceText,
  setFont,
} from "@/lib/office";

/* ─── Types ─── */
interface Message {
  id: string;
  role: "user" | "agent" | "system";
  content: string;
  timestamp: Date;
  status?: "pending" | "success" | "error";
}

interface AIAction {
  type: string;
  params: Record<string, any>;
}

interface AIResponse {
  message: string;
  actions: AIAction[];
  error?: string;
}

/* ─── Quick Action Chips ─── */
const QUICK_ACTIONS = [
  { label: "📄 Get doc text", prompt: "Get the full document text" },
  { label: "🖍️ Highlight all", prompt: "Highlight all text in yellow" },
  { label: "✍️ Insert text", prompt: 'Insert a paragraph saying "Hello from AI Agent!"' },
  { label: "🗑️ Delete paragraph 1", prompt: "Delete the first paragraph" },
  { label: "🎨 Bold headings", prompt: "Make all Heading 1 paragraphs bold" },
];

/* ─── Execute AI Actions on the document ─── */
async function executeActions(actions: AIAction[]): Promise<string[]> {
  // If not connected to Word, warn the user immediately
  if (!isOfficeReady()) {
    return ["⚠️ Not connected to Word — actions cannot be executed. Please open this add-in inside Microsoft Word."];
  }

  const results: string[] = [];

  for (const action of actions) {
    try {
      switch (action.type) {
        case "get_document_text": {
          const text = await getDocumentText();
          results.push(`📄 Document text:\n\`\`\`\n${text.slice(0, 500)}${text.length > 500 ? "\n..." : ""}\n\`\`\``);
          break;
        }
        case "insert_text":
          await insertText(action.params.text);
          results.push(`✅ Inserted: "${action.params.text}"`);
          break;
        case "delete_paragraphs":
          await deleteParagraphs(action.params.startIndex, action.params.endIndex);
          results.push(`✅ Deleted paragraphs ${action.params.startIndex + 1}–${action.params.endIndex + 1}`);
          break;
        case "delete_pages":
          await deletePages(action.params.startPage, action.params.endPage);
          results.push(`✅ Deleted pages ${action.params.startPage}–${action.params.endPage}`);
          break;
        case "keep_only_pages_with_text":
          await keepOnlyParagraphsWithText(action.params.searchText);
          results.push(`✅ Kept only content containing "${action.params.searchText}"`);
          break;
        case "delete_all_except_name":
          await keepOnlyParagraphsWithText(action.params.name);
          results.push(`✅ Kept only content containing "${action.params.name}"`);
          break;
        case "highlight_all":
          await highlightAllText(action.params.color || "Yellow");
          results.push(`✅ Highlighted all text in ${action.params.color || "Yellow"}`);
          break;
        case "format_by_style":
          await formatByStyle(action.params.targetStyle, action.params.formatting);
          results.push(`✅ Formatted ${action.params.targetStyle} paragraphs`);
          break;
        case "replace_text":
          await replaceText(action.params.find, action.params.replace);
          results.push(`✅ Replaced "${action.params.find}" → "${action.params.replace}"`);
          break;
        case "set_font":
          await setFont(action.params);
          results.push(`✅ Font updated`);
          break;
        default:
          results.push(`⚠️ Unknown action: ${action.type}`);
      }
    } catch (err: any) {
      results.push(`❌ ${action.type} failed: ${err.message}`);
    }
  }

  return results;
}

/* ─── Component ─── */
export default function TaskpaneChat() {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: "welcome",
      role: "agent",
      content:
        "Hey! 👋 I'm your Office Agent powered by **AI**. Tell me what you'd like to do with your document — delete pages, format text, highlight content, find & replace, and more.",
      timestamp: new Date(),
    },
  ]);
  const [input, setInput] = useState("");
  const [isThinking, setIsThinking] = useState(false);
  const [officeReady, setOfficeReady] = useState(false);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    initializeOffice().then((isConnected) => setOfficeReady(isConnected));
  }, []);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isThinking]);

  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    e.target.style.height = "auto";
    e.target.style.height = Math.min(e.target.scrollHeight, 120) + "px";
  };

  const callAI = useCallback(async (userMessage: string): Promise<AIResponse> => {
    try {
      let documentContext: string | undefined;
      try {
        documentContext = await getDocumentText();
        if (documentContext.startsWith("[Demo")) {
          documentContext = undefined;
        }
      } catch {
        // ignore
      }

      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message: userMessage, documentContext }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || err.message || `HTTP ${res.status}`);
      }

      return await res.json();
    } catch (err: any) {
      return {
        message: `⚠️ ${err.message || "Failed to reach AI backend"}`,
        actions: [],
        error: err.message,
      };
    }
  }, []);

  const sendMessage = useCallback(
    async (text?: string) => {
      const content = text || input.trim();
      if (!content || isThinking) return;

      const userMsg: Message = {
        id: Date.now().toString(),
        role: "user",
        content,
        timestamp: new Date(),
      };

      setMessages((prev) => [...prev, userMsg]);
      setInput("");
      if (inputRef.current) inputRef.current.style.height = "auto";
      setIsThinking(true);

      const aiResponse = await callAI(content);

      let actionResults: string[] = [];
      if (aiResponse.actions && aiResponse.actions.length > 0) {
        actionResults = await executeActions(aiResponse.actions);
      }

      let agentContent = aiResponse.message || "Done!";
      if (actionResults.length > 0) {
        agentContent += "\n\n" + actionResults.join("\n");
      }

      const agentMsg: Message = {
        id: (Date.now() + 1).toString(),
        role: "agent",
        content: agentContent,
        timestamp: new Date(),
        status: aiResponse.error ? "error" : "success",
      };

      setMessages((prev) => [...prev, agentMsg]);
      setIsThinking(false);
    },
    [input, isThinking, callAI]
  );

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  };

  return (
    <div className="flex flex-col h-screen" style={{ background: "var(--bg-primary)" }}>
      {/* ── Header ── */}
      <header
        className="glass-panel flex items-center gap-3 px-4 py-3 shrink-0"
        style={{ borderBottom: "1px solid var(--border-subtle)" }}
      >
        <div
          className="w-9 h-9 rounded-xl flex items-center justify-center text-lg font-bold"
          style={{
            background: "linear-gradient(135deg, var(--accent), #818cf8)",
            color: "#fff",
          }}
        >
          ⚡
        </div>
        <div className="flex-1 min-w-0">
          <h1 className="text-sm font-semibold" style={{ color: "var(--text-primary)" }}>
            Office Agent
          </h1>
          <p className="text-xs" style={{ color: "var(--text-secondary)" }}>
            {officeReady ? "🟢 Connected to Word" : "🟡 Demo mode"} · AI Agent
          </p>
        </div>
        <div
          className="w-2 h-2 rounded-full"
          style={{
            background: officeReady ? "var(--success)" : "var(--warning)",
            boxShadow: officeReady
              ? "0 0 8px rgba(34,197,94,0.5)"
              : "0 0 8px rgba(245,158,11,0.5)",
          }}
        />
      </header>

      {/* ── Chat Messages ── */}
      <div className="flex-1 overflow-y-auto px-3 py-4 space-y-3">
        <AnimatePresence initial={false}>
          {messages.map((msg) => (
            <motion.div
              key={msg.id}
              initial={{ opacity: 0, y: 10, scale: 0.97 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              transition={{ duration: 0.25, ease: "easeOut" }}
              className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
            >
              <div
                className={`max-w-[85%] px-4 py-2.5 text-sm leading-relaxed whitespace-pre-wrap ${msg.role === "user" ? "chat-bubble-user" : "chat-bubble-agent"
                  }`}
              >
                {msg.content}
              </div>
            </motion.div>
          ))}
        </AnimatePresence>

        {isThinking && (
          <motion.div
            initial={{ opacity: 0, y: 6 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0 }}
            className="flex justify-start"
          >
            <div className="chat-bubble-agent px-4 py-3 flex items-center gap-1.5">
              <span className="typing-dot" />
              <span className="typing-dot" />
              <span className="typing-dot" />
            </div>
          </motion.div>
        )}

        <div ref={chatEndRef} />
      </div>

      {/* ── Quick Actions ── */}
      <div className="px-3 pb-2 flex flex-wrap gap-1.5">
        {QUICK_ACTIONS.map((action) => (
          <button
            key={action.label}
            onClick={() => sendMessage(action.prompt)}
            disabled={isThinking}
            className="text-xs px-3 py-1.5 rounded-full transition-all duration-200 cursor-pointer disabled:opacity-40"
            style={{
              background: "var(--bg-surface)",
              color: "var(--text-secondary)",
              border: "1px solid var(--border-subtle)",
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.background = "var(--accent-soft)";
              e.currentTarget.style.borderColor = "var(--border-accent)";
              e.currentTarget.style.color = "var(--text-primary)";
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.background = "var(--bg-surface)";
              e.currentTarget.style.borderColor = "var(--border-subtle)";
              e.currentTarget.style.color = "var(--text-secondary)";
            }}
          >
            {action.label}
          </button>
        ))}
      </div>

      {/* ── Input Bar ── */}
      <div className="px-3 pb-3 shrink-0">
        <div
          className="input-glow flex items-end gap-2 px-3 py-2 rounded-2xl transition-all duration-200"
          style={{
            background: "var(--bg-surface)",
            border: "1px solid var(--border-subtle)",
          }}
        >
          <textarea
            ref={inputRef}
            rows={1}
            value={input}
            onChange={handleInputChange}
            onKeyDown={handleKeyDown}
            placeholder="Tell me what to do..."
            disabled={isThinking}
            className="flex-1 bg-transparent border-none outline-none resize-none text-sm placeholder:text-[var(--text-muted)] disabled:opacity-50"
            style={{
              color: "var(--text-primary)",
              maxHeight: "120px",
            }}
          />
          <button
            onClick={() => sendMessage()}
            disabled={!input.trim() || isThinking}
            className="w-8 h-8 rounded-xl flex items-center justify-center shrink-0 transition-all duration-200 cursor-pointer disabled:opacity-30 disabled:cursor-not-allowed"
            style={{
              background: input.trim()
                ? "linear-gradient(135deg, var(--accent), #818cf8)"
                : "var(--bg-surface-hover)",
            }}
          >
            <svg
              width="16"
              height="16"
              viewBox="0 0 24 24"
              fill="none"
              stroke="white"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <line x1="22" y1="2" x2="11" y2="13" />
              <polygon points="22 2 15 22 11 13 2 9 22 2" />
            </svg>
          </button>
        </div>
        <p
          className="text-center mt-2 text-[10px]"
          style={{ color: "var(--text-muted)" }}
        >
          Phase 3 · SiliconFlow → Mistral → Gemini
        </p>
      </div>
    </div>
  );
}