"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import GradientText from "@/components/reactbits/GradientText";
import ShinyText from "@/components/reactbits/ShinyText";
import SpotlightCard from "@/components/reactbits/SpotlightCard";
import ClickSpark from "@/components/reactbits/ClickSpark";
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

/* ─── Quick Actions ─── */
const QUICK_ACTIONS = [
  { label: "📄 Get text", prompt: "Get the full document text" },
  { label: "🖍️ Highlight", prompt: "Highlight all text in yellow" },
  { label: "✍️ Insert", prompt: 'Insert a paragraph saying "Hello from AI Agent!"' },
  { label: "🗑️ Delete ¶1", prompt: "Delete the first paragraph" },
  { label: "🎨 Bold H1", prompt: "Make all Heading 1 paragraphs bold" },
];

/* ─── Execute AI Actions on the document ─── */
async function executeActions(actions: AIAction[]): Promise<string[]> {
  if (!isOfficeReady()) {
    return ["⚠️ Not connected to Word — actions cannot run."];
  }
  const results: string[] = [];
  for (const action of actions) {
    try {
      switch (action.type) {
        case "get_document_text": {
          const text = await getDocumentText();
          results.push(`📄 Document text:\n\`\`\`\n${text.slice(0, 500)}${text.length > 500 ? "\n…" : ""}\n\`\`\``);
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
          results.push(`✅ Kept only "${action.params.searchText}"`);
          break;
        case "delete_all_except_name":
          await keepOnlyParagraphsWithText(action.params.name);
          results.push(`✅ Kept only "${action.params.name}"`);
          break;
        case "highlight_all":
          await highlightAllText(action.params.color || "Yellow");
          results.push(`✅ Highlighted in ${action.params.color || "Yellow"}`);
          break;
        case "format_by_style":
          await formatByStyle(action.params.targetStyle, action.params.formatting);
          results.push(`✅ Formatted ${action.params.targetStyle}`);
          break;
        case "replace_text":
          await replaceText(action.params.find, action.params.replace);
          results.push(`✅ "${action.params.find}" → "${action.params.replace}"`);
          break;
        case "set_font":
          await setFont(action.params);
          results.push(`✅ Font updated`);
          break;
        default:
          results.push(`⚠️ Unknown: ${action.type}`);
      }
    } catch (err: any) {
      results.push(`❌ ${action.type}: ${err.message}`);
    }
  }
  return results;
}

/* ─── Animation Variants ─── */
const bubbleVariants = {
  initial: { opacity: 0, y: 12, scale: 0.97 },
  animate: {
    opacity: 1,
    y: 0,
    scale: 1,
    transition: { type: "spring" as const, stiffness: 420, damping: 32, mass: 0.8 },
  },
  exit: { opacity: 0, scale: 0.95, transition: { duration: 0.18 } },
};

const chipStagger = {
  initial: { opacity: 0, y: 6 },
  animate: (i: number) => ({
    opacity: 1,
    y: 0,
    transition: { delay: 0.3 + i * 0.05, duration: 0.3, ease: [0.16, 1, 0.3, 1] as const },
  }),
};

/* ─── Send Arrow Icon ─── */
function SendIcon() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <path d="M5 12h14M12 5l7 7-7 7" />
    </svg>
  );
}

/* ─── Paperclip Icon ─── */
function AttachIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" className="opacity-35">
      <path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48" />
    </svg>
  );
}

/* ═══════════════════════════════════════
   Component
   ═══════════════════════════════════════ */
export default function TaskpaneChat() {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: "welcome",
      role: "agent",
      content:
        "Hello! I'm your Office Agent — an AI assistant for Word. I can help you format, edit, highlight, find & replace, and manipulate your document. What would you like to do?",
      timestamp: new Date(),
    },
  ]);
  const [input, setInput] = useState("");
  const [isThinking, setIsThinking] = useState(false);
  const [officeReady, setOfficeReady] = useState(false);
  const chatEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    initializeOffice().then((ok) => setOfficeReady(ok));
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
        if (documentContext.startsWith("[Demo")) documentContext = undefined;
      } catch { /* ignore */ }

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
      return { message: `⚠️ ${err.message || "Couldn't reach the AI"}`, actions: [], error: err.message };
    }
  }, []);

  const sendMessage = useCallback(
    async (text?: string) => {
      const content = text || input.trim();
      if (!content || isThinking) return;

      setMessages((prev) => [...prev, { id: Date.now().toString(), role: "user", content, timestamp: new Date() }]);
      setInput("");
      if (inputRef.current) inputRef.current.style.height = "auto";
      setIsThinking(true);

      const aiResponse = await callAI(content);
      let actionResults: string[] = [];
      if (aiResponse.actions?.length) actionResults = await executeActions(aiResponse.actions);

      let agentContent = aiResponse.message || "Done!";
      if (actionResults.length) agentContent += "\n\n" + actionResults.join("\n");

      setMessages((prev) => [
        ...prev,
        {
          id: (Date.now() + 1).toString(),
          role: "agent",
          content: agentContent,
          timestamp: new Date(),
          status: aiResponse.error ? "error" : "success",
        },
      ]);
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

  /* ─── Render ─── */
  return (
    <div className="relative flex flex-col h-screen overflow-hidden" style={{ background: "var(--bg-base)" }}>
      {/* ── Ambient Orbs ── */}
      <div className="ambient-bg">
        <div className="ambient-orb ambient-orb--lavender" />
        <div className="ambient-orb ambient-orb--purple" />
        <div className="ambient-orb ambient-orb--peach" />
      </div>

      {/* ═══════ Header ═══════ */}
      <header className="taskpane-header relative z-10 flex items-center gap-3 px-4 py-3 shrink-0">
        <div className="logo-mark">⚡</div>

        <div className="flex-1 min-w-0">
          <GradientText
            className="!mx-0"
            colors={["#7B8CFF", "#C8B6FF", "#7BC8FF", "#7B8CFF"]}
            animationSpeed={7}
          >
            <span className="text-[14px] font-semibold tracking-tight">Office Agent</span>
          </GradientText>

          <div className="flex items-center gap-1.5 mt-0.5">
            <div className={`status-dot ${officeReady ? "status-dot--connected" : "status-dot--demo"}`} />
            <span
              className="text-[10px] font-medium"
              style={{ color: "var(--text-muted)" }}
            >
              {officeReady ? "Connected to Word" : "Demo mode"}
            </span>
          </div>
        </div>
      </header>

      <hr className="divider-soft relative z-10" />

      {/* ═══════ Chat Area ═══════ */}
      <div className="relative z-10 flex-1 overflow-y-auto px-4 py-5 space-y-4">
        <AnimatePresence initial={false}>
          {messages.map((msg) => (
            <motion.div
              key={msg.id}
              variants={bubbleVariants}
              initial="initial"
              animate="animate"
              exit="exit"
              className={`flex ${msg.role === "user" ? "justify-end" : "justify-start"}`}
            >
              {msg.role === "user" ? (
                /* ── User Bubble ── */
                <div className="bubble-user max-w-[82%] px-4 py-3 text-[13px] leading-relaxed whitespace-pre-wrap font-medium">
                  {msg.content}
                </div>
              ) : (
                /* ── Agent Card ── */
                <div className="flex items-start gap-2.5 max-w-[88%]">
                  <div className="agent-avatar mt-0.5">⚡</div>
                  <SpotlightCard
                    className="bubble-agent flex-1 px-4 py-3 text-[13px] leading-relaxed whitespace-pre-wrap"
                    spotlightColor="rgba(123, 140, 255, 0.06)"
                    spotlightSize={85}
                  >
                    {msg.content}
                  </SpotlightCard>
                </div>
              )}
            </motion.div>
          ))}
        </AnimatePresence>

        {/* ── Thinking Indicator ── */}
        {isThinking && (
          <motion.div
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0 }}
            className="flex items-start gap-2.5"
          >
            <div className="agent-avatar">⚡</div>
            <div className="bubble-agent px-5 py-3.5 flex items-center gap-2">
              <span className="typing-dot" />
              <span className="typing-dot" />
              <span className="typing-dot" />
            </div>
          </motion.div>
        )}

        <div ref={chatEndRef} />
      </div>

      {/* ═══════ Quick Actions ═══════ */}
      <div className="relative z-10 px-4 pb-2.5 flex flex-wrap gap-1.5">
        {QUICK_ACTIONS.map((action, i) => (
          <motion.button
            key={action.label}
            variants={chipStagger}
            initial="initial"
            animate="animate"
            custom={i}
            onClick={() => sendMessage(action.prompt)}
            disabled={isThinking}
            className="action-chip"
          >
            {action.label}
          </motion.button>
        ))}
      </div>

      {/* ═══════ Floating Input Dock ═══════ */}
      <div className="relative z-10 px-3 pb-3 shrink-0">
        <ClickSpark sparkColor="#7B8CFF" sparkRadius={22} sparkCount={10} duration={450}>
          <div className="input-dock-wrap">
            <div className="input-dock relative z-[1] flex items-end gap-2 px-4 py-2.5">
              {/* Attach hint */}
              <button
                className="shrink-0 w-7 h-7 flex items-center justify-center rounded-lg hover:bg-[var(--bg-inset)] transition-colors cursor-pointer"
                title="Attach file"
                type="button"
              >
                <AttachIcon />
              </button>

              <textarea
                ref={inputRef}
                rows={1}
                value={input}
                onChange={handleInputChange}
                onKeyDown={handleKeyDown}
                placeholder="Ask me anything…"
                disabled={isThinking}
                className="flex-1 bg-transparent border-none outline-none resize-none text-[13px] font-medium disabled:opacity-40"
                style={{
                  color: "var(--text-primary)",
                  maxHeight: "120px",
                  caretColor: "var(--accent)",
                }}
              />

              {/* Send */}
              <button
                onClick={() => sendMessage()}
                disabled={!input.trim() || isThinking}
                className="send-btn w-8 h-8 flex items-center justify-center shrink-0"
              >
                <SendIcon />
              </button>
            </div>
          </div>
        </ClickSpark>

        {/* Footer */}
        <div className="flex items-center justify-center mt-2.5 mb-0.5">
          <ShinyText
            text="Office Agent AI · Powered by Gemini"
            speed={5}
            color="#C8C8D0"
            shineColor="#AEAEB8"
            className="text-[9px] font-semibold tracking-[0.08em] uppercase"
          />
        </div>
      </div>
    </div>
  );
}