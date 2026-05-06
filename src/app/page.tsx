"use client";

import { motion } from "framer-motion";
import Link from "next/link";

export default function Home() {
  return (
    <main
      className="flex flex-col items-center justify-center min-h-screen px-6 text-center"
      style={{ background: "var(--bg-primary)" }}
    >
      {/* Ambient glow */}
      <div
        className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[500px] h-[500px] rounded-full pointer-events-none"
        style={{
          background: "radial-gradient(circle, rgba(99,102,241,0.08) 0%, transparent 70%)",
        }}
      />

      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.6, ease: "easeOut" }}
        className="relative z-10"
      >
        {/* Icon */}
        <div
          className="w-20 h-20 rounded-3xl mx-auto mb-8 flex items-center justify-center text-4xl animate-pulse-glow"
          style={{
            background: "linear-gradient(135deg, var(--accent), #818cf8)",
          }}
        >
          ⚡
        </div>

        <h1
          className="text-4xl font-bold mb-4 tracking-tight"
          style={{ color: "var(--text-primary)" }}
        >
          Office Agent{" "}
          <span
            style={{
              background: "linear-gradient(135deg, var(--accent), #a78bfa)",
              WebkitBackgroundClip: "text",
              WebkitTextFillColor: "transparent",
            }}
          >
            AI
          </span>
        </h1>

        <p
          className="text-lg mb-8 max-w-md mx-auto leading-relaxed"
          style={{ color: "var(--text-secondary)" }}
        >
          Manipulate Word documents using natural language.
          Delete pages, format text, highlight content — all from a chat interface.
        </p>

        <div className="flex gap-4 justify-center flex-wrap">
          <Link href="/taskpane">
            <motion.button
              whileHover={{ scale: 1.04 }}
              whileTap={{ scale: 0.97 }}
              className="px-6 py-3 rounded-xl font-semibold text-sm text-white cursor-pointer transition-shadow"
              style={{
                background: "linear-gradient(135deg, var(--accent), #818cf8)",
                boxShadow: "var(--shadow-glow)",
              } as React.CSSProperties}
            >
              Open Taskpane Demo →
            </motion.button>
          </Link>

          <a
            href="https://github.com"
            target="_blank"
            rel="noopener noreferrer"
          >
            <motion.button
              whileHover={{ scale: 1.04 }}
              whileTap={{ scale: 0.97 }}
              className="px-6 py-3 rounded-xl font-semibold text-sm cursor-pointer transition-all"
              style={{
                background: "var(--bg-surface)",
                color: "var(--text-secondary)",
                border: "1px solid var(--border-subtle)",
              } as React.CSSProperties}
            >
              GitHub
            </motion.button>
          </a>
        </div>
      </motion.div>

      {/* Architecture chips */}
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ delay: 0.4, duration: 0.6 }}
        className="relative z-10 mt-16 flex flex-wrap gap-3 justify-center"
      >
        {["Next.js", "Office.js", "Cloudflare Workers", "Groq / Gemini", "Upstash Redis"].map(
          (tech) => (
            <span
              key={tech}
              className="text-xs px-3 py-1.5 rounded-full"
              style={{
                background: "var(--accent-soft)",
                color: "var(--accent)",
                border: "1px solid var(--border-accent)",
              }}
            >
              {tech}
            </span>
          )
        )}
      </motion.div>
    </main>
  );
}
