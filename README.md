<div align="center">

# ⚡ Office Agent AI

**AI-powered Word document manipulation through natural language.**

*Delete pages, format text, find & replace, highlight — all from a chat interface inside Microsoft Word.*

[![Next.js](https://img.shields.io/badge/Next.js-16-black?style=flat-square&logo=next.js)](https://nextjs.org)
[![React](https://img.shields.io/badge/React-19-61DAFB?style=flat-square&logo=react)](https://react.dev)
[![TypeScript](https://img.shields.io/badge/TypeScript-5-3178C6?style=flat-square&logo=typescript)](https://typescriptlang.org)
[![Office.js](https://img.shields.io/badge/Office.js-Word_Add--in-D83B01?style=flat-square&logo=microsoft-office)](https://learn.microsoft.com/office/dev/add-ins)
[![Tailwind](https://img.shields.io/badge/Tailwind_CSS-4-06B6D4?style=flat-square&logo=tailwindcss)](https://tailwindcss.com)

</div>

---

## ✨ What is this?

Office Agent AI is a **Microsoft Word add-in** that lets you control your documents using plain English. Instead of navigating menus and ribbons, just type what you want:

> *"Delete pages 5 through 20"*
> *"Highlight all text in yellow"*
> *"Replace all instances of 'draft' with 'final'"*
> *"Make all Heading 1 paragraphs bold"*

The AI understands your intent, maps it to document operations, and executes them instantly — all within a beautiful taskpane that lives inside Word.

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────┐
│                    Microsoft Word                       │
│  ┌───────────────────────────────────────────────────┐  │
│  │              Office Agent Taskpane                │  │
│  │                                                   │  │
│  │   ┌─────────┐    ┌──────────┐    ┌────────────┐  │  │
│  │   │  React  │───▶│ Next.js  │───▶│  AI API    │  │  │
│  │   │  Chat   │    │  Server  │    │  Cascade   │  │  │
│  │   │   UI    │◀───│  (API)   │◀───│            │  │  │
│  │   └────┬────┘    └──────────┘    └────────────┘  │  │
│  │        │                          SiliconFlow     │  │
│  │        │                          → Mistral       │  │
│  │        ▼                          → Gemini        │  │
│  │   ┌─────────┐                                     │  │
│  │   │Office.js│──── Document Manipulation ──────▶   │  │
│  │   └─────────┘                                     │  │
│  └───────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────┘
```

---

## 🎨 Design

The taskpane UI is built with a **2026 premium aesthetic** inspired by Apple, Linear, and Notion AI:

- **Warm white + lavender palette** — calm, intelligent, no harsh contrasts
- **Soft glassmorphism** — frosted surfaces with layered depth
- **Ambient gradient orbs** — subtle floating lavender/periwinkle background
- **Spring-physics animations** — every message bounces in naturally
- **Animated gradient text** — the header title flows with color
- **Mouse-tracking spotlight** — agent messages glow where you hover
- **Click spark particles** — satisfying micro-interaction on send
- **Floating dock input** — pill-shaped input bar with animated gradient border

Built with [**ReactBits**](https://reactbits.dev) components (`GradientText`, `ShinyText`, `SpotlightCard`, `ClickSpark`) + [Framer Motion](https://motion.dev) + [Tailwind CSS v4](https://tailwindcss.com).

---

## 🧠 AI Provider Cascade

The API uses a **fallback cascade** for maximum reliability:

| Priority | Provider | Model |
|----------|----------|-------|
| 1st | **SiliconFlow** | MiniMax-M2.1 |
| 2nd | **Mistral** | mistral-small-latest |
| 3rd | **Gemini** | gemini-2.0-flash-lite → 2.5-flash → 2.0-flash |

If the first provider hits a rate limit or fails, it automatically falls through to the next. Gemini itself has an internal 3-model fallback.

---

## 📋 Supported Actions

| Action | What it does | Example prompt |
|--------|-------------|----------------|
| `get_document_text` | Read the full document | *"Get the document text"* |
| `insert_text` | Append a paragraph | *"Add a paragraph saying 'Hello World'"* |
| `delete_paragraphs` | Delete by paragraph index | *"Delete the first 3 paragraphs"* |
| `delete_pages` | Delete a page range | *"Delete pages 2 through 10"* |
| `keep_only_pages_with_text` | Keep only matching sections | *"Keep only the section about John"* |
| `highlight_all` | Highlight all text | *"Highlight everything in cyan"* |
| `format_by_style` | Format paragraphs by Word style | *"Make all Heading 1 bold and red"* |
| `replace_text` | Find & replace | *"Replace 'draft' with 'final'"* |
| `set_font` | Set font properties globally | *"Set font to Arial 12pt"* |

---

## 🚀 Getting Started

### Prerequisites

- **Node.js** 18+
- **Microsoft Word** (desktop or web)
- At least one AI API key (SiliconFlow, Mistral, or Gemini)

### 1. Clone & Install

```bash
git clone https://github.com/your-username/office-agent-ai.git
cd office-agent-ai
npm install
```

### 2. Configure API Keys

Create a `.env.local` file in the project root:

```env
# Provider API Keys (priority order: SiliconFlow → Mistral → Gemini)

# 1. SiliconFlow — https://cloud.siliconflow.cn
SILICONFLOW_API_KEY=your_key_here

# 2. Mistral — https://console.mistral.ai
MISTRAL_API_KEY=your_key_here

# 3. Gemini — https://aistudio.google.com/apikey
GEMINI_API_KEY=your_key_here
```

> **Tip:** You only need *one* key to get started. The cascade will skip unconfigured providers.

### 3. Generate HTTPS Certificates

Office add-ins require HTTPS. Generate local certs:

```bash
mkdir certificates
npx office-addin-dev-certs install --machine
```

Then copy the generated `.pem` files to the `certificates/` directory as `localhost-key.pem` and `localhost.pem`.

### 4. Start Development Server

```bash
npm run dev
```

The server starts at **https://localhost:3000**. The taskpane is at `/taskpane`.

### 5. Sideload into Word

#### Word Online (easiest)
1. Open [Word Online](https://www.office.com/launch/word)
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Upload the `manifest.xml` file from the project root
4. The "Office Agent AI" button appears in the Home tab

#### Word Desktop
1. Open Word → **Insert → Get Add-ins → Upload My Add-in**
2. Upload `manifest.xml`
3. Click **Open Agent** in the Home ribbon

---

## 📁 Project Structure

```
msofficeagent/
├── manifest.xml                    # Office add-in manifest
├── certificates/                   # HTTPS certs for dev
├── src/
│   ├── app/
│   │   ├── layout.tsx              # Root layout (Inter + JetBrains Mono)
│   │   ├── globals.css             # 2026 design system
│   │   ├── page.tsx                # Root redirect → /taskpane
│   │   ├── api/chat/route.ts       # AI API (provider cascade)
│   │   └── taskpane/
│   │       ├── page.tsx            # Taskpane route
│   │       └── TaskpaneChat.tsx    # Main chat UI
│   ├── lib/
│   │   └── office.ts              # Office.js document operations
│   └── components/
│       └── reactbits/             # ReactBits UI components
│           ├── GradientText.tsx
│           ├── ShinyText.tsx
│           ├── SpotlightCard.tsx
│           └── ClickSpark.tsx
└── public/assets/                 # Add-in icons (16/32/80px)
```

---

## 🛠️ Tech Stack

| Layer | Technology |
|-------|-----------|
| **Framework** | Next.js 16 (App Router + Turbopack) |
| **UI** | React 19, Framer Motion 12, ReactBits |
| **Styling** | Tailwind CSS 4, PostCSS |
| **Language** | TypeScript 5 |
| **AI** | SiliconFlow, Mistral, Google Gemini |
| **Office** | Office.js (Word Add-in API) |
| **Compiler** | React Compiler (babel-plugin) |

---

## 📝 Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start dev server with HTTPS |
| `npm run build` | Production build |
| `npm run start` | Start production server |
| `npm run lint` | Run ESLint |

---

## ⚠️ Important Notes

- **HTTPS is required** — Office add-ins only load over HTTPS, even in development
- **Demo mode** — When opened outside Word (e.g., in a browser), the app runs in demo mode with mock data
- **Rate limits** — The AI cascade handles rate limits automatically by falling through to the next provider
- **Word Online vs Desktop** — Some operations (like page-level deletion) use heuristic paragraph counting since Word Online doesn't expose page metadata

---

<div align="center">

**Built with ❤️ and AI**

</div>
