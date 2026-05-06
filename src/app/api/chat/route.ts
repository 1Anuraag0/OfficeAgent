import { GoogleGenAI } from "@google/genai";

/* ─── System Prompt — the brain of the agent ─── */
const SYSTEM_PROMPT = `You are Office Agent AI — an expert assistant that helps users manipulate Microsoft Word documents using natural language.

You MUST respond with a valid JSON object. No markdown, no explanation outside the JSON.

## Response Schema

{
  "message": "A friendly human-readable summary of what you did or will do",
  "actions": [
    {
      "type": "action_type",
      "params": { ... }
    }
  ]
}

## Available Actions

1. **get_document_text** — Read the document. params: {}
2. **insert_text** — Insert text at document end. params: { "text": "string" }
3. **delete_paragraphs** — Delete paragraphs by index (0-based). params: { "startIndex": number, "endIndex": number }
4. **delete_pages** — Delete a range of pages (1-based). params: { "startPage": number, "endPage": number }
5. **keep_only_pages_with_text** — Delete all content EXCEPT pages/sections containing a specific name or string. params: { "searchText": "string" }
6. **highlight_all** — Highlight all text. params: { "color": "Yellow" | "Green" | "Cyan" | "Pink" | "Red" }
7. **format_by_style** — Format paragraphs matching a Word style. params: { "targetStyle": "Heading1" | "Heading2" | "Normal", "formatting": { "bold": boolean?, "italic": boolean?, "color": string?, "size": number? } }
8. **replace_text** — Find and replace text. params: { "find": "string", "replace": "string" }
9. **set_font** — Set font properties on entire document body. params: { "name": string?, "size": number?, "color": string?, "bold": boolean?, "italic": boolean? }
10. **delete_all_except_name** — Keep only paragraphs/sections where a person's name appears. params: { "name": "string" }

## Rules

- If the user says "delete pages 1 to 126" or similar, use **delete_pages** with startPage and endPage.
- If the user says "delete all pages except where name = X" or "keep only X's section", use **keep_only_pages_with_text** with the name/text to keep.
- If the user says "delete all paragraphs except those containing X", use **delete_all_except_name**.
- If the user's request is ambiguous, ask a clarifying question in the "message" field and return an empty "actions" array.
- If the user asks something unrelated to Word documents, politely redirect them.
- Always be concise and friendly.
- You can chain multiple actions in a single response.`;

/* ─── Provider Definitions ─── */
type ProviderResult = { text: string; provider: string };

async function callSiliconFlow(userPrompt: string): Promise<ProviderResult> {
  const key = process.env.SILICONFLOW_API_KEY;
  if (!key || key === "YOUR_SILICONFLOW_KEY_HERE") {
    throw new Error("SiliconFlow API key not configured");
  }

  const res = await fetch("https://api.siliconflow.cn/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "MiniMaxAI/MiniMax-M2.1",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userPrompt },
      ],
      temperature: 0.3,
      max_tokens: 1024,
      response_format: { type: "json_object" },
    }),
  });

  if (!res.ok) {
    const err = await res.text().catch(() => "");
    throw new Error(`SiliconFlow ${res.status}: ${err.slice(0, 200)}`);
  }

  const data = await res.json();
  const text = data.choices?.[0]?.message?.content ?? "";
  if (!text) throw new Error("SiliconFlow returned empty response");
  return { text, provider: "SiliconFlow" };
}

async function callMistral(userPrompt: string): Promise<ProviderResult> {
  const key = process.env.MISTRAL_API_KEY;
  if (!key || key === "YOUR_MISTRAL_KEY_HERE") {
    throw new Error("Mistral API key not configured");
  }

  const res = await fetch("https://api.mistral.ai/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${key}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "mistral-small-latest",
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: userPrompt },
      ],
      temperature: 0.3,
      max_tokens: 1024,
      response_format: { type: "json_object" },
    }),
  });

  if (!res.ok) {
    const err = await res.text().catch(() => "");
    throw new Error(`Mistral ${res.status}: ${err.slice(0, 200)}`);
  }

  const data = await res.json();
  const text = data.choices?.[0]?.message?.content ?? "";
  if (!text) throw new Error("Mistral returned empty response");
  return { text, provider: "Mistral" };
}

async function callGemini(userPrompt: string): Promise<ProviderResult> {
  const key = process.env.GEMINI_API_KEY;
  if (!key || key === "YOUR_API_KEY_HERE") {
    throw new Error("Gemini API key not configured");
  }

  const ai = new GoogleGenAI({ apiKey: key });
  const MODELS = ["gemini-2.0-flash-lite", "gemini-2.5-flash-preview-05-20", "gemini-2.0-flash"];
  let lastError: any = null;

  for (const model of MODELS) {
    try {
      const response = await ai.models.generateContent({
        model,
        contents: userPrompt,
        config: {
          systemInstruction: SYSTEM_PROMPT,
          temperature: 0.3,
          maxOutputTokens: 1024,
          responseMimeType: "application/json",
        },
      });

      const text = response.text ?? "";
      if (!text) throw new Error("Gemini returned empty response");
      return { text, provider: `Gemini (${model})` };
    } catch (err: any) {
      lastError = err;
      console.warn(`[API /chat] Gemini ${model} failed:`, err.message);
      continue;
    }
  }

  throw lastError || new Error("All Gemini models failed");
}

/* ─── Provider Cascade ─── */
const PROVIDERS: Array<{
  name: string;
  call: (prompt: string) => Promise<ProviderResult>;
}> = [
    { name: "SiliconFlow", call: callSiliconFlow },
    { name: "Mistral", call: callMistral },
    { name: "Gemini", call: callGemini },
  ];

/* ─── Route Handler ─── */
export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { message, documentContext } = body as {
      message: string;
      documentContext?: string;
    };

    if (!message) {
      return Response.json({ error: "Message is required" }, { status: 400 });
    }

    let userPrompt = message;
    if (documentContext) {
      userPrompt = `[Current document content (first 2000 chars)]:\n${documentContext.slice(0, 2000)}\n\n[User request]: ${message}`;
    }

    let result: ProviderResult | null = null;
    let lastError: any = null;

    for (const provider of PROVIDERS) {
      try {
        console.log(`[API /chat] Trying ${provider.name}...`);
        result = await provider.call(userPrompt);
        console.log(`[API /chat] ✅ Success via ${result.provider}`);
        break;
      } catch (err: any) {
        lastError = err;
        console.warn(`[API /chat] ❌ ${provider.name} failed:`, err.message);
        continue;
      }
    }

    if (!result) {
      throw lastError || new Error("All AI providers failed");
    }

    let parsed;
    try {
      parsed = JSON.parse(result.text);
    } catch {
      parsed = {
        message: result.text || "I had trouble processing that. Could you rephrase?",
        actions: [],
      };
    }

    parsed._provider = result.provider;

    return Response.json(parsed, {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
      },
    });
  } catch (error: any) {
    console.error("[API /chat] Error:", error);
    return Response.json(
      {
        message: `⚠️ AI Error: ${error.message || "Something went wrong"}`,
        actions: [],
      },
      { status: 500 }
    );
  }
}

export async function OPTIONS() {
  return new Response(null, {
    status: 204,
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    },
  });
}