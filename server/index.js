import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import express from "express";
import multer from "multer";
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });
const PORT = Number(process.env.PORT || 3117);
const CONFIG_PATH = path.resolve("app.config.json");
const LOCAL_CONFIG_PATH = path.resolve("app.config.local.json");
const CACHE_DIR = path.resolve(".cache", "llm");
const jobs = new Map();

app.use(express.json({ limit: "2mb" }));

app.get("/api/health", async (_req, res) => {
  const config = await loadConfig();
  res.json({ ok: true, port: PORT, provider: config.provider });
});

app.get("/api/config", async (_req, res) => {
  try {
    const config = await loadConfig();
    res.json({ ok: true, config: publicConfig(config) });
  } catch (error) {
    res.status(500).json({ ok: false, error: error.message });
  }
});

app.get("/api/provider/models", async (_req, res) => {
  try {
    const config = await loadConfig();
    if (config.provider !== "lmstudio") {
      return res.json({ ok: true, provider: config.provider, models: [{ id: activeProvider(config).model }] });
    }
    const baseUrl = normalizeOpenAiBaseUrl(activeProvider(config).baseUrl);
    try {
      const response = await fetch(`${baseUrl}/models`);
      if (!response.ok) throw new Error(`LM Studio returned ${response.status}`);
      const data = await response.json();
      return res.json({ ok: true, provider: config.provider, baseUrl, models: data.data || [] });
    } catch (error) {
      return res.json({
        ok: true,
        provider: config.provider,
        baseUrl,
        models: [{ id: activeProvider(config).model }],
        warning: `Could not list LM Studio models: ${error.message}`
      });
    }
  } catch (error) {
    res.status(502).json({ ok: false, error: error.message });
  }
});

app.post("/api/summarize", upload.single("transcript"), async (req, res) => {
  const id = crypto.randomUUID();
  const config = await loadConfig();
  const options = parseOptions(req.body, config);
  const file = req.file;
  if (!file) return res.status(400).json({ error: "Transcript file is required." });

  jobs.set(id, {
    id,
    status: "queued",
    progress: 0,
    message: "Queued",
    createdAt: new Date().toISOString(),
    result: null,
    error: null
  });

  runJob(id, file, options).catch((error) => {
    updateJob(id, { status: "failed", error: error.message, message: "Failed" });
  });

  res.json({ id });
});

app.post("/api/analyze-chunks", upload.single("transcript"), async (req, res) => {
  try {
    const config = await loadConfig();
    const options = parseOptions(req.body, config);
    const file = req.file;
    if (!file) return res.status(400).json({ error: "Transcript file is required." });
    const transcriptText = decodeText(file.buffer);
    const cues = parseTranscript(transcriptText);
    const speakers = [...new Set(cues.map((cue) => cue.speaker).filter(Boolean))].sort();
    const filtered = filterCues(cues, options.speaker);
    if (!filtered.length) {
      return res.status(400).json({
        error: `No transcript lines found for speaker "${options.speaker}".`,
        speakers
      });
    }
    const chunks = chunkCues(filtered, options.chunkMinutes, options.maxChunkChars);
    const chunkStats = chunks.map((chunk, index) => ({
      index: index + 1,
      start: chunk.start,
      end: chunk.end,
      chars: chunk.text.length,
      cueCount: chunk.cueCount
    }));
    res.json({
      ok: true,
      sourceName: file.originalname,
      speakers,
      cueCount: filtered.length,
      chunkCount: chunks.length,
      maxChunkChars: options.maxChunkChars,
      largestChunkChars: Math.max(...chunkStats.map((chunk) => chunk.chars)),
      chunkStats
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ error: "Job not found." });
  res.json(job);
});

app.get("/api/jobs/:id/download/:format", async (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job?.result) return res.status(404).json({ error: "Result not found." });
  const name = safeFileName(job.result.title || "transcript-summary");
  if (req.params.format === "md") {
    res.setHeader("Content-Type", "text/markdown; charset=utf-8");
    res.setHeader("Content-Disposition", `attachment; filename="${name}.md"`);
    return res.send(job.result.markdown);
  }
  if (req.params.format === "docx") {
    const buffer = await buildDocx(job.result);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${name}.docx"`);
    return res.send(buffer);
  }
  res.status(400).json({ error: "Unsupported format." });
});

const distDir = path.resolve("dist");
app.use(express.static(distDir));
app.get(/.*/, async (_req, res) => {
  res.sendFile(path.join(distDir, "index.html"));
});

async function runJob(id, file, options) {
  const config = await loadConfig();
  const provider = activeProvider(config);
  updateJob(id, { status: "running", progress: 4, message: "Parsing transcript" });
  const transcriptText = decodeText(file.buffer);
  const cues = parseTranscript(transcriptText);
  const speakers = [...new Set(cues.map((cue) => cue.speaker).filter(Boolean))].sort();
  const filtered = filterCues(cues, options.speaker);
  if (!filtered.length) {
    throw new Error(`No transcript lines found for speaker "${options.speaker}". Detected speakers: ${speakers.join(", ") || "none"}`);
  }

  const chunks = chunkCues(filtered, options.chunkMinutes, options.maxChunkChars);
  const chunkStats = chunks.map((chunk, index) => ({
    index: index + 1,
    start: chunk.start,
    end: chunk.end,
    chars: chunk.text.length,
    cueCount: chunk.cueCount
  }));
  updateJob(id, { chunkStats, message: `Created ${chunks.length} chunks` });
  const chunkSummaries = [];

  for (let index = 0; index < chunks.length; index += 1) {
    updateJob(id, {
      progress: Math.round(8 + (index / chunks.length) * 62),
      message: `Summarizing chunk ${index + 1} of ${chunks.length} (${chunks[index].text.length} chars)`
    });
    chunkSummaries.push(await summarizeChunk({ chunk: chunks[index], index, total: chunks.length, config, provider, detail: options.detail }));
  }

  updateJob(id, { progress: 76, message: "Merging topics and highlights" });
  const merged = await mergeSummaries({ chunkSummaries, config, provider, detail: options.detail, speaker: options.speaker });
  const result = normalizeResult(merged, {
    title: `Session Summary - ${file.originalname}`,
    sourceName: file.originalname,
    speaker: options.speaker,
    speakers,
    cueCount: filtered.length,
    chunkCount: chunks.length,
    provider: config.provider,
    model: provider.model
  });
  result.markdown = renderMarkdown(result);

  updateJob(id, { status: "completed", progress: 100, message: "Completed", result });
}

function parseOptions(body, config) {
  const defaults = config.defaults || {};
  return {
    speaker: String(body.speaker || defaults.speaker || "").trim(),
    detail: String(body.detail || defaults.detail || "medium"),
    chunkMinutes: clamp(Number(body.chunkMinutes || defaults.chunkMinutes || 10), 2, 60),
    maxChunkChars: clamp(Number(body.maxChunkChars || defaults.maxChunkChars || 9000), 2500, 20000)
  };
}

function normalizeOpenAiBaseUrl(raw) {
  const cleaned = String(raw || "http://127.0.0.1:11435").replace(/\/+$/, "");
  return cleaned.endsWith("/v1") ? cleaned : `${cleaned}/v1`;
}

function normalizeGeminiBaseUrl(raw) {
  return String(raw || "https://generativelanguage.googleapis.com/v1beta").replace(/\/+$/, "");
}

async function loadConfig() {
  const base = await readJson(CONFIG_PATH, null);
  if (!base) throw new Error(`Missing config file: ${CONFIG_PATH}`);
  const local = await readJson(LOCAL_CONFIG_PATH, {});
  const config = mergeDeep(base, local);
  if (!["lmstudio", "gemini"].includes(config.provider)) {
    throw new Error(`Unsupported provider "${config.provider}". Use "lmstudio" or "gemini".`);
  }
  activeProvider(config);
  return config;
}

async function readJson(filePath, fallback) {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return fallback;
    throw new Error(`Could not read ${path.basename(filePath)}: ${error.message}`);
  }
}

function mergeDeep(base, override) {
  if (!override || typeof override !== "object") return base;
  const output = Array.isArray(base) ? [...base] : { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value && typeof value === "object" && !Array.isArray(value)) {
      output[key] = mergeDeep(output[key] || {}, value);
    } else {
      output[key] = value;
    }
  }
  return output;
}

function activeProvider(config) {
  const provider = config.providers?.[config.provider];
  if (!provider) throw new Error(`Provider "${config.provider}" is not configured.`);
  if (!provider.model) throw new Error(`Provider "${config.provider}" is missing a model.`);
  return provider;
}

function publicConfig(config) {
  return {
    provider: config.provider,
    defaults: config.defaults || {},
    activeProvider: {
      ...activeProvider(config),
      apiKey: activeProvider(config).apiKey ? "configured" : "",
      resolvedApiKey: config.provider === "gemini" && Boolean(resolveGeminiApiKey(activeProvider(config)))
    }
  };
}

function resolveGeminiApiKey(provider) {
  return provider.apiKey || process.env[provider.apiKeyEnv || "GEMINI_API_KEY"] || "";
}

function decodeText(buffer) {
  const text = buffer.toString("utf8");
  return text.replace(/\uFEFF/g, "").replace(/â€¦/g, "...").replace(/â€™/g, "'").replace(/â€œ|â€/g, '"');
}

function parseTranscript(text) {
  if (/WEBVTT/i.test(text) || /\d{2}:\d{2}:\d{2}[.,]\d{3}\s*-->/m.test(text)) return parseTimedTranscript(text);
  return parsePlainTranscript(text);
}

function parseTimedTranscript(text) {
  const lines = text.split(/\r?\n/);
  const cues = [];
  for (let i = 0; i < lines.length; i += 1) {
    const match = lines[i].match(/(\d{2}:\d{2}:\d{2})[.,]\d{3}\s*-->\s*(\d{2}:\d{2}:\d{2})[.,]\d{3}/);
    if (!match) continue;
    const start = match[1];
    const end = match[2];
    const textLines = [];
    i += 1;
    while (i < lines.length && lines[i].trim() !== "") {
      const line = lines[i].trim();
      if (!/^\d+$/.test(line)) textLines.push(line);
      i += 1;
    }
    const raw = textLines.join(" ").trim();
    if (!raw) continue;
    const speakerMatch = raw.match(/^([^:]{2,80}):\s*(.+)$/);
    cues.push({
      start,
      end,
      speaker: speakerMatch ? speakerMatch[1].trim() : "",
      text: speakerMatch ? speakerMatch[2].trim() : raw
    });
  }
  return cues;
}

function parsePlainTranscript(text) {
  const cues = [];
  const lines = text.split(/\r?\n/);
  let currentTime = "00:00:00";
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const timeMatch = trimmed.match(/(\d{2}:\d{2}:\d{2})/);
    if (timeMatch) currentTime = timeMatch[1];
    const speakerMatch = trimmed.match(/^(?:\[\d{2}:\d{2}:\d{2}\]\s*)?([^:]{2,80}):\s*(.+)$/);
    if (speakerMatch) {
      cues.push({ start: currentTime, end: currentTime, speaker: speakerMatch[1].trim(), text: speakerMatch[2].trim() });
    }
  }
  return mergeAdjacentCues(cues);
}

function mergeAdjacentCues(cues) {
  const merged = [];
  for (const cue of cues) {
    const previous = merged[merged.length - 1];
    if (previous && previous.speaker === cue.speaker && seconds(cue.start) - seconds(previous.end) <= 12) {
      previous.end = cue.end;
      previous.text = `${previous.text} ${cue.text}`.trim();
    } else {
      merged.push({ ...cue });
    }
  }
  return merged;
}

function filterCues(cues, speaker) {
  if (!speaker) return cues;
  const target = speaker.toLowerCase();
  return cues.filter((cue) => cue.speaker.toLowerCase().includes(target));
}

function chunkCues(cues, chunkMinutes, maxChars) {
  const segments = splitOversizedCues(cues, Math.max(1000, maxChars - 500));
  const chunks = [];
  let current = [];
  let startSeconds = seconds(segments[0].start);
  let chars = 0;
  for (const cue of segments) {
    const cueTextLength = formatCue(cue).length;
    const tooLong = seconds(cue.start) - startSeconds >= chunkMinutes * 60;
    const tooManyChars = chars + cueTextLength > maxChars;
    if (current.length && (tooLong || tooManyChars)) {
      chunks.push(buildChunk(current));
      current = [];
      startSeconds = seconds(cue.start);
      chars = 0;
    }
    current.push(cue);
    chars += cueTextLength;
  }
  if (current.length) chunks.push(buildChunk(current));
  return chunks;
}

function buildChunk(cues) {
  return {
    start: cues[0].start,
    end: cues[cues.length - 1].end,
    cueCount: cues.length,
    text: cues.map(formatCue).join("\n")
  };
}

function formatCue(cue) {
  return `[${cue.start}-${cue.end}] ${cue.text}`;
}

function splitOversizedCues(cues, maxCueChars) {
  return cues.flatMap((cue) => {
    if (formatCue(cue).length <= maxCueChars) return [cue];
    const pieces = splitText(cue.text, maxCueChars - 32);
    return pieces.map((text, index) => ({
      ...cue,
      text: pieces.length > 1 ? `${text} (continued part ${index + 1}/${pieces.length})` : text
    }));
  });
}

function splitText(text, maxChars) {
  const sentences = text.match(/[^.!?]+[.!?]+|[^.!?]+$/g) || [text];
  const pieces = [];
  let current = "";
  for (const sentence of sentences.map((item) => item.trim()).filter(Boolean)) {
    if ((current.length + sentence.length + 1) <= maxChars) {
      current = `${current} ${sentence}`.trim();
      continue;
    }
    if (current) pieces.push(current);
    if (sentence.length <= maxChars) {
      current = sentence;
    } else {
      const words = sentence.split(/\s+/);
      current = "";
      for (const word of words) {
        if ((current.length + word.length + 1) > maxChars && current) {
          pieces.push(current);
          current = "";
        }
        current = `${current} ${word}`.trim();
      }
    }
  }
  if (current) pieces.push(current);
  return pieces;
}

async function summarizeChunk({ chunk, index, total, config, provider, detail }) {
  const prompt = `You are summarizing a technical teaching transcript chunk.

Use only the tutor's teaching content in the transcript. Ignore greetings, logistics, student chatter, and unrelated Q&A unless the tutor uses it to teach a concept.

Return strict JSON with this shape:
{
  "topics": [
    {
      "name": "short topic name",
      "timestamp": "HH:MM:SS-HH:MM:SS",
      "definition": "brief definition",
      "description": "what was taught, with timestamp references included naturally"
    }
  ],
  "highlights": [
    { "timestamp": "HH:MM:SS-HH:MM:SS", "summary": "short highlight for quick review" }
  ]
}

Detail level: ${detail}
Chunk ${index + 1} of ${total}: ${chunk.start}-${chunk.end}

Transcript:
${chunk.text}`;
  return callProviderJson({ config, provider, prompt, cacheKey: hash(`${config.provider}:${provider.model}:chunk:${detail}:${prompt}`), fallback: { topics: [], highlights: [] } });
}

async function mergeSummaries({ chunkSummaries, config, provider, detail, speaker }) {
  const prompt = `Consolidate these chunk summaries into one study document.

Tutor/speaker: ${speaker || "selected speaker"}
Merge duplicate or overlapping topics. Keep useful timestamp ranges. Keep the output concise but complete.

Return strict JSON:
{
  "title": "Live Session Summary: ...",
  "highlights": [
    { "timestamp": "HH:MM:SS-HH:MM:SS", "summary": "what to watch for quick review" }
  ],
  "topics": [
    {
      "name": "topic name",
      "timestamp": "HH:MM:SS-HH:MM:SS",
      "definition": "brief definition",
      "description": "topic description with timestamps"
    }
  ],
  "reviewPath": ["short review recommendation"]
}

Detail level: ${detail}

Chunk summaries:
${JSON.stringify(chunkSummaries, null, 2)}`;
  return callProviderJson({ config, provider, prompt, cacheKey: hash(`${config.provider}:${provider.model}:merge:${detail}:${prompt}`), fallback: null });
}

async function callProviderJson({ config, provider, prompt, cacheKey, fallback }) {
  await fs.mkdir(CACHE_DIR, { recursive: true });
  const cacheFile = path.join(CACHE_DIR, `${cacheKey}.json`);
  try {
    return JSON.parse(await fs.readFile(cacheFile, "utf8"));
  } catch {}

  const content = config.provider === "gemini"
    ? await callGemini({ provider, prompt })
    : await callLmStudio({ provider, prompt });
  const parsed = parseJsonFromModel(content) || fallback;
  if (!parsed) throw new Error(`Model did not return parseable JSON. First 300 chars: ${content.slice(0, 300)}`);
  await fs.writeFile(cacheFile, JSON.stringify(parsed, null, 2));
  return parsed;
}

async function callLmStudio({ provider, prompt }) {
  const baseUrl = normalizeOpenAiBaseUrl(provider.baseUrl);
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: provider.model,
      temperature: 0.2,
      messages: [
        { role: "system", content: "You produce accurate structured summaries from transcripts. Return JSON only." },
        { role: "user", content: prompt }
      ]
    })
  });
  if (!response.ok) throw new Error(`LM Studio chat failed with ${response.status}: ${await response.text()}`);
  const data = await response.json();
  return data.choices?.[0]?.message?.content || "";
}

async function callGemini({ provider, prompt }) {
  const apiKey = resolveGeminiApiKey(provider);
  if (!apiKey) {
    throw new Error(`Gemini API key is missing. Set ${provider.apiKeyEnv || "GEMINI_API_KEY"} or provider.gemini.apiKey in app.config.local.json.`);
  }
  const baseUrl = normalizeGeminiBaseUrl(provider.baseUrl);
  const model = encodeURIComponent(provider.model);
  const response = await fetch(`${baseUrl}/models/${model}:generateContent`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-goog-api-key": apiKey
    },
    body: JSON.stringify({
      contents: [
        {
          role: "user",
          parts: [
            {
              text: `You produce accurate structured summaries from transcripts. Return JSON only.\n\n${prompt}`
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: "application/json"
      }
    })
  });
  if (!response.ok) throw new Error(`Gemini generateContent failed with ${response.status}: ${await response.text()}`);
  const data = await response.json();
  return (data.candidates?.[0]?.content?.parts || []).map((part) => part.text || "").join("\n");
}

function parseJsonFromModel(content) {
  try {
    return JSON.parse(content);
  } catch {}
  const fenced = content.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced) {
    try {
      return JSON.parse(fenced[1]);
    } catch {}
  }
  const first = content.indexOf("{");
  const last = content.lastIndexOf("}");
  if (first >= 0 && last > first) {
    try {
      return JSON.parse(content.slice(first, last + 1));
    } catch {}
  }
  return null;
}

function normalizeResult(result, meta) {
  const normalized = result || { topics: [], highlights: [], reviewPath: [] };
  return {
    ...meta,
    title: normalized.title || meta.title,
    highlights: Array.isArray(normalized.highlights) ? normalized.highlights : [],
    topics: Array.isArray(normalized.topics) ? normalized.topics : [],
    reviewPath: Array.isArray(normalized.reviewPath) ? normalized.reviewPath : []
  };
}

function renderMarkdown(result) {
  const lines = [
    `# ${result.title}`,
    "",
    `Source transcript: ${result.sourceName}`,
    `Tutor focus: ${result.speaker || "Selected speaker only"}`,
    `Processed cues: ${result.cueCount}; chunks: ${result.chunkCount}`,
    "",
    "## Session Highlights For Quick Review",
    "",
    "| Timestamp | Highlight |",
    "| --- | --- |",
    ...result.highlights.map((item) => `| ${item.timestamp || ""} | ${escapeMd(item.summary || "")} |`),
    "",
    "## Topics Covered With Brief Definitions",
    ""
  ];
  result.topics.forEach((topic, index) => {
    lines.push(`### ${index + 1}. ${topic.name || "Topic"}`, "");
    lines.push(`**Timestamp:** ${topic.timestamp || ""}`, "");
    lines.push(`**Definition:** ${topic.definition || ""}`, "");
    lines.push(`**Description:** ${topic.description || ""}`, "");
  });
  if (result.reviewPath.length) {
    lines.push("## Suggested Review Path", "");
    result.reviewPath.forEach((item, index) => lines.push(`${index + 1}. ${item}`));
    lines.push("");
  }
  return lines.join("\n");
}

async function buildDocx(result) {
  const children = [
    new Paragraph({ text: result.title, heading: "Title" }),
    new Paragraph(`Source transcript: ${result.sourceName}`),
    new Paragraph(`Tutor focus: ${result.speaker || "Selected speaker only"}`),
    new Paragraph({ text: "Session Highlights For Quick Review", heading: "Heading1" }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [cell("Timestamp", true), cell("Highlight", true)] }),
        ...result.highlights.map((item) => new TableRow({ children: [cell(item.timestamp || ""), cell(item.summary || "")] }))
      ]
    }),
    new Paragraph({ text: "Topics Covered With Brief Definitions", heading: "Heading1" })
  ];
  result.topics.forEach((topic, index) => {
    children.push(new Paragraph({ text: `${index + 1}. ${topic.name || "Topic"}`, heading: "Heading2" }));
    children.push(new Paragraph({ children: [new TextRun({ text: "Timestamp: ", bold: true }), new TextRun(topic.timestamp || "")] }));
    children.push(new Paragraph({ children: [new TextRun({ text: "Definition: ", bold: true }), new TextRun(topic.definition || "")] }));
    children.push(new Paragraph({ children: [new TextRun({ text: "Description: ", bold: true }), new TextRun(topic.description || "")] }));
  });
  if (result.reviewPath.length) {
    children.push(new Paragraph({ text: "Suggested Review Path", heading: "Heading1" }));
    result.reviewPath.forEach((item, index) => children.push(new Paragraph(`${index + 1}. ${item}`)));
  }
  return Packer.toBuffer(new Document({ sections: [{ children }] }));
}

function cell(text, bold = false) {
  return new TableCell({ children: [new Paragraph({ children: [new TextRun({ text, bold })] })] });
}

function updateJob(id, patch) {
  jobs.set(id, { ...jobs.get(id), ...patch, updatedAt: new Date().toISOString() });
}

function seconds(timestamp) {
  const [h, m, s] = timestamp.split(":").map(Number);
  return h * 3600 + m * 60 + s;
}

function hash(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function escapeMd(value) {
  return String(value).replace(/\|/g, "\\|").replace(/\n/g, " ");
}

function safeFileName(value) {
  return String(value).replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").slice(0, 120);
}

function clamp(value, min, max) {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, value));
}

app.listen(PORT, "127.0.0.1", () => {
  console.log(`API server running at http://127.0.0.1:${PORT}`);
});
