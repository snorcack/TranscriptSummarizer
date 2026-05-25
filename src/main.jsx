import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

function App() {
  const [file, setFile] = useState(null);
  const [speaker, setSpeaker] = useState("Sadhna Singh");
  const [appConfig, setAppConfig] = useState(null);
  const [models, setModels] = useState([]);
  const [detail, setDetail] = useState("medium");
  const [chunkMinutes, setChunkMinutes] = useState(10);
  const [maxChunkChars, setMaxChunkChars] = useState(9000);
  const [connection, setConnection] = useState({ status: "idle", message: "Not checked" });
  const [job, setJob] = useState(null);
  const [error, setError] = useState("");
  const pollRef = useRef(null);

  const canRun = file && speaker.trim() && appConfig?.activeProvider?.model && !["queued", "running"].includes(job?.status);

  useEffect(() => {
    loadConfig();
    return () => clearInterval(pollRef.current);
  }, []);

  async function loadConfig() {
    setConnection({ status: "checking", message: "Loading provider config..." });
    setError("");
    try {
      const configResponse = await fetch("/api/config");
      const configData = await configResponse.json();
      if (!configResponse.ok || !configData.ok) throw new Error(configData.error || "Unable to load app config.");
      setAppConfig(configData.config);
      const defaults = configData.config.defaults || {};
      setSpeaker(defaults.speaker || "Sadhna Singh");
      setDetail(defaults.detail || "medium");
      setChunkMinutes(defaults.chunkMinutes || 10);
      setMaxChunkChars(defaults.maxChunkChars || 9000);

      const response = await fetch("/api/provider/models");
      const data = await response.json();
      if (!response.ok || !data.ok) throw new Error(data.error || "Unable to check provider.");
      setModels(data.models);
      setConnection({
        status: data.warning ? "warn" : "ok",
        message: data.warning || `Using ${configData.config.provider}: ${configData.config.activeProvider.model}`
      });
    } catch (err) {
      setModels([]);
      setConnection({ status: "bad", message: err.message });
    }
  }

  async function startSummary() {
    if (!canRun) return;
    setError("");
    setJob(null);
    const form = new FormData();
    form.append("transcript", file);
    form.append("speaker", speaker);
    form.append("detail", detail);
    form.append("chunkMinutes", String(chunkMinutes));
    form.append("maxChunkChars", String(maxChunkChars));

    try {
      const response = await fetch("/api/summarize", { method: "POST", body: form });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "Unable to start summary job.");
      pollJob(data.id);
    } catch (err) {
      setError(err.message);
    }
  }

  function pollJob(id) {
    clearInterval(pollRef.current);
    const load = async () => {
      const response = await fetch(`/api/jobs/${id}`);
      const data = await response.json();
      if (!response.ok) {
        clearInterval(pollRef.current);
        setError(data.error || "Job not found.");
        return;
      }
      setJob(data);
      if (["completed", "failed"].includes(data.status)) clearInterval(pollRef.current);
    };
    load();
    pollRef.current = setInterval(load, 1200);
  }

  const resultStats = useMemo(() => {
    if (!job?.result) return null;
    return `${job.result.topics.length} topics, ${job.result.highlights.length} highlights, ${job.result.chunkCount} chunks`;
  }, [job]);

  return (
    <main className="shell">
      <section className="topbar">
        <div>
          <h1>Transcript Summarizer</h1>
          <p>Local React app for tutor-only summaries using the configured model provider.</p>
        </div>
        <StatusBadge status={connection.status}>{connection.message}</StatusBadge>
      </section>

      <section className="workspace">
        <aside className="panel settings">
          <h2>Input</h2>
          <label className="fileDrop">
            <input
              type="file"
              accept=".txt,.vtt,.srt,text/plain"
              onChange={(event) => setFile(event.target.files?.[0] || null)}
            />
            <span>{file ? file.name : "Choose transcript file"}</span>
          </label>

          <Field label="Tutor speaker">
            <input value={speaker} onChange={(event) => setSpeaker(event.target.value)} placeholder="Sadhna Singh" />
          </Field>

          <h2>Provider</h2>
          <div className="providerBox">
            <div>
              <span>Active</span>
              <strong>{appConfig?.provider || "Loading"}</strong>
            </div>
            <div>
              <span>Model</span>
              <strong>{appConfig?.activeProvider?.model || "Loading"}</strong>
            </div>
            {appConfig?.provider === "lmstudio" && (
              <div>
                <span>Base URL</span>
                <strong>{appConfig.activeProvider.baseUrl}</strong>
              </div>
            )}
            {appConfig?.provider === "gemini" && (
              <div>
                <span>API key</span>
                <strong>{appConfig.activeProvider.resolvedApiKey ? "Configured" : "Missing"}</strong>
              </div>
            )}
            <button type="button" className="secondary" onClick={loadConfig}>Reload Config</button>
          </div>

          <h2>Summary</h2>
          <Field label="Detail">
            <div className="segmented">
              {["short", "medium", "detailed"].map((item) => (
                <button key={item} type="button" className={detail === item ? "active" : ""} onClick={() => setDetail(item)}>
                  {item}
                </button>
              ))}
            </div>
          </Field>

          <Field label={`Chunk minutes: ${chunkMinutes}`}>
            <input type="range" min="5" max="20" value={chunkMinutes} onChange={(event) => setChunkMinutes(Number(event.target.value))} />
          </Field>

          <Field label="Max chunk characters">
            <input type="number" min="3000" max="16000" step="500" value={maxChunkChars} onChange={(event) => setMaxChunkChars(Number(event.target.value))} />
          </Field>

          <button className="primary" type="button" disabled={!canRun} onClick={startSummary}>
            Generate Summary
          </button>
          {error && <p className="error">{error}</p>}
        </aside>

        <section className="panel output">
          <div className="outputHeader">
            <div>
              <h2>Output</h2>
              <p>{job ? `${job.message} ${resultStats ? `• ${resultStats}` : ""}` : "Upload a transcript and generate a local summary."}</p>
            </div>
            {job?.result && (
              <div className="downloads">
                <a href={`/api/jobs/${job.id}/download/md`}>Markdown</a>
                <a href={`/api/jobs/${job.id}/download/docx`}>DOCX</a>
              </div>
            )}
          </div>

          {job && <Progress value={job.progress || 0} status={job.status} />}
          {job?.status === "failed" && <p className="error">{job.error}</p>}
          {job?.chunkStats?.length > 0 && <ChunkStats stats={job.chunkStats} />}

          {job?.result ? <SummaryPreview result={job.result} /> : <EmptyState />}
        </section>
      </section>
    </main>
  );
}

function Field({ label, children }) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}

function StatusBadge({ status, children }) {
  return <div className={`status ${status}`}>{children}</div>;
}

function Progress({ value, status }) {
  return (
    <div className="progressBlock">
      <div className="progressMeta">
        <span>{status}</span>
        <span>{value}%</span>
      </div>
      <div className="progress"><div style={{ width: `${value}%` }} /></div>
    </div>
  );
}

function EmptyState() {
  return (
    <div className="empty">
      <h3>Ready when your transcript is.</h3>
      <p>The app will keep only the selected speaker, preserve timestamps, summarize chunks locally, then merge topics and highlights.</p>
    </div>
  );
}

function ChunkStats({ stats }) {
  const largest = Math.max(...stats.map((item) => item.chars));
  return (
    <section className="chunkStats">
      <div className="chunkStatsHeader">
        <h2>Chunk Diagnostics</h2>
        <span>{stats.length} chunks • largest {largest.toLocaleString()} chars</span>
      </div>
      <div className="chunkGrid">
        {stats.map((chunk) => (
          <div className="chunkPill" key={chunk.index} title={`${chunk.start}-${chunk.end}, ${chunk.cueCount} cues`}>
            <strong>{chunk.index}</strong>
            <span>{chunk.chars.toLocaleString()}</span>
          </div>
        ))}
      </div>
    </section>
  );
}

function SummaryPreview({ result }) {
  return (
    <article className="summary">
      <h1>{result.title}</h1>
      <p className="muted">Speaker: {result.speaker} • Processed cues: {result.cueCount}</p>

      <h2>Session Highlights</h2>
      <div className="highlightList">
        {result.highlights.map((item, index) => (
          <div className="highlight" key={`${item.timestamp}-${index}`}>
            <strong>{item.timestamp}</strong>
            <span>{item.summary}</span>
          </div>
        ))}
      </div>

      <h2>Topics Covered</h2>
      <div className="topicList">
        {result.topics.map((topic, index) => (
          <section className="topic" key={`${topic.name}-${index}`}>
            <div className="topicTitle">
              <h3>{index + 1}. {topic.name}</h3>
              <span>{topic.timestamp}</span>
            </div>
            <p><strong>Definition:</strong> {topic.definition}</p>
            <p>{topic.description}</p>
          </section>
        ))}
      </div>

      {result.reviewPath?.length > 0 && (
        <>
          <h2>Suggested Review Path</h2>
          <ol>
            {result.reviewPath.map((item, index) => <li key={index}>{item}</li>)}
          </ol>
        </>
      )}
    </article>
  );
}

createRoot(document.getElementById("root")).render(<App />);
