import { useEffect, useMemo, useState } from "react";
import { deployTemplate, loadTemplates } from "./dataverse";
import type { DeploymentResult, TemplateRecord } from "./types";

type SortKey = "name" | "createdon" | "modifiedon" | "associatedentitytypecode";
const date = (value?: string) => value ? new Date(value).toLocaleString() : "";
const connectionName = (connection: unknown, fallback: string) => {
  const c = connection as Record<string, unknown> | null; return String(c?.name ?? c?.connectionName ?? c?.environmentName ?? fallback);
};
function resultsText(rows: DeploymentResult[]) {
  return rows.map(r => `${r.success ? "✓" : "✗"} ${r.template}\nTarget: ${r.target}\nAction: ${r.action} | Bindings: ${r.bindings ?? "—"}\n${r.message}`).join("\n\n");
}

export function App() {
  const [sourceName, setSourceName] = useState("Active connection"); const [targetName, setTargetName] = useState("Secondary connection");
  const [templates, setTemplates] = useState<TemplateRecord[]>([]); const [selected, setSelected] = useState<Set<string>>(new Set());
  const [results, setResults] = useState<DeploymentResult[]>([]); const [busy, setBusy] = useState(false); const [ready, setReady] = useState(false); const [status, setStatus] = useState("Loading connections…");
  const [sort, setSort] = useState<{ key: SortKey; direction: 1 | -1 }>({ key: "name", direction: 1 });
  const sorted = useMemo(() => [...templates].sort((a, b) => String(a[sort.key] ?? "").localeCompare(String(b[sort.key] ?? "")) * sort.direction), [templates, sort]);
  const changeSort = (key: SortKey) => setSort(current => current.key === key ? { key, direction: current.direction === 1 ? -1 : 1 } : { key, direction: 1 });
  const arrow = (key: SortKey) => sort.key === key ? (sort.direction === 1 ? " ▲" : " ▼") : "";
  const refresh = async () => { setBusy(true); setStatus("Loading Word templates…"); try { const rows = await loadTemplates(); setTemplates(rows); setSelected(new Set()); setStatus(`${rows.length} Word template${rows.length === 1 ? "" : "s"} loaded.`); } catch (e) { setStatus(`Could not load templates: ${e instanceof Error ? e.message : String(e)}`); } finally { setBusy(false); } };
  useEffect(() => { void (async () => { try { const [source, target] = await Promise.all([window.toolboxAPI.connections.getActiveConnection(), window.toolboxAPI.connections.getSecondaryConnection()]); if (!source) throw new Error("Select an active source connection."); if (!target) throw new Error("Select a secondary target connection."); if (source.id === target.id || source.url.replace(/\/$/, "").toLowerCase() === target.url.replace(/\/$/, "").toLowerCase()) throw new Error("The target connection cannot be the same as the source connection."); setSourceName(source.name); setTargetName(target.name); setReady(true); await refresh(); } catch (e) { setStatus(`Connection setup failed: ${e instanceof Error ? e.message : String(e)}`); } })(); }, []);
  const deploy = async () => { const chosen = templates.filter(t => selected.has(t.documenttemplateid)); if (!chosen.length) return; setBusy(true); setResults([]); const rows: DeploymentResult[] = []; for (const template of chosen) { setStatus(`Deploying ${template.name}…`); const result = await deployTemplate(template, targetName); rows.push(result); setResults([...rows]); } setStatus(`Deployment complete: ${rows.filter(x => x.success).length} succeeded, ${rows.filter(x => !x.success).length} failed.`); setBusy(false); };
  const toggleAll = () => setSelected(selected.size === templates.length ? new Set() : new Set(templates.map(t => t.documenttemplateid)));
  const copy = async () => { const text = resultsText(results); if (window.toolboxAPI?.utils?.copyToClipboard) await window.toolboxAPI.utils.copyToClipboard(text); else await navigator.clipboard.writeText(text); setStatus("Deployment results copied."); };
  return <main>
    <header><div><h1>Document Template Deployment Manager</h1><p>Safely rebind and deploy Dataverse Word templates.</p></div><button className="secondary" onClick={refresh} disabled={busy || !ready}>Refresh</button></header>
    <section className="connections"><div><span>Source</span><strong>{sourceName}</strong></div><div className="flow">→</div><div><span>Target</span><strong>{targetName}</strong></div></section>
    <section className="panel"><div className="panel-title"><div><h2>Word document templates</h2><small>{selected.size} selected</small></div><button onClick={deploy} disabled={busy || !ready || selected.size === 0}>Deploy selected</button></div>
      <div className="table-wrap"><table><thead><tr><th className="check"><input aria-label="Select all" type="checkbox" checked={templates.length > 0 && selected.size === templates.length} onChange={toggleAll}/></th><th onClick={() => changeSort("name")}>Name{arrow("name")}</th><th onClick={() => changeSort("createdon")}>Created{arrow("createdon")}</th><th onClick={() => changeSort("modifiedon")}>Modified{arrow("modifiedon")}</th><th onClick={() => changeSort("associatedentitytypecode")}>Associated table{arrow("associatedentitytypecode")}</th></tr></thead>
      <tbody>{sorted.map(t => <tr key={t.documenttemplateid}><td className="check"><input aria-label={`Select ${t.name}`} type="checkbox" checked={selected.has(t.documenttemplateid)} onChange={() => setSelected(current => { const next = new Set(current); next.has(t.documenttemplateid) ? next.delete(t.documenttemplateid) : next.add(t.documenttemplateid); return next; })}/></td><td><strong>{t.name}</strong></td><td>{date(t.createdon)}</td><td>{date(t.modifiedon)}</td><td><code>{t.associatedentitytypecode}</code></td></tr>)}</tbody></table></div>
    </section>
    <section className="panel"><div className="panel-title"><div><h2>Deployment results</h2><small>{results.length ? `${results.filter(x => x.success).length} successful` : "No deployment run yet"}</small></div>{results.length > 0 && <button className="secondary" onClick={copy}>Copy results</button>}</div>
      {results.length === 0 ? <div className="empty">Results will appear here after deployment.</div> : <div className="table-wrap"><table><thead><tr><th>Target</th><th>Template</th><th>Action</th><th>Status</th><th>Bindings</th><th>Result / error</th></tr></thead><tbody>{results.map((r, i) => <tr key={`${r.template}-${i}`}><td>{r.target}</td><td>{r.template}</td><td>{r.action}</td><td><span className={r.success ? "success" : "failure"}>{r.success ? "Success" : "Failed"}</span></td><td>{r.bindings ?? "—"}</td><td>{r.message}</td></tr>)}</tbody></table></div>}
    </section><footer aria-live="polite">{busy && <span className="spinner"/>}{status}</footer>
  </main>;
}
