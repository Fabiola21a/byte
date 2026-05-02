"use client"

import { useState, useEffect, useMemo, useCallback, useRef } from "react"
import * as XLSX from "xlsx"
import { createClient } from "@supabase/supabase-js"

// ─────────────────────────────────────────
// SUPABASE CLIENT
// ─────────────────────────────────────────
const supabase = createClient(
  "https://rzvcxrftygrddxfjmmqy.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ6dmN4cmZ0eWdyZGR4ZmptbXF5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc2NTgyODcsImV4cCI6MjA5MzIzNDI4N30.kWRQmYepHsM9YVZchFNvTGGaNsxuFgP0fYmuxOq-igY"
)

// ─────────────────────────────────────────
// TYPES
// ─────────────────────────────────────────
type Status = "prospecting" | "contacted" | "replied" | "meeting" | "negotiation" | "closed" | "no_reply" | "skip"
type View = "dashboard" | "leads" | "pipeline" | "sequences" | "campaigns"

interface Lead {
  id: string
  company: string
  contact?: string
  role?: string
  email?: string
  linkedin?: string
  country?: string
  industry?: string
  size?: string
  status: Status
  open_roles?: string
  growth_signal?: string
  main_role?: string
  remote_signal?: string
  match_type?: string
  notes?: string
  source?: string
  created_at?: string
  updated_at?: string
}

interface Activity {
  id: string
  lead_id: string
  type: string
  description?: string
  created_at: string
}

interface EmailSequence {
  id: string
  lead_id: string
  step: number
  status: string
  sent_at?: string
  subject?: string
  body?: string
}

interface EmailTemplate {
  id: string
  step: number
  name: string
  day_offset: number
  type?: string
  subject?: string
  body?: string
  notes?: string
}

// ─────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────
const STAGES: { id: Status; label: string; color: string; bg: string }[] = [
  { id: "prospecting",  label: "Prospecting",  color: "#64748b", bg: "rgba(100,116,139,0.12)" },
  { id: "contacted",    label: "Contacted",    color: "#38bdf8", bg: "rgba(56,189,248,0.12)" },
  { id: "replied",      label: "Replied",      color: "#34d399", bg: "rgba(52,211,153,0.12)" },
  { id: "meeting",      label: "Meeting",      color: "#fbbf24", bg: "rgba(251,191,36,0.12)" },
  { id: "negotiation",  label: "Negotiation",  color: "#a78bfa", bg: "rgba(167,139,250,0.12)" },
  { id: "closed",       label: "Closed",       color: "#fb923c", bg: "rgba(251,146,60,0.12)" },
  { id: "no_reply",     label: "No Reply",     color: "#475569", bg: "rgba(71,85,105,0.12)" },
  { id: "skip",         label: "Skip",         color: "#f87171", bg: "rgba(248,113,113,0.12)" },
]



// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2) }
function stageOf(id: Status) { return STAGES.find(s => s.id === id) || STAGES[0] }
function initials(s?: string) { return (s || "?").split(" ").slice(0, 2).map(w => w[0] || "").join("").toUpperCase() }
function firstName(n?: string) { return n ? n.trim().split(" ")[0] : "there" }

const AV = ["#38bdf8","#34d399","#fb923c","#fbbf24","#a78bfa","#f43f5e","#4ade80","#60a5fa"]
function avColor(s: string) { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) % AV.length; return AV[h] }

function renderEmail(body: string, lead: Lead) {
  return body
    .replace(/\{\{first_name\}\}/g, firstName(lead.contact))
    .replace(/\{\{company\}\}/g, lead.company)
    .replace(/\{\{open_roles\}\}/g, lead.open_roles || "several")
    .replace(/\{\{growth_signal\}\}/g, lead.growth_signal || "is growing fast")
    .replace(/\{\{main_role\}\}/g, lead.main_role || "the open role")
}

// ─────────────────────────────────────────
// AVATAR
// ─────────────────────────────────────────
function Avatar({ name, size = 28 }: { name: string; size?: number }) {
  const color = avColor(name)
  return (
    <div style={{ width: size, height: size, borderRadius: 8, background: color + "22", color, display: "flex", alignItems: "center", justifyContent: "center", fontSize: size * 0.36, fontWeight: 700, flexShrink: 0, fontFamily: "monospace" }}>
      {initials(name)}
    </div>
  )
}

// ─────────────────────────────────────────
// STATUS PILL
// ─────────────────────────────────────────
function StatusPill({ status }: { status: Status }) {
  const s = stageOf(status)
  return (
    <span style={{ fontSize: 10, fontFamily: "monospace", padding: "2px 8px", borderRadius: 5, background: s.bg, color: s.color, border: `1px solid ${s.color}30`, fontWeight: 600, whiteSpace: "nowrap" }}>
      {s.label}
    </span>
  )
}

// ─────────────────────────────────────────
// MAIN CRM
// ─────────────────────────────────────────
export default function ByteCRM() {
  const [leads, setLeads] = useState<Lead[]>([])
  const [loading, setLoading] = useState(true)
  const [view, setView] = useState<View>("dashboard")
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [addOpen, setAddOpen] = useState(false)
  const [detailOpen, setDetailOpen] = useState(false)
  const [detTab, setDetTab] = useState("info")
  const [search, setSearch] = useState("")
  const [filterStatus, setFilterStatus] = useState("all")
  const [activities, setActivities] = useState<Activity[]>([])
  const [emailSequences, setEmailSequences] = useState<EmailSequence[]>([])
  const [saving, setSaving] = useState(false)
  const [toast, setToast] = useState<string | null>(null)
  const [dbStatus, setDbStatus] = useState<"connecting" | "connected" | "error">("connecting")
const [importing, setImporting] = useState(false)
  const [importProgress, setImportProgress] = useState("")
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [expandedActivity, setExpandedActivity] = useState<string | null>(null)
  const [activityEmailContent, setActivityEmailContent] = useState<{ subject: string; body: string } | null>(null)
  const [loadingEmailContent, setLoadingEmailContent] = useState(false)
  const [lastUpdated, setLastUpdated] = useState<Date | null>(null)
  const [secondsAgo, setSecondsAgo] = useState(0)
  const [emailTemplates, setEmailTemplates] = useState<EmailTemplate[]>([])
  const [loadingTemplates, setLoadingTemplates] = useState(false)
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [isMobile, setIsMobile] = useState(false)
  const [collapsedPipelineCols, setCollapsedPipelineCols] = useState<Set<string>>(new Set())

  const [form, setForm] = useState({
    company: "", contact: "", role: "", email: "", linkedin: "",
    country: "US", industry: "", size: "", open_roles: "", growth_signal: "",
    main_role: "", remote_signal: "", match_type: "fully_matched", notes: "", source: "lessie"
  })

  const ff = (k: string) => (v: string) => setForm(f => ({ ...f, [k]: v }))

// ── FETCH LEADS ──
  const fetchLeads = useCallback(async (silent = false) => {
    if (!silent) setLoading(true)
    const { data, error } = await supabase
      .from("recruit_leads")
      .select("*")
      .order("created_at", { ascending: false })
    if (error) { setDbStatus("error"); console.error(error) }
    else { setLeads(data || []); setDbStatus("connected"); setLastUpdated(new Date()); setSecondsAgo(0) }
    if (!silent) setLoading(false)
  }, [])

  useEffect(() => { fetchLeads() }, [fetchLeads])

  // ── MOBILE DETECTION ──
  useEffect(() => {
    const checkMobile = () => setIsMobile(window.innerWidth < 768)
    checkMobile()
    window.addEventListener("resize", checkMobile)
    return () => window.removeEventListener("resize", checkMobile)
  }, [])

  // ── AUTO-REFRESH EVERY 30 SECONDS ──
  useEffect(() => {
    const autoRefresh = setInterval(() => {
      fetchLeads(true) // silent refresh
    }, 30000)
    return () => clearInterval(autoRefresh)
  }, [fetchLeads])

  // ── UPDATE SECONDS AGO COUNTER ──
  useEffect(() => {
    const ticker = setInterval(() => {
      if (lastUpdated) {
        setSecondsAgo(Math.floor((Date.now() - lastUpdated.getTime()) / 1000))
      }
    }, 1000)
    return () => clearInterval(ticker)
  }, [lastUpdated])

  // ── FETCH ACTIVITIES ──
  const fetchActivities = useCallback(async (leadId: string) => {
    const { data } = await supabase
      .from("recruit_activity")
      .select("*")
      .eq("lead_id", leadId)
      .order("created_at", { ascending: false })
    setActivities(data || [])
  }, [])

// ── FETCH EMAIL SEQUENCES ──
  const fetchEmailSequences = useCallback(async (leadId: string) => {
    const { data } = await supabase
      .from("recruit_email_sequence")
      .select("*")
      .eq("lead_id", leadId)
      .order("step", { ascending: true })
    setEmailSequences(data || [])
  }, [])

  // ── FETCH EMAIL TEMPLATES ──
  const fetchEmailTemplates = useCallback(async () => {
    setLoadingTemplates(true)
    const { data, error } = await supabase
      .from("recruit_email_templates")
      .select("*")
      .order("step", { ascending: true })
    if (error) {
      console.error("Error fetching templates:", error)
    }
    setEmailTemplates(data || [])
    setLoadingTemplates(false)
  }, [])

  useEffect(() => { fetchEmailTemplates() }, [fetchEmailTemplates])

  // ── FETCH EMAIL CONTENT FOR ACTIVITY ──
  const fetchEmailContentForActivity = useCallback(async (activityId: string, leadId: string, description: string) => {
    if (expandedActivity === activityId) {
      setExpandedActivity(null)
      setActivityEmailContent(null)
      return
    }
    
    setExpandedActivity(activityId)
    setLoadingEmailContent(true)
    setActivityEmailContent(null)
    
    // Extract step number from description (e.g., "Email Step 1 enviado" -> 1)
    const stepMatch = description?.match(/Step\s*(\d+)/i) || description?.match(/Email\s*#?(\d+)/i) || description?.match(/(\d+)/i)
    const stepNumber = stepMatch ? parseInt(stepMatch[1], 10) : 1
    
    const { data } = await supabase
      .from("recruit_email_sequence")
      .select("subject, body")
      .eq("lead_id", leadId)
      .eq("step", stepNumber)
      .maybeSingle()
    
    if (data) {
      setActivityEmailContent({ subject: data.subject || "", body: data.body || "" })
    } else {
      setActivityEmailContent({ subject: "Email nao encontrado", body: "Nao foi possivel encontrar o conteudo deste email." })
    }
    setLoadingEmailContent(false)
  }, [expandedActivity])

  // ── SHOW TOAST ──
  function showToast(msg: string) {
    setToast(msg)
    setTimeout(() => setToast(null), 3000)
  }


  // ── IMPORT EXCEL ──
  async function handleImportExcel(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setImporting(true)
    setImportProgress("Lendo arquivo...")

    try {
      const buffer = await file.arrayBuffer()
      const wb = XLSX.read(buffer, { type: "array" })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows: any[] = XLSX.utils.sheet_to_json(ws)

      setImportProgress(`${rows.length} leads encontrados. Importando...`)

      const leads = rows
        .filter(r => r["Empresa"] || r["company"] || r["Company"])
        .map(r => ({
          id: uid(),
          company:       r["Empresa"]        || r["Company"]       || r["company"]       || "",
          contact:       r["Nome"]           || r["Contact"]       || r["contact"]       || r["Contato"] || "",
          role:          r["Cargo (Título)"] || r["Role"]          || r["role"]          || r["Cargo"] || "",
          email:         r["Email"]          || r["email"]         || "",
          linkedin:      r["LinkedIn URL"]   || r["linkedin"]      || "",
          country:       r["País"]           || r["Country"]       || r["country"]       || "US",
          open_roles:    r["Vagas Abertas"]  || r["open_roles"]    || r["Vagas"]         || "",
          growth_signal: r["Sinal de Crescimento"] || r["growth_signal"] || r["Notas / Crescimento"] || "",
          main_role:     r["Cargo Contratando"] || r["main_role"]  || "",
          remote_signal: r["Global/Remote Hiring"] || r["remote_signal"] || "",
          match_type:    (r["Match"] || "").toLowerCase().includes("fully") ? "fully_matched" : "partial_match",
          notes:         r["Notas"]          || r["notes"]         || "",
          source:        "lessie",
          status:        "prospecting" as Status,
        }))
        .filter(l => l.company)

      let imported = 0
      let skipped = 0

      for (const lead of leads) {
        // Verificar duplicata por email ou empresa+contato
        let isDuplicate = false

        if (lead.email) {
          const { data: existing } = await supabase
            .from("recruit_leads")
            .select("id")
            .eq("email", lead.email)
            .maybeSingle()
          if (existing) isDuplicate = true
        } else {
          const { data: existing } = await supabase
            .from("recruit_leads")
            .select("id")
            .eq("company", lead.company)
            .eq("contact", lead.contact || "")
            .maybeSingle()
          if (existing) isDuplicate = true
        }

        if (isDuplicate) {
          skipped++
          setImportProgress(`Importando... ${imported + skipped}/${leads.length} (${skipped} duplicatas ignoradas)`)
          continue
        }

        const { error } = await supabase.from("recruit_leads").insert([lead])
        if (!error) {
          await supabase.from("recruit_activity").insert([{
            lead_id: lead.id, type: "lead_created",
            description: `Lead importado via Excel — ${lead.company}`
          }])
          imported++
          setImportProgress(`Importando... ${imported + skipped}/${leads.length}`)
        }
      }

      await fetchLeads()
      showToast(`✅ ${imported} importados${skipped > 0 ? ` · ${skipped} duplicatas ignoradas` : ""}!`)
      setImportProgress("")
    } catch (err) {
      showToast("❌ Erro ao ler o arquivo Excel.")
    }

    setImporting(false)
    if (fileInputRef.current) fileInputRef.current.value = ""
  }

  // ── SAVE LEAD ──
  async function saveLead() {
    if (!form.company.trim()) return
    setSaving(true)
    const lead: Lead = { id: uid(), ...form, status: "prospecting" }
    const { error } = await supabase.from("recruit_leads").insert([lead])
    if (!error) {
      await supabase.from("recruit_activity").insert([{
        lead_id: lead.id, type: "lead_created",
        description: `Lead criado — ${lead.company}`
      }])
      await fetchLeads()
      setAddOpen(false)
      setForm({ company: "", contact: "", role: "", email: "", linkedin: "", country: "US", industry: "", size: "", open_roles: "", growth_signal: "", main_role: "", remote_signal: "", match_type: "fully_matched", notes: "", source: "lessie" })
      showToast("✅ Lead salvo no Supabase!")
    } else {
      showToast("❌ Erro ao salvar: " + error.message)
    }
    setSaving(false)
  }

  // ── UPDATE STATUS ──
  async function updateStatus(id: string, status: Status) {
    const { error } = await supabase.from("recruit_leads").update({ status }).eq("id", id)
    if (!error) {
      await supabase.from("recruit_activity").insert([{
        lead_id: id, type: "status_changed",
        description: `Status atualizado para ${stageOf(status).label}`
      }])
      await fetchLeads()
      if (selectedId === id) fetchActivities(id)
      showToast(`🔄 Status → ${stageOf(status).label}`)
    }
  }

  // ── DELETE ──
  async function deleteLead(id: string) {
    if (!confirm("Remover este lead permanentemente?")) return
    await supabase.from("recruit_leads").delete().eq("id", id)
    await fetchLeads()
    setDetailOpen(false)
    setSelectedId(null)
    showToast("🗑 Lead removido.")
  }

// ── OPEN DETAIL ──
  function openDetail(id: string) {
    setSelectedId(id)
    setDetTab("info")
    setDetailOpen(true)
    setExpandedActivity(null)
    setActivityEmailContent(null)
    fetchActivities(id)
    fetchEmailSequences(id)
  }

  const sel = leads.find(l => l.id === selectedId)

  // ── TOGGLE PIPELINE COLUMN (MOBILE) ──
  const togglePipelineCol = (stageId: string) => {
    setCollapsedPipelineCols(prev => {
      const next = new Set(prev)
      if (next.has(stageId)) next.delete(stageId)
      else next.add(stageId)
      return next
    })
  }

  // ── COMPUTED ──
  const byStage = useMemo(() => {
    const m: Record<string, number> = {}
    STAGES.forEach(s => m[s.id] = leads.filter(l => l.status === s.id).length)
    return m
  }, [leads])

  const filtered = useMemo(() => leads.filter(l => {
    const ok = filterStatus === "all" || l.status === filterStatus
    const q = search.toLowerCase()
    return ok && (!q || [l.company, l.contact, l.email, l.country].some(v => v?.toLowerCase().includes(q)))
  }), [leads, filterStatus, search])

  const convRate = leads.length ? Math.round(((byStage.meeting || 0) + (byStage.negotiation || 0) + (byStage.closed || 0)) / leads.length * 100) : 0

  // ─────────────────────────────────────────
  // STYLES
  // ─────────────────────────────────────────
  const C = {
    bg: "#06080b", s1: "#0c1117", s2: "#111820", s3: "#18222e",
    border: "#1e293b", border2: "#243650",
    cyan: "#38bdf8", green: "#34d399", orange: "#fb923c",
    yellow: "#fbbf24", purple: "#a78bfa", red: "#f87171",
    text: "#e2e8f0", muted: "#64748b", dim: "#334155",
  }
  const panel = { background: C.s1, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" as const }
  const panelHd = { borderBottom: `1px solid ${C.border}`, padding: "12px 16px", display: "flex" as const, alignItems: "center" as const, justifyContent: "space-between" as const }

  function Btn({ children, onClick, variant = "ghost", size = "md", disabled = false }: any) {
    const base: any = { border: "none", borderRadius: 8, cursor: disabled ? "not-allowed" : "pointer", display: "inline-flex", alignItems: "center", gap: 5, fontWeight: 600, fontFamily: "inherit", padding: size === "sm" ? "4px 10px" : "7px 14px", fontSize: size === "sm" ? 11 : 12, opacity: disabled ? 0.5 : 1 }
    const variants: any = {
      primary: { background: C.cyan, color: "#000" },
      ghost: { background: C.s3, border: `1px solid ${C.border}`, color: C.text },
      green: { background: "rgba(52,211,153,0.15)", color: C.green, border: `1px solid rgba(52,211,153,0.25)` },
      danger: { background: "rgba(248,113,113,0.12)", color: C.red, border: `1px solid rgba(248,113,113,0.25)` },
    }
    return <button style={{ ...base, ...variants[variant] }} onClick={onClick} disabled={disabled}>{children}</button>
  }

  function FInput({ label, value, onChange, placeholder, type = "text", full = false }: any) {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 4, gridColumn: full ? "1/-1" : undefined }}>
        <label style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px" }}>{label}</label>
        <input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} type={type}
          style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 11px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none", width: "100%" }} />
      </div>
    )
  }

  function FSelect({ label, value, onChange, options, full = false }: any) {
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 4, gridColumn: full ? "1/-1" : undefined }}>
        <label style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px" }}>{label}</label>
        <select value={value} onChange={e => onChange(e.target.value)}
          style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 11px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none", cursor: "pointer" }}>
          {options.map((o: string) => <option key={o} value={o} style={{ background: C.s2 }}>{o}</option>)}
        </select>
      </div>
    )
  }

  // ─────────────────────────────────────────
  // RENDER
  // ─────────────────────���───────────────────
  return (
    <div style={{ display: "flex", height: "100vh", background: C.bg, color: C.text, fontFamily: "'DM Sans', system-ui, sans-serif", fontSize: 13, overflow: "hidden" }}>

{/* ── SIDEBAR OVERLAY (MOBILE) ── */}
      {isMobile && sidebarOpen && (
        <div 
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.6)", zIndex: 50 }}
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* ── SIDEBAR ── */}
      <aside style={{ 
        width: 210, 
        flexShrink: 0, 
        background: C.s1, 
        borderRight: `1px solid ${C.border}`, 
        display: "flex", 
        flexDirection: "column",
        ...(isMobile ? {
          position: "fixed",
          left: sidebarOpen ? 0 : -220,
          top: 0,
          bottom: 0,
          zIndex: 60,
          transition: "left 0.25s ease"
        } : {})
      }}>
        <div style={{ padding: "20px 18px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 900, color: C.cyan, letterSpacing: -0.5 }}>BYTE<span style={{ color: C.green }}>.</span></div>
            <div style={{ fontSize: 9, letterSpacing: "2.5px", textTransform: "uppercase", color: C.dim, marginTop: 4, fontFamily: "monospace" }}>CRM · Outreach</div>
          </div>
          {isMobile && (
            <button onClick={() => setSidebarOpen(false)} style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 6, width: 28, height: 28, cursor: "pointer", color: C.muted, fontSize: 14 }}>✕</button>
          )}
        </div>

        <nav style={{ flex: 1, padding: "12px 8px" }}>
          {[
            { id: "dashboard" as View, label: "Dashboard", icon: "⚡" },
            { id: "leads" as View, label: "Leads", icon: "🎯", badge: leads.length },
            { id: "pipeline" as View, label: "Pipeline", icon: "📊" },
            { id: "sequences" as View, label: "Sequencias", icon: "✉️" },
            { id: "campaigns" as View, label: "Campanhas", icon: "🚀" },
          ].map(item => (
            <button key={item.id} onClick={() => { setView(item.id); if (isMobile) setSidebarOpen(false) }}
              style={{ width: "100%", display: "flex", alignItems: "center", gap: 9, padding: "8px 10px", borderRadius: 8, cursor: "pointer", background: view === item.id ? "rgba(56,189,248,0.08)" : "transparent", color: view === item.id ? C.cyan : C.muted, border: view === item.id ? `1px solid rgba(56,189,248,0.15)` : "1px solid transparent", fontSize: 12.5, fontWeight: 500, fontFamily: "inherit", marginBottom: 2 }}>
              <span>{item.icon}</span>
              <span style={{ flex: 1, textAlign: "left" }}>{item.label}</span>
              {item.badge !== undefined && item.badge > 0 && (
                <span style={{ fontSize: 9, fontFamily: "monospace", fontWeight: 700, background: "rgba(56,189,248,0.2)", color: C.cyan, padding: "1px 6px", borderRadius: 8 }}>{item.badge}</span>
              )}
            </button>
          ))}
        </nav>

        <div style={{ padding: "12px 14px", borderTop: `1px solid ${C.border}` }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, background: C.s3, borderRadius: 8, padding: "8px 12px", border: `1px solid ${C.border}` }}>
            <div style={{ width: 7, height: 7, borderRadius: "50%", background: dbStatus === "connected" ? C.green : dbStatus === "error" ? C.red : C.yellow, flexShrink: 0, animation: "pulse 2s infinite" }} />
            <div>
              <div style={{ fontSize: 10, color: C.muted }}>Supabase</div>
              <div style={{ fontSize: 9, fontFamily: "monospace", color: dbStatus === "connected" ? C.green : dbStatus === "error" ? C.red : C.yellow, fontWeight: 700 }}>
                {dbStatus === "connected" ? "CONNECTED" : dbStatus === "error" ? "ERROR" : "CONNECTING..."}
              </div>
            </div>
          </div>
        </div>
      </aside>

      {/* ── MAIN ── */}
      <main style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>

{/* Topbar */}
        <div style={{ padding: isMobile ? "10px 12px" : "12px 24px", borderBottom: `1px solid ${C.border}`, background: C.s1 + "cc", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0, gap: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {isMobile && (
              <button onClick={() => setSidebarOpen(true)} style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 6, width: 32, height: 32, cursor: "pointer", color: C.text, fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }}>☰</button>
            )}
            <div style={{ fontWeight: 700, fontSize: isMobile ? 14 : 16 }}>
              {{ dashboard: "Dashboard", leads: "Leads", pipeline: "Pipeline", sequences: "Sequencias", campaigns: "Campanhas" }[view]}
              {!isMobile && (
                <span style={{ color: C.cyan, fontWeight: 400, fontSize: 13, opacity: .7, marginLeft: 8 }}>
                  {{ dashboard: "Overview", leads: "Database", pipeline: "Kanban", sequences: "Cold Email", campaigns: "A/B Testing" }[view]}
                </span>
              )}
            </div>
          </div>
          <div style={{ display: "flex", gap: isMobile ? 4 : 8, alignItems: "center" }}>
            {!isMobile && lastUpdated && (
              <span style={{ fontSize: 10, fontFamily: "monospace", color: C.muted, marginRight: 4 }}>
                Atualizado: {secondsAgo < 60 ? `${secondsAgo}s atrás` : `${Math.floor(secondsAgo / 60)}m atrás`}
              </span>
            )}
            <Btn onClick={() => fetchLeads()}>{isMobile ? "↻" : "↻ Sync"}</Btn>
            <Btn onClick={() => fileInputRef.current?.click()} disabled={importing}>
              {importing ? (isMobile ? "..." : importProgress || "Importando...") : (isMobile ? "📥" : "📥 Importar Excel")}
            </Btn>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleImportExcel} style={{ display: "none" }} />
            <Btn variant="primary" onClick={() => setAddOpen(true)}>{isMobile ? "＋" : "＋ Novo Lead"}</Btn>
          </div>
        </div>

{/* Content */}
        <div style={{ flex: 1, overflowY: "auto", padding: isMobile ? 12 : 24 }}>

          {loading ? (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100%", color: C.muted, fontSize: 14 }}>
              Carregando leads do Supabase...
            </div>
          ) : (
            <>
              {/* ══ DASHBOARD ══ */}
              {view === "dashboard" && (
                <div>
{/* KPIs */}
                  <div style={{ display: "grid", gridTemplateColumns: isMobile ? "repeat(2,1fr)" : "repeat(6,1fr)", gap: isMobile ? 8 : 12, marginBottom: 20 }}>
                    {[
                      { label: "Total Leads", val: leads.length, color: C.cyan, icon: "🎯" },
                      { label: "Contacted", val: byStage.contacted || 0, color: C.green, icon: "✉️" },
                      { label: "Replies", val: byStage.replied || 0, color: "#22d3ee", icon: "↩️" },
                      { label: "Em Meeting", val: (byStage.meeting || 0) + (byStage.negotiation || 0), color: C.yellow, icon: "📅" },
                      { label: "Fechados", val: byStage.closed || 0, color: C.orange, icon: "🏆" },
                      { label: "Conversao", val: convRate + "%", color: C.purple, icon: "📈" },
                    ].map(k => (
                      <div key={k.label} style={{ ...panel, position: "relative" }}>
                        <div style={{ position: "absolute", top: 0, left: 0, right: 0, height: 2, background: k.color, borderRadius: "12px 12px 0 0" }} />
                        <div style={{ padding: isMobile ? 12 : 16 }}>
                          <div style={{ position: "absolute", top: isMobile ? 10 : 14, right: isMobile ? 10 : 14, fontSize: isMobile ? 14 : 18, opacity: .2 }}>{k.icon}</div>
                          <div style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1.5px", marginBottom: isMobile ? 4 : 8 }}>{k.label}</div>
                          <div style={{ fontSize: isMobile ? 22 : 28, fontWeight: 900, color: k.color, lineHeight: 1 }}>{k.val}</div>
                        </div>
                      </div>
                    ))}
                  </div>

{/* Pipeline bars */}
                  <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.2fr 1fr", gap: 16, marginBottom: 20 }}>
                    <div style={panel}>
                      <div style={panelHd}><span style={{ fontWeight: 700 }}>Pipeline por Etapa</span></div>
                      <div style={{ padding: 16 }}>
                        {STAGES.slice(0, 6).map(s => (
                          <div key={s.id} style={{ marginBottom: 12 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, marginBottom: 5 }}>
                              <span style={{ fontWeight: 600 }}>{s.label}</span>
                              <span style={{ fontFamily: "monospace", color: C.muted }}>{byStage[s.id] || 0}</span>
                            </div>
                            <div style={{ height: 5, background: C.s3, borderRadius: 4, overflow: "hidden" }}>
                              <div style={{ height: "100%", borderRadius: 4, background: s.color, width: leads.length ? `${((byStage[s.id] || 0) / leads.length) * 100}%` : "0%", transition: "width .5s" }} />
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div style={panel}>
                      <div style={panelHd}><span style={{ fontWeight: 700 }}>Leads Recentes</span><Btn size="sm" variant="ghost" onClick={() => setView("leads")}>Ver todos →</Btn></div>
                      {leads.slice(0, 6).map(l => (
                        <div key={l.id} onClick={() => openDetail(l.id)}
                          style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 16px", cursor: "pointer", borderBottom: `1px solid ${C.border}40` }}
                          onMouseEnter={e => (e.currentTarget.style.background = C.s2)}
                          onMouseLeave={e => (e.currentTarget.style.background = "transparent")}>
                          <Avatar name={l.company} size={28} />
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ fontWeight: 600, fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{l.company}</div>
                            <div style={{ fontSize: 10, color: C.muted }}>{l.contact || "—"}</div>
                          </div>
                          <StatusPill status={l.status} />
                        </div>
                      ))}
                      {leads.length === 0 && <div style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 12 }}>Nenhum lead ainda. Adicione o primeiro!</div>}
                    </div>
                  </div>
                </div>
              )}

{/* ══ LEADS ══ */}
              {view === "leads" && (
                <div style={panel}>
                  <div style={{ ...panelHd, flexWrap: "wrap", gap: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 12px", flex: 1, maxWidth: isMobile ? "100%" : 300 }}>
                      <span style={{ color: C.muted }}>🔍</span>
                      <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Buscar..."
                        style={{ background: "transparent", border: "none", outline: "none", color: C.text, fontSize: 13, fontFamily: "inherit", width: "100%" }} />
                    </div>
                    <div style={{ display: "flex", gap: 8, width: isMobile ? "100%" : "auto" }}>
                      <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)}
                        style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 10px", color: C.text, fontSize: 12, fontFamily: "inherit", outline: "none", flex: isMobile ? 1 : "none" }}>
                        <option value="all">Todos</option>
                        {STAGES.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
                      </select>
                      <Btn variant="primary" onClick={() => setAddOpen(true)}>{isMobile ? "＋" : "＋ Lead"}</Btn>
                    </div>
                  </div>
                  
                  {/* Mobile: Cards view */}
                  {isMobile ? (
                    <div style={{ padding: 8, display: "flex", flexDirection: "column", gap: 8 }}>
                      {filtered.length === 0 && <div style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 12 }}>Nenhum lead encontrado.</div>}
                      {filtered.map(l => (
                        <div key={l.id} onClick={() => openDetail(l.id)}
                          style={{ background: C.s2, border: `1px solid ${C.border}`, borderRadius: 10, padding: 12, cursor: "pointer" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                            <Avatar name={l.company} size={32} />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontWeight: 700, fontSize: 13, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{l.company}</div>
                              <div style={{ fontSize: 11, color: C.muted }}>{l.contact || "—"}</div>
                            </div>
                            <StatusPill status={l.status} />
                          </div>
                          {l.open_roles && (
                            <div style={{ fontSize: 11, color: C.orange, fontWeight: 600, marginTop: 4 }}>Vagas: {l.open_roles}</div>
                          )}
                        </div>
                      ))}
                    </div>
                  ) : (
                    /* Desktop: Table view */
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                            {["Empresa", "Contato", "Email", "Vagas", "Growth Signal", "Status", ""].map(h => (
                              <th key={h} style={{ padding: "8px 14px", textAlign: "left", fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1.5px", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {filtered.length === 0 && (
                            <tr><td colSpan={7} style={{ padding: 40, textAlign: "center", color: C.muted }}>Nenhum lead encontrado.</td></tr>
                          )}
                          {filtered.map(l => (
                            <tr key={l.id} onClick={() => openDetail(l.id)}
                              style={{ borderBottom: `1px solid ${C.border}40`, cursor: "pointer" }}
                              onMouseEnter={e => (e.currentTarget.style.background = C.s2)}
                              onMouseLeave={e => (e.currentTarget.style.background = "transparent")}>
                              <td style={{ padding: "10px 14px" }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  <Avatar name={l.company} size={26} />
                                  <div>
                                    <div style={{ fontWeight: 600, fontSize: 12 }}>{l.company}</div>
                                    <div style={{ fontSize: 10, color: C.muted }}>{l.industry || ""}</div>
                                  </div>
                                </div>
                              </td>
                              <td style={{ padding: "10px 14px" }}>
                                <div style={{ fontSize: 12 }}>{l.contact || "—"}</div>
                                <div style={{ fontSize: 10, color: C.muted }}>{l.role || ""}</div>
                              </td>
                              <td style={{ padding: "10px 14px", fontFamily: "monospace", fontSize: 10, color: C.muted }}>{l.email || "—"}</td>
                              <td style={{ padding: "10px 14px", fontSize: 11, color: C.orange, fontWeight: 600 }}>{l.open_roles || "—"}</td>
                              <td style={{ padding: "10px 14px", fontSize: 11, color: C.muted, maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{l.growth_signal || "—"}</td>
                              <td style={{ padding: "10px 14px" }}><StatusPill status={l.status} /></td>
                              <td style={{ padding: "10px 14px" }} onClick={e => e.stopPropagation()}>
                                <div style={{ display: "flex", gap: 4 }}>
                                  <Btn size="sm" variant="ghost" onClick={() => openDetail(l.id)}>👁</Btn>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}

{/* ══ PIPELINE ══ */}
              {view === "pipeline" && (
                <div style={{ display: isMobile ? "flex" : "grid", flexDirection: "column", gridTemplateColumns: "repeat(6,1fr)", gap: 10 }}>
                  {STAGES.slice(0, 6).map(s => {
                    const sl = leads.filter(l => l.status === s.id)
                    const isCollapsed = isMobile && collapsedPipelineCols.has(s.id)
                    return (
                      <div key={s.id} style={{ ...panel, display: "flex", flexDirection: "column", minHeight: isMobile ? "auto" : 300 }}>
                        <div 
                          onClick={() => isMobile && togglePipelineCol(s.id)}
                          style={{ padding: "10px 12px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", cursor: isMobile ? "pointer" : "default" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            {isMobile && (
                              <span style={{ fontSize: 10, color: C.muted }}>{isCollapsed ? "▶" : "▼"}</span>
                            )}
                            <span style={{ fontSize: 10, fontWeight: 700, fontFamily: "monospace", textTransform: "uppercase", letterSpacing: 1, color: s.color }}>{s.label}</span>
                          </div>
                          <span style={{ fontSize: 10, fontFamily: "monospace", background: C.s3, padding: "1px 7px", borderRadius: 8, color: C.muted }}>{sl.length}</span>
                        </div>
                        {!isCollapsed && (
                          <div style={{ padding: 8, display: "flex", flexDirection: "column", gap: 7, flex: 1 }}>
                            {sl.length === 0 && <div style={{ textAlign: "center", padding: 16, color: C.dim, fontSize: 10, fontFamily: "monospace" }}>VAZIO</div>}
                            {sl.map(l => (
                              <div key={l.id} onClick={() => openDetail(l.id)}
                                style={{ background: C.s2, border: `1px solid ${C.border}`, borderRadius: 8, padding: 10, cursor: "pointer" }}
                                onMouseEnter={e => (e.currentTarget.style.borderColor = C.cyan)}
                                onMouseLeave={e => (e.currentTarget.style.borderColor = C.border)}>
                                <div style={{ fontWeight: 600, fontSize: 12, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 3 }}>{l.company}</div>
                                <div style={{ fontSize: 10, color: C.muted, marginBottom: 4 }}>{l.contact || "—"}</div>
                                {l.open_roles && <span style={{ fontSize: 9, fontFamily: "monospace", background: C.s3, color: C.orange, padding: "1px 5px", borderRadius: 4 }}>{l.open_roles}</span>}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}

{/* ══ SEQUENCES ══ */}
              {view === "sequences" && (
                <div style={{ maxWidth: isMobile ? "100%" : 720 }}>
                  <p style={{ color: C.muted, fontSize: 12, marginBottom: 20, lineHeight: 1.6 }}>
                    Templates de email cadastrados no Supabase. Emails personalizados com variaveis reais do lead. Disparados pelo n8n via SMTP.
                  </p>
                  {loadingTemplates ? (
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 60, color: C.muted }}>
                      <div style={{ width: 32, height: 32, border: `3px solid ${C.border}`, borderTopColor: C.cyan, borderRadius: "50%", animation: "spin 1s linear infinite", marginBottom: 12 }} />
                      <div style={{ fontSize: 12 }}>Carregando templates...</div>
                      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
                    </div>
                  ) : emailTemplates.length === 0 ? (
                    <div style={{ ...panel, padding: 40, textAlign: "center" }}>
                      <div style={{ fontSize: 28, marginBottom: 12 }}>📧</div>
                      <div style={{ fontSize: 13, color: C.muted }}>Nenhum template encontrado na tabela recruit_email_templates.</div>
                    </div>
                  ) : (
                    emailTemplates.map((t, i) => {
                      const stepColors = ["#0d47a1","#1b5e20","#e65100","#4a148c","#004d40","#37474f","#c62828","#6a1b9a"]
                      const color = stepColors[i % stepColors.length]
                      return (
                        <div key={t.id} style={{ ...panel, marginBottom: 10 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 0, background: color + "22", borderBottom: `1px solid ${C.border}`, padding: "10px 16px" }}>
                            <div style={{ width: 24, height: 24, borderRadius: 6, background: color, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 700, color: "#fff", marginRight: 10, flexShrink: 0 }}>{t.step}</div>
                            <div style={{ flex: 1 }}>
                              <div style={{ fontWeight: 700, fontSize: 13 }}>{t.name}</div>
                              <div style={{ fontSize: 10, fontFamily: "monospace", color: C.muted, marginTop: 2 }}>
                                Dia {t.day_offset} · {t.type || "email"} {t.subject ? `· Subject: "${t.subject}"` : ""}
                              </div>
                            </div>
                          </div>
                          <div style={{ padding: 14 }}>
                            {t.body && (
                              <pre style={{ background: C.s3, borderRadius: 8, padding: "12px 14px", fontSize: 12, color: C.text, whiteSpace: "pre-wrap", lineHeight: 1.7, margin: "0 0 8px", border: `1px solid ${C.border}` }}>{t.body}</pre>
                            )}
                            {t.notes && <div style={{ fontSize: 10, color: C.muted, fontStyle: "italic" }}>// {t.notes}</div>}
                          </div>
                        </div>
                      )
                    })
                  )}
                </div>
              )}

{/* ══ CAMPAIGNS ══ */}
              {view === "campaigns" && (
                <div style={{ maxWidth: isMobile ? "100%" : 720 }}>
                  <div style={{ ...panel, marginBottom: 16 }}>
                    <div style={{ ...panelHd, flexWrap: "wrap", gap: 8 }}><span style={{ fontWeight: 700 }}>🚀 Batch #1 — US Talent Leaders</span><StatusPill status="prospecting" /></div>
                    <div style={{ display: "grid", gridTemplateColumns: isMobile ? "repeat(2,1fr)" : "repeat(4,1fr)", gap: 0 }}>
                      {[
                        { label: "Total Leads", val: leads.length, color: C.cyan },
                        { label: "Contacted", val: (byStage.contacted || 0) + (byStage.replied || 0), color: C.green },
                        { label: "Replies", val: byStage.replied || 0, color: C.yellow },
                        { label: "Meetings", val: (byStage.meeting || 0), color: C.orange },
                      ].map((k, i) => (
                        <div key={k.label} style={{ padding: isMobile ? "12px 14px" : "16px 20px", borderRight: isMobile && i % 2 === 1 ? "none" : `1px solid ${C.border}`, borderBottom: isMobile && i < 2 ? `1px solid ${C.border}` : "none" }}>
                          <div style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 6 }}>{k.label}</div>
                          <div style={{ fontSize: isMobile ? 20 : 24, fontWeight: 800, color: k.color }}>{k.val}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                  <div style={{ ...panel }}>
                    <div style={panelHd}><span style={{ fontWeight: 700 }}>Cronograma de Envio</span></div>
                    <div style={{ padding: 16 }}>
{[
                        { dia: "Dia 0",  step: "Email #1", hora: "09:00 EST", desc: "Trigger event + pergunta genuina" },
                        { dia: "Dia 3",  step: "Email #2", hora: "10:00 EST", desc: "Revelar zero cost — mesmo thread" },
                        { dia: "Dia 7",  step: "Email #3", hora: "09:30 EST", desc: "Angulo de custo — pergunta direta" },
                        { dia: "Dia 10", step: "LinkedIn", hora: "11:00 EST", desc: "Conexao + mensagem curta (manual)" },
                        { dia: "Dia 14", step: "Email #4", hora: "09:00 EST", desc: "CTA minimo — so pede o JD" },
                        { dia: "Dia 21", step: "Email #5", hora: "10:00 EST", desc: "Break-up humanizado" },
                      ].map((row, i) => (
                        <div key={row.dia} style={{ display: "flex", alignItems: isMobile ? "flex-start" : "center", flexDirection: isMobile ? "column" : "row", gap: isMobile ? 4 : 12, padding: "10px 0", borderBottom: i < 5 ? `1px solid ${C.border}40` : "none" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <span style={{ fontSize: 10, fontFamily: "monospace", color: C.cyan, fontWeight: 700, width: 50, flexShrink: 0 }}>{row.dia}</span>
                            <span style={{ fontSize: 11, fontWeight: 600 }}>{row.step}</span>
                            {!isMobile && <span style={{ fontSize: 10, color: C.muted, fontFamily: "monospace" }}>{row.hora}</span>}
                          </div>
                          <span style={{ fontSize: 11, color: C.muted, marginLeft: isMobile ? 58 : 0 }}>{row.desc}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </main>

{/* ══ ADD LEAD MODAL ══ */}
      {addOpen && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.75)", backdropFilter: "blur(6px)", zIndex: 100, display: "flex", alignItems: isMobile ? "stretch" : "center", justifyContent: "center" }}
          onClick={e => e.target === e.currentTarget && setAddOpen(false)}>
          <div style={{ background: C.s1, border: isMobile ? "none" : `1px solid ${C.border2}`, borderRadius: isMobile ? 0 : 14, width: isMobile ? "100%" : 600, height: isMobile ? "100%" : "auto", maxHeight: isMobile ? "100%" : "90vh", overflowY: "auto" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span style={{ fontWeight: 700, fontSize: 15 }}>Novo Lead</span>
              <button onClick={() => setAddOpen(false)} style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 7, width: 28, height: 28, cursor: "pointer", color: C.muted, fontSize: 14 }}>✕</button>
            </div>
            <div style={{ padding: isMobile ? 16 : 20, display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 12 }}>
              <FInput label="Empresa *" value={form.company} onChange={ff("company")} placeholder="Acme Corp" />
              <FInput label="País" value={form.country} onChange={ff("country")} placeholder="US" />
              <FInput label="Contato" value={form.contact} onChange={ff("contact")} placeholder="John Smith" />
              <FInput label="Cargo" value={form.role} onChange={ff("role")} placeholder="Head of Talent" />
              <FInput label="Email" value={form.email} onChange={ff("email")} placeholder="john@acme.com" type="email" />
              <FInput label="LinkedIn" value={form.linkedin} onChange={ff("linkedin")} placeholder="linkedin.com/in/john" />
              <FInput label="Vagas Abertas" value={form.open_roles} onChange={ff("open_roles")} placeholder="25+" />
              <FInput label="Cargo Principal Contratando" value={form.main_role} onChange={ff("main_role")} placeholder="Senior Engineer" />
              <FInput label="Growth Signal" value={form.growth_signal} onChange={ff("growth_signal")} placeholder="grew headcount 22% in 2025" full />
              <FInput label="Remote Signal" value={form.remote_signal} onChange={ff("remote_signal")} placeholder="Remote-first / Global hiring" />
              <FSelect label="Match Type" value={form.match_type} onChange={ff("match_type")} options={["fully_matched", "partial_match"]} />
              <FSelect label="Indústria" value={form.industry} onChange={ff("industry")} options={["", "Technology", "Finance", "Healthcare", "E-commerce", "Manufacturing", "Consulting", "Media", "Education", "Logistics", "Other"]} />
              <FSelect label="Tamanho" value={form.size} onChange={ff("size")} options={["", "1-10", "11-50", "51-200", "201-500", "501-1000", "1000+"]} />
              <FSelect label="Fonte" value={form.source} onChange={ff("source")} options={["lessie", "apollo", "linkedin", "manual"]} />
              <div style={{ gridColumn: "1/-1", display: "flex", flexDirection: "column", gap: 4 }}>
                <label style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px" }}>Notas</label>
                <textarea value={form.notes} onChange={e => setForm(f => ({ ...f, notes: e.target.value }))} placeholder="Observações, pontos de dor..." rows={3}
                  style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 11px", color: C.text, fontSize: 13, fontFamily: "inherit", outline: "none", resize: "vertical" }} />
              </div>
            </div>
            <div style={{ padding: "12px 20px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <Btn onClick={() => setAddOpen(false)}>Cancelar</Btn>
              <Btn variant="primary" onClick={saveLead} disabled={saving}>{saving ? "Salvando..." : "💾 Salvar no Supabase"}</Btn>
            </div>
          </div>
        </div>
      )}

{/* ══ DETAIL MODAL ══ */}
      {detailOpen && sel && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.75)", backdropFilter: "blur(6px)", zIndex: 100, display: "flex", alignItems: isMobile ? "stretch" : "center", justifyContent: "center" }}
          onClick={e => e.target === e.currentTarget && setDetailOpen(false)}>
          <div style={{ background: C.s1, border: isMobile ? "none" : `1px solid ${C.border2}`, borderRadius: isMobile ? 0 : 14, width: isMobile ? "100%" : 680, height: isMobile ? "100%" : "auto", maxHeight: isMobile ? "100%" : "90vh", overflowY: "auto" }}>
            <div style={{ padding: isMobile ? "12px 16px" : "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flex: 1, minWidth: 0 }}>
                <Avatar name={sel.company} size={isMobile ? 36 : 40} />
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: isMobile ? 14 : 16, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sel.company}</div>
                  <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>{sel.contact} · {sel.country}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <select value={sel.status} onChange={e => updateStatus(sel.id, e.target.value as Status)}
                  style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 8, padding: "6px 10px", color: stageOf(sel.status).color, fontSize: 12, fontFamily: "inherit", outline: "none", cursor: "pointer" }}>
                  {STAGES.map(s => <option key={s.id} value={s.id} style={{ color: C.text }}>{s.label}</option>)}
                </select>
                <button onClick={() => setDetailOpen(false)} style={{ background: C.s3, border: `1px solid ${C.border}`, borderRadius: 7, width: 28, height: 28, cursor: "pointer", color: C.muted, fontSize: 14 }}>✕</button>
              </div>
            </div>

{/* Tabs */}
            <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, padding: "0 20px" }}>
              {["info", "history"].map(t => (
                <button key={t} onClick={() => setDetTab(t)}
                  style={{ padding: "10px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", background: "transparent", border: "none", borderBottom: detTab === t ? `2px solid ${C.cyan}` : "2px solid transparent", color: detTab === t ? C.cyan : C.muted, fontFamily: "inherit" }}>
                  {{ info: "Informacoes", history: "Historico" }[t]}
                </button>
              ))}
            </div>

<div style={{ padding: isMobile ? 16 : 20 }}>
              {detTab === "info" && (
                <div>
                  <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: isMobile ? 10 : 14, marginBottom: 14 }}>
                    {[
                      ["Email", sel.email || "—"],
                      ["LinkedIn", sel.linkedin ? "Ver perfil" : "—"],
                      ["Vagas Abertas", sel.open_roles || "—"],
                      ["Cargo Contratando", sel.main_role || "—"],
                      ["Remote Signal", sel.remote_signal || "—"],
                      ["Match Type", sel.match_type || "—"],
                      ["Indústria", sel.industry || "—"],
                      ["Tamanho", sel.size || "—"],
                      ["Fonte", sel.source || "—"],
                      ["Status", stageOf(sel.status).label],
                    ].map(([k, v]) => (
                      <div key={k}>
                        <div style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 3 }}>{k}</div>
                        <div style={{ fontSize: 13, fontWeight: 500, color: k === "LinkedIn" && sel.linkedin ? C.cyan : C.text }}>
                          {k === "LinkedIn" && sel.linkedin
                            ? <a
                                href={sel.linkedin.startsWith("http") ? sel.linkedin : `https://${sel.linkedin}`}
                                target="_blank"
                                rel="noopener noreferrer"
                                style={{ color: C.cyan, textDecoration: "none" }}
                              >Ver perfil →</a>
                            : v}
                        </div>
                      </div>
                    ))}
                  </div>
                  {sel.growth_signal && (
                    <div style={{ background: C.s3, borderRadius: 8, padding: "10px 12px", marginBottom: 10 }}>
                      <div style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 4 }}>Growth Signal</div>
                      <div style={{ fontSize: 12, color: C.green }}>{sel.growth_signal}</div>
                    </div>
                  )}
                  {sel.notes && <div style={{ background: C.s3, borderRadius: 8, padding: "10px 12px", fontSize: 12, color: C.muted, lineHeight: 1.6 }}>{sel.notes}</div>}
                  <div style={{ marginTop: 14, display: "flex", gap: 8 }}>
                    <Btn variant="danger" onClick={() => deleteLead(sel.id)}>🗑 Remover</Btn>
                  </div>
                </div>
              )}

{detTab === "history" && (
                <div style={{ display: "flex", flexDirection: "column", gap: 10, maxHeight: 400, overflowY: "auto" }}>
                  {activities.length === 0 && <div style={{ color: C.muted, fontSize: 12, textAlign: "center", padding: 20 }}>Nenhuma atividade registrada.</div>}
                  {activities.map(a => {
                    const isEmailSent = a.type === "email_sent"
                    const isExpanded = expandedActivity === a.id
                    return (
                      <div key={a.id}>
                        <div 
                          onClick={() => isEmailSent && fetchEmailContentForActivity(a.id, a.lead_id, a.description || "")}
                          style={{ 
                            display: "flex", 
                            gap: 10, 
                            padding: "8px 10px",
                            borderRadius: 8,
                            cursor: isEmailSent ? "pointer" : "default",
                            background: isExpanded ? C.s3 : "transparent",
                            border: isEmailSent ? `1px solid ${isExpanded ? C.cyan : C.border}` : "1px solid transparent",
                            transition: "all 0.15s ease"
                          }}
                        >
                          <div style={{ 
                            width: 10, 
                            height: 10, 
                            borderRadius: "50%", 
                            background: isEmailSent ? C.green : C.cyan, 
                            flexShrink: 0, 
                            marginTop: 3 
                          }} />
                          <div style={{ flex: 1 }}>
                            <div style={{ fontSize: 12, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                              {a.description || a.type}
                              {isEmailSent && (
                                <span style={{ fontSize: 10, color: C.muted }}>
                                  {isExpanded ? "[-]" : "[+]"}
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: 10, color: C.muted, fontFamily: "monospace", marginTop: 2 }}>
                              {new Date(a.created_at).toLocaleString("pt-BR")}
                            </div>
                          </div>
                        </div>
                        
                        {/* Expandable email content */}
                        {isExpanded && isEmailSent && (
                          <div style={{ marginLeft: 20, marginTop: 8, marginBottom: 8 }}>
                            {loadingEmailContent ? (
                              <div style={{ color: C.muted, fontSize: 12, padding: 12 }}>Carregando email...</div>
                            ) : activityEmailContent ? (
                              <div style={{ background: C.s2, borderRadius: 8, border: `1px solid ${C.border}`, overflow: "hidden" }}>
                                <div style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}` }}>
                                  <div style={{ fontSize: 9, fontFamily: "monospace", color: C.muted, textTransform: "uppercase", letterSpacing: "1px", marginBottom: 4 }}>Assunto</div>
                                  <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{activityEmailContent.subject}</div>
                                </div>
                                <pre style={{ 
                                  background: "#0a0d12", 
                                  margin: 0, 
                                  padding: "14px 16px", 
                                  fontSize: 12, 
                                  fontFamily: "'JetBrains Mono', 'Fira Code', monospace", 
                                  color: C.text, 
                                  whiteSpace: "pre-wrap", 
                                  lineHeight: 1.7,
                                  maxHeight: 200,
                                  overflowY: "auto"
                                }}>
                                  {activityEmailContent.body}
                                </pre>
                              </div>
                            ) : null}
                          </div>
                        )}
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ══ TOAST ══ */}
      {toast && (
        <div style={{ position: "fixed", bottom: 22, right: 22, background: C.s1, border: `1px solid ${C.green}`, borderRadius: 10, padding: "12px 18px", fontSize: 13, zIndex: 200, boxShadow: "0 8px 28px rgba(0,0,0,.5)" }}>
          {toast}
        </div>
      )}

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }
        ::-webkit-scrollbar { width: 5px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 3px; }
      `}</style>
    </div>
  )
}
