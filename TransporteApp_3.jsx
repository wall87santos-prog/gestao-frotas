import React, { useState, useEffect, useMemo, useRef } from "react";
import { Truck, Users, HardHat, ClipboardList, LayoutDashboard, Plus, X, Search, Camera, FileText, Trash2, Pencil, ArrowRightLeft, AlertTriangle, CheckCircle2, Clock, ChevronRight, BarChart3, Download, ShieldCheck, ChevronDown, Lock, Upload, FileSpreadsheet, ArrowLeft, ArrowRight, CircleAlert, ScanLine, Loader2, Sparkles, LogOut, Building2, Mail, KeyRound, UserPlus, CreditCard, Printer, Wrench, Fuel, AlertOctagon, DollarSign, Gauge, TrendingUp, Menu } from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, LineChart, Line, Legend } from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------
   Perfis de acesso (espelha as policies do banco em schema.sql)
   admin     -> cadastra, edita, exclui, vê relatórios
   operador  -> cadastra e registra movimentação, não edita/exclui
   consulta  -> só visualiza
--------------------------------------------------------- */
const ROLES = {
  admin: { label: "Administrador", canCreate: true, canEdit: true, canDelete: true },
  operador: { label: "Operador de pátio", canCreate: true, canEdit: false, canDelete: false },
  consulta: { label: "Consulta", canCreate: false, canEdit: false, canDelete: false },
};

/* ---------------------------------------------------------
   Tokens
   BG_APP #10151C · SIDEBAR #151C26 · SURFACE #FFFFFF
   ACCENT (safety amber) #FF8A1E · ACCENT_DEEP #C9600A
   OK #1F9D6B · WARN #E0A400 · DANGER #E0503C
   INK #10151C · MUTED #6B7480 · LINE #E7E9EC
--------------------------------------------------------- */

const uid = () => Math.random().toString(36).slice(2, 10);
const todayISO = () => new Date().toISOString().slice(0, 10);

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  const now = new Date(todayISO() + "T00:00:00");
  return Math.round((d - now) / 86400000);
}

function docStatus(dateStr) {
  if (!dateStr) return { label: "Sem data", tone: "muted" };
  const days = daysUntil(dateStr);
  if (days < 0) return { label: "Vencido", tone: "danger" };
  if (days <= 30) return { label: `Vence em ${days}d`, tone: "warn" };
  return { label: "Em dia", tone: "ok" };
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

// Redimensiona/comprime a foto antes de enviar para leitura por IA
// (fotos de celular costumam vir grandes demais para a chamada)
function fileToCompressedDataUrl(file, maxDim = 1024, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          const scale = maxDim / Math.max(width, height);
          width = Math.round(width * scale);
          height = Math.round(height * scale);
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = () => reject(new Error("Não consegui abrir essa imagem."));
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/* ---------------------------------------------------------
   OCR de documentos (lê a foto e extrai os dados via IA)
--------------------------------------------------------- */

const OCR_PROMPTS = {
  crlv: `Você está vendo a foto de um documento de veículo (CRLV) brasileiro. Extraia os dados visíveis e responda APENAS com um JSON válido, sem nenhum texto antes ou depois, exatamente neste formato: {"placa": "", "modelo": "", "ano": "", "numero_documento": "", "validade": ""}. A validade deve estar em formato ISO AAAA-MM-DD (se não houver validade explícita, deixe vazio). Se algum campo não estiver legível, deixe como string vazia.`,
  cnh: `Tarefa: transcrição de texto impresso (OCR simples), NÃO reconhecimento facial nem verificação de identidade. Contexto: sistema interno de gestão de frota de uma transportadora, cadastrando dados administrativos de motoristas funcionários, com consentimento deles, para controle de validade de habilitação.

A imagem é a foto de uma Carteira Nacional de Habilitação (CNH) brasileira. Ignore completamente a foto/rosto da pessoa — não descreva, não comente, não analise. Sua única tarefa é ler e transcrever o texto impresso nos campos administrativos do documento, como faria um leitor de OCR. Responda APENAS com um JSON válido, sem nenhum texto antes ou depois, exatamente neste formato: {"nome": "", "numero_registro": "", "categoria": "", "validade": ""}. A validade deve estar em formato ISO AAAA-MM-DD. Se algum campo não estiver legível, deixe como string vazia.`,
  generico: `Você está vendo a foto de um documento. Extraia, se houver, o número do documento e a data de validade. Responda APENAS com um JSON válido, sem nenhum texto antes ou depois, exatamente neste formato: {"numero": "", "validade": ""}. A validade deve estar em formato ISO AAAA-MM-DD.`,
};

async function callClaudeVision(mediaType, base64, prompt) {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      messages: [
        {
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
            { type: "text", text: prompt },
          ],
        },
      ],
    }),
  });
  if (!response.ok) {
    let detail = "";
    try { detail = (await response.json())?.error?.message || ""; } catch {}
    const err = new Error(`A API respondeu ${response.status}${detail ? " — " + detail : ""}`);
    err.status = response.status;
    throw err;
  }
  return response.json();
}

async function extractDocumentData(dataUrl, kind) {
  const match = dataUrl.match(/^data:(.*?);base64,(.*)$/);
  if (!match) throw new Error("Não consegui ler o arquivo da foto.");
  const [, mediaType, base64] = match;
  const prompt = OCR_PROMPTS[kind] || OCR_PROMPTS.generico;

  const maxAttempts = 3;
  let data;
  let lastErr;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      data = await callClaudeVision(mediaType, base64, prompt);
      lastErr = null;
      break;
    } catch (err) {
      lastErr = err;
      const isRateLimit = err.status === 429;
      const transient = !err.status || err.status >= 500 || isRateLimit;
      if (!transient || attempt === maxAttempts) break;
      await new Promise((r) => setTimeout(r, isRateLimit ? 4000 : 900 * attempt));
    }
  }
  if (lastErr) {
    if (lastErr.status === 429) {
      throw new Error("O limite de uso da IA foi atingido no momento. Espere um pouco (cerca de 1 minuto) e tente novamente, ou preencha os campos manualmente.");
    }
    if (!lastErr.status) {
      throw new Error("Não foi possível conectar à IA agora (instabilidade momentânea do ambiente). Tente novamente em instantes ou preencha os campos manualmente.");
    }
    throw new Error(lastErr?.message || "Falha de conexão ao chamar a IA.");
  }

  const text = (data.content || []).map((b) => b.text || "").join("").trim();
  if (!text) throw new Error("A IA não retornou texto (resposta vazia).");

  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end === -1 || end < start) {
    const looksLikeRefusal = /not able to|não posso|não consigo|cannot extract|i'm unable/i.test(text);
    if (looksLikeRefusal) {
      throw new Error("A IA não conseguiu processar esta foto de documento pessoal. Tente novamente com outra foto (bem iluminada, só o documento) ou preencha os campos manualmente.");
    }
    throw new Error("Resposta da IA não veio em formato JSON: " + text.slice(0, 120));
  }
  const jsonSlice = text.slice(start, end + 1);
  try {
    return JSON.parse(jsonSlice);
  } catch (err) {
    throw new Error("Não consegui interpretar a resposta da IA.");
  }
}

function OcrButton({ onFile, status, title = "Ler documento com IA" }) {
  const ref = useRef(null);
  return (
    <>
      <button
        type="button"
        onClick={() => ref.current?.click()}
        disabled={status?.loading}
        className="p-2 rounded-lg flex items-center justify-center shrink-0"
        style={{ background: "#F1E9FB", color: "#7C3AED" }}
        title={title}
      >
        {status?.loading ? <Loader2 size={13} className="animate-spin" /> : <ScanLine size={13} />}
      </button>
      <input
        ref={ref}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ""; }}
      />
    </>
  );
}

/* ---------------------------------------------------------
   Small UI atoms
--------------------------------------------------------- */

function StatusPill({ tone, children }) {
  const map = {
    ok: { bg: "#E7F7EF", fg: "#1F9D6B", dot: "#1F9D6B" },
    warn: { bg: "#FFF6E0", fg: "#8A5F00", dot: "#E0A400" },
    danger: { bg: "#FDEBE8", fg: "#B3372A", dot: "#E0503C" },
    muted: { bg: "#EEF0F2", fg: "#6B7480", dot: "#9AA1AA" },
  };
  const s = map[tone] || map.muted;
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-medium"
      style={{ background: s.bg, color: s.fg }}
    >
      <span className="w-1.5 h-1.5 rounded-full" style={{ background: s.dot }} />
      {children}
    </span>
  );
}

function Avatar({ src, fallback, size = 44 }) {
  return (
    <div
      className="rounded-xl overflow-hidden flex items-center justify-center shrink-0"
      style={{ width: size, height: size, background: "#EEF0F2", border: "1px solid #E7E9EC" }}
    >
      {src ? (
        <img src={src} alt="" className="w-full h-full object-cover" />
      ) : (
        <span className="text-sm font-semibold" style={{ color: "#9AA1AA" }}>
          {fallback}
        </span>
      )}
    </div>
  );
}

function PhotoPicker({ label, value, onChange }) {
  const inputRef = useRef(null);
  return (
    <div>
      <p className="text-xs font-medium mb-1.5" style={{ color: "#6B7480" }}>{label}</p>
      <div className="flex items-center gap-3">
        <Avatar src={value} fallback={<Camera size={16} />} size={56} />
        <div className="flex gap-2">
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            className="text-xs font-medium px-3 py-1.5 rounded-lg"
            style={{ background: "#EEF0F2", color: "#10151C" }}
          >
            {value ? "Trocar foto" : "Adicionar foto"}
          </button>
          {value && (
            <button
              type="button"
              onClick={() => onChange(null)}
              className="text-xs font-medium px-3 py-1.5 rounded-lg"
              style={{ background: "#FDEBE8", color: "#B3372A" }}
            >
              Remover
            </button>
          )}
        </div>
        <input
          ref={inputRef}
          type="file"
          accept="image/*"
          className="hidden"
          onChange={async (e) => {
            const f = e.target.files?.[0];
            if (f) onChange(await fileToDataUrl(f));
            e.target.value = "";
          }}
        />
      </div>
    </div>
  );
}

function MultiPhotoPicker({ label, values = [], onChange, angleHints }) {
  const inputRef = useRef(null);

  const addFiles = async (fileList) => {
    const files = Array.from(fileList || []);
    if (files.length === 0) return;
    const dataUrls = await Promise.all(files.map(fileToDataUrl));
    onChange([...values, ...dataUrls.map((src, i) => ({ id: uid(), src, angulo: angleHints?.[values.length + i] || "" }))]);
  };
  const removeAt = (id) => onChange(values.filter((p) => p.id !== id));
  const setAngle = (id, angulo) => onChange(values.map((p) => (p.id === id ? { ...p, angulo } : p)));

  return (
    <div>
      <div className="flex items-center justify-between mb-1.5">
        <p className="text-xs font-medium" style={{ color: "#6B7480" }}>{label}</p>
        <span className="text-xs" style={{ color: "#9AA1AA" }}>{values.length} foto{values.length === 1 ? "" : "s"}</span>
      </div>
      <div className="grid grid-cols-3 sm:grid-cols-4 gap-2.5">
        {values.map((p) => (
          <div key={p.id} className="relative group">
            <div className="rounded-lg overflow-hidden" style={{ aspectRatio: "4/3", background: "#EEF0F2", border: "1px solid #E7E9EC" }}>
              <img src={p.src} alt="" className="w-full h-full object-cover" />
            </div>
            <input
              value={p.angulo}
              onChange={(e) => setAngle(p.id, e.target.value)}
              placeholder="Ângulo (ex: frente)"
              className="w-full mt-1 text-[11px] rounded-md px-1.5 py-1 outline-none"
              style={{ background: "#F5F6F8", border: "1px solid #E7E9EC", color: "#10151C" }}
            />
            <button
              type="button"
              onClick={() => removeAt(p.id)}
              className="absolute top-1 right-1 w-5 h-5 rounded-full flex items-center justify-center"
              style={{ background: "rgba(16,21,28,0.7)" }}
            >
              <X size={11} color="#FFFFFF" />
            </button>
          </div>
        ))}
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="rounded-lg flex flex-col items-center justify-center gap-1"
          style={{ aspectRatio: "4/3", background: "#F5F6F8", border: "1px dashed #C7CBD1" }}
        >
          <Camera size={16} style={{ color: "#9AA1AA" }} />
          <span className="text-[11px] font-medium" style={{ color: "#9AA1AA" }}>Adicionar</span>
        </button>
      </div>
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        multiple
        className="hidden"
        onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }}
      />
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium mb-1.5" style={{ color: "#6B7480" }}>{label}</span>
      {children}
    </label>
  );
}

const inputClass =
  "w-full rounded-lg px-3 py-2 text-sm outline-none transition-colors";
const inputStyle = { background: "#F5F6F8", border: "1px solid #E7E9EC", color: "#10151C" };

function TextInput(props) {
  return <input {...props} className={inputClass} style={inputStyle} />;
}
function SelectInput(props) {
  return <select {...props} className={inputClass} style={inputStyle} />;
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div
      className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-4 overflow-y-auto"
      style={{ background: "rgba(16,21,28,0.55)" }}
      onClick={onClose}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        className={`w-full ${wide ? "max-w-2xl" : "max-w-md"} rounded-2xl my-8`}
        style={{ background: "#FFFFFF" }}
      >
        <div className="flex items-center justify-between px-6 py-4" style={{ borderBottom: "1px solid #E7E9EC" }}>
          <h3 className="text-base font-semibold" style={{ color: "#10151C" }}>{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-lg" style={{ background: "#F5F6F8" }}>
            <X size={16} />
          </button>
        </div>
        <div className="px-6 py-5">{children}</div>
      </div>
    </div>
  );
}

function EmptyState({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-16 px-6">
      <div className="w-12 h-12 rounded-full flex items-center justify-center mb-3" style={{ background: "#EEF0F2" }}>
        <Icon size={20} style={{ color: "#9AA1AA" }} />
      </div>
      <p className="text-sm font-medium" style={{ color: "#10151C" }}>{title}</p>
      <p className="text-xs mt-1" style={{ color: "#9AA1AA" }}>{subtitle}</p>
    </div>
  );
}

/* ---------------------------------------------------------
   Main App
--------------------------------------------------------- */

const NAV = [
  { key: "dashboard", label: "Painel", icon: LayoutDashboard },
  { key: "trucks", label: "Caminhões", icon: Truck },
  { key: "drivers", label: "Motoristas", icon: Users },
  { key: "helpers", label: "Ajudantes", icon: HardHat },
  { key: "loadings", label: "Carregamentos", icon: ClipboardList },
  { key: "access", label: "Controle de Acesso", icon: ScanLine },
  { key: "maintenance", label: "Manutenção", icon: Wrench },
  { key: "fuel", label: "Abastecimento", icon: Fuel },
  { key: "fines", label: "Multas", icon: AlertOctagon },
  { key: "reports", label: "Relatórios", icon: BarChart3 },
];

const emptyMaintenance = {
  id: "",
  placa: "",
  tipo: "Preventiva",
  descricao: "",
  data: "",
  km: "",
  custo: "",
  oficina: "",
  proximaData: "",
  proximoKm: "",
};

const emptyFuel = {
  id: "",
  placa: "",
  motorista: "",
  data: "",
  litros: "",
  valorTotal: "",
  kmAtual: "",
  posto: "",
};

const emptyFine = {
  id: "",
  placa: "",
  motorista: "",
  data: "",
  infracao: "",
  valor: "",
  pontos: "",
  vencimento: "",
  status: "Pendente",
};

function maintenanceAlertStatus(m) {
  // alerta por data prevista ou por km restante até a próxima manutenção
  if (m.proximaData) {
    const s = docStatus(m.proximaData);
    if (s.tone !== "ok") return s;
  }
  return null;
}

/* ---------------------------------------------------------
   Autenticação (protótipo)
   Simula o modelo de empresas.sql: uma empresa (tenant) com
   usuários vinculados por papel. Senha não é criptografada de
   verdade — isso é só para demonstrar o fluxo; em produção o
   login roda no backend (Supabase Auth), como no schema.sql.
--------------------------------------------------------- */

function encodePass(str) {
  try { return btoa(unescape(encodeURIComponent(str))); } catch { return btoa(str); }
}
function decodePass(str) {
  try { return decodeURIComponent(escape(atob(str))); } catch { return atob(str); }
}

async function storageSetWithRetry(key, value, shared, tries = 3) {
  let lastErr;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await window.storage.set(key, value, shared);
      if (res) return res;
      lastErr = new Error("Resposta vazia do armazenamento");
    } catch (err) {
      lastErr = err;
    }
    await new Promise((r) => setTimeout(r, 500 + i * 400));
  }
  throw new Error(`Falha ao salvar "${key}": ${lastErr?.message || "erro desconhecido"}`);
}

async function loadAuthData() {
  try {
    const res = await window.storage.get("auth_dados", false);
    if (res?.value) {
      const parsed = JSON.parse(res.value);
      return { empresas: parsed.empresas || [], usuarios: parsed.usuarios || [] };
    }
  } catch {}
  return { empresas: [], usuarios: [] };
}

function AuthScreen({ onAuthenticated }) {
  const [mode, setMode] = useState("login"); // login | signup
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const [loginForm, setLoginForm] = useState({ email: "", senha: "" });
  const [signupForm, setSignupForm] = useState({ empresa: "", nome: "", email: "", senha: "", confirmar: "" });

  const handleLogin = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      const { empresas, usuarios } = await loadAuthData();
      const email = loginForm.email.trim().toLowerCase();
      const usuario = usuarios.find((u) => u.email.toLowerCase() === email);
      if (!usuario || decodePass(usuario.senha) !== loginForm.senha) {
        setError("E-mail ou senha incorretos.");
        setLoading(false);
        return;
      }
      const empresa = empresas.find((e) => e.id === usuario.empresaId);
      await storageSetWithRetry("sessao", JSON.stringify({ usuarioId: usuario.id }), false);
      onAuthenticated({ usuario, empresa });
    } catch (err) {
      console.error("Erro ao entrar:", err);
      setError("Não consegui verificar seu login agora: " + (err?.message || "erro desconhecido"));
    }
    setLoading(false);
  };

  const handleSignup = async (e) => {
    e.preventDefault();
    setError("");
    if (!signupForm.empresa || !signupForm.nome || !signupForm.email || !signupForm.senha) {
      setError("Preencha todos os campos.");
      return;
    }
    if (signupForm.senha.length < 4) {
      setError("A senha precisa ter pelo menos 4 caracteres.");
      return;
    }
    if (signupForm.senha !== signupForm.confirmar) {
      setError("As senhas não conferem.");
      return;
    }
    setLoading(true);
    try {
      const { empresas, usuarios } = await loadAuthData();
      const email = signupForm.email.trim().toLowerCase();
      if (usuarios.some((u) => u.email.toLowerCase() === email)) {
        setError("Já existe uma conta com esse e-mail.");
        setLoading(false);
        return;
      }
      const empresa = { id: uid(), nome: signupForm.empresa.trim(), criadoEm: todayISO() };
      const usuario = {
        id: uid(),
        empresaId: empresa.id,
        nome: signupForm.nome.trim(),
        email,
        senha: encodePass(signupForm.senha),
        papel: "admin",
      };
      await storageSetWithRetry("auth_dados", JSON.stringify({ empresas: [...empresas, empresa], usuarios: [...usuarios, usuario] }), false);
      await storageSetWithRetry("sessao", JSON.stringify({ usuarioId: usuario.id }), false);
      onAuthenticated({ usuario, empresa });
    } catch (err) {
      console.error("Erro ao criar conta:", err);
      setError("Não consegui criar sua conta agora: " + (err?.message || "erro desconhecido"));
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen w-full flex" style={{ background: "#151C26" }}>
      <div className="hidden lg:flex flex-1 flex-col justify-between p-12" style={{ background: "linear-gradient(155deg, #151C26 0%, #1B2531 60%, #20293A 100%)" }}>
        <div className="flex items-center gap-2.5">
          <div className="w-9 h-9 rounded-lg flex items-center justify-center" style={{ background: "#FF8A1E" }}>
            <Truck size={18} color="#151C26" strokeWidth={2.5} />
          </div>
          <p className="font-display font-semibold text-lg" style={{ color: "#FFFFFF" }}>Frota+</p>
        </div>
        <div className="max-w-sm">
          <p className="font-display text-3xl font-semibold leading-snug" style={{ color: "#FFFFFF" }}>
            Gestão completa da sua transportadora, em um só lugar.
          </p>
          <p className="text-sm mt-3" style={{ color: "#8891A0" }}>
            Caminhões, motoristas, ajudantes, documentos e controle de carregamento — cada empresa com seus próprios dados e acessos.
          </p>
        </div>
        <p className="text-xs" style={{ color: "#5B6472" }}>© {new Date().getFullYear()} Frota+</p>
      </div>

      <div className="flex-1 flex items-center justify-center p-6">
        <div className="w-full max-w-sm">
          <div className="lg:hidden flex items-center gap-2.5 mb-8 justify-center">
            <div className="w-8 h-8 rounded-lg flex items-center justify-center" style={{ background: "#FF8A1E" }}>
              <Truck size={16} color="#151C26" strokeWidth={2.5} />
            </div>
            <p className="font-display font-semibold text-base" style={{ color: "#FFFFFF" }}>Frota+</p>
          </div>

          <div className="rounded-2xl p-6" style={{ background: "#FFFFFF" }}>
            <div className="flex rounded-lg p-1 mb-5" style={{ background: "#EEF0F2" }}>
              {[{ k: "login", l: "Entrar" }, { k: "signup", l: "Criar empresa" }].map((t) => (
                <button
                  key={t.k}
                  onClick={() => { setMode(t.k); setError(""); }}
                  className="flex-1 text-sm font-medium py-2 rounded-md"
                  style={{ background: mode === t.k ? "#FFFFFF" : "transparent", color: mode === t.k ? "#10151C" : "#6B7480" }}
                >
                  {t.l}
                </button>
              ))}
            </div>

            {mode === "login" ? (
              <form onSubmit={handleLogin} className="space-y-3.5">
                <Field label="E-mail">
                  <TextInput type="email" required value={loginForm.email} onChange={(e) => setLoginForm({ ...loginForm, email: e.target.value })} placeholder="voce@empresa.com" />
                </Field>
                <Field label="Senha">
                  <TextInput type="password" required value={loginForm.senha} onChange={(e) => setLoginForm({ ...loginForm, senha: e.target.value })} placeholder="••••••" />
                </Field>
                {error && <p className="flex items-center gap-1.5 text-xs" style={{ color: "#B3372A" }}><CircleAlert size={13} /> {error}</p>}
                <button type="submit" disabled={loading} className="w-full text-sm font-medium py-2.5 rounded-lg mt-1" style={{ background: "#FF8A1E", color: "#151C26" }}>
                  {loading ? "Entrando…" : "Entrar"}
                </button>
              </form>
            ) : (
              <form onSubmit={handleSignup} className="space-y-3.5">
                <Field label="Nome da empresa">
                  <TextInput required value={signupForm.empresa} onChange={(e) => setSignupForm({ ...signupForm, empresa: e.target.value })} placeholder="Sua transportadora" />
                </Field>
                <Field label="Seu nome">
                  <TextInput required value={signupForm.nome} onChange={(e) => setSignupForm({ ...signupForm, nome: e.target.value })} placeholder="Nome completo" />
                </Field>
                <Field label="E-mail">
                  <TextInput type="email" required value={signupForm.email} onChange={(e) => setSignupForm({ ...signupForm, email: e.target.value })} placeholder="voce@empresa.com" />
                </Field>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="Senha">
                    <TextInput type="password" required value={signupForm.senha} onChange={(e) => setSignupForm({ ...signupForm, senha: e.target.value })} placeholder="••••••" />
                  </Field>
                  <Field label="Confirmar">
                    <TextInput type="password" required value={signupForm.confirmar} onChange={(e) => setSignupForm({ ...signupForm, confirmar: e.target.value })} placeholder="••••••" />
                  </Field>
                </div>
                {error && <p className="flex items-center gap-1.5 text-xs" style={{ color: "#B3372A" }}><CircleAlert size={13} /> {error}</p>}
                <button type="submit" disabled={loading} className="w-full flex items-center justify-center gap-1.5 text-sm font-medium py-2.5 rounded-lg mt-1" style={{ background: "#FF8A1E", color: "#151C26" }}>
                  <UserPlus size={15} /> {loading ? "Criando…" : "Criar conta e entrar"}
                </button>
              </form>
            )}
          </div>
          <button
            onClick={() => onAuthenticated({
              usuario: { id: "demo", nome: "Visitante", email: "", papel: "admin" },
              empresa: { id: "demo", nome: "Modo demonstração" },
              demo: true,
            })}
            className="w-full text-center text-xs font-medium mt-3 py-2"
            style={{ color: "#8891A0" }}
          >
            Continuar sem criar conta (modo teste, não salva)
          </button>
          <p className="text-center text-[11px] mt-1" style={{ color: "#5B6472" }}>
            Protótipo — login e dados ficam salvos neste navegador, sem criptografia real.
          </p>
        </div>
      </div>
    </div>
  );
}

export default function TransporteApp() {
  const [tab, setTab] = useState("dashboard");
  const [trucks, setTrucks] = useState([]);
  const [drivers, setDrivers] = useState([]);
  const [helpers, setHelpers] = useState([]);
  const [loadings, setLoadings] = useState([]);
  const [maintenances, setMaintenances] = useState([]);
  const [fuels, setFuels] = useState([]);
  const [fines, setFines] = useState([]);
  const [companyLogo, setCompanyLogo] = useState(null);
  const [ready, setReady] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [role, setRole] = useState("admin");
  const perms = ROLES[role];
  const [mobileNavOpen, setMobileNavOpen] = useState(false);

  const [authReady, setAuthReady] = useState(false);
  const [session, setSession] = useState(null); // { usuario, empresa }

  useEffect(() => {
    (async () => {
      try {
        const s = await window.storage.get("sessao", false).catch(() => null);
        if (s?.value) {
          const { usuarioId } = JSON.parse(s.value);
          const { empresas, usuarios } = await loadAuthData();
          const usuario = usuarios.find((u) => u.id === usuarioId);
          const empresa = usuario ? empresas.find((e) => e.id === usuario.empresaId) : null;
          if (usuario && empresa) {
            setSession({ usuario, empresa });
            setRole(usuario.papel);
          }
        }
      } catch {}
      setAuthReady(true);
    })();
  }, []);

  const handleAuthenticated = ({ usuario, empresa }) => {
    setSession({ usuario, empresa });
    setRole(usuario.papel);
  };

  const handleLogout = async () => {
    try { await window.storage.set("sessao", "", false); } catch {}
    setSession(null);
    setTab("dashboard");
  };

  useEffect(() => {
    if (!session) return;
    (async () => {
      try {
        const keys = ["trucks", "drivers", "helpers", "loadings", "maintenances", "fuels", "fines", "companyLogo"];
        const results = await Promise.all(
          keys.map(async (k) => {
            try {
              const r = await window.storage.get(k, false);
              return r ? JSON.parse(r.value) : (k === "companyLogo" ? null : []);
            } catch {
              return k === "companyLogo" ? null : [];
            }
          })
        );
        setTrucks(results[0]);
        setDrivers(results[1]);
        setHelpers(results[2]);
        setLoadings(results[3]);
        setMaintenances(results[4]);
        setFuels(results[5]);
        setFines(results[6]);
        setCompanyLogo(results[7]);
      } catch {
        setSaveError(true);
      }
      setReady(true);
    })();
  }, [session]);

  const persist = async (key, value) => {
    try {
      const res = await storageSetWithRetry(key, JSON.stringify(value), false, 4);
      if (!res) {
        setSaveError(true);
      } else {
        setSaveError(false);
      }
    } catch {
      setSaveError(true);
    }
  };

  // se uma tentativa de salvar falhar, tenta de novo sozinho em segundo plano
  // (rede instável costuma se resolver em alguns segundos)
  useEffect(() => {
    if (!saveError || !ready) return;
    const t = setTimeout(() => { retryAllPersist(); }, 6000);
    return () => clearTimeout(t);
  }, [saveError, ready]);

  useEffect(() => { if (ready) persist("trucks", trucks); }, [trucks, ready]);
  useEffect(() => { if (ready) persist("drivers", drivers); }, [drivers, ready]);
  useEffect(() => { if (ready) persist("helpers", helpers); }, [helpers, ready]);
  useEffect(() => { if (ready) persist("loadings", loadings); }, [loadings, ready]);
  useEffect(() => { if (ready) persist("maintenances", maintenances); }, [maintenances, ready]);
  useEffect(() => { if (ready) persist("fuels", fuels); }, [fuels, ready]);
  useEffect(() => { if (ready) persist("fines", fines); }, [fines, ready]);
  useEffect(() => { if (ready) persist("companyLogo", companyLogo); }, [companyLogo, ready]);

  const retryAllPersist = async () => {
    setSaveError(false);
    await Promise.all([
      persist("trucks", trucks),
      persist("drivers", drivers),
      persist("helpers", helpers),
      persist("loadings", loadings),
      persist("maintenances", maintenances),
      persist("fuels", fuels),
      persist("fines", fines),
      persist("companyLogo", companyLogo),
    ]);
  };

  const globalStyles = (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');
      .font-display { font-family: 'Space Grotesk', system-ui, sans-serif; }
      ::selection { background: #FFE3C2; }
      input:focus, select:focus, textarea:focus { border-color: #FF8A1E !important; background: #FFFFFF !important; }
      @media print {
        body * { visibility: hidden !important; }
        .badge-print-area, .badge-print-area * { visibility: visible !important; }
        .badge-print-sheet, .badge-print-sheet * { visibility: visible !important; }
        .badge-print-area { position: fixed !important; top: 50%; left: 50%; transform: translate(-50%, -50%) !important; box-shadow: none !important; }
        .badge-print-sheet { position: fixed !important; top: 0; left: 0; margin: 0 !important; box-shadow: none !important; width: 100% !important; }
        .badge-print-sheet .badge-print-area { position: static !important; transform: none !important; }
      }
    `}</style>
  );

  if (!authReady) {
    return (
      <div className="w-full min-h-screen flex items-center justify-center" style={{ background: "#151C26" }}>
        {globalStyles}
        <Loader2 size={22} className="animate-spin" style={{ color: "#FF8A1E" }} />
      </div>
    );
  }

  if (!session) {
    return (
      <div style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
        {globalStyles}
        <AuthScreen onAuthenticated={handleAuthenticated} />
      </div>
    );
  }

  return (
    <div className="w-full min-h-screen flex" style={{ background: "#F5F6F8", fontFamily: "'Inter', system-ui, sans-serif" }}>
      {globalStyles}

      <Sidebar
        tab={tab}
        setTab={setTab}
        counts={{ trucks: trucks.length, drivers: drivers.length, helpers: helpers.length, loadings: loadings.length, maintenance: maintenances.length, fuel: fuels.length, fines: fines.length }}
        role={role}
        setRole={setRole}
        session={session}
        onLogout={handleLogout}
        companyLogo={companyLogo}
        setCompanyLogo={setCompanyLogo}
      />

      {mobileNavOpen && (
        <MobileNav
          tab={tab}
          setTab={(k) => { setTab(k); setMobileNavOpen(false); }}
          counts={{ trucks: trucks.length, drivers: drivers.length, helpers: helpers.length, loadings: loadings.length, maintenance: maintenances.length, fuel: fuels.length, fines: fines.length }}
          role={role}
          session={session}
          onLogout={() => { setMobileNavOpen(false); handleLogout(); }}
          onClose={() => setMobileNavOpen(false)}
          companyLogo={companyLogo}
          setCompanyLogo={setCompanyLogo}
        />
      )}

      <main className="flex-1 min-w-0">
        <TopBar tab={tab} saveError={saveError} role={role} setTab={setTab} onOpenMenu={() => setMobileNavOpen(true)} onRetrySave={retryAllPersist} />
        <div className="max-w-6xl mx-auto px-6 sm:px-8 py-6">
          {!ready ? (
            <div className="py-24 text-center text-sm" style={{ color: "#9AA1AA" }}>Carregando dados…</div>
          ) : tab === "dashboard" ? (
            <Dashboard trucks={trucks} drivers={drivers} helpers={helpers} loadings={loadings} maintenances={maintenances} fuels={fuels} fines={fines} setTab={setTab} />
          ) : tab === "trucks" ? (
            <TrucksView trucks={trucks} setTrucks={setTrucks} perms={perms} />
          ) : tab === "drivers" ? (
            <DriversView drivers={drivers} setDrivers={setDrivers} perms={perms} empresaNome={session?.empresa?.nome} companyLogo={companyLogo} />
          ) : tab === "helpers" ? (
            <HelpersView helpers={helpers} setHelpers={setHelpers} perms={perms} empresaNome={session?.empresa?.nome} companyLogo={companyLogo} />
          ) : tab === "loadings" ? (
            <LoadingsView loadings={loadings} setLoadings={setLoadings} trucks={trucks} drivers={drivers} perms={perms} />
          ) : tab === "access" ? (
            <AccessControlView
              loadings={loadings}
              setLoadings={setLoadings}
              trucks={trucks}
              drivers={drivers}
              helpers={helpers}
              perms={perms}
            />
          ) : tab === "maintenance" ? (
            <MaintenanceView maintenances={maintenances} setMaintenances={setMaintenances} trucks={trucks} perms={perms} />
          ) : tab === "fuel" ? (
            <FuelView fuels={fuels} setFuels={setFuels} trucks={trucks} drivers={drivers} perms={perms} />
          ) : tab === "fines" ? (
            <FinesView fines={fines} setFines={setFines} trucks={trucks} drivers={drivers} perms={perms} />
          ) : (
            <ReportsView loadings={loadings} trucks={trucks} drivers={drivers} />
          )}
        </div>
      </main>
    </div>
  );
}

function CompanyLogoBadge({ logo, onChange, size = 32 }) {
  const inputRef = useRef(null);
  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const dataUrl = await fileToCompressedDataUrl(file, 400, 0.85);
      onChange(dataUrl);
    } catch {
      // falha silenciosa — usuário pode tentar novamente
    }
  };
  return (
    <div className="relative shrink-0" style={{ width: size, height: size }}>
      <div
        className="w-full h-full rounded-lg flex items-center justify-center overflow-hidden"
        style={{ background: logo ? "#FFFFFF" : "#FF8A1E" }}
      >
        {logo ? (
          <img src={logo} alt="Logo da empresa" className="w-full h-full object-contain p-0.5" />
        ) : (
          <Truck size={16} color="#151C26" strokeWidth={2.5} />
        )}
      </div>
      <button
        onClick={() => inputRef.current?.click()}
        title="Definir logo da empresa"
        className="absolute -bottom-1 -right-1 w-4 h-4 rounded-full flex items-center justify-center"
        style={{ background: "#1D2733", border: "1.5px solid #151C26" }}
      >
        <Pencil size={8} style={{ color: "#FF8A1E" }} />
      </button>
      <input ref={inputRef} type="file" accept="image/*" onChange={handleFile} className="hidden" />
    </div>
  );
}

function Sidebar({ tab, setTab, counts, role, setRole, session, onLogout, companyLogo, setCompanyLogo }) {
  const [roleOpen, setRoleOpen] = useState(false);
  return (
    <aside className="hidden md:flex w-60 shrink-0 flex-col" style={{ background: "#151C26" }}>
      <div className="px-6 py-6">
        <div className="flex items-center gap-2.5">
          <CompanyLogoBadge logo={companyLogo} onChange={setCompanyLogo} />
          <div className="min-w-0">
            <p className="font-display font-semibold text-sm leading-none truncate" style={{ color: "#FFFFFF" }}>{session?.empresa?.nome || "Frota+"}</p>
            <p className="text-[11px] mt-0.5" style={{ color: "#7C8797" }}>Gestão de transportes</p>
          </div>
        </div>
      </div>
      <nav className="flex-1 px-3 space-y-1">
        {NAV.map(({ key, label, icon: Icon }) => {
          const active = tab === key;
          const count = counts[key];
          return (
            <button
              key={key}
              onClick={() => setTab(key)}
              className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm transition-colors"
              style={{
                background: active ? "#FF8A1E" : "transparent",
                color: active ? "#151C26" : "#C4CAD3",
                fontWeight: active ? 600 : 500,
              }}
            >
              <Icon size={16} strokeWidth={2.2} />
              <span className="flex-1 text-left">{label}</span>
              {typeof count === "number" && (
                <span
                  className="text-[11px] px-1.5 py-0.5 rounded-md"
                  style={{ background: active ? "rgba(21,28,38,0.15)" : "#232C39", color: active ? "#151C26" : "#7C8797" }}
                >
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      <div className="px-3 pb-3 relative">
        <p className="px-3 text-[10px] font-medium uppercase tracking-wide mb-1.5" style={{ color: "#5B6472" }}>Simular perfil (teste)</p>
        <button
          onClick={() => setRoleOpen((v) => !v)}
          className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg"
          style={{ background: "#1D2733" }}
        >
          <ShieldCheck size={15} style={{ color: "#FF8A1E" }} />
          <span className="flex-1 text-left text-xs font-medium" style={{ color: "#FFFFFF" }}>{ROLES[role].label}</span>
          <ChevronDown size={14} style={{ color: "#7C8797" }} />
        </button>
        {roleOpen && (
          <div className="absolute bottom-full left-3 right-3 mb-1.5 rounded-lg overflow-hidden" style={{ background: "#1D2733", border: "1px solid #2A3644" }}>
            {Object.entries(ROLES).map(([key, r]) => (
              <button
                key={key}
                onClick={() => { setRole(key); setRoleOpen(false); }}
                className="w-full text-left px-3 py-2.5 text-xs"
                style={{ color: role === key ? "#FF8A1E" : "#C4CAD3", fontWeight: role === key ? 600 : 500 }}
              >
                {r.label}
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="px-3 pb-3 pt-3 flex items-center gap-2.5" style={{ borderTop: "1px solid #232C39" }}>
        <Avatar src={null} fallback={(session?.usuario?.nome || "U").slice(0, 1).toUpperCase()} size={32} />
        <div className="min-w-0 flex-1">
          <p className="text-xs font-medium truncate" style={{ color: "#FFFFFF" }}>{session?.usuario?.nome || "Usuário"}</p>
          <p className="text-[11px] truncate" style={{ color: "#7C8797" }}>{session?.usuario?.email || ""}</p>
        </div>
        <button onClick={onLogout} className="p-1.5 rounded-lg shrink-0" style={{ background: "#1D2733" }} title="Sair">
          <LogOut size={13} style={{ color: "#C4CAD3" }} />
        </button>
      </div>
      <div className="px-6 py-3 text-[11px]" style={{ color: "#5B6472", borderTop: "1px solid #232C39" }}>
        Protótipo · dados salvos neste navegador
      </div>
    </aside>
  );
}

function MobileNav({ tab, setTab, counts, role, session, onLogout, onClose, companyLogo, setCompanyLogo }) {
  return (
    <div className="fixed inset-0 z-50 md:hidden flex">
      <div className="absolute inset-0" style={{ background: "rgba(16,21,28,0.55)" }} onClick={onClose} />
      <aside className="relative w-72 max-w-[80vw] h-full flex flex-col ml-0" style={{ background: "#151C26" }}>
        <div className="px-5 py-5 flex items-center justify-between" style={{ borderBottom: "1px solid #232C39" }}>
          <div className="flex items-center gap-2.5 min-w-0">
            <CompanyLogoBadge logo={companyLogo} onChange={setCompanyLogo} />
            <p className="font-display font-semibold text-sm truncate" style={{ color: "#FFFFFF" }}>{session?.empresa?.nome || "Frota+"}</p>
          </div>
          <button onClick={onClose} className="p-1.5 rounded-lg shrink-0" style={{ background: "#1D2733" }}>
            <X size={15} style={{ color: "#C4CAD3" }} />
          </button>
        </div>
        <nav className="flex-1 px-3 py-3 space-y-1 overflow-y-auto">
          {NAV.map(({ key, label, icon: Icon }) => {
            const active = tab === key;
            const count = counts[key];
            return (
              <button
                key={key}
                onClick={() => setTab(key)}
                className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm"
                style={{
                  background: active ? "#FF8A1E" : "transparent",
                  color: active ? "#151C26" : "#C4CAD3",
                  fontWeight: active ? 600 : 500,
                }}
              >
                <Icon size={16} strokeWidth={2.2} />
                <span className="flex-1 text-left">{label}</span>
                {typeof count === "number" && (
                  <span
                    className="text-[11px] px-1.5 py-0.5 rounded-md"
                    style={{ background: active ? "rgba(21,28,38,0.15)" : "#232C39", color: active ? "#151C26" : "#7C8797" }}
                  >
                    {count}
                  </span>
                )}
              </button>
            );
          })}
        </nav>
        <div className="px-3 pb-3 pt-3 flex items-center gap-2.5" style={{ borderTop: "1px solid #232C39" }}>
          <Avatar src={null} fallback={(session?.usuario?.nome || "U").slice(0, 1).toUpperCase()} size={32} />
          <div className="min-w-0 flex-1">
            <p className="text-xs font-medium truncate" style={{ color: "#FFFFFF" }}>{session?.usuario?.nome || "Usuário"}</p>
            <p className="text-[11px] truncate" style={{ color: "#7C8797" }}>{ROLES[role].label}</p>
          </div>
          <button onClick={onLogout} className="p-1.5 rounded-lg shrink-0" style={{ background: "#1D2733" }} title="Sair">
            <LogOut size={13} style={{ color: "#C4CAD3" }} />
          </button>
        </div>
      </aside>
    </div>
  );
}

function TopBar({ tab, saveError, role, setTab, onOpenMenu, onRetrySave }) {
  const titleMap = {
    dashboard: "Painel geral",
    trucks: "Caminhões",
    drivers: "Motoristas",
    helpers: "Ajudantes",
    loadings: "Controle de carregamento",
    access: "Controle de acesso por QR Code",
    maintenance: "Manutenção da frota",
    fuel: "Abastecimento",
    fines: "Multas e infrações",
    reports: "Relatórios de carregamento",
  };
  const subMap = {
    dashboard: "Visão consolidada da frota, equipe e movimentações",
    trucks: "Cadastro de veículos, documentos e fotos",
    drivers: "Cadastro de motoristas, CNH e documentos",
    helpers: "Cadastro de ajudantes e documentos",
    loadings: "Registro de entrada e saída de carregamento",
    access: "Leia o QR do crachá para registrar entrada ou saída rapidamente",
    maintenance: "Revisões, trocas de peças e alertas de manutenção",
    fuel: "Litros, custo e consumo por abastecimento",
    fines: "Controle de multas, pontuação e vencimento",
    reports: "Movimentações por período, caminhão e motorista",
  };
  return (
    <header className="sticky top-0 z-10 px-6 sm:px-8 py-5" style={{ background: "#F5F6F8", borderBottom: "1px solid #E7E9EC" }}>
      <div className="max-w-6xl mx-auto flex items-center justify-between gap-3">
        <div className="flex items-center gap-3 min-w-0">
          <button
            onClick={onOpenMenu}
            className="md:hidden shrink-0 w-9 h-9 rounded-lg flex items-center justify-center"
            style={{ background: "#EEF0F2" }}
            title="Abrir menu"
          >
            <Menu size={16} style={{ color: "#10151C" }} />
          </button>
          {tab !== "dashboard" && (
            <button
              onClick={() => setTab("dashboard")}
              className="shrink-0 w-9 h-9 rounded-lg flex items-center justify-center"
              style={{ background: "#EEF0F2" }}
              title="Voltar ao painel"
            >
              <ArrowLeft size={16} style={{ color: "#10151C" }} />
            </button>
          )}
          <div className="min-w-0">
            <h1 className="font-display text-xl font-semibold truncate" style={{ color: "#10151C" }}>{titleMap[tab]}</h1>
            <p className="text-sm mt-0.5 truncate" style={{ color: "#6B7480" }}>{subMap[tab]}</p>
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {role !== "admin" && (
            <span className="hidden sm:flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg" style={{ background: "#EEF0F2", color: "#6B7480" }}>
              <Lock size={12} /> {ROLES[role].label}
            </span>
          )}
          {saveError && (
            <button
              onClick={onRetrySave}
              className="flex items-center gap-1.5 text-xs px-3 py-1.5 rounded-lg"
              style={{ background: "#FDEBE8", color: "#B3372A" }}
            >
              <AlertTriangle size={12} /> Não salvou — tocar para tentar de novo
            </button>
          )}
        </div>
      </div>
    </header>
  );
}

/* ---------------------------------------------------------
   Dashboard
--------------------------------------------------------- */

function Dashboard({ trucks, drivers, helpers, loadings, maintenances = [], fuels = [], fines = [], setTab }) {
  const alerts = useMemo(() => {
    const items = [];
    trucks.forEach((t) => (t.documentos || []).forEach((d) => {
      const s = docStatus(d.validade);
      if (s.tone === "danger" || s.tone === "warn") items.push({ tipo: "Caminhão", nome: t.placa, doc: d.tipo, ...s });
    }));
    drivers.forEach((d) => {
      const s = docStatus(d.cnhValidade);
      if (s.tone === "danger" || s.tone === "warn") items.push({ tipo: "Motorista", nome: d.nome, doc: "CNH", ...s });
    });
    maintenances.forEach((m) => {
      const s = maintenanceAlertStatus(m);
      if (s) items.push({ tipo: "Manutenção", nome: m.placa, doc: m.descricao || "Próxima revisão", ...s });
    });
    return items.sort((a, b) => (a.tone === "danger" ? -1 : 1));
  }, [trucks, drivers, maintenances]);

  const recentLoadings = [...loadings].sort((a, b) => (b.data + b.hora).localeCompare(a.data + a.hora)).slice(0, 6);

  // Custos por caminhão no mês atual (manutenção + combustível)
  const custosPorCaminhao = useMemo(() => {
    const mesAtual = todayISO().slice(0, 7);
    const map = {};
    trucks.forEach((t) => { map[t.placa] = { placa: t.placa, manutencao: 0, combustivel: 0 }; });
    maintenances.filter((m) => (m.data || "").slice(0, 7) === mesAtual).forEach((m) => {
      if (!map[m.placa]) map[m.placa] = { placa: m.placa, manutencao: 0, combustivel: 0 };
      map[m.placa].manutencao += Number(m.custo) || 0;
    });
    fuels.filter((f) => (f.data || "").slice(0, 7) === mesAtual).forEach((f) => {
      if (!map[f.placa]) map[f.placa] = { placa: f.placa, manutencao: 0, combustivel: 0 };
      map[f.placa].combustivel += Number(f.valorTotal) || 0;
    });
    return Object.values(map)
      .map((c) => ({ ...c, total: c.manutencao + c.combustivel }))
      .filter((c) => c.total > 0)
      .sort((a, b) => b.total - a.total)
      .slice(0, 8);
  }, [trucks, maintenances, fuels]);

  // Produtividade por motorista (movimentações registradas)
  const produtividadeMotoristas = useMemo(() => {
    const map = {};
    loadings.forEach((l) => { map[l.motorista] = (map[l.motorista] || 0) + 1; });
    return Object.entries(map)
      .map(([motorista, total]) => ({ motorista, total }))
      .sort((a, b) => b.total - a.total)
      .slice(0, 8);
  }, [loadings]);

  const gastoMesTotal = useMemo(() => {
    const mesAtual = todayISO().slice(0, 7);
    const manut = maintenances.filter((m) => (m.data || "").slice(0, 7) === mesAtual).reduce((s, m) => s + (Number(m.custo) || 0), 0);
    const comb = fuels.filter((f) => (f.data || "").slice(0, 7) === mesAtual).reduce((s, f) => s + (Number(f.valorTotal) || 0), 0);
    const multas = fines.filter((f) => (f.data || "").slice(0, 7) === mesAtual).reduce((s, f) => s + (Number(f.valor) || 0), 0);
    return { manut, comb, multas, total: manut + comb + multas };
  }, [maintenances, fuels, fines]);

  const stats = [
    { label: "Caminhões", value: trucks.length, icon: Truck, tab: "trucks", bg: "#FFF1E0", fg: "#C9600A" },
    { label: "Motoristas", value: drivers.length, icon: Users, tab: "drivers", bg: "#E8F1FC", fg: "#2A66C9" },
    { label: "Ajudantes", value: helpers.length, icon: HardHat, tab: "helpers", bg: "#E7F7EF", fg: "#1F7A54" },
    { label: "Movimentações", value: loadings.length, icon: ClipboardList, tab: "loadings", bg: "#F1E9FB", fg: "#7C3AED" },
  ];

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {stats.map(({ label, value, icon: Icon, tab: t, bg, fg }) => (
          <button
            key={label}
            onClick={() => setTab(t)}
            className="text-left rounded-2xl p-4 transition-transform hover:-translate-y-0.5"
            style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}
          >
            <div className="w-8 h-8 rounded-lg flex items-center justify-center mb-3" style={{ background: bg }}>
              <Icon size={15} style={{ color: fg }} />
            </div>
            <p className="font-display text-2xl font-semibold" style={{ color: "#10151C" }}>{value}</p>
            <p className="text-xs mt-0.5" style={{ color: "#6B7480" }}>{label}</p>
          </button>
        ))}
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div className="rounded-2xl p-4" style={{ background: "#151C26" }}>
          <div className="flex items-center gap-1.5 mb-1"><DollarSign size={13} style={{ color: "#FF8A1E" }} /><p className="text-[11px]" style={{ color: "#9AA1AA" }}>Gasto total no mês</p></div>
          <p className="font-display text-lg font-semibold" style={{ color: "#FFFFFF" }}>R$ {gastoMesTotal.total.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
        </div>
        <button onClick={() => setTab("maintenance")} className="text-left rounded-2xl p-4" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-1.5 mb-1"><Wrench size={13} style={{ color: "#C9600A" }} /><p className="text-[11px]" style={{ color: "#9AA1AA" }}>Manutenção (mês)</p></div>
          <p className="font-display text-lg font-semibold" style={{ color: "#10151C" }}>R$ {gastoMesTotal.manut.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
        </button>
        <button onClick={() => setTab("fuel")} className="text-left rounded-2xl p-4" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-1.5 mb-1"><Fuel size={13} style={{ color: "#2A66C9" }} /><p className="text-[11px]" style={{ color: "#9AA1AA" }}>Combustível (mês)</p></div>
          <p className="font-display text-lg font-semibold" style={{ color: "#10151C" }}>R$ {gastoMesTotal.comb.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
        </button>
        <button onClick={() => setTab("fines")} className="text-left rounded-2xl p-4" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-1.5 mb-1"><AlertOctagon size={13} style={{ color: "#B3372A" }} /><p className="text-[11px]" style={{ color: "#9AA1AA" }}>Multas (mês)</p></div>
          <p className="font-display text-lg font-semibold" style={{ color: "#10151C" }}>R$ {gastoMesTotal.multas.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
        </button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-4">
        <div className="lg:col-span-3 rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Últimas movimentações</h2>
            <button onClick={() => setTab("loadings")} className="text-xs font-medium flex items-center gap-1" style={{ color: "#C9600A" }}>
              Ver tudo <ChevronRight size={13} />
            </button>
          </div>
          {recentLoadings.length === 0 ? (
            <EmptyState icon={ClipboardList} title="Nenhuma movimentação ainda" subtitle="Registros de entrada e saída aparecerão aqui" />
          ) : (
            <div className="space-y-2">
              {recentLoadings.map((l) => (
                <div key={l.id} className="flex items-center gap-3 py-2" style={{ borderBottom: "1px solid #F0F1F3" }}>
                  <span
                    className="w-7 h-7 rounded-lg flex items-center justify-center shrink-0"
                    style={{ background: l.tipo === "Entrada" ? "#E7F7EF" : "#FFF1E0", color: l.tipo === "Entrada" ? "#1F9D6B" : "#C9600A" }}
                  >
                    <ArrowRightLeft size={13} />
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-medium truncate" style={{ color: "#10151C" }}>{l.placa} · {l.motorista}</p>
                    <p className="text-xs" style={{ color: "#9AA1AA" }}>{l.data} às {l.hora}</p>
                  </div>
                  <StatusPill tone={l.tipo === "Entrada" ? "ok" : "warn"}>{l.tipo}</StatusPill>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="lg:col-span-2 rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-2 mb-4">
            <AlertTriangle size={15} style={{ color: "#E0503C" }} />
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Documentos e manutenções a acompanhar</h2>
          </div>
          {alerts.length === 0 ? (
            <div className="flex items-center gap-2 py-6 justify-center">
              <CheckCircle2 size={16} style={{ color: "#1F9D6B" }} />
              <p className="text-xs" style={{ color: "#6B7480" }}>Tudo em dia por aqui</p>
            </div>
          ) : (
            <div className="space-y-2.5">
              {alerts.slice(0, 6).map((a, i) => (
                <div key={i} className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <p className="text-sm truncate" style={{ color: "#10151C" }}>{a.doc} · {a.nome}</p>
                    <p className="text-xs" style={{ color: "#9AA1AA" }}>{a.tipo}</p>
                  </div>
                  <StatusPill tone={a.tone}>{a.label}</StatusPill>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div className="rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-2 mb-4">
            <DollarSign size={15} style={{ color: "#C9600A" }} />
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Custo por caminhão no mês</h2>
          </div>
          {custosPorCaminhao.length === 0 ? (
            <div className="py-10 text-center text-xs" style={{ color: "#9AA1AA" }}>Sem lançamentos de manutenção ou combustível neste mês</div>
          ) : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={custosPorCaminhao}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F0F1F3" vertical={false} />
                <XAxis dataKey="placa" tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={false} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={false} tickLine={false} width={40} />
                <Tooltip formatter={(v) => `R$ ${Number(v).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #E7E9EC" }} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="manutencao" name="Manutenção" stackId="a" fill="#C9600A" radius={[0, 0, 0, 0]} />
                <Bar dataKey="combustivel" name="Combustível" stackId="a" fill="#2A66C9" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div className="rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-2 mb-4">
            <TrendingUp size={15} style={{ color: "#1F7A54" }} />
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Produtividade por motorista</h2>
          </div>
          {produtividadeMotoristas.length === 0 ? (
            <div className="py-10 text-center text-xs" style={{ color: "#9AA1AA" }}>Sem movimentações registradas ainda</div>
          ) : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={produtividadeMotoristas} layout="vertical" margin={{ left: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#F0F1F3" horizontal={false} />
                <XAxis type="number" tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={false} tickLine={false} allowDecimals={false} />
                <YAxis type="category" dataKey="motorista" tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={false} tickLine={false} width={90} />
                <Tooltip formatter={(v) => `${v} movimentações`} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #E7E9EC" }} />
                <Bar dataKey="total" name="Movimentações" fill="#1F7A54" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Shared list scaffold
--------------------------------------------------------- */

function ListToolbar({ query, setQuery, placeholder, onAdd, addLabel, canCreate = true, onImport }) {
  return (
    <div className="flex items-center gap-3 mb-4">
      <div className="relative flex-1 max-w-xs">
        <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#9AA1AA" }} />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder={placeholder}
          className="w-full rounded-lg pl-8 pr-3 py-2 text-sm outline-none"
          style={inputStyle}
        />
      </div>
      {canCreate && onImport && (
        <button
          onClick={onImport}
          className="ml-auto flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg shrink-0"
          style={{ background: "#E8F1FC", color: "#2A66C9" }}
        >
          <Upload size={15} /> Importar planilha
        </button>
      )}
      {canCreate && (
        <button
          onClick={onAdd}
          className={`flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg shrink-0 ${onImport ? "" : "ml-auto"}`}
          style={{ background: "#FF8A1E", color: "#151C26" }}
        >
          <Plus size={15} /> {addLabel}
        </button>
      )}
    </div>
  );
}

function RowActions({ onEdit, onDelete, canEdit = true, canDelete = true }) {
  if (!canEdit && !canDelete) return null;
  return (
    <div className="flex items-center gap-1">
      {canEdit && (
        <button onClick={onEdit} className="p-1.5 rounded-lg" style={{ background: "#E8F1FC" }} title="Editar">
          <Pencil size={13} style={{ color: "#2A66C9" }} />
        </button>
      )}
      {canDelete && (
        <button onClick={onDelete} className="p-1.5 rounded-lg" style={{ background: "#FDEBE8" }} title="Excluir">
          <Trash2 size={13} style={{ color: "#B3372A" }} />
        </button>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   Importação de planilha (CSV / Excel)
--------------------------------------------------------- */

function parseSpreadsheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function guessColumn(header, key, label) {
  const h = header.map((x) => String(x).toLowerCase().trim());
  const target = [key.toLowerCase(), label.toLowerCase()];
  let idx = h.findIndex((c) => target.includes(c));
  if (idx === -1) idx = h.findIndex((c) => target.some((t) => c.includes(t) || t.includes(c)));
  return idx === -1 ? "" : String(idx);
}

function ImportModal({ title, fields, onClose, onImport }) {
  const [step, setStep] = useState("upload"); // upload | mapping | done
  const [fileName, setFileName] = useState("");
  const [rows, setRows] = useState([]); // raw arrays, rows[0] = header
  const [mapping, setMapping] = useState({}); // fieldKey -> column index (string)
  const [error, setError] = useState("");
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    setError("");
    try {
      const parsed = await parseSpreadsheet(file);
      if (!parsed.length || parsed.length < 2) {
        setError("A planilha precisa ter uma linha de cabeçalho e ao menos uma linha de dados.");
        return;
      }
      setFileName(file.name);
      setRows(parsed);
      const header = parsed[0];
      const guessed = {};
      fields.forEach((f) => { guessed[f.key] = guessColumn(header, f.key, f.label); });
      setMapping(guessed);
      setStep("mapping");
    } catch {
      setError("Não consegui ler esse arquivo. Tente um .xlsx ou .csv.");
    }
  };

  const header = rows[0] || [];
  const dataRows = rows.slice(1);

  const mappedPreview = useMemo(() => {
    return dataRows
      .map((r) => {
        const obj = {};
        fields.forEach((f) => {
          const colIdx = mapping[f.key];
          obj[f.key] = colIdx !== "" && colIdx !== undefined ? String(r[Number(colIdx)] ?? "").trim() : "";
        });
        return obj;
      })
      .filter((obj) => Object.values(obj).some((v) => v !== ""));
  }, [dataRows, mapping, fields]);

  const requiredOk = fields.filter((f) => f.required).every((f) => mapping[f.key] !== "" && mapping[f.key] !== undefined);

  const confirmImport = () => {
    onImport(mappedPreview);
  };

  return (
    <Modal title={title} onClose={onClose} wide>
      {step === "upload" && (
        <div>
          <div
            onClick={() => inputRef.current?.click()}
            className="rounded-xl flex flex-col items-center justify-center gap-2 py-10 cursor-pointer"
            style={{ background: "#F5F6F8", border: "1px dashed #C7CBD1" }}
          >
            <FileSpreadsheet size={22} style={{ color: "#9AA1AA" }} />
            <p className="text-sm font-medium" style={{ color: "#10151C" }}>Clique para escolher a planilha</p>
            <p className="text-xs" style={{ color: "#9AA1AA" }}>.xlsx, .xls ou .csv — com uma linha de cabeçalho</p>
          </div>
          <input
            ref={inputRef}
            type="file"
            accept=".csv,.xlsx,.xls"
            className="hidden"
            onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); e.target.value = ""; }}
          />
          {error && (
            <p className="flex items-center gap-1.5 text-xs mt-3" style={{ color: "#B3372A" }}>
              <CircleAlert size={13} /> {error}
            </p>
          )}
        </div>
      )}

      {step === "mapping" && (
        <div>
          <p className="text-xs mb-4" style={{ color: "#6B7480" }}>
            <FileSpreadsheet size={12} className="inline mr-1" style={{ color: "#9AA1AA" }} />
            {fileName} · {dataRows.length} linha{dataRows.length === 1 ? "" : "s"} encontrada{dataRows.length === 1 ? "" : "s"}. Confirme de qual coluna vem cada informação:
          </p>
          <div className="space-y-2.5 max-h-72 overflow-y-auto pr-1">
            {fields.map((f) => (
              <div key={f.key} className="grid grid-cols-[1fr_1.4fr] gap-3 items-center">
                <span className="text-sm" style={{ color: "#10151C" }}>
                  {f.label}{f.required && <span style={{ color: "#E0503C" }}> *</span>}
                </span>
                <SelectInput value={mapping[f.key] ?? ""} onChange={(e) => setMapping({ ...mapping, [f.key]: e.target.value })}>
                  <option value="">Não importar</option>
                  {header.map((h, i) => (
                    <option key={i} value={i}>{String(h) || `Coluna ${i + 1}`}</option>
                  ))}
                </SelectInput>
              </div>
            ))}
          </div>
          <div className="rounded-lg mt-4 p-3" style={{ background: "#F5F6F8" }}>
            <p className="text-xs font-medium mb-1" style={{ color: "#6B7480" }}>{mappedPreview.length} registro{mappedPreview.length === 1 ? "" : "s"} pronto{mappedPreview.length === 1 ? "" : "s"} para importar</p>
            {mappedPreview[0] && (
              <p className="text-xs truncate" style={{ color: "#9AA1AA" }}>
                Ex.: {fields.map((f) => mappedPreview[0][f.key]).filter(Boolean).slice(0, 3).join(" · ") || "—"}
              </p>
            )}
          </div>
          <div className="flex justify-between items-center pt-4">
            <button onClick={() => setStep("upload")} className="flex items-center gap-1 text-sm font-medium px-3 py-2 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>
              <ArrowLeft size={14} /> Trocar arquivo
            </button>
            <div className="flex gap-2">
              <button onClick={onClose} className="text-sm font-medium px-4 py-2 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>Cancelar</button>
              <button
                onClick={confirmImport}
                disabled={!requiredOk || mappedPreview.length === 0}
                className="flex items-center gap-1.5 text-sm font-medium px-4 py-2 rounded-lg"
                style={{ background: !requiredOk || mappedPreview.length === 0 ? "#F0D4B0" : "#FF8A1E", color: "#151C26" }}
              >
                Importar {mappedPreview.length || ""} <ArrowRight size={14} />
              </button>
            </div>
          </div>
        </div>
      )}
    </Modal>
  );
}

const emptyTruck = { id: null, placa: "", modelo: "", ano: "", capacidade: "", fotos: [], documentos: [{ tipo: "CRLV", numero: "", validade: "" }, { tipo: "Seguro", numero: "", validade: "" }] };

function TrucksView({ trucks, setTrucks, perms = ROLES.admin }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null);
  const [importing, setImporting] = useState(false);

  const filtered = trucks.filter((t) => (t.placa + t.modelo).toLowerCase().includes(query.toLowerCase()));

  const save = (truck) => {
    setTrucks((prev) => {
      const exists = prev.some((t) => t.id === truck.id);
      return exists ? prev.map((t) => (t.id === truck.id ? truck : t)) : [...prev, truck];
    });
    setEditing(null);
  };
  const remove = (id) => setTrucks((prev) => prev.filter((t) => t.id !== id));

  const truckImportFields = [
    { key: "placa", label: "Placa", required: true },
    { key: "modelo", label: "Modelo" },
    { key: "ano", label: "Ano" },
    { key: "capacidade", label: "Capacidade" },
  ];
  const handleImport = (rows) => {
    const novos = rows.filter((r) => r.placa).map((r) => ({
      ...emptyTruck,
      ...r,
      id: uid(),
      placa: r.placa.toUpperCase(),
      fotos: [],
      documentos: emptyTruck.documentos.map((d) => ({ ...d })),
    }));
    setTrucks((prev) => [...prev, ...novos]);
    setImporting(false);
  };

  return (
    <div>
      <ListToolbar
        query={query}
        setQuery={setQuery}
        placeholder="Buscar por placa ou modelo"
        onAdd={() => setEditing({ ...emptyTruck, id: uid() })}
        addLabel="Novo caminhão"
        canCreate={perms.canCreate}
        onImport={() => setImporting(true)}
      />
      {filtered.length === 0 ? (
        <EmptyState icon={Truck} title="Nenhum caminhão cadastrado" subtitle="Cadastre a frota para começar o controle" />
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
          {filtered.map((t) => {
            const worst = (t.documentos || []).reduce((acc, d) => {
              const s = docStatus(d.validade).tone;
              if (s === "danger") return "danger";
              if (s === "warn" && acc !== "danger") return "warn";
              return acc;
            }, "ok");
            return (
              <div key={t.id} className="rounded-2xl p-4" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
                <div className="flex items-start gap-3">
                  <Avatar src={t.fotos?.[0]?.src} fallback={<Truck size={16} />} />
                  <div className="min-w-0 flex-1">
                    <p className="font-mono text-sm font-bold tracking-wide" style={{ color: "#10151C" }}>{t.placa || "SEM PLACA"}</p>
                    <p className="text-xs truncate" style={{ color: "#6B7480" }}>{t.modelo}{t.ano ? ` · ${t.ano}` : ""}</p>
                  </div>
                  <RowActions onEdit={() => setEditing(t)} onDelete={() => remove(t.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
                </div>
                {t.fotos && t.fotos.length > 0 && (
                  <div className="flex items-center gap-1.5 mt-3">
                    {t.fotos.slice(0, 4).map((p, i) => (
                      <div key={p.id} className="rounded-md overflow-hidden shrink-0" style={{ width: 34, height: 34, border: "1px solid #E7E9EC" }}>
                        <img src={p.src} alt="" className="w-full h-full object-cover" />
                      </div>
                    ))}
                    {t.fotos.length > 4 && (
                      <span className="text-[11px] font-medium px-1.5" style={{ color: "#9AA1AA" }}>+{t.fotos.length - 4}</span>
                    )}
                  </div>
                )}
                <div className="flex items-center justify-between mt-3 pt-3" style={{ borderTop: "1px solid #F0F1F3" }}>
                  <span className="text-xs" style={{ color: "#9AA1AA" }}>{t.capacidade ? `${t.capacidade} de capacidade` : "—"}</span>
                  <StatusPill tone={worst}>{worst === "ok" ? "Docs em dia" : worst === "warn" ? "Vencendo" : "Vencido"}</StatusPill>
                </div>
              </div>
            );
          })}
        </div>
      )}
      {editing && <TruckForm truck={editing} onClose={() => setEditing(null)} onSave={save} />}
      {importing && (
        <ImportModal
          title="Importar caminhões de planilha"
          fields={truckImportFields}
          onClose={() => setImporting(false)}
          onImport={handleImport}
        />
      )}
    </div>
  );
}

function TruckForm({ truck, onClose, onSave }) {
  const [form, setForm] = useState(truck);
  const [ocrState, setOcrState] = useState({});
  const setDoc = (i, patch) => setForm((f) => ({ ...f, documentos: f.documentos.map((d, idx) => (idx === i ? { ...d, ...patch } : d)) }));
  const addDoc = () => setForm((f) => ({ ...f, documentos: [...f.documentos, { tipo: "", numero: "", validade: "" }] }));
  const removeDoc = (i) => setForm((f) => ({ ...f, documentos: f.documentos.filter((_, idx) => idx !== i) }));

  const handleOcrFile = async (i, file) => {
    setOcrState((s) => ({ ...s, [i]: { loading: true } }));
    try {
      const dataUrl = await fileToCompressedDataUrl(file);
      const tipoLower = (form.documentos[i].tipo || "").toLowerCase();
      const kind = tipoLower.includes("crlv") ? "crlv" : "generico";
      const result = await extractDocumentData(dataUrl, kind);
      setDoc(i, {
        numero: result.numero_documento || result.numero || form.documentos[i].numero,
        validade: result.validade || form.documentos[i].validade,
      });
      if (kind === "crlv") {
        setForm((f) => ({
          ...f,
          placa: f.placa || (result.placa ? String(result.placa).toUpperCase() : f.placa),
          modelo: f.modelo || result.modelo || f.modelo,
          ano: f.ano || result.ano || f.ano,
        }));
      }
      setOcrState((s) => ({ ...s, [i]: { loading: false, success: true } }));
    } catch (err) {
      console.error("OCR documento falhou:", err);
      setOcrState((s) => ({ ...s, [i]: { loading: false, error: err?.message || "Erro desconhecido" } }));
    }
  };

  return (
    <Modal title={truck.placa ? "Editar caminhão" : "Novo caminhão"} onClose={onClose} wide>
      <div className="space-y-4">
        <MultiPhotoPicker
          label="Fotos do veículo (todos os ângulos)"
          values={form.fotos || []}
          onChange={(fotos) => setForm({ ...form, fotos })}
          angleHints={["Frente", "Traseira", "Lateral esquerda", "Lateral direita", "Interior/cabine", "Placa"]}
        />
        <div className="grid grid-cols-2 gap-3">
          <Field label="Placa"><TextInput value={form.placa} onChange={(e) => setForm({ ...form, placa: e.target.value.toUpperCase() })} placeholder="ABC1D23" /></Field>
          <Field label="Ano"><TextInput value={form.ano} onChange={(e) => setForm({ ...form, ano: e.target.value })} placeholder="2022" /></Field>
          <Field label="Modelo"><TextInput value={form.modelo} onChange={(e) => setForm({ ...form, modelo: e.target.value })} placeholder="Volvo FH 540" /></Field>
          <Field label="Capacidade"><TextInput value={form.capacidade} onChange={(e) => setForm({ ...form, capacidade: e.target.value })} placeholder="24 ton" /></Field>
        </div>
        <div>
          <div className="flex items-center justify-between mb-2">
            <p className="text-xs font-medium" style={{ color: "#6B7480" }}>Documentos</p>
            <button onClick={addDoc} className="text-xs font-medium" style={{ color: "#C9600A" }}>+ adicionar</button>
          </div>
          <div className="flex items-center gap-1.5 mb-2">
            <Sparkles size={12} style={{ color: "#7C3AED" }} />
            <p className="text-[11px]" style={{ color: "#9AA1AA" }}>Toque no ícone roxo para fotografar o documento e preencher número/validade automaticamente</p>
          </div>
          <div className="space-y-2">
            {form.documentos.map((d, i) => (
              <div key={i}>
                <div className="grid grid-cols-[1.2fr_1fr_1fr_auto_auto] gap-2 items-center">
                  <TextInput value={d.tipo} onChange={(e) => setDoc(i, { tipo: e.target.value })} placeholder="Tipo (ex: CRLV)" />
                  <TextInput value={d.numero} onChange={(e) => setDoc(i, { numero: e.target.value })} placeholder="Número" />
                  <TextInput type="date" value={d.validade} onChange={(e) => setDoc(i, { validade: e.target.value })} />
                  <OcrButton onFile={(file) => handleOcrFile(i, file)} status={ocrState[i]} title="Fotografar e ler com IA" />
                  <button onClick={() => removeDoc(i)} className="p-2 rounded-lg" style={{ background: "#FDEBE8" }}><Trash2 size={13} style={{ color: "#B3372A" }} /></button>
                </div>
                {ocrState[i]?.error && (
                  <p className="flex items-center gap-1 text-[11px] mt-1" style={{ color: "#B3372A" }}><CircleAlert size={11} /> {ocrState[i].error}</p>
                )}
                {ocrState[i]?.success && (
                  <p className="flex items-center gap-1 text-[11px] mt-1" style={{ color: "#1F7A54" }}><CheckCircle2 size={11} /> Dados preenchidos — confira antes de salvar.</p>
                )}
              </div>
            ))}
          </div>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.placa} />
      </div>
    </Modal>
  );
}

function FormFooter({ onClose, onSave, disabled }) {
  return (
    <div className="flex justify-end gap-2 pt-2">
      <button onClick={onClose} className="text-sm font-medium px-4 py-2 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>Cancelar</button>
      <button onClick={onSave} disabled={disabled} className="text-sm font-medium px-4 py-2 rounded-lg" style={{ background: disabled ? "#F0D4B0" : "#FF8A1E", color: "#151C26" }}>Salvar</button>
    </div>
  );
}

/* ---------------------------------------------------------
   Drivers
--------------------------------------------------------- */

const emptyDriver = { id: null, nome: "", cpf: "", telefone: "", cnhNumero: "", cnhCategoria: "", cnhValidade: "", foto: null, documentos: [] };

/* ---------------------------------------------------------
   Crachás (impressão)
--------------------------------------------------------- */

function qrToken(kind, id) {
  return `FROTAPLUS|${kind}|${id}`;
}
function qrImageUrl(data, size = 120) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&margin=0&data=${encodeURIComponent(data)}`;
}

function Badge({ person, tipo, empresaNome, companyLogo }) {
  const isDriver = tipo === "motorista";
  const token = qrToken(isDriver ? "MOTORISTA" : "AJUDANTE", person.id);
  return (
    <div
      className="badge-print-area rounded-2xl overflow-hidden flex flex-col shrink-0"
      style={{ width: "220px", height: "360px", background: "#FFFFFF", border: "1px solid #E7E9EC", boxShadow: "0 1px 4px rgba(16,21,28,0.1)" }}
    >
      <div className="flex flex-col items-center pt-3 pb-3" style={{ background: "#151C26" }}>
        {companyLogo ? (
          <div className="w-9 h-9 rounded-lg overflow-hidden flex items-center justify-center mb-1.5" style={{ background: "#FFFFFF" }}>
            <img src={companyLogo} alt="" className="w-full h-full object-contain p-0.5" />
          </div>
        ) : (
          <div className="w-2.5 h-2.5 rounded-full mb-2" style={{ background: "#0B0F14", border: "2px solid #3A4553" }} />
        )}
        <p className="font-display text-[11px] font-semibold tracking-wide text-center px-2 truncate max-w-full" style={{ color: "#FFFFFF" }}>
          {(empresaNome || "FROTA+").toUpperCase()}
        </p>
      </div>
      <div className="flex-1 flex flex-col items-center px-3 pt-4">
        <div className="rounded-full overflow-hidden mb-2.5 flex items-center justify-center" style={{ width: 76, height: 76, background: "#EEF0F2", border: "3px solid #FF8A1E" }}>
          {person.foto ? (
            <img src={person.foto} alt="" className="w-full h-full object-cover" />
          ) : (
            <span className="text-xl font-semibold" style={{ color: "#9AA1AA" }}>{(person.nome || "?").slice(0, 1).toUpperCase()}</span>
          )}
        </div>
        <p className="font-display text-sm font-semibold text-center leading-snug px-1" style={{ color: "#10151C" }}>{person.nome || "—"}</p>
        <span
          className="mt-2 text-[10px] font-semibold px-2.5 py-1 rounded-full tracking-wide"
          style={{ background: isDriver ? "#E8F1FC" : "#E7F7EF", color: isDriver ? "#2A66C9" : "#1F7A54" }}
        >
          {isDriver ? "MOTORISTA" : "AJUDANTE"}
        </span>
        <div className="mt-1.5 text-center">
          {isDriver && person.cnhCategoria && <p className="text-[11px]" style={{ color: "#6B7480" }}>CNH categoria {person.cnhCategoria}</p>}
          {person.cpf && <p className="text-[11px] mt-0.5" style={{ color: "#9AA1AA" }}>CPF {person.cpf}</p>}
        </div>
        <div className="mt-2 mb-1 rounded-lg overflow-hidden" style={{ width: 72, height: 72, border: "1px solid #F0F1F3" }}>
          <img
            src={qrImageUrl(token, 144)}
            alt="QR de acesso"
            className="w-full h-full object-contain"
            onError={(e) => { e.target.style.display = "none"; }}
          />
        </div>
      </div>
      <div className="text-center py-1.5 text-[9px]" style={{ color: "#B7BCC3", borderTop: "1px solid #F0F1F3" }}>
        Crachá de identificação · escaneie para acesso
      </div>
    </div>
  );
}

function BadgePrintModal({ person, tipo, empresaNome, companyLogo, onClose }) {
  return (
    <Modal title={`Crachá — ${person.nome || "sem nome"}`} onClose={onClose}>
      <div className="flex flex-col items-center gap-4">
        <Badge person={person} tipo={tipo} empresaNome={empresaNome} companyLogo={companyLogo} />
        <p className="text-xs text-center" style={{ color: "#9AA1AA" }}>
          Confira os dados antes de imprimir. Sai em tamanho de cartão — dá para plastificar ou colocar no cordão.
        </p>
        <div className="flex gap-2 w-full">
          <button onClick={onClose} className="flex-1 text-sm font-medium px-4 py-2.5 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>Fechar</button>
          <button onClick={() => window.print()} className="flex-1 flex items-center justify-center gap-1.5 text-sm font-medium px-4 py-2.5 rounded-lg" style={{ background: "#FF8A1E", color: "#151C26" }}>
            <Printer size={15} /> Imprimir
          </button>
        </div>
      </div>
    </Modal>
  );
}

function BadgeSheetModal({ people, tipo, empresaNome, companyLogo, onClose }) {
  return (
    <Modal title="Imprimir crachás" onClose={onClose} wide>
      <div className="space-y-4">
        <p className="text-xs" style={{ color: "#6B7480" }}>{people.length} crachá{people.length === 1 ? "" : "s"} nesta lista. Confira antes de imprimir.</p>
        <div className="badge-print-sheet flex flex-wrap gap-4 justify-center max-h-96 overflow-y-auto p-1">
          {people.map((p) => (
            <Badge key={p.id} person={p} tipo={tipo} empresaNome={empresaNome} companyLogo={companyLogo} />
          ))}
        </div>
        <div className="flex gap-2 w-full">
          <button onClick={onClose} className="flex-1 text-sm font-medium px-4 py-2.5 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>Fechar</button>
          <button onClick={() => window.print()} className="flex-1 flex items-center justify-center gap-1.5 text-sm font-medium px-4 py-2.5 rounded-lg" style={{ background: "#FF8A1E", color: "#151C26" }}>
            <Printer size={15} /> Imprimir todos
          </button>
        </div>
      </div>
    </Modal>
  );
}

function DriversView({ drivers, setDrivers, perms = ROLES.admin, empresaNome, companyLogo }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null);
  const [importing, setImporting] = useState(false);
  const [badgePerson, setBadgePerson] = useState(null);
  const [sheetOpen, setSheetOpen] = useState(false);
  const filtered = drivers.filter((d) => d.nome.toLowerCase().includes(query.toLowerCase()));

  const save = (driver) => {
    setDrivers((prev) => (prev.some((d) => d.id === driver.id) ? prev.map((d) => (d.id === driver.id ? driver : d)) : [...prev, driver]));
    setEditing(null);
  };
  const remove = (id) => setDrivers((prev) => prev.filter((d) => d.id !== id));

  const driverImportFields = [
    { key: "nome", label: "Nome", required: true },
    { key: "cpf", label: "CPF" },
    { key: "telefone", label: "Telefone" },
    { key: "cnhNumero", label: "Número CNH" },
    { key: "cnhCategoria", label: "Categoria CNH" },
    { key: "cnhValidade", label: "Validade CNH" },
  ];
  const handleImport = (rows) => {
    const novos = rows.filter((r) => r.nome).map((r) => ({ ...emptyDriver, ...r, id: uid(), foto: null, documentos: [] }));
    setDrivers((prev) => [...prev, ...novos]);
    setImporting(false);
  };

  return (
    <div>
      <div className="flex items-center gap-3 mb-1">
        <ListToolbar
          query={query}
          setQuery={setQuery}
          placeholder="Buscar motorista"
          onAdd={() => setEditing({ ...emptyDriver, id: uid() })}
          addLabel="Novo motorista"
          canCreate={perms.canCreate}
          onImport={() => setImporting(true)}
        />
      </div>
      {filtered.length > 0 && (
        <div className="flex justify-end mb-3 -mt-2">
          <button
            onClick={() => setSheetOpen(true)}
            className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg"
            style={{ background: "#FFF1E0", color: "#C9600A" }}
          >
            <CreditCard size={13} /> Imprimir crachás ({filtered.length})
          </button>
        </div>
      )}
      {filtered.length === 0 ? (
        <EmptyState icon={Users} title="Nenhum motorista cadastrado" subtitle="Cadastre a equipe para vincular aos carregamentos" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          {filtered.map((d, i) => {
            const s = docStatus(d.cnhValidade);
            return (
              <div key={d.id} className="flex items-center gap-3 px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
                <Avatar src={d.foto} fallback={<Users size={15} />} size={40} />
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium truncate" style={{ color: "#10151C" }}>{d.nome}</p>
                  <p className="text-xs" style={{ color: "#9AA1AA" }}>CNH {d.cnhCategoria || "—"}{d.telefone ? ` · ${d.telefone}` : ""}</p>
                </div>
                <StatusPill tone={s.tone}>CNH {s.label}</StatusPill>
                <button onClick={() => setBadgePerson(d)} className="p-1.5 rounded-lg" style={{ background: "#FFF1E0" }} title="Imprimir crachá">
                  <CreditCard size={13} style={{ color: "#C9600A" }} />
                </button>
                <RowActions onEdit={() => setEditing(d)} onDelete={() => remove(d.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
              </div>
            );
          })}
        </div>
      )}
      {editing && <DriverForm driver={editing} onClose={() => setEditing(null)} onSave={save} />}
      {importing && (
        <ImportModal
          title="Importar motoristas de planilha"
          fields={driverImportFields}
          onClose={() => setImporting(false)}
          onImport={handleImport}
        />
      )}
      {badgePerson && <BadgePrintModal person={badgePerson} tipo="motorista" empresaNome={empresaNome} companyLogo={companyLogo} onClose={() => setBadgePerson(null)} />}
      {sheetOpen && <BadgeSheetModal people={filtered} tipo="motorista" empresaNome={empresaNome} companyLogo={companyLogo} onClose={() => setSheetOpen(false)} />}
    </div>
  );
}

function DriverForm({ driver, onClose, onSave }) {
  const [form, setForm] = useState(driver);
  const [cnhStatus, setCnhStatus] = useState(null);

  const handleCnhFile = async (file) => {
    setCnhStatus({ loading: true });
    try {
      const dataUrl = await fileToCompressedDataUrl(file);
      const result = await extractDocumentData(dataUrl, "cnh");
      setForm((f) => ({
        ...f,
        nome: f.nome || result.nome || f.nome,
        cnhNumero: result.numero_registro || result.numero || f.cnhNumero,
        cnhCategoria: result.categoria ? String(result.categoria).toUpperCase() : f.cnhCategoria,
        cnhValidade: result.validade || f.cnhValidade,
      }));
      setCnhStatus({ loading: false, success: true });
    } catch (err) {
      console.error("OCR CNH falhou:", err);
      setCnhStatus({ loading: false, error: err?.message || "Erro desconhecido" });
    }
  };

  return (
    <Modal title={driver.nome ? "Editar motorista" : "Novo motorista"} onClose={onClose} wide>
      <div className="space-y-4">
        <PhotoPicker label="Foto do motorista" value={form.foto} onChange={(v) => setForm({ ...form, foto: v })} />

        <div className="rounded-xl p-3 flex items-center gap-3" style={{ background: "#F8F5FE", border: "1px solid #E9DEFB" }}>
          <OcrButton onFile={handleCnhFile} status={cnhStatus} title="Fotografar CNH e ler com IA" />
          <div className="min-w-0 flex-1">
            <p className="text-xs font-medium" style={{ color: "#10151C" }}>Ler CNH com IA</p>
            <p className="text-[11px]" style={{ color: "#9AA1AA" }}>Fotografe a CNH e os campos abaixo são preenchidos automaticamente</p>
          </div>
          {cnhStatus?.error && <span className="text-[11px] shrink-0 max-w-[160px]" style={{ color: "#B3372A" }}>{cnhStatus.error}</span>}
          {cnhStatus?.success && <span className="text-[11px] shrink-0" style={{ color: "#1F7A54" }}>Preenchido, confira</span>}
        </div>

        <div className="grid grid-cols-2 gap-3">
          <Field label="Nome completo"><TextInput value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} /></Field>
          <Field label="CPF"><TextInput value={form.cpf} onChange={(e) => setForm({ ...form, cpf: e.target.value })} /></Field>
          <Field label="Telefone"><TextInput value={form.telefone} onChange={(e) => setForm({ ...form, telefone: e.target.value })} /></Field>
          <Field label="Nº da CNH"><TextInput value={form.cnhNumero} onChange={(e) => setForm({ ...form, cnhNumero: e.target.value })} /></Field>
          <Field label="Categoria CNH">
            <SelectInput value={form.cnhCategoria} onChange={(e) => setForm({ ...form, cnhCategoria: e.target.value })}>
              <option value="">Selecione</option>
              {["A", "B", "AB", "C", "D", "E"].map((c) => <option key={c} value={c}>{c}</option>)}
            </SelectInput>
          </Field>
          <Field label="Validade da CNH"><TextInput type="date" value={form.cnhValidade} onChange={(e) => setForm({ ...form, cnhValidade: e.target.value })} /></Field>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.nome} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Helpers
--------------------------------------------------------- */

const emptyHelper = { id: null, nome: "", cpf: "", telefone: "", foto: null };

function HelpersView({ helpers, setHelpers, perms = ROLES.admin, empresaNome, companyLogo }) {
  const [query, setQuery] = useState("");
  const [editing, setEditing] = useState(null);
  const [importing, setImporting] = useState(false);
  const [badgePerson, setBadgePerson] = useState(null);
  const [sheetOpen, setSheetOpen] = useState(false);
  const filtered = helpers.filter((h) => h.nome.toLowerCase().includes(query.toLowerCase()));

  const save = (helper) => {
    setHelpers((prev) => (prev.some((h) => h.id === helper.id) ? prev.map((h) => (h.id === helper.id ? helper : h)) : [...prev, helper]));
    setEditing(null);
  };
  const remove = (id) => setHelpers((prev) => prev.filter((h) => h.id !== id));

  const helperImportFields = [
    { key: "nome", label: "Nome", required: true },
    { key: "cpf", label: "CPF" },
    { key: "telefone", label: "Telefone" },
  ];
  const handleImport = (rows) => {
    const novos = rows.filter((r) => r.nome).map((r) => ({ ...emptyHelper, ...r, id: uid(), foto: null }));
    setHelpers((prev) => [...prev, ...novos]);
    setImporting(false);
  };

  return (
    <div>
      <ListToolbar
        query={query}
        setQuery={setQuery}
        placeholder="Buscar ajudante"
        onAdd={() => setEditing({ ...emptyHelper, id: uid() })}
        addLabel="Novo ajudante"
        canCreate={perms.canCreate}
        onImport={() => setImporting(true)}
      />
      {filtered.length > 0 && (
        <div className="flex justify-end mb-3 -mt-2">
          <button
            onClick={() => setSheetOpen(true)}
            className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg"
            style={{ background: "#FFF1E0", color: "#C9600A" }}
          >
            <CreditCard size={13} /> Imprimir crachás ({filtered.length})
          </button>
        </div>
      )}
      {filtered.length === 0 ? (
        <EmptyState icon={HardHat} title="Nenhum ajudante cadastrado" subtitle="Cadastre a equipe de apoio" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          {filtered.map((h, i) => (
            <div key={h.id} className="flex items-center gap-3 px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
              <Avatar src={h.foto} fallback={<HardHat size={15} />} size={40} />
              <div className="min-w-0 flex-1">
                <p className="text-sm font-medium truncate" style={{ color: "#10151C" }}>{h.nome}</p>
                <p className="text-xs" style={{ color: "#9AA1AA" }}>{h.telefone || "Sem telefone"}</p>
              </div>
              <button onClick={() => setBadgePerson(h)} className="p-1.5 rounded-lg" style={{ background: "#FFF1E0" }} title="Imprimir crachá">
                <CreditCard size={13} style={{ color: "#C9600A" }} />
              </button>
              <RowActions onEdit={() => setEditing(h)} onDelete={() => remove(h.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
            </div>
          ))}
        </div>
      )}
      {editing && <HelperForm helper={editing} onClose={() => setEditing(null)} onSave={save} />}
      {importing && (
        <ImportModal
          title="Importar ajudantes de planilha"
          fields={helperImportFields}
          onClose={() => setImporting(false)}
          onImport={handleImport}
        />
      )}
      {badgePerson && <BadgePrintModal person={badgePerson} tipo="ajudante" empresaNome={empresaNome} companyLogo={companyLogo} onClose={() => setBadgePerson(null)} />}
      {sheetOpen && <BadgeSheetModal people={filtered} tipo="ajudante" empresaNome={empresaNome} companyLogo={companyLogo} onClose={() => setSheetOpen(false)} />}
    </div>
  );
}

function HelperForm({ helper, onClose, onSave }) {
  const [form, setForm] = useState(helper);
  return (
    <Modal title={helper.nome ? "Editar ajudante" : "Novo ajudante"} onClose={onClose}>
      <div className="space-y-4">
        <PhotoPicker label="Foto" value={form.foto} onChange={(v) => setForm({ ...form, foto: v })} />
        <Field label="Nome completo"><TextInput value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} /></Field>
        <Field label="CPF"><TextInput value={form.cpf} onChange={(e) => setForm({ ...form, cpf: e.target.value })} /></Field>
        <Field label="Telefone"><TextInput value={form.telefone} onChange={(e) => setForm({ ...form, telefone: e.target.value })} /></Field>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.nome} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Loadings (entrada/saída)
--------------------------------------------------------- */

function LoadingsView({ loadings, setLoadings, trucks, drivers, perms = ROLES.admin }) {
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState(null);
  const [filterTipo, setFilterTipo] = useState("Todos");

  const sorted = [...loadings].sort((a, b) => (b.data + b.hora).localeCompare(a.data + a.hora));
  const filtered = filterTipo === "Todos" ? sorted : sorted.filter((l) => l.tipo === filterTipo);

  const add = (entry) => {
    setLoadings((prev) => [...prev, entry]);
    setOpen(false);
  };
  const update = (entry) => {
    setLoadings((prev) => prev.map((l) => (l.id === entry.id ? entry : l)));
    setEditing(null);
  };
  const remove = (id) => setLoadings((prev) => prev.filter((l) => l.id !== id));

  return (
    <div>
      <div className="flex items-center gap-3 mb-4">
        <div className="flex rounded-lg p-1" style={{ background: "#EEF0F2" }}>
          {["Todos", "Entrada", "Saída"].map((t) => (
            <button
              key={t}
              onClick={() => setFilterTipo(t)}
              className="text-xs font-medium px-3 py-1.5 rounded-md"
              style={{ background: filterTipo === t ? "#FFFFFF" : "transparent", color: filterTipo === t ? "#10151C" : "#6B7480" }}
            >
              {t}
            </button>
          ))}
        </div>
        {perms.canCreate && (
          <button
            onClick={() => setOpen(true)}
            className="ml-auto flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg shrink-0"
            style={{ background: "#FF8A1E", color: "#151C26" }}
          >
            <Plus size={15} /> Registrar movimentação
          </button>
        )}
      </div>

      {trucks.length === 0 || drivers.length === 0 ? (
        <p className="text-xs mb-4 px-1" style={{ color: "#9AA1AA" }}>Cadastre ao menos um caminhão e um motorista para registrar movimentações.</p>
      ) : null}

      {filtered.length === 0 ? (
        <EmptyState icon={ClipboardList} title="Nenhuma movimentação" subtitle="Os registros de entrada e saída aparecerão aqui" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="grid grid-cols-[100px_1fr_1fr_100px_72px] gap-2 px-4 py-2.5 text-xs font-medium" style={{ color: "#9AA1AA", borderBottom: "1px solid #E7E9EC" }}>
            <span>Data / hora</span><span>Caminhão</span><span>Motorista</span><span>Tipo</span><span />
          </div>
          {filtered.map((l, i) => (
            <div key={l.id} className="grid grid-cols-[100px_1fr_1fr_100px_72px] gap-2 items-center px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
              <div>
                <p className="text-sm font-medium" style={{ color: "#10151C" }}>{l.data.split("-").reverse().join("/")}</p>
                <p className="text-xs flex items-center gap-1" style={{ color: "#9AA1AA" }}><Clock size={11} />{l.hora}</p>
              </div>
              <p className="text-sm font-mono font-semibold" style={{ color: "#10151C" }}>{l.placa}</p>
              <p className="text-sm truncate" style={{ color: "#10151C" }}>{l.motorista}</p>
              <StatusPill tone={l.tipo === "Entrada" ? "ok" : "warn"}>{l.tipo}</StatusPill>
              <RowActions onEdit={() => setEditing(l)} onDelete={() => remove(l.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
            </div>
          ))}
        </div>
      )}

      {open && <LoadingForm onClose={() => setOpen(false)} onSave={add} trucks={trucks} drivers={drivers} />}
      {editing && <LoadingForm loading={editing} onClose={() => setEditing(null)} onSave={update} trucks={trucks} drivers={drivers} />}
    </div>
  );
}

function LoadingForm({ loading, onClose, onSave, trucks, drivers }) {
  const now = new Date();
  const [form, setForm] = useState(
    loading || {
      id: uid(),
      tipo: "Entrada",
      placa: trucks[0]?.placa || "",
      motorista: drivers[0]?.nome || "",
      data: todayISO(),
      hora: now.toTimeString().slice(0, 5),
    }
  );

  return (
    <Modal title={loading ? "Editar movimentação" : "Registrar movimentação"} onClose={onClose}>
      <div className="space-y-4">
        {(trucks.length === 0 || drivers.length === 0) && (
          <p className="text-xs rounded-lg px-3 py-2" style={{ background: "#FFF1E0", color: "#C9600A" }}>
            Cadastre ao menos um caminhão e um motorista antes de registrar a movimentação.
          </p>
        )}
        <div className="flex rounded-lg p-1" style={{ background: "#EEF0F2" }}>
          {["Entrada", "Saída"].map((t) => (
            <button
              key={t}
              onClick={() => setForm({ ...form, tipo: t })}
              className="flex-1 text-sm font-medium py-2 rounded-md"
              style={{ background: form.tipo === t ? "#FFFFFF" : "transparent", color: form.tipo === t ? "#10151C" : "#6B7480" }}
            >
              {t}
            </button>
          ))}
        </div>
        <Field label="Caminhão">
          <SelectInput value={form.placa} onChange={(e) => setForm({ ...form, placa: e.target.value })}>
            <option value="">Selecione</option>
            {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa} — {t.modelo}</option>)}
          </SelectInput>
        </Field>
        <Field label="Motorista">
          <SelectInput value={form.motorista} onChange={(e) => setForm({ ...form, motorista: e.target.value })}>
            <option value="">Selecione</option>
            {drivers.map((d) => <option key={d.id} value={d.nome}>{d.nome}</option>)}
          </SelectInput>
        </Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Data"><TextInput type="date" value={form.data} onChange={(e) => setForm({ ...form, data: e.target.value })} /></Field>
          <Field label="Hora"><TextInput type="time" value={form.hora} onChange={(e) => setForm({ ...form, hora: e.target.value })} /></Field>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.placa || !form.motorista} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Controle de Acesso (QR Code)
--------------------------------------------------------- */

function parseQrToken(raw) {
  const text = (raw || "").trim();
  const parts = text.split("|");
  if (parts.length === 3 && parts[0] === "FROTAPLUS") {
    return { kind: parts[1], id: parts[2] };
  }
  return null;
}

function AccessControlView({ loadings, setLoadings, trucks, drivers, helpers, perms = ROLES.admin }) {
  const [scanning, setScanning] = useState(false);
  const [scanSupported, setScanSupported] = useState(true);
  const [scanError, setScanError] = useState("");
  const [manualCode, setManualCode] = useState("");
  const [selectedPerson, setSelectedPerson] = useState(null); // { id, nome, tipo }
  const [placa, setPlaca] = useState(trucks[0]?.placa || "");
  const [tipoMov, setTipoMov] = useState("Entrada");
  const [query, setQuery] = useState("");
  const [justRegistered, setJustRegistered] = useState(null);

  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const detectLoopRef = useRef(null);

  const everyone = useMemo(() => {
    const d = drivers.map((p) => ({ ...p, tipo: "motorista" }));
    const h = helpers.map((p) => ({ ...p, tipo: "ajudante" }));
    return [...d, ...h];
  }, [drivers, helpers]);

  const findPerson = (kind, id) => {
    if (kind === "MOTORISTA") return drivers.find((d) => d.id === id);
    if (kind === "AJUDANTE") return helpers.find((h) => h.id === id);
    return null;
  };

  const stopScan = () => {
    setScanning(false);
    if (detectLoopRef.current) cancelAnimationFrame(detectLoopRef.current);
    if (streamRef.current) {
      streamRef.current.getTracks().forEach((t) => t.stop());
      streamRef.current = null;
    }
  };

  const handleDecoded = (text) => {
    const parsed = parseQrToken(text);
    if (!parsed) {
      setScanError("Código lido não é um crachá válido do Frota+.");
      return;
    }
    const person = findPerson(parsed.kind, parsed.id);
    if (!person) {
      setScanError("Crachá lido, mas a pessoa não foi encontrada no cadastro.");
      return;
    }
    setSelectedPerson({ ...person, tipo: parsed.kind === "MOTORISTA" ? "motorista" : "ajudante" });
    setScanError("");
    stopScan();
  };

  const startScan = async () => {
    setScanError("");
    if (typeof window === "undefined" || !("BarcodeDetector" in window)) {
      setScanSupported(false);
      return;
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      streamRef.current = stream;
      setScanning(true);
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
      }
      const detector = new window.BarcodeDetector({ formats: ["qr_code"] });
      const loop = async () => {
        if (!videoRef.current || !streamRef.current) return;
        try {
          const codes = await detector.detect(videoRef.current);
          if (codes && codes.length > 0) {
            handleDecoded(codes[0].rawValue);
            return;
          }
        } catch {
          // frame não decodificável — tenta o próximo
        }
        detectLoopRef.current = requestAnimationFrame(loop);
      };
      detectLoopRef.current = requestAnimationFrame(loop);
    } catch (err) {
      setScanError("Não foi possível acessar a câmera. Verifique a permissão do navegador ou use a busca manual abaixo.");
      setScanning(false);
    }
  };

  useEffect(() => () => stopScan(), []);

  const registrar = () => {
    if (!selectedPerson || !placa) return;
    const now = new Date();
    const entry = {
      id: uid(),
      tipo: tipoMov,
      placa,
      motorista: selectedPerson.tipo === "motorista" ? selectedPerson.nome : (drivers[0]?.nome || selectedPerson.nome),
      data: todayISO(),
      hora: now.toTimeString().slice(0, 5),
      viaQr: true,
      pessoaAcesso: selectedPerson.nome,
      pessoaTipo: selectedPerson.tipo,
    };
    setLoadings((prev) => [...prev, entry]);
    setJustRegistered(entry);
    setSelectedPerson(null);
    setManualCode("");
    setTimeout(() => setJustRegistered(null), 4000);
  };

  const recentAccess = [...loadings].filter((l) => l.viaQr).sort((a, b) => (b.data + b.hora).localeCompare(a.data + a.hora)).slice(0, 8);
  const filteredPeople = everyone.filter((p) => p.nome.toLowerCase().includes(query.toLowerCase()));

  if (trucks.length === 0) {
    return (
      <EmptyState icon={ScanLine} title="Cadastre um caminhão primeiro" subtitle="O controle de acesso precisa de ao menos um caminhão cadastrado" />
    );
  }

  return (
    <div className="grid grid-cols-1 lg:grid-cols-5 gap-4">
      <div className="lg:col-span-3 space-y-4">
        <div className="rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-2 mb-4">
            <Camera size={15} style={{ color: "#C9600A" }} />
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Leitor de QR do crachá</h2>
          </div>

          {!scanSupported && (
            <p className="text-xs rounded-lg px-3 py-2 mb-3" style={{ background: "#FFF1E0", color: "#C9600A" }}>
              Este navegador não suporta leitura de QR pela câmera. Use a busca manual abaixo — funciona igual, só sem escanear.
            </p>
          )}
          {scanError && (
            <p className="text-xs rounded-lg px-3 py-2 mb-3" style={{ background: "#FDEBE8", color: "#B3372A" }}>{scanError}</p>
          )}

          {scanning ? (
            <div className="space-y-3">
              <div className="rounded-xl overflow-hidden" style={{ background: "#0B0F14", aspectRatio: "4/3" }}>
                <video ref={videoRef} className="w-full h-full object-cover" muted playsInline />
              </div>
              <button onClick={stopScan} className="w-full text-sm font-medium px-4 py-2.5 rounded-lg" style={{ background: "#F5F6F8", color: "#10151C" }}>
                Cancelar leitura
              </button>
            </div>
          ) : (
            scanSupported && (
              <button
                onClick={startScan}
                className="w-full flex items-center justify-center gap-2 text-sm font-medium px-4 py-3 rounded-lg"
                style={{ background: "#FF8A1E", color: "#151C26" }}
              >
                <Camera size={16} /> Ativar câmera e escanear crachá
              </button>
            )
          )}
        </div>

        <div className="rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="flex items-center gap-2 mb-3">
            <Search size={14} style={{ color: "#6B7480" }} />
            <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Ou busque manualmente</h2>
          </div>
          <TextInput
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Nome do motorista ou ajudante"
            className={inputClass}
            style={inputStyle}
          />
          {query && (
            <div className="mt-2 rounded-lg overflow-hidden" style={{ border: "1px solid #E7E9EC" }}>
              {filteredPeople.length === 0 ? (
                <p className="text-xs px-3 py-2.5" style={{ color: "#9AA1AA" }}>Ninguém encontrado</p>
              ) : (
                filteredPeople.slice(0, 6).map((p, i) => (
                  <button
                    key={p.id}
                    onClick={() => { setSelectedPerson(p); setQuery(""); }}
                    className="w-full flex items-center gap-2.5 px-3 py-2.5 text-left"
                    style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}
                  >
                    <span className="text-sm" style={{ color: "#10151C" }}>{p.nome}</span>
                    <span className="text-[10px] ml-auto px-1.5 py-0.5 rounded" style={{ background: p.tipo === "motorista" ? "#E8F1FC" : "#E7F7EF", color: p.tipo === "motorista" ? "#2A66C9" : "#1F7A54" }}>
                      {p.tipo === "motorista" ? "Motorista" : "Ajudante"}
                    </span>
                  </button>
                ))
              )}
            </div>
          )}
        </div>

        {selectedPerson && (
          <div className="rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #FF8A1E" }}>
            <div className="flex items-center gap-3 mb-4">
              <div className="rounded-full overflow-hidden flex items-center justify-center shrink-0" style={{ width: 44, height: 44, background: "#EEF0F2" }}>
                {selectedPerson.foto ? <img src={selectedPerson.foto} className="w-full h-full object-cover" alt="" /> : <span className="text-sm font-semibold" style={{ color: "#9AA1AA" }}>{selectedPerson.nome.slice(0, 1).toUpperCase()}</span>}
              </div>
              <div className="min-w-0 flex-1">
                <p className="text-sm font-semibold truncate" style={{ color: "#10151C" }}>{selectedPerson.nome}</p>
                <p className="text-xs" style={{ color: "#9AA1AA" }}>{selectedPerson.tipo === "motorista" ? "Motorista" : "Ajudante"}</p>
              </div>
              <button onClick={() => setSelectedPerson(null)} className="p-1.5 rounded-lg shrink-0" style={{ background: "#F5F6F8" }}><X size={14} /></button>
            </div>
            <div className="space-y-3">
              <Field label="Caminhão">
                <SelectInput value={placa} onChange={(e) => setPlaca(e.target.value)}>
                  {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa} — {t.modelo}</option>)}
                </SelectInput>
              </Field>
              <div className="flex rounded-lg p-1" style={{ background: "#EEF0F2" }}>
                {["Entrada", "Saída"].map((t) => (
                  <button
                    key={t}
                    onClick={() => setTipoMov(t)}
                    className="flex-1 text-sm font-medium py-2 rounded-md"
                    style={{ background: tipoMov === t ? "#FFFFFF" : "transparent", color: tipoMov === t ? "#10151C" : "#6B7480" }}
                  >
                    {t}
                  </button>
                ))}
              </div>
              {perms.canCreate ? (
                <button onClick={registrar} className="w-full text-sm font-semibold px-4 py-3 rounded-lg" style={{ background: "#FF8A1E", color: "#151C26" }}>
                  Registrar {tipoMov.toLowerCase()}
                </button>
              ) : (
                <p className="text-xs text-center" style={{ color: "#9AA1AA" }}>Seu perfil não tem permissão para registrar acessos.</p>
              )}
            </div>
          </div>
        )}

        {justRegistered && (
          <div className="flex items-center gap-2 rounded-lg px-3 py-2.5" style={{ background: "#E7F7EF" }}>
            <CheckCircle2 size={15} style={{ color: "#1F9D6B" }} />
            <p className="text-xs" style={{ color: "#1F7A54" }}>
              {justRegistered.pessoaTipo === "motorista" ? "Motorista" : "Ajudante"} {justRegistered.pessoaAcesso} — {justRegistered.tipo.toLowerCase()} registrada às {justRegistered.hora}.
            </p>
          </div>
        )}
      </div>

      <div className="lg:col-span-2 rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
        <div className="flex items-center gap-2 mb-4">
          <ClipboardList size={15} style={{ color: "#7C3AED" }} />
          <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Últimos acessos via QR</h2>
        </div>
        {recentAccess.length === 0 ? (
          <EmptyState icon={ScanLine} title="Nenhum acesso registrado ainda" subtitle="Escaneie um crachá ou busque uma pessoa para começar" />
        ) : (
          <div className="space-y-2.5">
            {recentAccess.map((a) => (
              <div key={a.id} className="flex items-center gap-2.5 py-2" style={{ borderBottom: "1px solid #F0F1F3" }}>
                <span
                  className="w-7 h-7 rounded-lg flex items-center justify-center shrink-0"
                  style={{ background: a.tipo === "Entrada" ? "#E7F7EF" : "#FFF1E0", color: a.tipo === "Entrada" ? "#1F9D6B" : "#C9600A" }}
                >
                  <ArrowRightLeft size={13} />
                </span>
                <div className="min-w-0 flex-1">
                  <p className="text-sm font-medium truncate" style={{ color: "#10151C" }}>{a.pessoaAcesso} · {a.placa}</p>
                  <p className="text-xs" style={{ color: "#9AA1AA" }}>{a.data.split("-").reverse().join("/")} às {a.hora}</p>
                </div>
                <StatusPill tone={a.tipo === "Entrada" ? "ok" : "warn"}>{a.tipo}</StatusPill>
              </div>
            ))}
          </div>
        )}
        <p className="text-[11px] mt-4" style={{ color: "#9AA1AA" }}>
          Os crachás com QR ficam disponíveis em Motoristas e Ajudantes, no botão de imprimir crachá.
        </p>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   Manutenção
--------------------------------------------------------- */

function MaintenanceView({ maintenances, setMaintenances, trucks, perms = ROLES.admin }) {
  const [editing, setEditing] = useState(null);
  const [query, setQuery] = useState("");

  const sorted = [...maintenances].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  const filtered = sorted.filter((m) => (m.placa + m.descricao).toLowerCase().includes(query.toLowerCase()));

  const save = (m) => {
    setMaintenances((prev) => {
      const exists = prev.some((x) => x.id === m.id);
      return exists ? prev.map((x) => (x.id === m.id ? m : x)) : [...prev, m];
    });
    setEditing(null);
  };
  const remove = (id) => setMaintenances((prev) => prev.filter((m) => m.id !== id));

  return (
    <div>
      <ListToolbar
        query={query}
        setQuery={setQuery}
        placeholder="Buscar por placa ou descrição"
        onAdd={() => setEditing({ ...emptyMaintenance, id: uid(), data: todayISO() })}
        addLabel="Nova manutenção"
        canCreate={perms.canCreate}
      />
      {trucks.length === 0 && (
        <p className="text-xs mb-4 px-1" style={{ color: "#9AA1AA" }}>Cadastre ao menos um caminhão para registrar manutenções.</p>
      )}
      {filtered.length === 0 ? (
        <EmptyState icon={Wrench} title="Nenhuma manutenção registrada" subtitle="Revisões, trocas e reparos aparecerão aqui" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="grid grid-cols-[90px_80px_1fr_90px_100px_40px] gap-2 px-4 py-2.5 text-xs font-medium" style={{ color: "#9AA1AA", borderBottom: "1px solid #E7E9EC" }}>
            <span>Data</span><span>Placa</span><span>Descrição</span><span>Custo</span><span>Próxima</span><span />
          </div>
          {filtered.map((m, i) => {
            const alert = maintenanceAlertStatus(m);
            return (
              <div key={m.id} className="grid grid-cols-[90px_80px_1fr_90px_100px_40px] gap-2 items-center px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
                <span className="text-sm" style={{ color: "#10151C" }}>{m.data ? m.data.split("-").reverse().join("/") : "—"}</span>
                <span className="text-sm font-mono font-semibold" style={{ color: "#10151C" }}>{m.placa}</span>
                <div className="min-w-0">
                  <p className="text-sm truncate" style={{ color: "#10151C" }}>{m.descricao || "—"}</p>
                  <p className="text-xs" style={{ color: "#9AA1AA" }}>{m.tipo}{m.oficina ? ` · ${m.oficina}` : ""}</p>
                </div>
                <span className="text-sm" style={{ color: "#10151C" }}>{m.custo ? `R$ ${Number(m.custo).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}` : "—"}</span>
                {alert ? <StatusPill tone={alert.tone}>{alert.label}</StatusPill> : <span className="text-xs" style={{ color: "#9AA1AA" }}>{m.proximaData ? m.proximaData.split("-").reverse().join("/") : "—"}</span>}
                <RowActions onEdit={() => setEditing(m)} onDelete={() => remove(m.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
              </div>
            );
          })}
        </div>
      )}
      {editing && <MaintenanceForm maintenance={editing} trucks={trucks} onClose={() => setEditing(null)} onSave={save} />}
    </div>
  );
}

function MaintenanceForm({ maintenance, trucks, onClose, onSave }) {
  const [form, setForm] = useState(maintenance);
  return (
    <Modal title={maintenance.descricao ? "Editar manutenção" : "Nova manutenção"} onClose={onClose} wide>
      <div className="space-y-4">
        <Field label="Caminhão">
          <SelectInput value={form.placa} onChange={(e) => setForm({ ...form, placa: e.target.value })}>
            <option value="">Selecione</option>
            {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa} — {t.modelo}</option>)}
          </SelectInput>
        </Field>
        <div className="flex rounded-lg p-1" style={{ background: "#EEF0F2" }}>
          {["Preventiva", "Corretiva"].map((t) => (
            <button
              key={t}
              onClick={() => setForm({ ...form, tipo: t })}
              className="flex-1 text-sm font-medium py-2 rounded-md"
              style={{ background: form.tipo === t ? "#FFFFFF" : "transparent", color: form.tipo === t ? "#10151C" : "#6B7480" }}
            >
              {t}
            </button>
          ))}
        </div>
        <Field label="Descrição"><TextInput value={form.descricao} onChange={(e) => setForm({ ...form, descricao: e.target.value })} placeholder="Ex.: Troca de óleo e filtros" /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Data realizada"><TextInput type="date" value={form.data} onChange={(e) => setForm({ ...form, data: e.target.value })} /></Field>
          <Field label="Km no momento"><TextInput type="number" value={form.km} onChange={(e) => setForm({ ...form, km: e.target.value })} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Custo (R$)"><TextInput type="number" step="0.01" value={form.custo} onChange={(e) => setForm({ ...form, custo: e.target.value })} /></Field>
          <Field label="Oficina/Fornecedor"><TextInput value={form.oficina} onChange={(e) => setForm({ ...form, oficina: e.target.value })} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Próxima manutenção (data)"><TextInput type="date" value={form.proximaData} onChange={(e) => setForm({ ...form, proximaData: e.target.value })} /></Field>
          <Field label="Próxima manutenção (km)"><TextInput type="number" value={form.proximoKm} onChange={(e) => setForm({ ...form, proximoKm: e.target.value })} /></Field>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.placa || !form.descricao} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Abastecimento
--------------------------------------------------------- */

function FuelView({ fuels, setFuels, trucks, drivers, perms = ROLES.admin }) {
  const [editing, setEditing] = useState(null);

  const sorted = [...fuels].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  const totalMes = useMemo(() => {
    const mesAtual = todayISO().slice(0, 7);
    return fuels.filter((f) => (f.data || "").slice(0, 7) === mesAtual).reduce((sum, f) => sum + (Number(f.valorTotal) || 0), 0);
  }, [fuels]);

  const save = (f) => {
    setFuels((prev) => {
      const exists = prev.some((x) => x.id === f.id);
      return exists ? prev.map((x) => (x.id === f.id ? f : x)) : [...prev, f];
    });
    setEditing(null);
  };
  const remove = (id) => setFuels((prev) => prev.filter((f) => f.id !== id));

  return (
    <div>
      <div className="flex items-center gap-3 mb-4">
        <div className="rounded-xl px-4 py-2.5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <p className="text-[11px]" style={{ color: "#9AA1AA" }}>Gasto no mês</p>
          <p className="text-sm font-semibold" style={{ color: "#10151C" }}>R$ {totalMes.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
        </div>
        {perms.canCreate && (
          <button
            onClick={() => setEditing({ ...emptyFuel, id: uid(), data: todayISO(), placa: trucks[0]?.placa || "", motorista: drivers[0]?.nome || "" })}
            disabled={trucks.length === 0}
            className="ml-auto flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg shrink-0"
            style={{ background: trucks.length === 0 ? "#F0D4B0" : "#FF8A1E", color: "#151C26" }}
          >
            <Plus size={15} /> Novo abastecimento
          </button>
        )}
      </div>
      {trucks.length === 0 && (
        <p className="text-xs mb-4 px-1" style={{ color: "#9AA1AA" }}>Cadastre ao menos um caminhão para registrar abastecimentos.</p>
      )}
      {fuels.length === 0 ? (
        <EmptyState icon={Fuel} title="Nenhum abastecimento registrado" subtitle="Registros de combustível aparecerão aqui" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="grid grid-cols-[90px_80px_1fr_70px_90px_40px] gap-2 px-4 py-2.5 text-xs font-medium" style={{ color: "#9AA1AA", borderBottom: "1px solid #E7E9EC" }}>
            <span>Data</span><span>Placa</span><span>Posto / Motorista</span><span>Litros</span><span>Valor</span><span />
          </div>
          {sorted.map((f, i) => (
            <div key={f.id} className="grid grid-cols-[90px_80px_1fr_70px_90px_40px] gap-2 items-center px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.data ? f.data.split("-").reverse().join("/") : "—"}</span>
              <span className="text-sm font-mono font-semibold" style={{ color: "#10151C" }}>{f.placa}</span>
              <div className="min-w-0">
                <p className="text-sm truncate" style={{ color: "#10151C" }}>{f.posto || "—"}</p>
                <p className="text-xs truncate" style={{ color: "#9AA1AA" }}>{f.motorista}</p>
              </div>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.litros || "—"}</span>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.valorTotal ? `R$ ${Number(f.valorTotal).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}` : "—"}</span>
              <RowActions onEdit={() => setEditing(f)} onDelete={() => remove(f.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
            </div>
          ))}
        </div>
      )}
      {editing && <FuelForm fuel={editing} trucks={trucks} drivers={drivers} onClose={() => setEditing(null)} onSave={save} />}
    </div>
  );
}

function FuelForm({ fuel, trucks, drivers, onClose, onSave }) {
  const [form, setForm] = useState(fuel);
  return (
    <Modal title="Abastecimento" onClose={onClose} wide>
      <div className="space-y-4">
        {trucks.length === 0 && (
          <p className="text-xs rounded-lg px-3 py-2" style={{ background: "#FFF1E0", color: "#C9600A" }}>
            Cadastre ao menos um caminhão antes de registrar o abastecimento.
          </p>
        )}
        <div className="grid grid-cols-2 gap-3">
          <Field label="Caminhão">
            <SelectInput value={form.placa} onChange={(e) => setForm({ ...form, placa: e.target.value })}>
              <option value="">Selecione</option>
              {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa} — {t.modelo}</option>)}
            </SelectInput>
          </Field>
          <Field label="Motorista">
            <SelectInput value={form.motorista} onChange={(e) => setForm({ ...form, motorista: e.target.value })}>
              <option value="">—</option>
              {drivers.map((d) => <option key={d.id} value={d.nome}>{d.nome}</option>)}
            </SelectInput>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Data"><TextInput type="date" value={form.data} onChange={(e) => setForm({ ...form, data: e.target.value })} /></Field>
          <Field label="Posto"><TextInput value={form.posto} onChange={(e) => setForm({ ...form, posto: e.target.value })} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3">
          <Field label="Litros"><TextInput type="number" step="0.01" value={form.litros} onChange={(e) => setForm({ ...form, litros: e.target.value })} /></Field>
          <Field label="Valor total (R$)"><TextInput type="number" step="0.01" value={form.valorTotal} onChange={(e) => setForm({ ...form, valorTotal: e.target.value })} /></Field>
          <Field label="Km atual"><TextInput type="number" value={form.kmAtual} onChange={(e) => setForm({ ...form, kmAtual: e.target.value })} /></Field>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.placa || !form.litros} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Multas e infrações
--------------------------------------------------------- */

function FinesView({ fines, setFines, trucks, drivers, perms = ROLES.admin }) {
  const [editing, setEditing] = useState(null);
  const [filterStatus, setFilterStatus] = useState("Todas");

  const sorted = [...fines].sort((a, b) => (b.data || "").localeCompare(a.data || ""));
  const filtered = filterStatus === "Todas" ? sorted : sorted.filter((f) => f.status === filterStatus);

  const save = (f) => {
    setFines((prev) => {
      const exists = prev.some((x) => x.id === f.id);
      return exists ? prev.map((x) => (x.id === f.id ? f : x)) : [...prev, f];
    });
    setEditing(null);
  };
  const remove = (id) => setFines((prev) => prev.filter((f) => f.id !== id));

  return (
    <div>
      <div className="flex items-center gap-3 mb-4">
        <div className="flex rounded-lg p-1" style={{ background: "#EEF0F2" }}>
          {["Todas", "Pendente", "Paga", "Recorrida"].map((t) => (
            <button
              key={t}
              onClick={() => setFilterStatus(t)}
              className="text-xs font-medium px-3 py-1.5 rounded-md"
              style={{ background: filterStatus === t ? "#FFFFFF" : "transparent", color: filterStatus === t ? "#10151C" : "#6B7480" }}
            >
              {t}
            </button>
          ))}
        </div>
        {perms.canCreate && (
          <button
            onClick={() => setEditing({ ...emptyFine, id: uid(), data: todayISO(), placa: trucks[0]?.placa || "", motorista: drivers[0]?.nome || "" })}
            disabled={trucks.length === 0}
            className="ml-auto flex items-center gap-1.5 text-sm font-medium px-3.5 py-2 rounded-lg shrink-0"
            style={{ background: trucks.length === 0 ? "#F0D4B0" : "#FF8A1E", color: "#151C26" }}
          >
            <Plus size={15} /> Nova multa
          </button>
        )}
      </div>
      {trucks.length === 0 && (
        <p className="text-xs mb-4 px-1" style={{ color: "#9AA1AA" }}>Cadastre ao menos um caminhão para registrar multas.</p>
      )}
      {filtered.length === 0 ? (
        <EmptyState icon={AlertOctagon} title="Nenhuma multa registrada" subtitle="Infrações e multas aparecerão aqui" />
      ) : (
        <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <div className="grid grid-cols-[90px_80px_1fr_80px_100px_100px_40px] gap-2 px-4 py-2.5 text-xs font-medium" style={{ color: "#9AA1AA", borderBottom: "1px solid #E7E9EC" }}>
            <span>Data</span><span>Placa</span><span>Infração / Motorista</span><span>Valor</span><span>Vencimento</span><span>Status</span><span />
          </div>
          {filtered.map((f, i) => (
            <div key={f.id} className="grid grid-cols-[90px_80px_1fr_80px_100px_100px_40px] gap-2 items-center px-4 py-3" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.data ? f.data.split("-").reverse().join("/") : "—"}</span>
              <span className="text-sm font-mono font-semibold" style={{ color: "#10151C" }}>{f.placa}</span>
              <div className="min-w-0">
                <p className="text-sm truncate" style={{ color: "#10151C" }}>{f.infracao || "—"}</p>
                <p className="text-xs truncate" style={{ color: "#9AA1AA" }}>{f.motorista}{f.pontos ? ` · ${f.pontos} pts` : ""}</p>
              </div>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.valor ? `R$ ${Number(f.valor).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}` : "—"}</span>
              <span className="text-sm" style={{ color: "#10151C" }}>{f.vencimento ? f.vencimento.split("-").reverse().join("/") : "—"}</span>
              <StatusPill tone={f.status === "Paga" ? "ok" : f.status === "Recorrida" ? "warn" : "danger"}>{f.status}</StatusPill>
              <RowActions onEdit={() => setEditing(f)} onDelete={() => remove(f.id)} canEdit={perms.canEdit} canDelete={perms.canDelete} />
            </div>
          ))}
        </div>
      )}
      {editing && <FineForm fine={editing} trucks={trucks} drivers={drivers} onClose={() => setEditing(null)} onSave={save} />}
    </div>
  );
}

function FineForm({ fine, trucks, drivers, onClose, onSave }) {
  const [form, setForm] = useState(fine);
  return (
    <Modal title="Multa / infração" onClose={onClose} wide>
      <div className="space-y-4">
        <div className="grid grid-cols-2 gap-3">
          <Field label="Caminhão">
            <SelectInput value={form.placa} onChange={(e) => setForm({ ...form, placa: e.target.value })}>
              {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa} — {t.modelo}</option>)}
            </SelectInput>
          </Field>
          <Field label="Motorista">
            <SelectInput value={form.motorista} onChange={(e) => setForm({ ...form, motorista: e.target.value })}>
              <option value="">—</option>
              {drivers.map((d) => <option key={d.id} value={d.nome}>{d.nome}</option>)}
            </SelectInput>
          </Field>
        </div>
        <Field label="Descrição da infração"><TextInput value={form.infracao} onChange={(e) => setForm({ ...form, infracao: e.target.value })} placeholder="Ex.: Excesso de velocidade" /></Field>
        <div className="grid grid-cols-3 gap-3">
          <Field label="Data"><TextInput type="date" value={form.data} onChange={(e) => setForm({ ...form, data: e.target.value })} /></Field>
          <Field label="Valor (R$)"><TextInput type="number" step="0.01" value={form.valor} onChange={(e) => setForm({ ...form, valor: e.target.value })} /></Field>
          <Field label="Pontos na CNH"><TextInput type="number" value={form.pontos} onChange={(e) => setForm({ ...form, pontos: e.target.value })} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Vencimento"><TextInput type="date" value={form.vencimento} onChange={(e) => setForm({ ...form, vencimento: e.target.value })} /></Field>
          <Field label="Status">
            <SelectInput value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
              <option value="Pendente">Pendente</option>
              <option value="Paga">Paga</option>
              <option value="Recorrida">Recorrida</option>
            </SelectInput>
          </Field>
        </div>
        <FormFooter onClose={onClose} onSave={() => onSave(form)} disabled={!form.placa || !form.infracao} />
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------
   Relatórios
--------------------------------------------------------- */

function downloadCsv(rows) {
  const header = ["Data", "Hora", "Tipo", "Placa", "Motorista"];
  const lines = rows.map((r) => [r.data.split("-").reverse().join("/"), r.hora, r.tipo, r.placa, r.motorista].join(";"));
  const csv = [header.join(";"), ...lines].join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `relatorio-carregamentos-${todayISO()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function ReportsView({ loadings, trucks, drivers }) {
  const [filters, setFilters] = useState({ inicio: "", fim: "", placa: "Todos", motorista: "Todos", tipo: "Todos" });

  const filtered = useMemo(() => {
    return loadings.filter((l) => {
      if (filters.inicio && l.data < filters.inicio) return false;
      if (filters.fim && l.data > filters.fim) return false;
      if (filters.placa !== "Todos" && l.placa !== filters.placa) return false;
      if (filters.motorista !== "Todos" && l.motorista !== filters.motorista) return false;
      if (filters.tipo !== "Todos" && l.tipo !== filters.tipo) return false;
      return true;
    });
  }, [loadings, filters]);

  const entradas = filtered.filter((l) => l.tipo === "Entrada").length;
  const saidas = filtered.filter((l) => l.tipo === "Saída").length;
  const caminhoesAtivos = new Set(filtered.map((l) => l.placa)).size;

  const porDia = useMemo(() => {
    const map = {};
    filtered.forEach((l) => { map[l.data] = (map[l.data] || 0) + 1; });
    return Object.entries(map)
      .sort(([a], [b]) => a.localeCompare(b))
      .slice(-14)
      .map(([data, total]) => ({ dia: data.slice(5).split("-").reverse().join("/"), total }));
  }, [filtered]);

  const rankingMotoristas = useMemo(() => {
    const map = {};
    filtered.forEach((l) => { map[l.motorista] = (map[l.motorista] || 0) + 1; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, 5);
  }, [filtered]);

  const sorted = [...filtered].sort((a, b) => (b.data + b.hora).localeCompare(a.data + a.hora));

  return (
    <div className="space-y-5">
      <div className="rounded-2xl p-4 grid grid-cols-2 sm:grid-cols-5 gap-3" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
        <Field label="De">
          <TextInput type="date" value={filters.inicio} onChange={(e) => setFilters({ ...filters, inicio: e.target.value })} />
        </Field>
        <Field label="Até">
          <TextInput type="date" value={filters.fim} onChange={(e) => setFilters({ ...filters, fim: e.target.value })} />
        </Field>
        <Field label="Caminhão">
          <SelectInput value={filters.placa} onChange={(e) => setFilters({ ...filters, placa: e.target.value })}>
            <option>Todos</option>
            {trucks.map((t) => <option key={t.id} value={t.placa}>{t.placa}</option>)}
          </SelectInput>
        </Field>
        <Field label="Motorista">
          <SelectInput value={filters.motorista} onChange={(e) => setFilters({ ...filters, motorista: e.target.value })}>
            <option>Todos</option>
            {drivers.map((d) => <option key={d.id} value={d.nome}>{d.nome}</option>)}
          </SelectInput>
        </Field>
        <Field label="Tipo">
          <SelectInput value={filters.tipo} onChange={(e) => setFilters({ ...filters, tipo: e.target.value })}>
            <option>Todos</option>
            <option>Entrada</option>
            <option>Saída</option>
          </SelectInput>
        </Field>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {[
          { label: "Movimentações", value: filtered.length },
          { label: "Entradas", value: entradas },
          { label: "Saídas", value: saidas },
          { label: "Caminhões envolvidos", value: caminhoesAtivos },
        ].map((s) => (
          <div key={s.label} className="rounded-2xl p-4" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
            <p className="font-display text-2xl font-semibold" style={{ color: "#10151C" }}>{s.value}</p>
            <p className="text-xs mt-0.5" style={{ color: "#6B7480" }}>{s.label}</p>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-5 gap-4">
        <div className="lg:col-span-3 rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <h2 className="text-sm font-semibold mb-4" style={{ color: "#10151C" }}>Movimentações por dia</h2>
          {porDia.length === 0 ? (
            <EmptyState icon={BarChart3} title="Sem dados no período" subtitle="Ajuste os filtros ou registre movimentações" />
          ) : (
            <div style={{ width: "100%", height: 220 }}>
              <ResponsiveContainer>
                <BarChart data={porDia}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#F0F1F3" vertical={false} />
                  <XAxis dataKey="dia" tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={{ stroke: "#E7E9EC" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "#9AA1AA" }} axisLine={false} tickLine={false} allowDecimals={false} />
                  <Tooltip contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #E7E9EC" }} />
                  <Bar dataKey="total" fill="#FF8A1E" radius={[4, 4, 0, 0]} maxBarSize={28} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>

        <div className="lg:col-span-2 rounded-2xl p-5" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
          <h2 className="text-sm font-semibold mb-4" style={{ color: "#10151C" }}>Motoristas mais ativos</h2>
          {rankingMotoristas.length === 0 ? (
            <EmptyState icon={Users} title="Sem dados" subtitle="Nenhuma movimentação no período" />
          ) : (
            <div className="space-y-3">
              {rankingMotoristas.map(([nome, total], i) => (
                <div key={nome} className="flex items-center gap-3">
                  <span className="w-6 h-6 rounded-md flex items-center justify-center text-xs font-semibold shrink-0" style={{ background: "#FFF1E0", color: "#C9600A" }}>{i + 1}</span>
                  <span className="text-sm flex-1 truncate" style={{ color: "#10151C" }}>{nome}</span>
                  <span className="text-xs font-medium" style={{ color: "#9AA1AA" }}>{total} mov.</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="rounded-2xl overflow-hidden" style={{ background: "#FFFFFF", border: "1px solid #E7E9EC" }}>
        <div className="flex items-center justify-between px-5 py-4" style={{ borderBottom: "1px solid #E7E9EC" }}>
          <h2 className="text-sm font-semibold" style={{ color: "#10151C" }}>Registros do período ({sorted.length})</h2>
          <button
            onClick={() => downloadCsv(sorted)}
            disabled={sorted.length === 0}
            className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg"
            style={{ background: sorted.length === 0 ? "#F0F1F3" : "#E7F7EF", color: sorted.length === 0 ? "#B7BCC3" : "#1F7A54" }}
          >
            <Download size={13} /> Exportar CSV
          </button>
        </div>
        {sorted.length === 0 ? (
          <EmptyState icon={ClipboardList} title="Nenhum registro no período" subtitle="Ajuste os filtros acima" />
        ) : (
          <div className="max-h-96 overflow-y-auto">
            {sorted.map((l, i) => (
              <div key={l.id} className="grid grid-cols-[100px_1fr_1fr_100px] gap-2 items-center px-5 py-2.5" style={{ borderTop: i ? "1px solid #F0F1F3" : "none" }}>
                <p className="text-xs" style={{ color: "#6B7480" }}>{l.data.split("-").reverse().join("/")} {l.hora}</p>
                <p className="text-sm font-mono font-semibold" style={{ color: "#10151C" }}>{l.placa}</p>
                <p className="text-sm truncate" style={{ color: "#10151C" }}>{l.motorista}</p>
                <StatusPill tone={l.tipo === "Entrada" ? "ok" : "warn"}>{l.tipo}</StatusPill>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
