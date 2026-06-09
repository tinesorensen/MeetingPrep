import { useState, useEffect } from "react";
import * as msal from "@azure/msal-browser";

// ── MSAL config ──────────────────────────────────────────────────────────────
// Replace CLIENT_ID with your Azure App Registration client ID
const msalConfig = {
  auth: {
    clientId: "YOUR_AZURE_CLIENT_ID",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const GRAPH_SCOPES = ["Calendars.Read", "User.Read"];

// ── Translations ─────────────────────────────────────────────────────────────
const t = {
  da: {
    title: "Møde-prep generator",
    subtitle: "Generér en struktureret briefing til dit næste møde",
    loginBtn: "Log ind med Microsoft",
    logoutBtn: "Log ud",
    loggedInAs: "Logget ind som",
    upcomingMeetings: "Kommende møder (næste 7 dage)",
    loadingMeetings: "Henter møder fra Outlook…",
    noMeetings: "Ingen møder fundet de næste 7 dage.",
    clickToFill: "Klik på et møde for at udfylde formularen automatisk",
    who: "Hvem mødes du med?",
    whoPlaceholder: "Navn, titel, firma — hvad ved du om dem?",
    context: "Kontekst for mødet",
    contextPlaceholder: "Baggrund, historik, relationens karakter…",
    goal: "Dit mål med mødet",
    goalPlaceholder: "Hvad vil du gerne opnå?",
    sensitive: "Noget følsomt eller svært?",
    sensitivePlaceholder: "Konflikter, dårlige nyheder, følsomme emner…",
    notes: "Andre noter",
    notesPlaceholder: "Alt andet der er relevant…",
    generateBtn: "Generér briefing",
    generating: "Genererer…",
    copyBtn: "Kopiér",
    copied: "Kopieret!",
    newBtn: "Ny briefing",
    lang: "EN",
    required: "Udfyld felterne 'Hvem' og 'Mål' for at fortsætte.",
  },
  en: {
    title: "Meeting prep generator",
    subtitle: "Generate a structured briefing for your next meeting",
    loginBtn: "Sign in with Microsoft",
    logoutBtn: "Sign out",
    loggedInAs: "Signed in as",
    upcomingMeetings: "Upcoming meetings (next 7 days)",
    loadingMeetings: "Fetching meetings from Outlook…",
    noMeetings: "No meetings found in the next 7 days.",
    clickToFill: "Click a meeting to auto-fill the form",
    who: "Who are you meeting?",
    whoPlaceholder: "Name, title, company — what do you know about them?",
    context: "Meeting context",
    contextPlaceholder: "Background, history, nature of the relationship…",
    goal: "Your goal for the meeting",
    goalPlaceholder: "What do you want to achieve?",
    sensitive: "Anything sensitive or difficult?",
    sensitivePlaceholder: "Conflicts, bad news, sensitive topics…",
    notes: "Other notes",
    notesPlaceholder: "Anything else relevant…",
    generateBtn: "Generate briefing",
    generating: "Generating…",
    copyBtn: "Copy",
    copied: "Copied!",
    newBtn: "New briefing",
    lang: "DA",
    required: "Fill in 'Who' and 'Goal' to continue.",
  },
};

// ── Prompt builder ────────────────────────────────────────────────────────────
function buildPrompt(fields, lang) {
  if (lang === "da") {
    return `Du er en erfaren kommunikationsrådgiver. Generér en struktureret møde-briefing baseret på disse oplysninger:

Hvem: ${fields.who}
Kontekst: ${fields.context || "Ikke angivet"}
Mål: ${fields.goal}
Følsomt/svært: ${fields.sensitive || "Intet angivet"}
Andre noter: ${fields.notes || "Ingen"}

Skriv briefingen på dansk. Strukturér den med disse sektioner:

## Formål med mødet
[2-3 sætninger om hvad mødet handler om og hvad succes ser ud som]

## Talking points (3-5 konkrete)
[Konkrete punkter at tage op — ikke vage]

## Spørgsmål de måske stiller mig
[3-4 spørgsmål de sandsynligvis stiller, med korte svarforslag]

## Spørgsmål jeg bør stille
[3-4 spørgsmål der giver mig nyttig information eller driver mødet fremad]

## Undgå / vær opmærksom på
[2-3 konkrete faldgruber eller ting at navigere varsomt]

## Succes ser sådan ud
[1-2 sætninger: hvad er det konkrete resultat hvis mødet går godt?]

Vær direkte og konkret. Ingen tomme vendinger.`;
  } else {
    return `You are an experienced communications advisor. Generate a structured meeting briefing based on this information:

Who: ${fields.who}
Context: ${fields.context || "Not specified"}
Goal: ${fields.goal}
Sensitive/difficult: ${fields.sensitive || "Nothing noted"}
Other notes: ${fields.notes || "None"}

Write the briefing in English. Structure it with these sections:

## Meeting purpose
[2-3 sentences on what the meeting is about and what success looks like]

## Talking points (3-5 concrete)
[Concrete points to raise — not vague]

## Questions they may ask me
[3-4 questions they'll likely ask, with short suggested answers]

## Questions I should ask
[3-4 questions that give me useful information or move the meeting forward]

## Avoid / watch out for
[2-3 concrete pitfalls or things to navigate carefully]

## Success looks like
[1-2 sentences: what is the concrete outcome if the meeting goes well?]

Be direct and concrete. No empty phrases.`;
  }
}

// ── Main component ────────────────────────────────────────────────────────────
export default function MeetingPrep() {
  const [lang, setLang] = useState("da");
  const [account, setAccount] = useState(null);
  const [meetings, setMeetings] = useState([]);
  const [loadingMeetings, setLoadingMeetings] = useState(false);
  const [fields, setFields] = useState({
    who: "",
    context: "",
    goal: "",
    sensitive: "",
    notes: "",
  });
  const [briefing, setBriefing] = useState("");
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState("");

  const tx = t[lang];

  // ── MSAL init ──
  useEffect(() => {
    msalInstance.initialize().then(() => {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        setAccount(accounts[0]);
      }
    });
  }, []);

  // ── Fetch calendar when logged in ──
  useEffect(() => {
    if (account) fetchMeetings();
  }, [account]);

  async function login() {
    try {
      await msalInstance.initialize();
      const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
      setAccount(result.account);
    } catch (e) {
      setError("Login fejlede: " + e.message);
    }
  }

  async function logout() {
    await msalInstance.initialize();
    await msalInstance.logoutPopup();
    setAccount(null);
    setMeetings([]);
  }

  async function getToken() {
    await msalInstance.initialize();
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account,
    });
    return result.accessToken;
  }

  async function fetchMeetings() {
    setLoadingMeetings(true);
    try {
      const token = await getToken();
      const now = new Date().toISOString();
      const end = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
      const url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${now}&endDateTime=${end}&$orderby=start/dateTime&$top=20&$select=subject,start,end,attendees,bodyPreview,organizer`;
      const res = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data = await res.json();
      setMeetings(data.value || []);
    } catch (e) {
      setError("Kunne ikke hente møder: " + e.message);
    }
    setLoadingMeetings(false);
  }

  function selectMeeting(meeting) {
    const attendees = (meeting.attendees || [])
      .filter((a) => a.type !== "resource")
      .map((a) => a.emailAddress?.name || a.emailAddress?.address)
      .filter(Boolean)
      .join(", ");

    setFields((f) => ({
      ...f,
      who: attendees || f.who,
      context: meeting.bodyPreview
        ? meeting.bodyPreview.slice(0, 300)
        : f.context,
    }));
  }

  function formatTime(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    return d.toLocaleString(lang === "da" ? "da-DK" : "en-GB", {
      weekday: "short",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  async function generate() {
    if (!fields.who.trim() || !fields.goal.trim()) {
      setError(tx.required);
      return;
    }
    setError("");
    setLoading(true);
    setBriefing("");

    try {
      // Uses /api/briefing Vercel serverless endpoint (API key stays on server)
      // For local testing without the endpoint, swap to direct Anthropic call below
      const res = await fetch("/api/briefing", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt: buildPrompt(fields, lang) }),
      });

      if (!res.ok) throw new Error(`API error ${res.status}`);
      const data = await res.json();
      setBriefing(data.content);
    } catch (e) {
      setError("Fejl: " + e.message);
    }
    setLoading(false);
  }

  async function copy() {
    await navigator.clipboard.writeText(briefing);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  function reset() {
    setBriefing("");
    setFields({ who: "", context: "", goal: "", sensitive: "", notes: "" });
    setError("");
  }

  // ── Styles ──
  const card = {
    background: "#fff",
    borderRadius: 12,
    padding: "1.5rem",
    boxShadow: "0 1px 4px rgba(0,0,0,.08)",
    marginBottom: "1rem",
  };

  const inputStyle = {
    width: "100%",
    padding: "0.6rem 0.8rem",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    fontSize: 14,
    fontFamily: "inherit",
    resize: "vertical",
    boxSizing: "border-box",
  };

  const btnPrimary = {
    background: "#005eb8",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    padding: "0.65rem 1.4rem",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    marginRight: 8,
  };

  const btnSecondary = {
    background: "#f3f4f6",
    color: "#374151",
    border: "1px solid #d1d5db",
    borderRadius: 8,
    padding: "0.65rem 1.2rem",
    fontSize: 14,
    cursor: "pointer",
  };

  const label = {
    display: "block",
    fontSize: 13,
    fontWeight: 600,
    color: "#374151",
    marginBottom: 4,
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#f9fafb",
        fontFamily: "'Inter', system-ui, sans-serif",
        padding: "2rem 1rem",
      }}
    >
      <div style={{ maxWidth: 680, margin: "0 auto" }}>
        {/* Header */}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
            marginBottom: "1.5rem",
          }}
        >
          <div>
            <h1
              style={{
                fontSize: 22,
                fontWeight: 700,
                color: "#111827",
                margin: 0,
              }}
            >
              {tx.title}
            </h1>
            <p style={{ fontSize: 13, color: "#6b7280", marginTop: 4 }}>
              {tx.subtitle}
            </p>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <button
              onClick={() => setLang(lang === "da" ? "en" : "da")}
              style={btnSecondary}
            >
              {tx.lang}
            </button>
            {account ? (
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 12, color: "#6b7280" }}>
                  {tx.loggedInAs}
                </div>
                <div style={{ fontSize: 13, fontWeight: 600 }}>
                  {account.name}
                </div>
                <button
                  onClick={logout}
                  style={{ ...btnSecondary, fontSize: 12, padding: "0.3rem 0.8rem", marginTop: 4 }}
                >
                  {tx.logoutBtn}
                </button>
              </div>
            ) : (
              <button onClick={login} style={btnPrimary}>
                {tx.loginBtn}
              </button>
            )}
          </div>
        </div>

        {/* Outlook calendar */}
        {account && (
          <div style={card}>
            <h2
              style={{
                fontSize: 14,
                fontWeight: 600,
                color: "#374151",
                margin: "0 0 0.75rem",
              }}
            >
              {tx.upcomingMeetings}
            </h2>
            {loadingMeetings ? (
              <p style={{ fontSize: 13, color: "#9ca3af" }}>
                {tx.loadingMeetings}
              </p>
            ) : meetings.length === 0 ? (
              <p style={{ fontSize: 13, color: "#9ca3af" }}>{tx.noMeetings}</p>
            ) : (
              <>
                <p style={{ fontSize: 12, color: "#9ca3af", marginBottom: 8 }}>
                  {tx.clickToFill}
                </p>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {meetings.map((m, i) => (
                    <button
                      key={i}
                      onClick={() => selectMeeting(m)}
                      style={{
                        background: "#f9fafb",
                        border: "1px solid #e5e7eb",
                        borderRadius: 8,
                        padding: "0.6rem 0.9rem",
                        textAlign: "left",
                        cursor: "pointer",
                        fontSize: 13,
                      }}
                    >
                      <div style={{ fontWeight: 600, color: "#111827" }}>
                        {m.subject}
                      </div>
                      <div style={{ color: "#6b7280", fontSize: 12, marginTop: 2 }}>
                        {formatTime(m.start?.dateTime)}
                      </div>
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
        )}

        {/* Form */}
        {!briefing && (
          <div style={card}>
            {[
              { key: "who", rows: 2 },
              { key: "context", rows: 3 },
              { key: "goal", rows: 2 },
              { key: "sensitive", rows: 2 },
              { key: "notes", rows: 2 },
            ].map(({ key, rows }) => (
              <div key={key} style={{ marginBottom: "1rem" }}>
                <label style={label}>{tx[key]}</label>
                <textarea
                  rows={rows}
                  value={fields[key]}
                  onChange={(e) =>
                    setFields((f) => ({ ...f, [key]: e.target.value }))
                  }
                  placeholder={tx[`${key}Placeholder`]}
                  style={inputStyle}
                />
              </div>
            ))}

            {error && (
              <p style={{ color: "#dc2626", fontSize: 13, marginBottom: 12 }}>
                {error}
              </p>
            )}

            <button
              onClick={generate}
              disabled={loading}
              style={{ ...btnPrimary, opacity: loading ? 0.6 : 1 }}
            >
              {loading ? tx.generating : tx.generateBtn}
            </button>
          </div>
        )}

        {/* Output */}
        {briefing && (
          <div style={card}>
            <div
              style={{
                whiteSpace: "pre-wrap",
                fontSize: 14,
                lineHeight: 1.7,
                color: "#1f2937",
                marginBottom: "1.25rem",
              }}
            >
              {briefing}
            </div>
            <div>
              <button onClick={copy} style={btnPrimary}>
                {copied ? tx.copied : tx.copyBtn}
              </button>
              <button onClick={reset} style={btnSecondary}>
                {tx.newBtn}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

