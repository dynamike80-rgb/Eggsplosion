import { useState, useEffect, useMemo, useRef } from "react";
import {
  ShoppingBag, Users, Package, Plus, X, Check, Search,
  MessageCircle, Trash2, Pencil, ChevronRight, TrendingUp,
  MinusCircle, PlusCircle, Upload, Download, ArrowUp, ArrowDown,
  Route, CircleCheck, BarChart3, Gift, Heart, Star, Send
} from "lucide-react";
import * as XLSX from "xlsx";

/* ---------- design tokens: white & yellow, cheerful ---------- */
const T = {
  bg: "#FFFFFF",
  cream: "#FFF8E6",
  surface: "#FFFFFF",
  headerFrom: "#FFD65C",
  headerTo: "#F6B93D",
  yolk: "#F5A623",
  yolkDeep: "#DE8E10",
  yolkPale: "#FFE9B8",
  blush: "#F79A8A",
  ink: "#2B2118",
  inkSoft: "#8C7F68",
  line: "#F0E4C8",
  danger: "#E2574C",
};

const FONTS_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@500;600;700;800&family=Nunito:wght@400;500;600;700&display=swap');`;

/* ---------- opslag via Netlify Function (proxy naar Supabase) ---------- */
/* De browser praat alleen met /.netlify/functions/data — hetzelfde domein als de app,
   dus geen cross-origin verzoek meer dat door sommige netwerken/browsers geblokkeerd wordt. */
const DATA_ENDPOINT = "/.netlify/functions/data";

async function loadKey(key, fallback) {
  try {
    const res = await fetch(`${DATA_ENDPOINT}?key=${encodeURIComponent(key)}`);
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      return { value: fallback, failed: true, errorDetail: `${res.status} — ${body.slice(0, 200)}` };
    }
    const data = await res.json();
    if (data.value === null || data.value === undefined) return { value: fallback, failed: false };
    return { value: data.value, failed: false };
  } catch (err) {
    return { value: fallback, failed: true, errorDetail: String(err) };
  }
}
async function saveKey(key, value) {
  try {
    const res = await fetch(DATA_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ key, value }),
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      return { ok: false, errorDetail: `${res.status} — ${body.slice(0, 200)}` };
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, errorDetail: String(err) };
  }
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

function formatEuro(n) {
  return new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR" }).format(n || 0);
}
function formatDateTime(ts) {
  return new Date(ts).toLocaleString("nl-NL", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
}

function isOnVacation(c) {
  if (!c.vacationStart || !c.vacationEnd) return false;
  const today = new Date().toISOString().slice(0, 10);
  return today >= c.vacationStart && today <= c.vacationEnd;
}

/* ---------- seizoenscorrectie voor de inkoopvoorspelling ---------- */
/* Let op: gebaseerd op maar ±1 jaar aan historische data — vooral bruikbaar
   als indicatie, niet als harde waarheid. Groeit hopelijk betrouwbaarder
   naarmate er meer seizoenen aan data bijkomen. */
function getEasterSunday(year) {
  // Meeus/Jones/Butcher-algoritme
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

function getSeasonalFactor(date) {
  const easter = getEasterSunday(date.getFullYear());
  const dayMs = 24 * 3600 * 1000;
  const daysFromEaster = (date.getTime() - easter.getTime()) / dayMs;
  if (daysFromEaster >= -7 && daysFromEaster <= 7) {
    return { factor: 1.25, label: "Paasperiode" };
  }
  const month = date.getMonth() + 1; // 1-12
  if (month >= 11 || month <= 4) {
    return { factor: 1.15, label: "winter" };
  }
  if (month >= 6 && month <= 8) {
    return { factor: 0.85, label: "zomer" };
  }
  return { factor: 1, label: null };
}

/* ---------- WhatsApp templates ---------- */
function greet(naam) {
  return naam ? `Hoi familie ${naam}` : "Hoi";
}

function vakantieBericht(naam, start, end) {
  const fmt = (d) => new Date(d).toLocaleDateString("nl-NL", { day: "numeric", month: "long" });
  let periode = "";
  if (start && end) periode = ` van ${fmt(start)} tot ${fmt(end)}`;
  else if (start) periode = ` vanaf ${fmt(start)}`;
  else if (end) periode = ` tot ${fmt(end)}`;
  return `${greet(naam)}, ik ben${periode} afwezig, dus dan kom ik niet langs met eieren. Daarna weer gewoon!`;
}

function weerBericht(naam, opmerking) {
  return `${greet(naam)}, door het weer: ${opmerking || "ik kom vandaag op een ander tijdstip of een andere dag langs"}. Tot dan!`;
}

const WA_TEMPLATES = {
  herinnering: (naam) =>
    `${greet(naam)}, ik kom morgen (vrijdag) weer langs met verse eieren! Laat het weten door op dit bericht te reageren als er iets bijzonders is, bijvoorbeeld niet thuis of een grote bestelling.`,
  bedankjeVerkoop: (naam) =>
    `${greet(naam)}, ik ben ei-genlijk best blij met jou als klant. Bedankt voor de bestelling en tot volgende keer! 🥚`,
  bedankjeFooi: (naam) =>
    `${greet(naam)}, ik wilde je even bedanken voor je fooi de laatste tijd — heel lief van je!`,
};

const WA_TEMPLATE_LABELS = {
  herinnering: "Herinnering (morgen langs)",
  bedankjeVerkoop: "Bedankje na verkoop",
  bedankjeFooi: "Bedankje voor fooi",
};

function waLink(phone, text) {
  const digits = (phone || "").replace(/[^\d+]/g, "").replace(/^0/, "31");
  return `https://wa.me/${digits}?text=${encodeURIComponent(text)}`;
}

/* ---------- klant status ---------- */
const STATUS_META = {
  klant: { label: "Klant", bg: T.yolkPale, color: T.yolkDeep },
  kans: { label: "Kans", bg: "#EFEAE0", color: "#7A7159" },
  geen_interesse: { label: "Geen interesse", bg: "#F2E2DD", color: "#B05C48" },
};

/* ---------- excel import ---------- */
const HEADER_ALIASES = {
  name: ["naam", "name"],
  postalCode: ["postcode", "postal code", "postalcode", "zip"],
  houseNumber: ["huisnummer", "huisnr", "nummer", "housenumber"],
  addition: ["toevoeging", "toev", "addition"],
  street: ["straat", "street"],
  city: ["plaats", "woonplaats", "stad", "city"],
  phone: ["telefoon", "telefoonnummer", "whatsapp", "phone", "mobiel"],
  status: ["status", "type"],
};

function normalizeHeader(h) {
  return String(h || "").toLowerCase().trim();
}

function mapRow(rawRow) {
  const lowerKeys = Object.keys(rawRow).reduce((acc, k) => {
    acc[normalizeHeader(k)] = rawRow[k];
    return acc;
  }, {});
  const c = { id: uid(), name: "", postalCode: "", houseNumber: "", addition: "", street: "", city: "", phone: "", status: "klant" };
  for (const field of Object.keys(HEADER_ALIASES)) {
    for (const alias of HEADER_ALIASES[field]) {
      if (lowerKeys[alias] !== undefined && lowerKeys[alias] !== "") {
        c[field] = String(lowerKeys[alias]).trim();
        break;
      }
    }
  }
  if (c.status) {
    const s = c.status.toLowerCase();
    if (s.includes("geen") || s.includes("interesse")) c.status = "geen_interesse";
    else if (s.includes("kans") || s.includes("twijfel")) c.status = "kans";
    else c.status = "klant";
  } else {
    c.status = "klant";
  }
  return c;
}

function parseImportFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const mapped = rows
          .map(mapRow)
          .filter((c) => c.street || c.name || c.postalCode);
        resolve(mapped);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseRhythmFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const mapped = rows.map((raw) => {
          const lowerKeys = Object.keys(raw).reduce((acc, k) => {
            acc[normalizeHeader(k)] = raw[k];
            return acc;
          }, {});
          return {
            street: String(lowerKeys["straat"] || "").trim().toLowerCase(),
            houseNumber: String(lowerKeys["huisnummer"] || "").trim(),
            addition: String(lowerKeys["toevoeging"] || "").trim().toLowerCase(),
            rhythm: String(lowerKeys["ritme"] || "").trim().toLowerCase(),
            avgEggs: parseFloat(lowerKeys["gemaantal"]) || 0,
            histLastPurchase: String(lowerKeys["laatsteaankoopooit"] || "").trim(),
          };
        }).filter((r) => r.street);
        resolve(mapped);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function downloadExampleFile() {
  const rows = [
    { Naam: "Fam. Jansen", Postcode: "3441 AB", Huisnummer: "12", Toevoeging: "", Straat: "Dorpsstraat", Plaats: "Woerden", Telefoon: "0612345678", Status: "Klant" },
    { Naam: "Fam. de Vries", Postcode: "3441 CD", Huisnummer: "14", Toevoeging: "a", Straat: "Dorpsstraat", Plaats: "Woerden", Telefoon: "", Status: "Kans" },
  ];
  const sheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Klanten");
  XLSX.writeFile(workbook, "egg-sellerate-klanten-voorbeeld.xlsx");
}

/* ---------- mascot: happy exploding egg (uploaded illustration) ---------- */
const MASCOT_SRC = "data:image/webp;base64,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";

function EggMascot({ size = 48, burst = false }) {
  return (
    <img
      src={MASCOT_SRC}
      alt="Egg-sellerate mascotte"
      draggable={false}
      width={size}
      height={Math.round(size * (276 / 480))}
      style={{
        display: "block",
        objectFit: "contain",
        animation: burst ? "mascotPop 0.9s ease-out" : undefined,
      }}
    />
  );
}

/* ---------- main app ---------- */
export default function EggSellerateApp() {
  const [tab, setTab] = useState("verkoop");
  const [loading, setLoading] = useState(true);
  const [customers, setCustomers] = useState([]);
  const [sales, setSales] = useState([]);
  const [purchases, setPurchases] = useState([]);
  const [extras, setExtras] = useState([]);
  const [settings, setSettings] = useState({ pricePerEgg: 0.35 });
  const [saveError, setSaveError] = useState("");
  const [loadError, setLoadError] = useState(false);

  useEffect(() => {
    (async () => {
      const [c, s, p, ex, set] = await Promise.all([
        loadKey("customers", []),
        loadKey("sales", []),
        loadKey("purchases", []),
        loadKey("extras", []),
        loadKey("settings", { pricePerEgg: 0.35 }),
      ]);
      const failedOne = [c, s, p, ex, set].find((r) => r.failed);
      setLoadError(failedOne ? (failedOne.errorDetail || true) : false);
      const customersData = c.value;
      const needsMigration = customersData.some((x) => x.routeOrder === undefined);
      const migrated = needsMigration ? customersData.map((x, i) => ({ ...x, routeOrder: x.routeOrder ?? i })) : customersData;
      setCustomers(migrated);
      setSales(s.value);
      setPurchases(p.value);
      setExtras(ex.value);
      setSettings(set.value);
      setLoading(false);
      if (needsMigration) saveKey("customers", migrated);
    })();
  }, []);

  async function persist(key, value, setter) {
    setter(value);
    const result = await saveKey(key, value);
    if (!result.ok) setSaveError(`Opslaan is niet gelukt: ${result.errorDetail || "onbekende fout"}`);
    else setSaveError("");
  }

  const stock = useMemo(() => {
    const bought = purchases.reduce((sum, p) => sum + p.eggCount, 0);
    const sold = sales.reduce((sum, s) => sum + s.eggCount, 0);
    return bought - sold;
  }, [purchases, sales]);

  const revenueThisWeek = useMemo(() => {
    const weekAgo = Date.now() - 7 * 24 * 3600 * 1000;
    return sales.filter((s) => s.ts >= weekAgo).reduce((sum, s) => sum + s.amount, 0);
  }, [sales]);

  if (loading) {
    return (
      <Shell>
        <div style={{ padding: 50, textAlign: "center", color: T.inkSoft, fontFamily: "'Nunito',sans-serif" }}>
          <EggMascot size={40} />
          <div style={{ marginTop: 10 }}>Bezig met laden…</div>
        </div>
      </Shell>
    );
  }

  return (
    <Shell>
      <Header stock={stock} revenue={revenueThisWeek} />
      {loadError && (
        <div style={styles.errorBanner}>
          Kon opgeslagen gegevens niet ophalen. Dit lijkt dan leeg, maar is mogelijk een laadfout — ververs de pagina voor je iets nieuws invoert.
          {typeof loadError === "string" && <div style={{ marginTop: 4, fontFamily: "monospace", fontSize: 11 }}>{loadError}</div>}
        </div>
      )}
      {saveError && <div style={styles.errorBanner}>{saveError}</div>}
      <div style={styles.content}>
        {tab === "verkoop" && (
          <VerkoopTab
            customers={customers}
            sales={sales}
            settings={settings}
            onLogSale={(sale) => persist("sales", [...sales, sale], setSales)}
            onLogExtra={(extra) => persist("extras", [...extras, extra], setExtras)}
          />
        )}
        {tab === "klanten" && (
          <KlantenTab
            customers={customers}
            onAdd={(c) => persist("customers", [...customers, { ...c, routeOrder: customers.length }], setCustomers)}
            onAddMany={(list) =>
              persist(
                "customers",
                [...customers, ...list.map((c, i) => ({ ...c, routeOrder: customers.length + i }))],
                setCustomers
              )
            }
            onUpdate={(c) =>
              persist("customers", customers.map((x) => (x.id === c.id ? c : x)), setCustomers)
            }
            onDelete={(id) => persist("customers", customers.filter((x) => x.id !== id), setCustomers)}
            onReorder={(list) => persist("customers", list, setCustomers)}
            onImportRhythm={(list) => persist("customers", list, setCustomers)}
          />
        )}
        {tab === "voorraad" && (
          <VoorraadTab
            stock={stock}
            purchases={purchases}
            sales={sales}
            extras={extras}
            customers={customers}
            settings={settings}
            onLogPurchase={(p) => persist("purchases", [...purchases, p], setPurchases)}
            onUpdateSettings={(s) => persist("settings", s, setSettings)}
            onUpdateSale={(s) => persist("sales", sales.map((x) => (x.id === s.id ? s : x)), setSales)}
            onDeleteSale={(id) => persist("sales", sales.filter((x) => x.id !== id), setSales)}
            onUpdatePurchase={(p) => persist("purchases", purchases.map((x) => (x.id === p.id ? p : x)), setPurchases)}
            onDeletePurchase={(id) => persist("purchases", purchases.filter((x) => x.id !== id), setPurchases)}
            onUpdateExtra={(e) => persist("extras", extras.map((x) => (x.id === e.id ? e : x)), setExtras)}
            onDeleteExtra={(id) => persist("extras", extras.filter((x) => x.id !== id), setExtras)}
            onImportAll={async (data) => {
              if (data.customers) await persist("customers", data.customers, setCustomers);
              if (data.sales) await persist("sales", data.sales, setSales);
              if (data.purchases) await persist("purchases", data.purchases, setPurchases);
              if (data.extras) await persist("extras", data.extras, setExtras);
              if (data.settings) await persist("settings", data.settings, setSettings);
            }}
          />
        )}
        {tab === "berichten" && (
          <BerichtenTab customers={customers} settings={settings} onUpdateSettings={(s) => persist("settings", s, setSettings)} />
        )}
        {tab === "stats" && (
          <StatsTab customers={customers} sales={sales} purchases={purchases} extras={extras} settings={settings} />
        )}
      </div>
      <TabBar tab={tab} setTab={setTab} />
    </Shell>
  );
}

/* ---------- shell / layout ---------- */
function Shell({ children }) {
  return (
    <div style={styles.appOuter}>
      <style>{`
        ${FONTS_IMPORT}
        * { box-sizing: border-box; }
        button { font-family: inherit; cursor: pointer; }
        input, select { font-family: inherit; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-thumb { background: ${T.line}; border-radius: 3px; }
        .eggsellerate-phone {
          height: 100vh;
          height: 100dvh;
        }
        @media (max-width: 480px) {
          .eggsellerate-phone {
            border-radius: 0;
            border: none;
            box-shadow: none;
            max-width: 100%;
          }
        }
        @keyframes sparkleBurst {
          0% { opacity: 0; transform: scale(0.4); }
          40% { opacity: 1; transform: scale(1.3); }
          100% { opacity: 0; transform: scale(1.8); }
        }
        @keyframes capFly {
          0% { transform: translate(0,0) rotate(0deg); opacity: 1; }
          100% { transform: translate(-14px,-20px) rotate(-50deg); opacity: 0; }
        }
        @keyframes capFly2 {
          0% { transform: translate(0,0) rotate(0deg); opacity: 1; }
          100% { transform: translate(14px,-20px) rotate(50deg); opacity: 0; }
        }
        @keyframes yolkPop {
          0% { transform: scale(1); }
          40% { transform: scale(1.25); }
          100% { transform: scale(1); }
        }
        @keyframes popIn {
          0% { transform: scale(0.7); opacity: 0; }
          100% { transform: scale(1); opacity: 1; }
        }
        @keyframes mascotPop {
          0% { transform: scale(0.6) rotate(-6deg); opacity: 0.5; }
          55% { transform: scale(1.15) rotate(3deg); opacity: 1; }
          100% { transform: scale(1) rotate(0deg); opacity: 1; }
        }
      `}</style>
      <div className="eggsellerate-phone" style={styles.phone}>{children}</div>
    </div>
  );
}

function Header({ stock, revenue }) {
  return (
    <div style={styles.header}>
      <div style={styles.headerTop}>
        <div>
          <div style={styles.headerTitle}>Egg-sellerate</div>
          <div style={styles.headerSlogan}>slimme app voor Egg-sperts</div>
        </div>
        <EggMascot size={50} />
      </div>
      <div style={styles.headerStats}>
        <div style={styles.headerStatCard}>
          <div style={styles.headerStatNum}>{stock}</div>
          <div style={styles.headerStatLabel}>eieren op voorraad</div>
        </div>
        <div style={styles.headerStatCard}>
          <div style={styles.headerStatNum}>{formatEuro(revenue)}</div>
          <div style={styles.headerStatLabel}>omzet deze week</div>
        </div>
      </div>
    </div>
  );
}

function TabBar({ tab, setTab }) {
  const items = [
    { id: "verkoop", label: "Verkoop", icon: ShoppingBag },
    { id: "klanten", label: "Klanten", icon: Users },
    { id: "voorraad", label: "Voorraad", icon: Package },
    { id: "berichten", label: "Berichten", icon: Send },
    { id: "stats", label: "Stats", icon: BarChart3 },
  ];
  return (
    <div style={styles.tabBar}>
      {items.map(({ id, label, icon: Icon }) => {
        const active = tab === id;
        return (
          <button
            key={id}
            onClick={() => setTab(id)}
            style={{ ...styles.tabBtn, color: active ? T.yolkDeep : T.inkSoft }}
          >
            <Icon size={20} strokeWidth={active ? 2.4 : 1.8} />
            <span style={{ fontWeight: active ? 700 : 500 }}>{label}</span>
          </button>
        );
      })}
    </div>
  );
}

/* ---------- Verkoop tab ---------- */
function VerkoopTab({ customers, sales, settings, onLogSale, onLogExtra }) {
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("alle");
  const [activeCustomer, setActiveCustomer] = useState(null);
  const [showExtra, setShowExtra] = useState(false);

  const eligible = useMemo(() => customers.filter((c) => c.status !== "geen_interesse"), [customers]);

  const statusCounts = useMemo(() => ({
    alle: eligible.length,
    klant: eligible.filter((c) => c.status !== "kans").length,
    kans: eligible.filter((c) => c.status === "kans").length,
  }), [eligible]);

  const byStatus = useMemo(() => {
    if (statusFilter === "alle") return eligible;
    if (statusFilter === "kans") return eligible.filter((c) => c.status === "kans");
    return eligible.filter((c) => c.status !== "kans");
  }, [eligible, statusFilter]);

  const routeSorted = useMemo(
    () => [...byStatus].sort((a, b) => (a.routeOrder ?? 9999) - (b.routeOrder ?? 9999)),
    [byStatus]
  );

  const soldTodayIds = useMemo(() => {
    const start = new Date();
    start.setHours(0, 0, 0, 0);
    return new Set(sales.filter((s) => s.ts >= start.getTime() && s.customerId).map((s) => s.customerId));
  }, [sales]);

  const lastSaleByCustomer = useMemo(() => {
    const map = {};
    sales.forEach((s) => {
      if (!s.customerId) return;
      if (!map[s.customerId] || s.ts > map[s.customerId].ts) map[s.customerId] = s;
    });
    return map;
  }, [sales]);

  const nextCustomer = useMemo(
    () =>
      [...eligible]
        .sort((a, b) => (a.routeOrder ?? 9999) - (b.routeOrder ?? 9999))
        .find((c) => !soldTodayIds.has(c.id) && !isOnVacation(c)),
    [eligible, soldTodayIds]
  );

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    const base = query ? byStatus : routeSorted;
    if (!q) return base;
    return base.filter((c) =>
      `${c.name} ${c.street} ${c.houseNumber} ${c.city}`.toLowerCase().includes(q)
    );
  }, [byStatus, routeSorted, query]);

  return (
    <div>
      {nextCustomer && !query && (
        <div style={styles.nextCard}>
          <div style={styles.nextLabel}>
            <Route size={13} /> Volgende in de looproute
          </div>
          <div style={styles.nextName}>{nextCustomer.name || `${nextCustomer.street} ${nextCustomer.houseNumber}${nextCustomer.addition ? `-${nextCustomer.addition}` : ""}`}</div>
          <div style={styles.nextAddress}>
            {nextCustomer.street} {nextCustomer.houseNumber}{nextCustomer.addition ? `-${nextCustomer.addition}` : ""}, {nextCustomer.postalCode} {nextCustomer.city}
          </div>
          <button style={styles.nextBtn} onClick={() => setActiveCustomer(nextCustomer)}>
            Verkocht <ChevronRight size={16} />
          </button>
        </div>
      )}

      <div style={styles.searchRow}>
        <Search size={17} color={T.inkSoft} />
        <input
          placeholder="Zoek klant op naam of straat…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={styles.searchInput}
        />
      </div>

      <div style={styles.filterRow}>
        {[
          { id: "alle", label: `Alle (${statusCounts.alle})` },
          { id: "klant", label: `Klant (${statusCounts.klant})` },
          { id: "kans", label: `Kans (${statusCounts.kans})` },
        ].map((f) => (
          <button
            key={f.id}
            onClick={() => setStatusFilter(f.id)}
            style={{
              ...styles.filterChip,
              background: statusFilter === f.id ? T.ink : "transparent",
              color: statusFilter === f.id ? "#fff" : T.inkSoft,
              borderColor: statusFilter === f.id ? T.ink : T.line,
            }}
          >
            {f.label}
          </button>
        ))}
      </div>

      <button style={styles.adhocBtn} onClick={() => setActiveCustomer("adhoc")}>
        <Plus size={16} /> Snelle verkoop (geen vaste klant)
      </button>

      <button style={styles.extraBtn} onClick={() => setShowExtra(true)}>
        <Star size={16} /> Extra verkoop (bv. actie of stroopwafels)
      </button>

      {filtered.length === 0 && (
        <EmptyState
          title={customers.length === 0 ? "Nog geen klanten" : "Geen klant gevonden"}
          text={
            customers.length === 0
              ? "Voeg je eerste klant toe bij het tabblad Klanten, of log hieronder een snelle verkoop."
              : "Probeer een andere zoekterm, of voeg deze klant toe."
          }
        />
      )}

      <div style={styles.list}>
        {filtered.map((c) => {
          const visited = soldTodayIds.has(c.id);
          const onVacation = isOnVacation(c);
          const lastSale = lastSaleByCustomer[c.id];
          const huisnr = `${c.houseNumber}${c.addition ? `-${c.addition}` : ""}`;
          return (
            <div key={c.id} style={{ ...styles.row, opacity: visited || onVacation ? 0.6 : 1 }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <div style={styles.rowTitle}>{c.name || `${c.street} ${huisnr}`}</div>
                  {c.status === "kans" ? (
                    <span style={{ ...styles.statusBadge, background: STATUS_META.kans.bg, color: STATUS_META.kans.color }}>
                      Kans
                    </span>
                  ) : (
                    <span style={{ ...styles.statusBadge, background: STATUS_META.klant.bg, color: STATUS_META.klant.color }}>
                      Klant
                    </span>
                  )}
                  {visited && (
                    <span style={styles.visitedBadge}>
                      <CircleCheck size={11} /> Vandaag geweest
                    </span>
                  )}
                  {onVacation && (
                    <span style={{ ...styles.statusBadge, background: "#E7E2F5", color: "#6B5FA8" }}>
                      Op vakantie tot {new Date(c.vacationEnd).toLocaleDateString("nl-NL", { day: "numeric", month: "short" })}
                    </span>
                  )}
                </div>
                <div style={styles.rowSub}>
                  {c.street} {huisnr}, {c.postalCode} {c.city}
                </div>
                {lastSale && (
                  <div style={styles.rowSub}>
                    Laatst: {formatDateTime(lastSale.ts)} · {lastSale.eggCount} st.
                  </div>
                )}
                {!lastSale && c.rhythm && (
                  <div style={{ ...styles.rowSub, fontStyle: "italic" }}>
                    Meestal: {c.rhythm}{c.avgEggs ? `, ~${c.avgEggs} st.` : ""} (historie)
                  </div>
                )}
              </div>
              {c.phone && (
                <a
                  href={waLink(c.phone, WA_TEMPLATES.herinnering(c.name))}
                  target="_blank"
                  rel="noreferrer"
                  style={styles.iconBtn}
                  title="WhatsApp"
                  onClick={(e) => e.stopPropagation()}
                >
                  <MessageCircle size={17} color={T.yolkDeep} />
                </a>
              )}
              <button style={styles.saleBtn} onClick={() => setActiveCustomer(c)}>
                Verkocht <ChevronRight size={15} />
              </button>
            </div>
          );
        })}
      </div>

      {activeCustomer && (
        <SaleModal
          customer={activeCustomer === "adhoc" ? null : activeCustomer}
          settings={settings}
          onClose={() => setActiveCustomer(null)}
          onLogSale={(sale) => onLogSale(sale)}
        />
      )}

      {showExtra && (
        <ExtraModal
          onClose={() => setShowExtra(false)}
          onConfirm={(extra) => {
            onLogExtra(extra);
            setShowExtra(false);
          }}
        />
      )}
    </div>
  );
}

function SaleModal({ customer, settings, onClose, onLogSale }) {
  const [count, setCount] = useState(6);
  const [amount, setAmount] = useState((6 * settings.pricePerEgg).toFixed(2));
  const [amountTouched, setAmountTouched] = useState(false);
  const [showTip, setShowTip] = useState(false);
  const [tip, setTip] = useState("");
  const [celebrating, setCelebrating] = useState(false);
  const [showThanks, setShowThanks] = useState(false);

  useEffect(() => {
    if (!amountTouched) setAmount((count * settings.pricePerEgg).toFixed(2));
  }, [count]); // eslint-disable-line

  const presets = [6, 10, 12, 18, 20, 30];
  const tipPresets = [0.5, 1, 2, 5];

  function handleConfirm() {
    const sale = {
      id: uid(),
      customerId: customer ? customer.id : null,
      customerLabel: customer ? customer.name || `${customer.street} ${customer.houseNumber}` : "Losse verkoop",
      eggCount: count,
      amount: parseFloat(amount) || 0,
      tip: parseFloat(tip) || 0,
      ts: Date.now(),
    };
    setCelebrating(true);
    setTimeout(() => {
      onLogSale(sale);
      if (customer?.phone) setShowThanks(true);
      else onClose();
    }, 750);
  }

  if (showThanks) {
    return (
      <ModalOverlay onClose={onClose}>
        <div style={styles.celebrateBox}>
          <EggMascot size={70} />
          <div style={{ fontWeight: 700, fontSize: 15, marginTop: 8, marginBottom: 14 }}>
            Verkoop gelogd! Nog een bedankje sturen?
          </div>
          <a
            href={waLink(customer.phone, WA_TEMPLATES.bedankjeVerkoop(customer.name))}
            target="_blank"
            rel="noreferrer"
            style={styles.confirmBtn}
            onClick={onClose}
          >
            <MessageCircle size={16} /> Bedankje sturen
          </a>
          <button style={styles.cancelLink} onClick={onClose}>Nee, klaar</button>
        </div>
      </ModalOverlay>
    );
  }

  if (celebrating) {
    return (
      <ModalOverlay onClose={() => {}}>
        <div style={styles.celebrateBox}>
          <EggMascot size={90} burst />
          <div style={styles.celebrateText}>Egg-sellerated!</div>
          <div style={{ color: T.inkSoft, fontSize: 13.5 }}>{count} eieren verkocht</div>
        </div>
      </ModalOverlay>
    );
  }

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>
        {customer ? customer.name || `${customer.street} ${customer.houseNumber}` : "Snelle verkoop"}
      </div>
      <div style={styles.modalLabel}>Aantal eieren</div>
      <div style={styles.stepperRow}>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => Math.max(1, n - 1))}>
          <MinusCircle size={22} color={T.ink} />
        </button>
        <div style={styles.stepperNum}>{count}</div>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => n + 1)}>
          <PlusCircle size={22} color={T.ink} />
        </button>
      </div>
      <div style={styles.presetRow}>
        {presets.map((p) => (
          <button
            key={p}
            onClick={() => setCount(p)}
            style={{
              ...styles.presetBtn,
              background: count === p ? T.yolk : "transparent",
              color: count === p ? "#fff" : T.ink,
              borderColor: count === p ? T.yolk : T.line,
            }}
          >
            {p}
          </button>
        ))}
      </div>

      <div style={styles.modalLabel}>Bedrag</div>
      <div style={styles.amountRow}>
        <span style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 20, color: T.inkSoft }}>€</span>
        <input
          type="number"
          step="0.01"
          value={amount}
          onChange={(e) => {
            setAmountTouched(true);
            setAmount(e.target.value);
          }}
          style={styles.amountInput}
        />
      </div>

      {!showTip ? (
        <button style={styles.tipToggle} onClick={() => setShowTip(true)}>
          <Gift size={14} /> Fooi ontvangen? Voeg toe
        </button>
      ) : (
        <div style={{ marginBottom: 10 }}>
          <div style={styles.modalLabel}>Fooi (optioneel)</div>
          <div style={styles.amountRow}>
            <span style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 20, color: T.inkSoft }}>€</span>
            <input
              type="number"
              step="0.01"
              value={tip}
              onChange={(e) => setTip(e.target.value)}
              placeholder="0,00"
              style={styles.amountInput}
            />
          </div>
          <div style={{ ...styles.presetRow, marginBottom: 6 }}>
            {tipPresets.map((p) => (
              <button
                key={p}
                onClick={() => setTip(String(p))}
                style={{
                  ...styles.presetBtn,
                  background: parseFloat(tip) === p ? T.yolkDeep : "transparent",
                  color: parseFloat(tip) === p ? "#fff" : T.ink,
                  borderColor: parseFloat(tip) === p ? T.yolkDeep : T.line,
                }}
              >
                +€{p}
              </button>
            ))}
          </div>
        </div>
      )}

      <button style={styles.confirmBtn} onClick={handleConfirm}>
        <Check size={18} /> Verkoop loggen
      </button>
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

function ExtraModal({ initial, onClose, onConfirm, onDelete }) {
  const [naam, setNaam] = useState(initial?.naam || "");
  const [aantal, setAantal] = useState(initial?.aantal ?? 1);
  const [bedrag, setBedrag] = useState(initial ? initial.bedrag.toFixed(2) : "");

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>{initial ? "Actie bewerken" : "Extra verkoop"}</div>
      <div style={styles.formNote}>
        Voor eenmalige of periodieke extra's naast de eierverkoop, zoals een stroopwafelactie.
      </div>
      <FormField label="Wat heb je verkocht?" value={naam} onChange={(e) => setNaam(e.target.value)} placeholder="bijv. Stroopwafels" autoFocus />
      <div style={styles.modalLabel}>Aantal</div>
      <div style={styles.stepperRow}>
        <button style={styles.stepperBtn} onClick={() => setAantal((n) => Math.max(1, n - 1))}>
          <MinusCircle size={22} color={T.ink} />
        </button>
        <div style={styles.stepperNum}>{aantal}</div>
        <button style={styles.stepperBtn} onClick={() => setAantal((n) => n + 1)}>
          <PlusCircle size={22} color={T.ink} />
        </button>
      </div>
      <div style={styles.modalLabel}>Bedrag</div>
      <div style={styles.amountRow}>
        <span style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 20, color: T.inkSoft }}>€</span>
        <input type="number" step="0.01" value={bedrag} onChange={(e) => setBedrag(e.target.value)} placeholder="0,00" style={styles.amountInput} />
      </div>
      <button
        style={styles.confirmBtn}
        onClick={() => {
          if (!naam.trim()) return;
          onConfirm({
            id: initial ? initial.id : uid(),
            naam: naam.trim(),
            aantal,
            bedrag: parseFloat(bedrag) || 0,
            ts: initial ? initial.ts : Date.now(),
          });
        }}
      >
        <Check size={18} /> {initial ? "Wijzigingen opslaan" : "Loggen"}
      </button>
      {initial && onDelete && <DeleteConfirmButton onDelete={() => onDelete(initial.id)} label="actie" />}
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

/* ---------- Klanten tab ---------- */
function KlantenTab({ customers, onAdd, onAddMany, onUpdate, onDelete, onReorder, onImportRhythm }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [waFor, setWaFor] = useState(null);
  const [filter, setFilter] = useState("alle");
  const [importMsg, setImportMsg] = useState("");
  const [showRoute, setShowRoute] = useState(false);
  const fileInputRef = useRef(null);
  const rhythmInputRef = useRef(null);

  const counts = useMemo(() => ({
    alle: customers.length,
    klant: customers.filter((c) => c.status !== "kans" && c.status !== "geen_interesse").length,
    kans: customers.filter((c) => c.status === "kans").length,
    geen_interesse: customers.filter((c) => c.status === "geen_interesse").length,
  }), [customers]);

  const filtered = useMemo(() => {
    if (filter === "alle") return customers;
    return customers.filter((c) => (c.status || "klant") === filter);
  }, [customers, filter]);

  async function handleFileChange(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const imported = await parseImportFile(file);
      if (imported.length === 0) {
        setImportMsg("Geen bruikbare rijen gevonden in dit bestand.");
        return;
      }
      const existingKeys = new Set(
        customers.map((c) => `${c.postalCode}|${c.houseNumber}|${c.street}`.toLowerCase())
      );
      const fresh = imported.filter(
        (c) => !existingKeys.has(`${c.postalCode}|${c.houseNumber}|${c.street}`.toLowerCase())
      );
      const skipped = imported.length - fresh.length;
      onAddMany(fresh);
      setImportMsg(
        `${fresh.length} klant${fresh.length === 1 ? "" : "en"} geïmporteerd` +
        (skipped > 0 ? `, ${skipped} overgeslagen (leken al te bestaan).` : ".")
      );
    } catch {
      setImportMsg("Kon dit bestand niet lezen. Gebruik het voorbeeldbestand als leidraad.");
    }
  }

  async function handleRhythmFileChange(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    try {
      const rows = await parseRhythmFile(file);
      const keyOf = (street, nr, add) => `${street}|${nr}|${add}`.toLowerCase();
      const byKey = {};
      rows.forEach((r) => { byKey[keyOf(r.street, r.houseNumber, r.addition)] = r; });
      let matched = 0;
      const updated = customers.map((c) => {
        const key = keyOf(c.street, c.houseNumber, c.addition || "");
        const hit = byKey[key];
        if (!hit) return c;
        matched++;
        return { ...c, rhythm: hit.rhythm, avgEggs: hit.avgEggs, histLastPurchase: hit.histLastPurchase };
      });
      onImportRhythm(updated);
      setImportMsg(
        `Koopritme gekoppeld aan ${matched} van de ${rows.length} adressen uit het bestand.` +
        (rows.length - matched > 0 ? " De rest kon niet gematcht worden op straat/huisnummer." : "")
      );
    } catch {
      setImportMsg("Kon dit koopritme-bestand niet lezen.");
    }
  }

  return (
    <div>
      <button style={styles.primaryBtn} onClick={() => { setEditing(null); setShowForm(true); }}>
        <Plus size={17} /> Nieuwe klant
      </button>

      <div style={styles.importRow}>
        <button style={styles.secondaryBtn} onClick={() => fileInputRef.current?.click()}>
          <Upload size={15} /> Importeer Excel
        </button>
        <button style={styles.secondaryBtnGhost} onClick={downloadExampleFile}>
          <Download size={15} /> Voorbeeldbestand
        </button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          onChange={handleFileChange}
          style={{ display: "none" }}
        />
      </div>
      {importMsg && <div style={styles.importMsg}>{importMsg}</div>}

      <div style={styles.importRow}>
        <button style={styles.secondaryBtnGhost} onClick={() => rhythmInputRef.current?.click()}>
          <Upload size={15} /> Importeer koopritme
        </button>
        <input
          ref={rhythmInputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          onChange={handleRhythmFileChange}
          style={{ display: "none" }}
        />
      </div>

      <button style={styles.secondaryBtnGhost2} onClick={() => setShowRoute(true)}>
        <Route size={15} /> Looproute aanpassen
      </button>

      <div style={styles.filterRow}>
        {[
          { id: "alle", label: `Alle (${counts.alle})` },
          { id: "klant", label: `Klanten (${counts.klant})` },
          { id: "kans", label: `Kansen (${counts.kans})` },
          { id: "geen_interesse", label: `Geen interesse (${counts.geen_interesse})` },
        ].map((f) => (
          <button
            key={f.id}
            onClick={() => setFilter(f.id)}
            style={{
              ...styles.filterChip,
              background: filter === f.id ? T.ink : "transparent",
              color: filter === f.id ? "#fff" : T.inkSoft,
              borderColor: filter === f.id ? T.ink : T.line,
            }}
          >
            {f.label}
          </button>
        ))}
      </div>

      {filtered.length === 0 && (
        <EmptyState
          title={customers.length === 0 ? "Nog geen klanten" : "Niets in dit filter"}
          text={
            customers.length === 0
              ? "Voeg de eerste klant toe op basis van postcode en huisnummer, of importeer een Excel-bestand."
              : "Probeer een ander filter."
          }
        />
      )}

      <div style={styles.list}>
        {filtered.map((c) => {
          const meta = STATUS_META[c.status] || STATUS_META.klant;
          return (
            <div key={c.id} style={{ ...styles.row, alignItems: "flex-start" }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <div style={styles.rowTitle}>{c.name || "(zonder naam)"}</div>
                  <span style={{ ...styles.statusBadge, background: meta.bg, color: meta.color }}>{meta.label}</span>
                  {isOnVacation(c) && (
                    <span style={{ ...styles.statusBadge, background: "#E7E2F5", color: "#6B5FA8" }}>
                      Vakantie tot {new Date(c.vacationEnd).toLocaleDateString("nl-NL", { day: "numeric", month: "short" })}
                    </span>
                  )}
                </div>
                <div style={styles.rowSub}>
                  {c.street} {c.houseNumber}{c.addition ? `-${c.addition}` : ""}, {c.postalCode} {c.city}
                </div>
                {c.phone && <div style={styles.rowSub}>{c.phone}</div>}
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                {c.phone && (
                  <button style={styles.iconBtn} onClick={() => setWaFor(c)} title="WhatsApp">
                    <MessageCircle size={17} color={T.yolkDeep} />
                  </button>
                )}
                <button style={styles.iconBtn} onClick={() => { setEditing(c); setShowForm(true); }}>
                  <Pencil size={16} color={T.inkSoft} />
                </button>
                <button style={styles.iconBtn} onClick={() => onDelete(c.id)}>
                  <Trash2 size={16} color={T.danger} />
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {showForm && (
        <CustomerForm
          initial={editing}
          onClose={() => setShowForm(false)}
          onSave={(c) => {
            editing ? onUpdate(c) : onAdd(c);
            setShowForm(false);
          }}
        />
      )}

      {waFor && (
        <ModalOverlay onClose={() => setWaFor(null)}>
          <div style={styles.modalTitle}>Bericht naar {waFor.name || "klant"}</div>
          {Object.entries(WA_TEMPLATES).map(([key, fn]) => (
            <a
              key={key}
              href={waLink(waFor.phone, fn(waFor.name))}
              target="_blank"
              rel="noreferrer"
              style={styles.waOption}
              onClick={() => setWaFor(null)}
            >
              <MessageCircle size={16} color={T.yolkDeep} />
              <div>
                <div style={{ fontWeight: 700 }}>{WA_TEMPLATE_LABELS[key] || key}</div>
                <div style={{ fontSize: 12.5, color: T.inkSoft }}>{fn(waFor.name)}</div>
              </div>
            </a>
          ))}
          <button style={styles.cancelLink} onClick={() => setWaFor(null)}>Annuleren</button>
        </ModalOverlay>
      )}
      {showRoute && (
        <RouteModal customers={customers} statusFilter={filter} onClose={() => setShowRoute(false)} onSave={(list) => { onReorder(list); setShowRoute(false); }} />
      )}
    </div>
  );
}

function RouteModal({ customers, statusFilter = "alle", onClose, onSave }) {
  const routable = useMemo(
    () =>
      customers
        .filter((c) => c.status !== "geen_interesse")
        .sort((a, b) => (a.routeOrder ?? 9999) - (b.routeOrder ?? 9999)),
    [customers]
  );
  const [order, setOrder] = useState(routable.map((c) => c.id));
  const [query, setQuery] = useState("");
  const byId = useMemo(() => Object.fromEntries(customers.map((c) => [c.id, c])), [customers]);

  function move(visIdx, dir) {
    const targetVisIdx = visIdx + dir;
    if (targetVisIdx < 0 || targetVisIdx >= visible.length) return;
    const globalA = visible[visIdx].i;
    const globalB = visible[targetVisIdx].i;
    const newOrder = [...order];
    [newOrder[globalA], newOrder[globalB]] = [newOrder[globalB], newOrder[globalA]];
    setOrder(newOrder);
  }

  function moveToEdge(index, edge) {
    const newOrder = [...order];
    const [item] = newOrder.splice(index, 1);
    if (edge === "top") newOrder.unshift(item);
    else newOrder.push(item);
    setOrder(newOrder);
  }

  function handleSave() {
    const positions = Object.fromEntries(order.map((id, i) => [id, i]));
    const updated = customers.map((c) =>
      positions[c.id] !== undefined ? { ...c, routeOrder: positions[c.id] } : c
    );
    onSave(updated);
  }

  const q = query.trim().toLowerCase();
  const visible = order
    .map((id, i) => ({ id, i }))
    .filter(({ id }) => {
      const c = byId[id];
      if (!c) return false;
      if (statusFilter !== "alle" && (c.status || "klant") !== statusFilter) return false;
      if (!q) return true;
      return `${c.name} ${c.street} ${c.houseNumber} ${c.city}`.toLowerCase().includes(q);
    });

  const filterLabels = { klant: "Klanten", kans: "Kansen", geen_interesse: "Geen interesse" };

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>Looproute aanpassen</div>
      {statusFilter !== "alle" && (
        <div style={{ ...styles.statusBadge, background: T.yolkPale, color: T.yolkDeep, display: "inline-block", marginBottom: 10 }}>
          Gefilterd op: {filterLabels[statusFilter] || statusFilter}
        </div>
      )}
      <div style={styles.formNote}>
        Zet de klanten in de volgorde waarin je ze onderweg tegenkomt. De app toont daarna vanzelf wie de volgende is.
        {order.length > 20 && " Zoek een adres op en gebruik \u201cbovenaan/onderaan\u201d om het snel te verplaatsen."}
      </div>
      <div style={styles.searchRow}>
        <Search size={17} color={T.inkSoft} />
        <input
          placeholder="Zoek adres in de route…"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={styles.searchInput}
        />
      </div>
      <div style={styles.routeList}>
        {visible.length === 0 && (
          <div style={{ fontSize: 13, color: T.inkSoft, padding: "10px 4px" }}>
            {statusFilter === "geen_interesse"
              ? "Klanten met \u2018Geen interesse\u2019 maken geen deel uit van de looproute."
              : "Niets gevonden."}
          </div>
        )}
        {visible.map(({ id, i }, visIdx) => {
          const c = byId[id];
          if (!c) return null;
          const huisnr = `${c.houseNumber}${c.addition ? `-${c.addition}` : ""}`;
          const onVacation = isOnVacation(c);
          return (
            <div key={id} style={{ ...styles.routeRow, opacity: onVacation ? 0.55 : 1 }}>
              <div style={styles.routeNum}>{i + 1}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <div style={styles.rowTitle}>{c.name || `${c.street} ${huisnr}`}</div>
                  {onVacation && (
                    <span style={{ ...styles.statusBadge, background: "#E7E2F5", color: "#6B5FA8" }}>Vakantie</span>
                  )}
                </div>
                <div style={styles.rowSub}>{c.street} {huisnr}</div>
              </div>
              <div style={{ display: "flex", gap: 4 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <button style={styles.routeArrow} disabled={visIdx === 0} onClick={() => move(visIdx, -1)} title="Eén omhoog">
                    <ArrowUp size={14} color={visIdx === 0 ? T.line : T.ink} />
                  </button>
                  <button style={styles.routeArrow} disabled={visIdx === visible.length - 1} onClick={() => move(visIdx, 1)} title="Eén omlaag">
                    <ArrowDown size={14} color={visIdx === visible.length - 1 ? T.line : T.ink} />
                  </button>
                </div>
                {order.length > 8 && (
                  <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                    <button style={styles.routeJumpBtn} disabled={i === 0} onClick={() => moveToEdge(i, "top")}>
                      Bovenaan
                    </button>
                    <button style={styles.routeJumpBtn} disabled={i === order.length - 1} onClick={() => moveToEdge(i, "bottom")}>
                      Onderaan
                    </button>
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
      <button style={styles.confirmBtn} onClick={handleSave}>
        <Check size={18} /> Volgorde opslaan
      </button>
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

function CustomerForm({ initial, onClose, onSave }) {
  const [form, setForm] = useState(
    initial || { id: uid(), name: "", postalCode: "", houseNumber: "", addition: "", street: "", city: "", phone: "", status: "klant", vacationStart: "", vacationEnd: "" }
  );
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>{initial ? "Klant bewerken" : "Nieuwe klant"}</div>
      <div style={styles.formNote}>
        Automatisch adres opzoeken werkt niet in deze omgeving — vul postcode en huisnummer handmatig aan.
      </div>

      <div style={styles.modalLabel}>Status</div>
      <div style={styles.statusToggleRow}>
        {["klant", "kans", "geen_interesse"].map((s) => (
          <button
            key={s}
            onClick={() => setForm({ ...form, status: s })}
            style={{
              ...styles.statusToggleBtn,
              background: form.status === s ? T.yolk : "transparent",
              color: form.status === s ? "#fff" : T.ink,
              borderColor: form.status === s ? T.yolk : T.line,
            }}
          >
            {STATUS_META[s].label}
          </button>
        ))}
      </div>

      <FormField label="Naam (optioneel)" value={form.name} onChange={set("name")} placeholder="Fam. Jansen" autoFocus />
      <div style={{ display: "flex", gap: 8 }}>
        <FormField label="Postcode" value={form.postalCode} onChange={set("postalCode")} placeholder="1234 AB" style={{ flex: 1 }} />
        <FormField label="Huisnr." value={form.houseNumber} onChange={set("houseNumber")} placeholder="12" style={{ width: 70 }} />
        <FormField label="Toev." value={form.addition} onChange={set("addition")} placeholder="a" style={{ width: 60 }} />
      </div>
      <FormField label="Straat" value={form.street} onChange={set("street")} placeholder="Dorpsstraat" />
      <FormField label="Plaats" value={form.city} onChange={set("city")} placeholder="Woerden" />
      <FormField label="Telefoon / WhatsApp (optioneel)" value={form.phone} onChange={set("phone")} placeholder="06 12345678" />

      <div style={styles.modalLabel}>Vakantie (optioneel)</div>
      <div style={styles.formNote}>
        Tijdens deze periode wordt de klant automatisch overgeslagen als "volgende in de looproute" — zonder dat je de status hoeft te wijzigen.
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <FormField label="Van" type="date" value={form.vacationStart || ""} onChange={set("vacationStart")} style={{ flex: 1 }} />
        <FormField label="Tot" type="date" value={form.vacationEnd || ""} onChange={set("vacationEnd")} style={{ flex: 1 }} />
      </div>
      {(form.vacationStart || form.vacationEnd) && (
        <button
          style={styles.cancelLink}
          onClick={() => setForm({ ...form, vacationStart: "", vacationEnd: "" })}
        >
          Vakantie wissen
        </button>
      )}

      <button
        style={styles.confirmBtn}
        onClick={() => {
          if (!form.street && !form.name) return;
          onSave(form);
        }}
      >
        <Check size={18} /> Opslaan
      </button>
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

function FormField({ label, style, ...props }) {
  return (
    <div style={{ marginBottom: 12, ...style }}>
      <label style={styles.fieldLabel}>{label}</label>
      <input {...props} style={styles.fieldInput} />
    </div>
  );
}

/* ---------- Voorraad tab ---------- */
function VoorraadTab({
  stock, purchases, sales, extras, customers, settings, onLogPurchase, onUpdateSettings,
  onUpdateSale, onDeleteSale, onUpdatePurchase, onDeletePurchase,
  onUpdateExtra, onDeleteExtra, onImportAll,
}) {
  const [showBuy, setShowBuy] = useState(false);
  const [priceInput, setPriceInput] = useState(settings.pricePerEgg);
  const [histAvgInput, setHistAvgInput] = useState(settings.historicalWeeklyAvg || "");
  const [editSale, setEditSale] = useState(null);
  const [editPurchase, setEditPurchase] = useState(null);
  const [editExtra, setEditExtra] = useState(null);

  const ledger = useMemo(() => {
    const items = [
      ...purchases.map((p) => ({ ...p, type: "purchase" })),
      ...sales.map((s) => ({ ...s, type: "sale" })),
    ];
    return items.sort((a, b) => b.ts - a.ts).slice(0, 25);
  }, [purchases, sales]);

  const recentExtras = useMemo(() => [...extras].sort((a, b) => b.ts - a.ts).slice(0, 15), [extras]);

  return (
    <div>
      <div style={styles.stockCard}>
        <TrendingUp size={20} color={T.yolkDeep} />
        <div>
          <div style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 28, fontWeight: 700, color: T.ink }}>{stock}</div>
          <div style={{ fontSize: 12.5, color: T.inkSoft }}>eieren op voorraad</div>
        </div>
      </div>

      <button style={styles.primaryBtn} onClick={() => setShowBuy(true)}>
        <Plus size={17} /> Inkoop registreren
      </button>

      <div style={styles.settingsRow}>
        <label style={styles.fieldLabel}>Verkoopprijs per ei</label>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <input
            type="number"
            step="0.01"
            value={priceInput}
            onChange={(e) => setPriceInput(e.target.value)}
            style={{ ...styles.fieldInput, width: 90 }}
          />
          <button
            style={styles.smallBtn}
            onClick={() => onUpdateSettings({ ...settings, pricePerEgg: parseFloat(priceInput) || 0 })}
          >
            Opslaan
          </button>
        </div>
      </div>

      <div style={styles.settingsRow}>
        <label style={styles.fieldLabel}>Historisch weekgemiddelde (voor inkoopvoorspelling)</label>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <input
            type="number"
            value={histAvgInput}
            onChange={(e) => setHistAvgInput(e.target.value)}
            placeholder="bijv. 332"
            style={{ ...styles.fieldInput, width: 90 }}
          />
          <button
            style={styles.smallBtn}
            onClick={() => onUpdateSettings({ ...settings, historicalWeeklyAvg: parseFloat(histAvgInput) || 0 })}
          >
            Opslaan
          </button>
        </div>
      </div>

      <BackupSection
        customers={customers}
        sales={sales}
        purchases={purchases}
        extras={extras}
        settings={settings}
        onImportAll={onImportAll}
      />

      <div style={styles.sectionLabel}>Recente boekingen</div>
      <div style={styles.helperNote}>Tik op een boeking om deze te wijzigen of te verwijderen.</div>
      {ledger.length === 0 && <EmptyState title="Nog geen boekingen" text="Inkoop en verkopen verschijnen hier." />}
      <div style={styles.list}>
        {ledger.map((item) => (
          <button
            key={item.id}
            style={{ ...styles.row, ...styles.rowClickable }}
            onClick={() => (item.type === "purchase" ? setEditPurchase(item) : setEditSale(item))}
          >
            <div style={{ flex: 1, textAlign: "left" }}>
              <div style={styles.rowTitle}>
                {item.type === "purchase" ? "Inkoop bij Egg-sellent" : item.customerLabel}
              </div>
              <div style={styles.rowSub}>{formatDateTime(item.ts)}</div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontWeight: 700, color: item.type === "purchase" ? T.yolkDeep : T.ink }}>
                {item.type === "purchase" ? "+" : "-"}
                {item.eggCount} eieren
              </div>
              {item.type === "sale" && (
                <div style={{ fontSize: 12.5, color: T.inkSoft }}>
                  {formatEuro(item.amount)}
                  {item.tip > 0 && <span style={{ color: T.yolkDeep }}> +{formatEuro(item.tip)} fooi</span>}
                </div>
              )}
            </div>
            <Pencil size={14} color={T.inkSoft} />
          </button>
        ))}
      </div>

      {showBuy && (
        <PurchaseModal onClose={() => setShowBuy(false)} onConfirm={(p) => { onLogPurchase(p); setShowBuy(false); }} />
      )}

      {editSale && (
        <EditSaleModal
          sale={editSale}
          onClose={() => setEditSale(null)}
          onSave={(s) => { onUpdateSale(s); setEditSale(null); }}
          onDelete={(id) => { onDeleteSale(id); setEditSale(null); }}
        />
      )}

      {editPurchase && (
        <EditPurchaseModal
          purchase={editPurchase}
          onClose={() => setEditPurchase(null)}
          onSave={(p) => { onUpdatePurchase(p); setEditPurchase(null); }}
          onDelete={(id) => { onDeletePurchase(id); setEditPurchase(null); }}
        />
      )}

      <div style={styles.sectionLabel}>
        <Star size={13} style={{ marginRight: 4, verticalAlign: -2 }} />
        Acties &amp; extra's
      </div>
      <div style={styles.helperNote}>Losse acties, zoals een stroopwafelverkoop, tel je hier apart mee — niet bij de eierstatistieken.</div>
      {recentExtras.length === 0 && (
        <EmptyState title="Nog geen acties" text="Log een extra verkoop via het tabblad Verkoop." />
      )}
      <div style={styles.list}>
        {recentExtras.map((e) => (
          <button key={e.id} style={{ ...styles.row, ...styles.rowClickable }} onClick={() => setEditExtra(e)}>
            <div style={{ flex: 1, textAlign: "left" }}>
              <div style={styles.rowTitle}>{e.naam}</div>
              <div style={styles.rowSub}>{formatDateTime(e.ts)} · {e.aantal}x</div>
            </div>
            <div style={{ fontWeight: 700 }}>{formatEuro(e.bedrag)}</div>
            <Pencil size={14} color={T.inkSoft} />
          </button>
        ))}
      </div>

      {editExtra && (
        <ExtraModal
          initial={editExtra}
          onClose={() => setEditExtra(null)}
          onConfirm={(e) => { onUpdateExtra(e); setEditExtra(null); }}
          onDelete={(id) => { onDeleteExtra(id); setEditExtra(null); }}
        />
      )}
    </div>
  );
}

function EditSaleModal({ sale, onClose, onSave, onDelete }) {
  const [count, setCount] = useState(sale.eggCount);
  const [amount, setAmount] = useState(sale.amount.toFixed(2));
  const [tip, setTip] = useState(sale.tip ? sale.tip.toFixed(2) : "");

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>Verkoop bewerken</div>
      <div style={styles.modalLabel}>{sale.customerLabel} · {formatDateTime(sale.ts)}</div>
      <div style={{ height: 10 }} />
      <div style={styles.modalLabel}>Aantal eieren</div>
      <div style={styles.stepperRow}>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => Math.max(1, n - 1))}>
          <MinusCircle size={22} color={T.ink} />
        </button>
        <div style={styles.stepperNum}>{count}</div>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => n + 1)}>
          <PlusCircle size={22} color={T.ink} />
        </button>
      </div>
      <div style={styles.modalLabel}>Bedrag</div>
      <div style={styles.amountRow}>
        <span style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 20, color: T.inkSoft }}>€</span>
        <input type="number" step="0.01" value={amount} onChange={(e) => setAmount(e.target.value)} style={styles.amountInput} />
      </div>
      <div style={styles.modalLabel}>Fooi (optioneel)</div>
      <div style={styles.amountRow}>
        <span style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 20, color: T.inkSoft }}>€</span>
        <input type="number" step="0.01" value={tip} onChange={(e) => setTip(e.target.value)} placeholder="0,00" style={styles.amountInput} />
      </div>
      <button
        style={styles.confirmBtn}
        onClick={() => onSave({ ...sale, eggCount: count, amount: parseFloat(amount) || 0, tip: parseFloat(tip) || 0 })}
      >
        <Check size={18} /> Wijzigingen opslaan
      </button>
      <DeleteConfirmButton onDelete={() => onDelete(sale.id)} label="verkoop" />
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

function EditPurchaseModal({ purchase, onClose, onSave, onDelete }) {
  const [count, setCount] = useState(purchase.eggCount);
  const [cost, setCost] = useState(purchase.cost ? String(purchase.cost) : "");

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>Inkoop bewerken</div>
      <div style={styles.modalLabel}>{formatDateTime(purchase.ts)}</div>
      <div style={{ height: 10 }} />
      <div style={styles.modalLabel}>Aantal eieren besteld</div>
      <div style={styles.stepperRow}>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => Math.max(6, n - 6))}>
          <MinusCircle size={22} color={T.ink} />
        </button>
        <div style={styles.stepperNum}>{count}</div>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => n + 6)}>
          <PlusCircle size={22} color={T.ink} />
        </button>
      </div>
      <FormField label="Inkoopkosten (optioneel)" value={cost} onChange={(e) => setCost(e.target.value)} placeholder="bijv. 18,50" />
      <button
        style={styles.confirmBtn}
        onClick={() => onSave({ ...purchase, eggCount: count, cost: parseFloat(cost) || 0 })}
      >
        <Check size={18} /> Wijzigingen opslaan
      </button>
      <DeleteConfirmButton onDelete={() => onDelete(purchase.id)} label="inkoop" />
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

function BulkSendModal({ title, note, needsCustomText, customPlaceholder, buildMessage, customers, onClose }) {
  const [customText, setCustomText] = useState("");
  const eligible = useMemo(
    () => customers.filter((c) => c.phone && c.status !== "geen_interesse" && !isOnVacation(c)),
    [customers]
  );
  const skippedCount = customers.filter((c) => c.status !== "geen_interesse" && !isOnVacation(c)).length - eligible.length;

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>{title}</div>
      {note && <div style={styles.formNote}>{note}</div>}
      {needsCustomText && (
        <FormField
          label="Opmerking (verschijnt in het bericht)"
          value={customText}
          onChange={(e) => setCustomText(e.target.value)}
          placeholder={customPlaceholder}
        />
      )}
      <div style={{ ...styles.formNote, background: T.cream }}>
        {eligible.length} klant{eligible.length === 1 ? "" : "en"} met telefoonnummer.
        {skippedCount > 0 && ` ${skippedCount} overgeslagen (geen nummer of op vakantie).`}
        {" "}Tik per klant op "Verstuur" — WhatsApp opent dan met het bericht klaar, jij verstuurt het zelf.
      </div>
      {eligible.length === 0 ? (
        <EmptyState title="Niemand om naar te sturen" text="Voeg telefoonnummers toe bij klanten om deze functie te gebruiken." />
      ) : (
        <div style={styles.list}>
          {eligible.map((c) => (
            <div key={c.id} style={styles.row}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={styles.rowTitle}>{c.name || `${c.street} ${c.houseNumber}`}</div>
                <div style={styles.rowSub}>{c.phone}</div>
              </div>
              <a
                href={waLink(c.phone, buildMessage(c, customText))}
                target="_blank"
                rel="noreferrer"
                style={styles.saleBtn}
              >
                Verstuur
              </a>
            </div>
          ))}
        </div>
      )}
      <button style={styles.cancelLink} onClick={onClose}>Sluiten</button>
    </ModalOverlay>
  );
}


function DeleteConfirmButton({ onDelete, label }) {
  const [confirming, setConfirming] = useState(false);
  if (!confirming) {
    return (
      <button style={styles.deleteLink} onClick={() => setConfirming(true)}>
        <Trash2 size={14} /> Deze {label} verwijderen
      </button>
    );
  }
  return (
    <button style={styles.deleteLinkConfirm} onClick={onDelete}>
      <Trash2 size={14} /> Zeker weten? Tik nogmaals om te verwijderen
    </button>
  );
}

function BackupSection({ customers, sales, purchases, extras, settings, onImportAll }) {
  const fileInputRef = useRef(null);
  const [confirmImport, setConfirmImport] = useState(null); // parsed data pending confirmation
  const [message, setMessage] = useState("");

  function handleExport() {
    const payload = {
      exportedAt: new Date().toISOString(),
      app: "Egg-sellerate",
      customers,
      sales,
      purchases,
      extras,
      settings,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const dateStr = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `egg-sellerate-backup-${dateStr}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function handleFileChange(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        if (!data.customers && !data.sales && !data.purchases && !data.extras) {
          setMessage("Dit lijkt geen geldig Egg-sellerate-back-upbestand.");
          return;
        }
        setConfirmImport(data);
        setMessage("");
      } catch {
        setMessage("Kon dit bestand niet lezen.");
      }
    };
    reader.readAsText(file);
  }

  return (
    <div style={styles.settingsRow}>
      <label style={styles.fieldLabel}>Back-up</label>
      <div style={{ fontSize: 12, color: T.inkSoft, marginBottom: 10, lineHeight: 1.4 }}>
        Bewaar al je gegevens (klanten, verkopen, inkoop, acties) als bestand — handig als reserve, of om mee te nemen naar een andere plek.
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button style={styles.secondaryBtn} onClick={handleExport}>
          <Download size={15} /> Exporteer alles
        </button>
        <button style={styles.secondaryBtnGhost} onClick={() => fileInputRef.current?.click()}>
          <Upload size={15} /> Importeer back-up
        </button>
        <input ref={fileInputRef} type="file" accept=".json" onChange={handleFileChange} style={{ display: "none" }} />
      </div>
      {message && <div style={{ ...styles.importMsg, marginTop: 10 }}>{message}</div>}

      {confirmImport && (
        <ModalOverlay onClose={() => setConfirmImport(null)}>
          <div style={styles.modalTitle}>Back-up terugzetten?</div>
          <div style={styles.formNote}>
            Dit vervangt al je huidige klanten, verkopen, inkoop en acties door de inhoud van dit bestand
            {confirmImport.exportedAt ? ` (geëxporteerd op ${new Date(confirmImport.exportedAt).toLocaleDateString("nl-NL")})` : ""}.
            Dit kan niet ongedaan gemaakt worden.
          </div>
          <button
            style={styles.confirmBtn}
            onClick={async () => {
              await onImportAll(confirmImport);
              setConfirmImport(null);
              setMessage("Back-up teruggezet.");
            }}
          >
            <Check size={18} /> Ja, vervangen
          </button>
          <button style={styles.cancelLink} onClick={() => setConfirmImport(null)}>Annuleren</button>
        </ModalOverlay>
      )}
    </div>
  );
}

function PurchaseModal({ onClose, onConfirm }) {
  const [count, setCount] = useState(60);
  const [cost, setCost] = useState("");

  return (
    <ModalOverlay onClose={onClose}>
      <div style={styles.modalTitle}>Inkoop registreren</div>
      <div style={styles.modalLabel}>Aantal eieren besteld</div>
      <div style={styles.stepperRow}>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => Math.max(6, n - 6))}>
          <MinusCircle size={22} color={T.ink} />
        </button>
        <div style={styles.stepperNum}>{count}</div>
        <button style={styles.stepperBtn} onClick={() => setCount((n) => n + 6)}>
          <PlusCircle size={22} color={T.ink} />
        </button>
      </div>
      <FormField
        label="Inkoopkosten (optioneel)"
        value={cost}
        onChange={(e) => setCost(e.target.value)}
        placeholder="bijv. 18,50"
      />
      <button
        style={styles.confirmBtn}
        onClick={() =>
          onConfirm({ id: uid(), eggCount: count, cost: parseFloat(cost) || 0, ts: Date.now() })
        }
      >
        <Check size={18} /> Opslaan
      </button>
      <button style={styles.cancelLink} onClick={onClose}>Annuleren</button>
    </ModalOverlay>
  );
}

/* ---------- Berichten tab ---------- */
function BerichtenTab({ customers, settings, onUpdateSettings }) {
  const [absenceStart, setAbsenceStart] = useState(settings.absenceStart || "");
  const [absenceEnd, setAbsenceEnd] = useState(settings.absenceEnd || "");
  const [bulkModal, setBulkModal] = useState(null); // "herinnering" | "afwezigheid" | "weer"

  const isThursdayAfternoon = useMemo(() => {
    const now = new Date();
    return now.getDay() === 4 && now.getHours() >= 12;
  }, []);

  return (
    <div>
      {isThursdayAfternoon && (
        <div style={styles.thursdayBanner}>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 13.5 }}>Het is donderdagmiddag 🥚</div>
            <div style={{ fontSize: 12, color: T.inkSoft }}>Goed moment om de herinnering te versturen.</div>
          </div>
          <button style={styles.smallBtn} onClick={() => setBulkModal("herinnering")}>Versturen</button>
        </div>
      )}

      <div style={styles.sectionLabel}>
        <Send size={13} style={{ marginRight: 4, verticalAlign: -2 }} />
        Berichten naar alle klanten
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 18 }}>
        <button style={styles.secondaryBtnGhost2} onClick={() => setBulkModal("herinnering")}>
          Stuur herinnering (morgen langs)
        </button>
        <button
          style={{ ...styles.secondaryBtnGhost2, opacity: settings.absenceStart || settings.absenceEnd ? 1 : 0.5 }}
          onClick={() => (settings.absenceStart || settings.absenceEnd) && setBulkModal("afwezigheid")}
        >
          Stuur afwezigheidsbericht
        </button>
        <button style={styles.secondaryBtnGhost2} onClick={() => setBulkModal("weer")}>
          Stuur weerbericht
        </button>
      </div>

      <div style={styles.settingsRow}>
        <label style={styles.fieldLabel}>Eigen afwezigheid (voor het afwezigheidsbericht)</label>
        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <input type="date" value={absenceStart} onChange={(e) => setAbsenceStart(e.target.value)} style={{ ...styles.fieldInput, flex: 1 }} />
          <input type="date" value={absenceEnd} onChange={(e) => setAbsenceEnd(e.target.value)} style={{ ...styles.fieldInput, flex: 1 }} />
        </div>
        <button
          style={styles.smallBtn}
          onClick={() => onUpdateSettings({ ...settings, absenceStart, absenceEnd })}
        >
          Opslaan
        </button>
      </div>

      <div style={styles.sectionLabel}>Sjablonen</div>
      <div style={styles.list}>
        {Object.entries(WA_TEMPLATE_LABELS).map(([key, label]) => (
          <div key={key} style={styles.row}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={styles.rowTitle}>{label}</div>
              <div style={styles.rowSub}>{WA_TEMPLATES[key]("...")}</div>
            </div>
          </div>
        ))}
      </div>

      {bulkModal === "herinnering" && (
        <BulkSendModal
          title="Herinnering versturen"
          note="Ga morgen (vrijdag) weer langs? Stuur dit naar iedereen die je gewend bent te bezoeken."
          buildMessage={(c) => WA_TEMPLATES.herinnering(c.name)}
          customers={customers}
          onClose={() => setBulkModal(null)}
        />
      )}
      {bulkModal === "afwezigheid" && (
        <BulkSendModal
          title="Afwezigheidsbericht versturen"
          note="Gebruikt de afwezigheidsdatum die je hierboven hebt ingesteld."
          buildMessage={(c) => vakantieBericht(c.name, settings.absenceStart, settings.absenceEnd)}
          customers={customers}
          onClose={() => setBulkModal(null)}
        />
      )}
      {bulkModal === "weer" && (
        <BulkSendModal
          title="Weerbericht versturen"
          note="Bijvoorbeeld bij storm of noodweer: laat weten dat je later, eerder of een andere dag komt."
          needsCustomText
          customPlaceholder="bijv. ik kom vandaag pas na 17:00 langs"
          buildMessage={(c, text) => weerBericht(c.name, text)}
          customers={customers}
          onClose={() => setBulkModal(null)}
        />
      )}
    </div>
  );
}

/* ---------- Statistieken tab ---------- */
function StatsTab({ customers, sales, purchases, extras, settings }) {
  const [waFor, setWaFor] = useState(null);
  const [period, setPeriod] = useState("week");

  const stats = useMemo(() => {
    const now = new Date();

    const periodStart = {
      week: now.getTime() - 7 * 24 * 3600 * 1000,
      maand: new Date(now.getFullYear(), now.getMonth(), 1).getTime(),
      jaar: new Date(now.getFullYear(), 0, 1).getTime(),
      totaal: 0,
    }[period];

    const salesInPeriod = sales.filter((s) => s.ts >= periodStart);
    const extrasInPeriod = extras.filter((e) => e.ts >= periodStart);
    const purchasesInPeriod = purchases.filter((p) => p.ts >= periodStart);

    const revenue = salesInPeriod.reduce((s, x) => s + x.amount, 0);
    const extraRevenue = extrasInPeriod.reduce((s, x) => s + x.bedrag, 0);
    const tips = salesInPeriod.reduce((s, x) => s + (x.tip || 0), 0);
    const eggsSold = salesInPeriod.reduce((s, x) => s + x.eggCount, 0);
    const eggsBought = purchasesInPeriod.reduce((s, x) => s + x.eggCount, 0);
    const cost = purchasesInPeriod.reduce((s, x) => s + (x.cost || 0), 0);
    const margin = revenue - cost;

    // laatste 8 weken omzet (altijd, los van periodefilter — dit is de trend)
    const weeks = [];
    for (let i = 7; i >= 0; i--) {
      const end = now.getTime() - i * 7 * 24 * 3600 * 1000;
      const start = end - 7 * 24 * 3600 * 1000;
      const weekSales = sales.filter((s) => s.ts >= start && s.ts < end);
      const wRevenue = weekSales.reduce((s, x) => s + x.amount, 0);
      const wEggs = weekSales.reduce((s, x) => s + x.eggCount, 0);
      weeks.push({ label: i === 0 ? "Nu" : `-${i}w`, revenue: wRevenue, eggs: wEggs });
    }
    const maxWeek = Math.max(1, ...weeks.map((w) => w.revenue));
    const weeksWithData = weeks.filter((w) => w.eggs > 0);
    const recentEggsAvg = weeksWithData.length
      ? weeksWithData.reduce((s, w) => s + w.eggs, 0) / weeksWithData.length
      : 0;

    // top klanten op fooi (all-time)
    const tipByCustomer = {};
    sales.forEach((s) => {
      if (!s.customerId || !s.tip) return;
      tipByCustomer[s.customerId] = (tipByCustomer[s.customerId] || 0) + s.tip;
    });
    const topTippers = Object.entries(tipByCustomer)
      .map(([id, tip]) => ({ customer: customers.find((c) => c.id === id), tip }))
      .filter((x) => x.customer)
      .sort((a, b) => b.tip - a.tip)
      .slice(0, 5);

    // top 5 klanten op omzet (all-time)
    const revenueByCustomer = {};
    sales.forEach((s) => {
      if (!s.customerId) return;
      revenueByCustomer[s.customerId] = (revenueByCustomer[s.customerId] || 0) + s.amount;
    });
    const topSpenders = Object.entries(revenueByCustomer)
      .map(([id, amount]) => ({ customer: customers.find((c) => c.id === id), amount }))
      .filter((x) => x.customer)
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 5);

    return { revenue, extraRevenue, tips, eggsSold, eggsBought, cost, margin, weeks, maxWeek, weeksWithData, recentEggsAvg, topTippers, topSpenders };
  }, [sales, purchases, extras, customers, period]);

  const periodLabels = { week: "deze week", maand: "deze maand", jaar: "dit jaar", totaal: "totaal" };

  return (
    <div>
      <div style={styles.filterRow}>
        {["week", "maand", "jaar", "totaal"].map((p) => (
          <button
            key={p}
            onClick={() => setPeriod(p)}
            style={{
              ...styles.filterChip,
              background: period === p ? T.ink : "transparent",
              color: period === p ? "#fff" : T.inkSoft,
              borderColor: period === p ? T.ink : T.line,
              textTransform: "capitalize",
            }}
          >
            {p}
          </button>
        ))}
      </div>

      <div style={styles.statsGrid}>
        <StatCard label={`Omzet eieren (${periodLabels[period]})`} value={formatEuro(stats.revenue)} />
        <StatCard label={`Omzet acties (${periodLabels[period]})`} value={formatEuro(stats.extraRevenue)} />
        <StatCard label="Eieren verkocht" value={stats.eggsSold} />
        <StatCard label="Eieren ingekocht" value={stats.eggsBought} />
        {stats.cost > 0 && <StatCard label="Winstmarge (eieren)" value={formatEuro(stats.margin)} accent />}
        <StatCard label="Fooi" value={formatEuro(stats.tips)} accent />
      </div>

      {(settings.historicalWeeklyAvg > 0 || stats.weeksWithData.length >= 2) && (() => {
        const useLive = stats.weeksWithData.length >= 2;
        const seasonal = getSeasonalFactor(new Date());
        const base = useLive ? stats.recentEggsAvg : settings.historicalWeeklyAvg;
        const predicted = useLive ? base : base * seasonal.factor;
        return (
          <div style={styles.stockCard}>
            <TrendingUp size={20} color={T.yolkDeep} />
            <div>
              <div style={{ fontFamily: "'Baloo 2',sans-serif", fontSize: 22, fontWeight: 700, color: T.ink }}>
                ~{Math.round(predicted)} eieren
              </div>
              <div style={{ fontSize: 12.5, color: T.inkSoft }}>
                Verwachte vraag komende week
                {useLive
                  ? ` (gemiddelde van ${stats.weeksWithData.length} recente weken)`
                  : ` (historisch gemiddelde${seasonal.label ? `, ${seasonal.label}correctie toegepast` : ""})`}
              </div>
              {!useLive && seasonal.label && (
                <div style={{ fontSize: 11, color: T.inkSoft, marginTop: 2, fontStyle: "italic" }}>
                  Gebaseerd op maar 1 jaar ervaring — check dit gaandeweg.
                </div>
              )}
            </div>
          </div>
        );
      })()}

      <div style={styles.sectionLabel}>Omzet per week (laatste 8 weken)</div>
      <div style={styles.weekChart}>
        {stats.weeks.map((w, i) => (
          <div key={i} style={styles.weekBarWrap}>
            <div style={{ ...styles.weekBar, height: `${Math.max(4, (w.revenue / stats.maxWeek) * 70)}px` }} />
            <div style={styles.weekLabel}>{w.label}</div>
          </div>
        ))}
      </div>

      <div style={styles.sectionLabel}>Top 5 klanten (omzet all-time)</div>
      {stats.topSpenders.length === 0 ? (
        <EmptyState title="Nog geen verkopen" text="Zodra er verkopen aan klanten worden gelogd, verschijnt de top 5 hier." />
      ) : (
        <div style={styles.list}>
          {stats.topSpenders.map(({ customer: c, amount }, i) => (
            <div key={c.id} style={styles.row}>
              <div style={styles.routeNum}>{i + 1}</div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={styles.rowTitle}>{c.name || `${c.street} ${c.houseNumber}`}</div>
                <div style={styles.rowSub}>{formatEuro(amount)} totale omzet</div>
              </div>
            </div>
          ))}
        </div>
      )}

      <div style={styles.sectionLabel}>
        <Heart size={13} style={{ marginRight: 4, verticalAlign: -2 }} />
        Speciale klanten (meeste fooi)
      </div>
      {stats.topTippers.length === 0 ? (
        <EmptyState title="Nog geen fooi ontvangen" text="Zodra iemand fooi geeft bij een verkoop, verschijnt diegene hier." />
      ) : (
        <div style={styles.list}>
          {stats.topTippers.map(({ customer: c, tip }) => (
            <div key={c.id} style={styles.row}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={styles.rowTitle}>{c.name || `${c.street} ${c.houseNumber}`}</div>
                <div style={styles.rowSub}>{formatEuro(tip)} totaal aan fooi</div>
              </div>
              {c.phone && (
                <button style={styles.iconBtn} onClick={() => setWaFor(c)} title="Bedankje sturen">
                  <Gift size={17} color={T.yolkDeep} />
                </button>
              )}
            </div>
          ))}
        </div>
      )}

      {waFor && (
        <ModalOverlay onClose={() => setWaFor(null)}>
          <div style={styles.modalTitle}>Bedankje naar {waFor.name || "klant"}</div>
          <a
            href={waLink(waFor.phone, WA_TEMPLATES.bedankjeFooi(waFor.name))}
            target="_blank"
            rel="noreferrer"
            style={styles.waOption}
            onClick={() => setWaFor(null)}
          >
            <Gift size={16} color={T.yolkDeep} />
            <div>
              <div style={{ fontWeight: 700 }}>Bedankje</div>
              <div style={{ fontSize: 12.5, color: T.inkSoft }}>{WA_TEMPLATES.bedankjeFooi(waFor.name)}</div>
            </div>
          </a>
          <button style={styles.cancelLink} onClick={() => setWaFor(null)}>Annuleren</button>
        </ModalOverlay>
      )}
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div style={{ ...styles.statCard, borderColor: accent ? T.yolk : T.line }}>
      <div style={{ ...styles.statCardValue, color: accent ? T.yolkDeep : T.ink }}>{value}</div>
      <div style={styles.statCardLabel}>{label}</div>
    </div>
  );
}


function EmptyState({ title, text }) {
  return (
    <div style={styles.empty}>
      <EggMascot size={44} />
      <div style={{ fontWeight: 700, marginTop: 8, marginBottom: 4 }}>{title}</div>
      <div style={{ fontSize: 13.5, color: T.inkSoft }}>{text}</div>
    </div>
  );
}

function ModalOverlay({ children, onClose }) {
  return (
    <div style={styles.overlay} onClick={onClose}>
      <div style={{ ...styles.modal, animation: "popIn 0.2s ease-out" }} onClick={(e) => e.stopPropagation()}>
        <button style={styles.modalClose} onClick={onClose}>
          <X size={18} color={T.inkSoft} />
        </button>
        {children}
      </div>
    </div>
  );
}

/* ---------- styles ---------- */
const styles = {
  appOuter: {
    minHeight: "100%",
    display: "flex",
    justifyContent: "center",
    background: "transparent",
    fontFamily: "'Nunito', sans-serif",
    color: T.ink,
  },
  phone: {
    width: "100%",
    maxWidth: 430,
    minHeight: 640,
    background: T.bg,
    display: "flex",
    flexDirection: "column",
    position: "relative",
    borderRadius: 22,
    overflow: "hidden",
    border: `1px solid ${T.line}`,
    boxShadow: "0 8px 30px rgba(222,142,16,0.15)",
  },
  header: {
    background: `linear-gradient(135deg, ${T.headerFrom} 0%, ${T.headerTo} 100%)`,
    color: T.ink,
    padding: "20px 20px 18px",
  },
  headerTop: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 },
  headerTitle: { fontFamily: "'Baloo 2', sans-serif", fontSize: 26, fontWeight: 800, lineHeight: 1 },
  headerSlogan: { fontSize: 12.5, fontWeight: 600, opacity: 0.75, marginTop: 3, fontStyle: "italic" },
  headerStats: { display: "flex", gap: 10 },
  headerStatCard: {
    flex: 1,
    background: "rgba(255,255,255,0.55)",
    borderRadius: 14,
    padding: "10px 12px",
  },
  headerStatNum: { fontFamily: "'Baloo 2', sans-serif", fontSize: 20, fontWeight: 700, lineHeight: 1.1 },
  headerStatLabel: { fontSize: 11, fontWeight: 600, opacity: 0.8, marginTop: 2 },
  content: { flex: 1, overflowY: "auto", padding: "16px", background: T.cream },
  errorBanner: { background: "#FBDCD8", color: T.danger, fontSize: 12.5, padding: "8px 16px" },
  searchRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: T.surface,
    border: `1.5px solid ${T.line}`,
    borderRadius: 999,
    padding: "10px 14px",
    marginBottom: 10,
  },
  searchInput: { border: "none", outline: "none", background: "transparent", flex: 1, fontSize: 14.5, color: T.ink },
  adhocBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    width: "100%",
    padding: "10px 12px",
    background: "transparent",
    border: `1.5px dashed ${T.yolk}`,
    borderRadius: 999,
    color: T.yolkDeep,
    fontSize: 13.5,
    fontWeight: 700,
    marginBottom: 16,
  },
  extraBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    width: "100%",
    padding: "10px 12px",
    background: "transparent",
    border: `1.5px dashed ${T.line}`,
    borderRadius: 999,
    color: T.inkSoft,
    fontSize: 13,
    fontWeight: 700,
    marginBottom: 16,
  },
  primaryBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    width: "100%",
    padding: "13px 12px",
    background: T.yolk,
    border: "none",
    borderRadius: 999,
    color: "#fff",
    fontSize: 14.5,
    fontWeight: 700,
    marginBottom: 16,
    boxShadow: "0 4px 10px rgba(245,166,35,0.35)",
  },
  list: { display: "flex", flexDirection: "column", gap: 8 },
  row: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: T.surface,
    border: `1.5px solid ${T.line}`,
    borderRadius: 14,
    padding: "10px 12px",
  },
  rowTitle: { fontWeight: 700, fontSize: 14.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  rowSub: { fontSize: 12.5, color: T.inkSoft, marginTop: 2 },
  saleBtn: {
    display: "flex",
    alignItems: "center",
    gap: 2,
    background: T.ink,
    color: "#fff",
    border: "none",
    borderRadius: 999,
    padding: "8px 12px",
    fontSize: 13,
    fontWeight: 700,
    whiteSpace: "nowrap",
  },
  iconBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: 34,
    height: 34,
    borderRadius: 999,
    border: `1.5px solid ${T.line}`,
    background: T.surface,
  },
  tabBar: {
    display: "flex",
    borderTop: `1.5px solid ${T.line}`,
    background: T.surface,
    flexShrink: 0,
  },
  tabBtn: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: 3,
    padding: "10px 0 12px",
    background: "transparent",
    border: "none",
    fontSize: 11.5,
  },
  overlay: {
    position: "absolute",
    inset: 0,
    background: "rgba(43,33,24,0.45)",
    display: "flex",
    alignItems: "flex-end",
    zIndex: 20,
  },
  modal: {
    width: "100%",
    background: "#fff",
    borderRadius: "24px 24px 0 0",
    padding: "22px 20px 26px",
    position: "relative",
    maxHeight: "85%",
    overflowY: "auto",
  },
  modalClose: { position: "absolute", top: 16, right: 16, background: "transparent", border: "none" },
  modalTitle: { fontFamily: "'Baloo 2',sans-serif", fontSize: 20, fontWeight: 700, marginBottom: 16 },
  modalLabel: { fontSize: 12.5, color: T.inkSoft, marginBottom: 6, fontWeight: 700 },
  stepperRow: { display: "flex", alignItems: "center", justifyContent: "center", gap: 20, marginBottom: 12 },
  stepperBtn: { background: "transparent", border: "none", display: "flex" },
  stepperNum: { fontFamily: "'Baloo 2',sans-serif", fontSize: 36, fontWeight: 700, minWidth: 60, textAlign: "center" },
  presetRow: { display: "flex", gap: 8, marginBottom: 18, justifyContent: "center", flexWrap: "wrap" },
  presetBtn: { border: "1.5px solid", borderRadius: 999, padding: "6px 14px", fontSize: 13.5, fontWeight: 700, background: "transparent" },
  amountRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    background: T.cream,
    border: `1.5px solid ${T.line}`,
    borderRadius: 14,
    padding: "8px 14px",
    marginBottom: 18,
  },
  amountInput: { border: "none", outline: "none", background: "transparent", fontSize: 20, flex: 1, fontFamily: "'Baloo 2',sans-serif" },
  confirmBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: 6,
    width: "100%",
    padding: "14px",
    background: T.yolk,
    color: "#fff",
    border: "none",
    borderRadius: 999,
    fontSize: 15,
    fontWeight: 700,
    marginBottom: 10,
    boxShadow: "0 4px 10px rgba(245,166,35,0.35)",
  },
  cancelLink: { display: "block", width: "100%", textAlign: "center", background: "transparent", border: "none", color: T.inkSoft, fontSize: 13.5, padding: 6, fontWeight: 600 },
  fieldLabel: { display: "block", fontSize: 12, color: T.inkSoft, marginBottom: 4, fontWeight: 700 },
  fieldInput: {
    width: "100%",
    border: `1.5px solid ${T.line}`,
    borderRadius: 12,
    padding: "9px 12px",
    fontSize: 14,
    background: "#fff",
    color: T.ink,
    outline: "none",
  },
  formNote: { fontSize: 12, color: T.yolkDeep, background: T.yolkPale, borderRadius: 12, padding: "8px 12px", marginBottom: 14, lineHeight: 1.4, fontWeight: 600 },
  waOption: { display: "flex", gap: 10, alignItems: "flex-start", padding: "10px 4px", borderBottom: `1px solid ${T.line}`, textDecoration: "none", color: T.ink },
  stockCard: {
    display: "flex",
    alignItems: "center",
    gap: 12,
    background: "#fff",
    border: `1.5px solid ${T.line}`,
    borderRadius: 16,
    padding: "14px 16px",
    marginBottom: 14,
  },
  settingsRow: { background: "#fff", border: `1.5px solid ${T.line}`, borderRadius: 16, padding: "12px 14px", marginBottom: 18 },
  smallBtn: { background: T.ink, color: "#fff", border: "none", borderRadius: 999, padding: "8px 14px", fontSize: 12.5, fontWeight: 700 },
  sectionLabel: { fontSize: 13, fontWeight: 700, color: T.inkSoft, marginBottom: 4 },
  helperNote: { fontSize: 11.5, color: T.inkSoft, marginBottom: 10 },
  rowClickable: { width: "100%", textAlign: "left", cursor: "pointer" },
  importRow: { display: "flex", gap: 8, marginBottom: 10 },
  secondaryBtn: {
    flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    padding: "9px 10px", background: T.ink, color: "#fff", border: "none",
    borderRadius: 999, fontSize: 12.5, fontWeight: 700,
  },
  secondaryBtnGhost: {
    flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    padding: "9px 10px", background: "transparent", color: T.inkSoft, border: `1.5px solid ${T.line}`,
    borderRadius: 999, fontSize: 12.5, fontWeight: 700,
  },
  secondaryBtnGhost2: {
    width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    padding: "10px 10px", background: "transparent", color: T.yolkDeep, border: `1.5px dashed ${T.yolk}`,
    borderRadius: 999, fontSize: 12.5, fontWeight: 700, marginBottom: 16,
  },
  importMsg: { fontSize: 12, color: T.yolkDeep, background: T.yolkPale, borderRadius: 10, padding: "8px 10px", marginBottom: 10 },
  filterRow: { display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" },
  filterChip: { border: "1.5px solid", borderRadius: 999, padding: "6px 12px", fontSize: 12, fontWeight: 700, background: "transparent" },
  statusBadge: { fontSize: 10.5, fontWeight: 700, borderRadius: 999, padding: "2px 8px" },
  visitedBadge: {
    display: "inline-flex", alignItems: "center", gap: 3, fontSize: 10.5, fontWeight: 700,
    borderRadius: 999, padding: "2px 8px", background: T.greenSoft || "#EFEAE0", color: T.inkSoft,
  },
  thursdayBanner: {
    display: "flex", alignItems: "center", gap: 8,
    background: T.yolkPale, border: `1.5px solid ${T.yolk}`, borderRadius: 14,
    padding: "10px 12px", marginBottom: 14,
  },
  nextCard: {
    background: `linear-gradient(135deg, ${T.headerFrom} 0%, ${T.headerTo} 100%)`,
    borderRadius: 18,
    padding: "14px 16px",
    marginBottom: 14,
  },
  nextLabel: { display: "flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 700, color: T.ink, opacity: 0.75, marginBottom: 6 },
  nextName: { fontFamily: "'Baloo 2',sans-serif", fontSize: 19, fontWeight: 700, color: T.ink },
  nextAddress: { fontSize: 12.5, color: T.ink, opacity: 0.75, marginBottom: 10 },
  nextBtn: {
    display: "flex", alignItems: "center", justifyContent: "center", gap: 4,
    background: T.ink, color: "#fff", border: "none", borderRadius: 999,
    padding: "10px 14px", fontSize: 13.5, fontWeight: 700, width: "100%",
  },
  routeList: { display: "flex", flexDirection: "column", gap: 6, marginBottom: 16, maxHeight: 320, overflowY: "auto" },
  routeRow: {
    display: "flex", alignItems: "center", gap: 10, background: T.cream,
    border: `1.5px solid ${T.line}`, borderRadius: 12, padding: "8px 10px",
  },
  routeNum: {
    width: 22, height: 22, borderRadius: 999, background: T.yolk, color: "#fff",
    fontSize: 11.5, fontWeight: 700, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
  },
  routeArrow: { background: "transparent", border: "none", padding: 2, display: "flex" },
  routeJumpBtn: {
    background: T.cream, border: `1px solid ${T.line}`, borderRadius: 6,
    padding: "3px 6px", fontSize: 9.5, fontWeight: 700, color: T.inkSoft, whiteSpace: "nowrap",
  },
  statusToggleRow: { display: "flex", gap: 6, marginBottom: 16 },
  statusToggleBtn: { flex: 1, border: "1.5px solid", borderRadius: 999, padding: "9px 2px", fontSize: 11.5, fontWeight: 700, background: "transparent", lineHeight: 1.2 },
  deleteLink: {
    display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    width: "100%", padding: "10px", background: "transparent", border: "none",
    color: T.danger, fontSize: 13, fontWeight: 700, marginBottom: 4,
  },
  deleteLinkConfirm: {
    display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    width: "100%", padding: "10px", background: "#FBDCD8", border: "none",
    borderRadius: 999, color: T.danger, fontSize: 13, fontWeight: 700, marginBottom: 4,
  },
  empty: { background: "#fff", border: `1.5px dashed ${T.line}`, borderRadius: 16, padding: "24px 16px", textAlign: "center", marginBottom: 12 },
  celebrateBox: { textAlign: "center", padding: "20px 0" },
  celebrateText: { fontFamily: "'Baloo 2',sans-serif", fontSize: 26, fontWeight: 800, color: T.yolkDeep, marginTop: 10 },
  tipToggle: {
    display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
    width: "100%", padding: "9px", background: "transparent", border: `1.5px dashed ${T.line}`,
    borderRadius: 999, color: T.inkSoft, fontSize: 12.5, fontWeight: 700, marginBottom: 14,
  },
  statsGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 18 },
  statCard: { background: "#fff", border: "1.5px solid", borderRadius: 14, padding: "12px 12px" },
  statCardValue: { fontFamily: "'Baloo 2',sans-serif", fontSize: 19, fontWeight: 700, lineHeight: 1.1 },
  statCardLabel: { fontSize: 11, color: T.inkSoft, marginTop: 3, fontWeight: 600 },
  weekChart: {
    display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 6,
    background: "#fff", border: `1.5px solid ${T.line}`, borderRadius: 16, padding: "14px 12px 10px",
    marginBottom: 18, height: 110,
  },
  weekBarWrap: { display: "flex", flexDirection: "column", alignItems: "center", gap: 6, flex: 1 },
  weekBar: { width: "100%", maxWidth: 20, background: T.yolk, borderRadius: 4 },
  weekLabel: { fontSize: 9.5, color: T.inkSoft, fontWeight: 600 },
};
