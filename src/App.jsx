import React, { useEffect, useMemo, useRef, useState } from "react";
import { Download } from "lucide-react";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://prqrgcstfboexfehpmnt.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBycXJnY3N0ZmJvZXhmZWhwbW50Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY0MjQyODIsImV4cCI6MjA5MjAwMDI4Mn0.cnK1AzYvH6jtvhLAnb8QWhz29j-6Y-d6wC6SgDOpQvg";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
const DAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"];

function normalizePhotoPath(photoUrl = "") {
  if (!photoUrl) return "";
  if (!photoUrl.startsWith("http://") && !photoUrl.startsWith("https://")) return photoUrl;

  try {
    const url = new URL(photoUrl);
    const marker = "/feedback-images/";
    const idx = url.pathname.indexOf(marker);
    if (idx !== -1) return decodeURIComponent(url.pathname.slice(idx + marker.length));
  } catch {
    return photoUrl;
  }

  return photoUrl;
}

async function openPhoto(photoUrl = "") {
  if (!photoUrl) return;

  const normalized = normalizePhotoPath(photoUrl);

  if (normalized.startsWith("http://") || normalized.startsWith("https://")) {
    window.open(normalized, "_blank", "noopener,noreferrer");
    return;
  }

  const { data, error } = await supabase.storage
    .from("feedback-images")
    .createSignedUrl(normalized, 60);

  if (error || !data?.signedUrl) {
    alert("Foto konnte nicht geöffnet werden. Bitte Bucket/Freigabe prüfen.");
    return;
  }

  window.open(data.signedUrl, "_blank", "noopener,noreferrer");
}

function getISOWeek(date = new Date()) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) {
    target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
  }
  return 1 + Math.ceil((firstThursday - target) / 604800000);
}

function extractDishes(comment = "") {
  const result = {};
  const regex = /(Montag|Dienstag|Mittwoch|Donnerstag|Freitag):\s*([^]*?)(?=Montag:|Dienstag:|Mittwoch:|Donnerstag:|Freitag:|$)/g;

  let match;
  while ((match = regex.exec(comment)) !== null) {
    result[match[1]] = match[2].trim();
  }

  return result;
}

function extractFreeComment(comment = "") {
  const marker = "Gerichte:";
  const idx = comment.indexOf(marker);

  if (idx !== -1) return comment.slice(0, idx).trim();

  return comment
    .replace(/^Top\s*/i, "")
    .replace(
      /(Montag|Dienstag|Mittwoch|Donnerstag|Freitag):\s*([^]*?)(?=Montag:|Dienstag:|Mittwoch:|Donnerstag:|Freitag:|$)/g,
      ""
    )
    .trim();
}

export default function WochenfeedbackDashboard() {
  const [entries, setEntries] = useState([]);
  const [archiveWeeks, setArchiveWeeks] = useState([]);
  const [kw, setKw] = useState(String(getISOWeek()));
  const [selectedKita, setSelectedKita] = useState("Alle");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [selectedEntry, setSelectedEntry] = useState(null);
  const [seenIds, setSeenIds] = useState([]);
  const [showFeedbackList, setShowFeedbackList] = useState(false);
  const [selectedDayFilter, setSelectedDayFilter] = useState(null);
  const [csvPreview, setCsvPreview] = useState("");
  const [excelPreview, setExcelPreview] = useState("");
  const [excelRowsPreview, setExcelRowsPreview] = useState([]);
  const excelTextRef = useRef(null);

  useEffect(() => {
    let active = true;

    async function loadArchiveWeeks() {
      const { data } = await supabase
        .from("feedback_entries")
        .select("kw")
        .order("kw", { ascending: false });

      if (!active) return;

      const uniqueWeeks = Array.from(new Set((data || []).map((item) => String(item.kw)).filter(Boolean)));
      setArchiveWeeks(uniqueWeeks);
    }

    loadArchiveWeeks();

    async function load() {
      setLoading(true);
      setError("");

      const { data, error } = await supabase
        .from("feedback_entries")
        .select("id, kita_name, ratings, comment, photo_url, kw")
        .eq("kw", Number(kw))
        .order("id", { ascending: false });

      if (!active) return;

      if (error) {
        setError(error.message);
        setEntries([]);
      } else {
        setEntries(data || []);
      }

      setLoading(false);
    }

    if (kw) load();
    return () => {
      active = false;
    };
  }, [kw]);

  useEffect(() => {
    const interval = setInterval(() => {
      const currentWeek = String(getISOWeek());
      setKw((prev) => (prev !== currentWeek ? currentWeek : prev));
    }, 60 * 60 * 1000);

    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    const storageKey = `wochenfeedback-seen-${kw}`;
    const raw = window.localStorage.getItem(storageKey);
    if (!raw) {
      setSeenIds([]);
      return;
    }
    try {
      setSeenIds(JSON.parse(raw));
    } catch {
      setSeenIds([]);
    }
  }, [kw]);

  const kitas = useMemo(() => {
    const names = Array.from(new Set(entries.map((e) => e.kita_name).filter(Boolean)));
    return ["Alle", ...names.sort((a, b) => a.localeCompare(b))];
  }, [entries]);

  const filteredEntries = useMemo(() => {
    if (selectedKita === "Alle") return entries;
    return entries.filter((entry) => entry.kita_name === selectedKita);
  }, [entries, selectedKita]);

  const unreadEntries = useMemo(() => {
    return filteredEntries.filter((entry) => !seenIds.includes(entry.id));
  }, [filteredEntries, seenIds]);

  const displayedUnreadEntries = useMemo(() => {
    if (!selectedDayFilter) return unreadEntries;
    return unreadEntries.filter((entry) => entry.ratings?.[selectedDayFilter] === "down");
  }, [unreadEntries, selectedDayFilter]);

  const summary = useMemo(() => {
    const dayStats = Object.fromEntries(DAYS.map((day) => [day, { up: 0, down: 0 }]));
    let upTotal = 0;
    let downTotal = 0;

    filteredEntries.forEach((entry) => {
      DAYS.forEach((day) => {
        const value = entry.ratings?.[day];
        if (value === "up") {
          dayStats[day].up += 1;
          upTotal += 1;
        }
        if (value === "down") {
          dayStats[day].down += 1;
          downTotal += 1;
        }
      });
    });

    const totalRatings = upTotal + downTotal;
    const positiveRate = totalRatings > 0 ? Math.round((upTotal / totalRatings) * 1000) / 10 : 0;

    let signalLabel = "Noch keine Daten";
    let signalClasses = "bg-slate-100 text-slate-700 border-slate-200";

    if (totalRatings > 0) {
      if (positiveRate >= 80) {
        signalLabel = "Sehr gute Woche";
        signalClasses = "bg-emerald-100 text-emerald-800 border-emerald-200";
      } else if (positiveRate >= 60) {
        signalLabel = "Gemischte Woche";
        signalClasses = "bg-amber-100 text-amber-800 border-amber-200";
      } else {
        signalLabel = "Auffällige Woche";
        signalClasses = "bg-red-100 text-red-700 border-red-200";
      }
    }

    return {
      upTotal,
      downTotal,
      positiveRate,
      dayStats,
      signalLabel,
      signalClasses,
    };
  }, [filteredEntries]);

  const markEntriesAsSeen = (ids) => {
    const uniqueNewIds = ids.filter((id) => !seenIds.includes(id));
    if (uniqueNewIds.length === 0) return;
    const updated = [...uniqueNewIds, ...seenIds];
    setSeenIds(updated);
    window.localStorage.setItem(`wochenfeedback-seen-${kw}`, JSON.stringify(updated));
  };

  const openEntry = (entry) => {
    setSelectedEntry(entry);
    markEntriesAsSeen([entry.id]);
  };

  const toggleUnreadList = () => {
    if (unreadEntries.length === 0) return;
    setSelectedDayFilter(null);
    setShowFeedbackList((prev) => !prev);
  };

  const openDayNegatives = (day) => {
    const unreadNegativeIds = unreadEntries
      .filter((entry) => entry.ratings?.[day] === "down")
      .map((entry) => entry.id);

    if (unreadNegativeIds.length === 0) return;

    setSelectedDayFilter(day);
    setShowFeedbackList(true);
  };

  const buildCsv = () => {
    const rows = filteredEntries.map((entry) => ({
      kita_name: entry.kita_name || "",
      kw: entry.kw || "",
      montag: entry.ratings?.Montag || "",
      dienstag: entry.ratings?.Dienstag || "",
      mittwoch: entry.ratings?.Mittwoch || "",
      donnerstag: entry.ratings?.Donnerstag || "",
      freitag: entry.ratings?.Freitag || "",
      comment: extractFreeComment(entry.comment || ""),
      gerichte: Object.entries(extractDishes(entry.comment || ""))
        .map(([day, dish]) => `${day}: ${dish}`)
        .join(" | "),
      photo_url: entry.photo_url || "",
    }));

    const headers = [
      "kita_name",
      "kw",
      "montag",
      "dienstag",
      "mittwoch",
      "donnerstag",
      "freitag",
      "comment",
      "gerichte",
      "photo_url",
    ];

    return [
      headers.join(","),
      ...rows.map((row) =>
        headers
          .map((header) => {
            const value = String(row[header] ?? "").replace(/"/g, '""');
            return `"${value}"`;
          })
          .join(",")
      ),
    ].join("\n");
  };

  const buildExcelRows = () => {
    return filteredEntries.map((entry) => {
      const dishes = extractDishes(entry.comment || "");
      return {
        einrichtung: entry.kita_name || "",
        kw: entry.kw || "",
        mo_bewertung: entry.ratings?.Montag || "",
        mo_gericht: dishes.Montag || "",
        di_bewertung: entry.ratings?.Dienstag || "",
        di_gericht: dishes.Dienstag || "",
        mi_bewertung: entry.ratings?.Mittwoch || "",
        mi_gericht: dishes.Mittwoch || "",
        do_bewertung: entry.ratings?.Donnerstag || "",
        do_gericht: dishes.Donnerstag || "",
        fr_bewertung: entry.ratings?.Freitag || "",
        fr_gericht: dishes.Freitag || "",
        kommentar: extractFreeComment(entry.comment || ""),
        bild: entry.photo_url ? "Ja" : "Nein",
      };
    });
  };

  const buildExcelText = () => {
    const rows = buildExcelRows();

    const headers = [
      "Einrichtung",
      "KW",
      "Mo Bewertung",
      "Mo Gericht",
      "Di Bewertung",
      "Di Gericht",
      "Mi Bewertung",
      "Mi Gericht",
      "Do Bewertung",
      "Do Gericht",
      "Fr Bewertung",
      "Fr Gericht",
      "Kommentar",
      "Bild vorhanden",
    ];

    return [
      headers,
      ...rows.map((row) => [
        row.einrichtung,
        row.kw,
        row.mo_bewertung,
        row.mo_gericht,
        row.di_bewertung,
        row.di_gericht,
        row.mi_bewertung,
        row.mi_gericht,
        row.do_bewertung,
        row.do_gericht,
        row.fr_bewertung,
        row.fr_gericht,
        row.kommentar,
        row.bild,
      ]),
    ]
      .map((row) =>
        row
          .map((cell) => String(cell ?? "").replace(/\t/g, " ").replace(/\n/g, " "))
          .join("\t")
      )
      .join("\n");
  };

  const downloadCsv = () => {
    alert("Direkter Download wird in dieser Vorschau blockiert. Bitte 'Excel-Inhalt kopieren' nutzen.");
  };

  const copyCsv = async () => {
    try {
      await navigator.clipboard.writeText(csvPreview || buildCsv());
      alert("CSV in Zwischenablage kopiert");
    } catch {
      alert("Kopieren hat nicht funktioniert");
    }
  };

  const copyExcel = async () => {
    const text = excelPreview || buildExcelText();

    try {
      // moderner Weg
      await navigator.clipboard.writeText(text);
      alert("Für Excel kopiert. Jetzt Excel öffnen, Zelle A1 anklicken und STRG+V drücken.");
    } catch (err) {
      // Fallback für blockierte Clipboard APIs
      try {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.focus();
        textarea.select();

        document.execCommand("copy");
        document.body.removeChild(textarea);

        alert("Für Excel kopiert (Fallback). Jetzt Excel öffnen und STRG+V drücken.");
      } catch {
        alert("Kopieren wurde vom Browser blockiert. Bitte manuell markieren und kopieren.");
      }
    }
  };

  const exportToCsv = () => {
    if (filteredEntries.length === 0) {
      alert("Keine Daten zum Exportieren vorhanden");
      return;
    }
    setCsvPreview(buildCsv());
    setExcelPreview(buildExcelText());
    setExcelRowsPreview(buildExcelRows());
  };

  const selectExcelText = () => {
    if (!excelTextRef.current) return;
    excelTextRef.current.focus();
    excelTextRef.current.select();
    alert("Excel-Text ist markiert. Jetzt STRG+C drücken und dann in Excel in A1 einfügen.");
  };

  return (
    <div className="min-h-screen bg-emerald-50 px-4 py-8">
      <div className="mx-auto max-w-6xl space-y-6">
        <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-emerald-100">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-sm font-medium uppercase tracking-wide text-emerald-700">Interne Auswertung</p>
              <h1 className="mt-1 text-3xl font-bold text-slate-900">Wochenfeedback Dashboard</h1>
              <p className="mt-2 text-slate-500">Alle Rückmeldungen auf einen Blick, mit klarer Ampellogik und ruhigem Überblick.</p>
            </div>

            <div className="flex flex-col gap-3 sm:flex-row sm:items-end">
              <label className="rounded-2xl bg-emerald-50 p-3 text-sm font-medium text-slate-700">
                Kalenderwoche
                <input
                  type="number"
                  min="1"
                  max="53"
                  value={kw}
                  onChange={(e) => setKw(e.target.value)}
                  className="mt-2 w-full rounded-xl border border-emerald-200 bg-white px-3 py-2 outline-none focus:border-emerald-400"
                />
              </label>

              <label className="rounded-2xl bg-emerald-50 p-3 text-sm font-medium text-slate-700">
                Einrichtung
                <select
                  value={selectedKita}
                  onChange={(e) => setSelectedKita(e.target.value)}
                  className="mt-2 w-full rounded-xl border border-emerald-200 bg-white px-3 py-2 outline-none focus:border-emerald-400"
                >
                  {kitas.map((kita) => (
                    <option key={kita} value={kita}>
                      {kita}
                    </option>
                  ))}
                </select>
              </label>

              <button
                type="button"
                onClick={exportToCsv}
                className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-3 text-sm font-medium text-white hover:bg-slate-800"
              >
                <Download className="h-4 w-4" />
                Export Excel
              </button>
            </div>
          </div>
        </div>

        {error && (
          <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-red-700">
            Fehler beim Laden: {error}
          </div>
        )}

        <div className="rounded-3xl bg-white p-4 shadow-sm ring-1 ring-emerald-100">
          <div className="flex flex-wrap items-center gap-3">
            <span className="text-sm font-medium text-slate-500">Archiv</span>

            {/* letzte 5 Wochen als Buttons */}
            {archiveWeeks.slice(0, 5).map((week) => (
              <button
                key={week}
                type="button"
                onClick={() => setKw(String(week))}
                className={`rounded-full px-3 py-1 text-sm font-medium transition ${String(kw) === String(week) ? "bg-emerald-600 text-white" : "bg-emerald-50 text-emerald-700 hover:bg-emerald-100"}`}
              >
                KW {week}
              </button>
            ))}

            {/* Dropdown für alle weiteren */}
            {archiveWeeks.length > 5 && (
              <select
                onChange={(e) => setKw(e.target.value)}
                className="rounded-full border border-emerald-200 bg-white px-3 py-1 text-sm text-emerald-700"
                value=""
              >
                <option value="">Archiv ▼</option>
                {archiveWeeks.slice(5).map((week) => (
                  <option key={week} value={week}>
                    KW {week}
                  </option>
                ))}
              </select>
            )}
          </div>
        </div>

        <div className="grid gap-4 md:grid-cols-5">
          <button
            type="button"
            onClick={toggleUnreadList}
            className="rounded-3xl bg-white p-5 text-left shadow-sm ring-1 ring-emerald-100 transition hover:bg-emerald-50"
          >
            <div className="flex items-center justify-between gap-3">
              <div className="text-sm text-slate-500">Rückmeldungen</div>
              {unreadEntries.length > 0 && (
                <div className="rounded-full bg-red-100 px-2 py-1 text-xs font-medium text-red-600">
                  Neu: {unreadEntries.length}
                </div>
              )}
            </div>
            <div className="mt-2 text-3xl font-bold text-slate-900">{filteredEntries.length}</div>
            <div className="mt-3 text-sm font-medium text-emerald-700">
              {showFeedbackList && !selectedDayFilter ? "Neue Rückmeldungen ausblenden" : "Neue Rückmeldungen anzeigen"}
            </div>
          </button>

          <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
            <div className="text-sm text-slate-500">Gesamt 👍</div>
            <div className="mt-2 text-3xl font-bold text-emerald-600">{summary.upTotal}</div>
          </div>

          <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
            <div className="text-sm text-slate-500">Gesamt 👎</div>
            <div className="mt-2 text-3xl font-bold text-red-500">{summary.downTotal}</div>
          </div>

          <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
            <div className="text-sm text-slate-500">Positivquote</div>
            <div className="mt-2 text-3xl font-bold text-slate-900">{summary.positiveRate}%</div>
          </div>

          <div className={`rounded-3xl border p-5 shadow-sm ${summary.signalClasses}`}>
            <div className="text-sm">Wochenampel</div>
            <div className="mt-2 text-2xl font-bold">{summary.signalLabel}</div>
          </div>
        </div>

        <div className="grid gap-4 md:grid-cols-5">
          {DAYS.map((day) => {
            const stats = summary.dayStats[day];
            const unreadNegativeCount = unreadEntries.filter((entry) => entry.ratings?.[day] === "down").length;
            const isAlert = unreadNegativeCount > 0;

            return (
              <button
                key={day}
                type="button"
                onClick={() => openDayNegatives(day)}
                className={`rounded-3xl p-5 text-left shadow-sm ring-1 transition ${isAlert ? "bg-red-100 ring-red-200 hover:bg-red-200" : "bg-white ring-emerald-100"}`}
              >
                <div className="flex items-center justify-between gap-3">
                  <div className="text-sm text-slate-500">{day}</div>
                  {isAlert && (
                    <div className="rounded-full bg-red-200 px-2 py-1 text-xs font-medium text-red-700">
                      Neu
                    </div>
                  )}
                </div>
                <div className="mt-2 text-lg font-bold text-slate-900">
                  👍 {stats.up} | 👎 {stats.down}
                </div>
              </button>
            );
          })}
        </div>

        {showFeedbackList && displayedUnreadEntries.length > 0 && (
          <div className="rounded-3xl bg-white p-6 shadow-sm ring-1 ring-emerald-100">
            <div className="flex items-center justify-between gap-3">
              <h2 className="text-xl font-bold text-slate-900">
                {selectedDayFilter ? `Neue Negativmeldungen: ${selectedDayFilter}` : "Neue Rückmeldungen"}
              </h2>
              <div className="flex items-center gap-3">
                <div className="text-sm text-slate-500">{displayedUnreadEntries.length} ungelesen</div>
                <button
                  type="button"
                  onClick={() => {
                    if (selectedDayFilter) {
                      const idsToMark = unreadEntries
                        .filter((entry) => entry.ratings?.[selectedDayFilter] === "down")
                        .map((entry) => entry.id);
                      markEntriesAsSeen(idsToMark);
                    }
                    setShowFeedbackList(false);
                    setSelectedDayFilter(null);
                  }}
                  className="rounded-xl bg-slate-100 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-200"
                >
                  Schließen
                </button>
              </div>
            </div>

            <div className="mt-4 space-y-3">
              {displayedUnreadEntries.map((entry) => {
                const freeComment = extractFreeComment(entry.comment || "");
                const dishes = extractDishes(entry.comment || "");
                const previewDish = DAYS.map((day) => dishes[day]).find(Boolean);

                return (
                  <div key={entry.id} className="rounded-2xl border border-emerald-100 bg-slate-50 px-4 py-4">
                    <div className="flex items-start justify-between gap-4">
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="font-semibold text-slate-900">{entry.kita_name || "Ohne Name"}</span>
                          <span className="rounded-full bg-emerald-100 px-2 py-0.5 text-xs font-medium text-emerald-700">KW {entry.kw}</span>
                          <span className="rounded-full bg-red-100 px-2 py-0.5 text-xs font-medium text-red-600">Neu</span>
                        </div>
                        <div className="mt-1 truncate text-sm text-slate-500">
                          {freeComment || previewDish || "Öffnen für Details"}
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => openEntry(entry)}
                        className="shrink-0 rounded-xl bg-slate-900 px-4 py-2 text-sm font-medium text-white hover:bg-slate-800"
                      >
                        Öffnen
                      </button>
                    </div>

                    <div className="mt-3 flex flex-wrap gap-2">
                      {DAYS.map((day) => {
                        const value = entry.ratings?.[day];
                        return (
                          <span
                            key={day}
                            className="rounded-full bg-white px-3 py-1 text-xs font-medium text-slate-700 ring-1 ring-emerald-100"
                          >
                            {day}: {value === "up" ? "👍" : value === "down" ? "👎" : "-"}
                          </span>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {csvPreview && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
            <div className="w-full max-w-5xl rounded-3xl bg-white p-6 shadow-2xl">
              <div className="flex items-start justify-between gap-4">
                <div>
                  <h3 className="text-2xl font-bold text-slate-900">Excel Export</h3>
                  <p className="mt-1 text-sm text-slate-500">Sauber vorbereitet für Excel. Am besten: In Excel kopieren, dann in Excel auf A1 klicken und einfügen.</p>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setCsvPreview("");
                    setExcelPreview("");
                    setExcelRowsPreview([]);
                  }}
                  className="rounded-xl bg-slate-100 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-200"
                >
                  Schließen
                </button>
              </div>

              <div className="mt-4 overflow-hidden rounded-2xl border border-slate-200">
                <div className="max-h-80 overflow-auto bg-white">
                  <table className="min-w-full text-sm">
                    <thead className="sticky top-0 bg-slate-50">
                      <tr className="text-left text-slate-700">
                        <th className="px-3 py-2 font-semibold">Einrichtung</th>
                        <th className="px-3 py-2 font-semibold">KW</th>
                        <th className="px-3 py-2 font-semibold">Mo</th>
                        <th className="px-3 py-2 font-semibold">Di</th>
                        <th className="px-3 py-2 font-semibold">Mi</th>
                        <th className="px-3 py-2 font-semibold">Do</th>
                        <th className="px-3 py-2 font-semibold">Fr</th>
                        <th className="px-3 py-2 font-semibold">Kommentar</th>
                        <th className="px-3 py-2 font-semibold">Bild</th>
                      </tr>
                    </thead>
                    <tbody>
                      {excelRowsPreview.map((row, idx) => (
                        <tr key={idx} className="border-t border-slate-100 text-slate-700">
                          <td className="px-3 py-2">{row.einrichtung}</td>
                          <td className="px-3 py-2">{row.kw}</td>
                          <td className="px-3 py-2">{row.mo_bewertung} {row.mo_gericht}</td>
                          <td className="px-3 py-2">{row.di_bewertung} {row.di_gericht}</td>
                          <td className="px-3 py-2">{row.mi_bewertung} {row.mi_gericht}</td>
                          <td className="px-3 py-2">{row.do_bewertung} {row.do_gericht}</td>
                          <td className="px-3 py-2">{row.fr_bewertung} {row.fr_gericht}</td>
                          <td className="px-3 py-2">{row.kommentar}</td>
                          <td className="px-3 py-2">{row.bild}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="mt-4 rounded-2xl border border-slate-200 bg-slate-50 p-3">
                <div className="mb-2 text-sm font-medium text-slate-700">Excel-Text für manuelles Kopieren</div>
                <textarea
                  ref={excelTextRef}
                  readOnly
                  value={excelPreview}
                  className="h-40 w-full rounded-xl border border-slate-200 bg-white p-3 font-mono text-xs outline-none"
                />
              </div>

              <div className="mt-4 flex flex-wrap gap-3">
                <button
                  type="button"
                  onClick={downloadCsv}
                  className="rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white hover:bg-slate-800"
                >
                  Download
                </button>
                <button
                  type="button"
                  onClick={selectExcelText}
                  className="rounded-xl bg-emerald-600 px-4 py-3 text-sm font-medium text-white hover:bg-emerald-700"
                >
                  Excel-Text markieren
                </button>
                <button
                  type="button"
                  onClick={copyExcel}
                  className="rounded-xl bg-emerald-100 px-4 py-3 text-sm font-medium text-emerald-800 hover:bg-emerald-200"
                >
                  In Excel kopieren
                </button>
              </div>
            </div>
          </div>
        )}

        {selectedEntry && (
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4">
            <div className="max-h-[90vh] w-full max-w-4xl overflow-y-auto rounded-3xl bg-white p-6 shadow-2xl">
              <div className="flex items-start justify-between gap-4">
                <div>
                  <h3 className="text-2xl font-bold text-slate-900">{selectedEntry.kita_name || "Ohne Name"}</h3>
                  <p className="mt-1 text-sm text-slate-500">KW {selectedEntry.kw}</p>
                </div>
                <button
                  type="button"
                  onClick={() => setSelectedEntry(null)}
                  className="rounded-xl bg-slate-100 px-3 py-2 text-sm font-medium text-slate-700 hover:bg-slate-200"
                >
                  Schließen
                </button>
              </div>

              <div className="mt-6 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                {DAYS.map((day) => {
                  const value = selectedEntry.ratings?.[day];
                  const dishes = extractDishes(selectedEntry.comment || "");
                  const dish = dishes[day];
                  return (
                    <div key={day} className="rounded-2xl bg-slate-50 px-4 py-3 ring-1 ring-slate-100">
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-medium text-slate-900">{day}</span>
                        <span>{value === "up" ? "👍" : value === "down" ? "👎" : "-"}</span>
                      </div>
                      <div className="mt-1 text-slate-500">{dish || "Kein Gericht"}</div>
                    </div>
                  );
                })}
              </div>

              {extractFreeComment(selectedEntry.comment || "") && (
                <div className="mt-4 rounded-2xl bg-slate-50 p-4 text-slate-700 ring-1 ring-slate-100">
                  {extractFreeComment(selectedEntry.comment || "")}
                </div>
              )}

              {selectedEntry.photo_url && (
                <div className="mt-4">
                  <button
                    type="button"
                    onClick={() => openPhoto(selectedEntry.photo_url)}
                    className="inline-flex items-center rounded-xl bg-slate-900 px-4 py-3 text-sm font-medium text-white hover:bg-slate-800"
                  >
                    Foto öffnen
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
