import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { toPng } from "html-to-image";
import pptxgen from "pptxgenjs";
import data from "./data/requests.json";
import "./styles.css";

const PHASES = [
  { label: "Foreslåtte initiativ", tone: "proposed" },
  { label: "Inntakskø", tone: "intake" },
  { label: "Løsningsdesign", tone: "design" },
  { label: "Tilbud", tone: "offer" },
  { label: "Gjennomføring", tone: "execution" },
  { label: "Avslutning", tone: "closing" },
];

const ALL_PHASES = [...PHASES, { label: "Uklassifisert", tone: "unclassified" }];

const CARDS_PER_COLUMN = 25;
const CARD_H = 0.185;
const CARD_GAP = 0.025;
const CARD_FONT = 6.5;

const SLIDE_COLORS = {
  proposed: { fill: "DDEFC8", accent: "608E36" },
  intake: { fill: "F8E5BF", accent: "B16B16" },
  design: { fill: "DCEBE7", accent: "37766B" },
  offer: { fill: "E8E3D4", accent: "8A7246" },
  execution: { fill: "DCE2EF", accent: "4F6B9D" },
  closing: { fill: "E4DED8", accent: "7C675B" },
};

function groupBy(items, key) {
  return items.reduce((acc, item) => {
    const value = item[key] || "Uten eier";
    acc[value] ??= [];
    acc[value].push(item);
    return acc;
  }, {});
}

function countByPhase(items) {
  return ALL_PHASES.map((phase) => ({
    ...phase,
    count: items.filter((item) => item.derivedPhase === phase.label).length,
  }));
}

function safeFileName(value) {
  return String(value)
    .normalize("NFKD")
    .replace(/[^\w.-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80)
    .toLowerCase();
}

function downloadDataUrl(dataUrl, fileName) {
  const link = document.createElement("a");
  link.download = fileName;
  link.href = dataUrl;
  link.click();
}

function applyPiFilter(items, selectedPIs) {
  if (selectedPIs.size === 0) return items;
  const groupHasPI = items.some((i) => i.pi);
  if (!groupHasPI) return items;
  return items.filter((i) => selectedPIs.has(i.pi));
}

function App() {
  const [selectedGroup, setSelectedGroup] = useState("");
  const [isExporting, setIsExporting] = useState(false);
  const [selectedPIs, setSelectedPIs] = useState(new Set());
  const [piFilterOpen, setPiFilterOpen] = useState(false);
  const boardRefs = useRef(new Map());
  const piDropdownRef = useRef(null);

  useEffect(() => {
    if (!piFilterOpen) return;
    function handleClickOutside(e) {
      if (!piDropdownRef.current?.contains(e.target)) setPiFilterOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [piFilterOpen]);

  const requests = data.requests ?? [];
  const groups = useMemo(() => {
    const grouped = groupBy(requests, "assignmentGroup");
    return Object.entries(grouped)
      .map(([name, items]) => ({
        name,
        items: [...items].sort((a, b) => a.number.localeCompare(b.number, "nb")),
      }))
      .sort((a, b) => b.items.length - a.items.length || a.name.localeCompare(b.name, "nb"));
  }, [requests]);

  const activeGroupName = selectedGroup || groups[0]?.name || "";
  const activeGroup = groups.find((group) => group.name === activeGroupName) ?? groups[0];
  const totals = countByPhase(requests);

  const piOptions = useMemo(() => {
    const values = [...new Set((activeGroup?.items ?? []).map((i) => i.pi).filter(Boolean))].sort();
    return values;
  }, [activeGroup]);

  const filteredActiveCount = activeGroup
    ? applyPiFilter(activeGroup.items, selectedPIs).length
    : 0;

  function handleSelectGroup(name) {
    setSelectedGroup(name);
    setSelectedPIs(new Set());
    setPiFilterOpen(false);
  }

  function togglePI(pi) {
    setSelectedPIs((prev) => {
      const next = new Set(prev);
      next.has(pi) ? next.delete(pi) : next.add(pi);
      return next;
    });
  }

  function setBoardRef(groupName, node) {
    if (!node) {
      boardRefs.current.delete(groupName);
      return;
    }
    boardRefs.current.set(groupName, node);
  }

  async function exportNodeAsPng(node, fileName) {
    await document.fonts?.ready;
    const dataUrl = await toPng(node, {
      cacheBust: true,
      pixelRatio: 2,
      backgroundColor: "#edf2e7",
      style: { margin: "0" },
    });
    downloadDataUrl(dataUrl, fileName);
  }

  async function exportSelectedPng() {
    if (!activeGroup) return;
    setIsExporting(true);
    try {
      const node = boardRefs.current.get(activeGroup.name);
      await exportNodeAsPng(node, `nye-tjenester-${safeFileName(activeGroup.name)}.png`);
    } finally {
      setIsExporting(false);
    }
  }

  async function exportAllPng() {
    setIsExporting(true);
    try {
      await new Promise((resolve) => window.setTimeout(resolve, 50));
      for (const group of groups) {
        const node = boardRefs.current.get(group.name);
        if (node) {
          await exportNodeAsPng(node, `nye-tjenester-${safeFileName(group.name)}.png`);
        }
      }
    } finally {
      setIsExporting(false);
    }
  }

  async function exportPowerPoint() {
    setIsExporting(true);
    try {
      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_WIDE";
      pptx.author = "Helse Nord IKT";
      pptx.subject = "Utviklingsradar kliniske og pasientrettede systemer";
      pptx.title = "Utviklingsradar kliniske og pasientrettede systemer";
      pptx.company = "Helse Nord IKT";
      pptx.lang = "nb-NO";
      pptx.theme = {
        headFontFace: "Aptos Display",
        bodyFontFace: "Aptos",
        lang: "nb-NO",
      };

      const piFilter = [...selectedPIs];
      for (const group of groups) {
        const exportGroup =
          selectedPIs.size > 0
            ? { ...group, items: applyPiFilter(group.items, selectedPIs) }
            : group;
        addGroupSlides(pptx, exportGroup, data.metadata, piFilter);
      }

      await pptx.writeFile({ fileName: "nye-tjenester-boards.pptx" });
    } finally {
      setIsExporting(false);
    }
  }

  if (requests.length === 0) {
    return (
      <main className="empty-state">
        <h1>Utviklingsradar kliniske og pasientrettede systemer</h1>
        <p>Ingen importerte saker funnet. Kjør `npm run import:data` med en .xlsx-fil i `input/` eller prosjektroten.</p>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <div className="ambient-grid" aria-hidden="true" />
      <header className="hero">
        <div>
          <p className="eyebrow">Helse Nord IKT</p>
          <h1>Utviklingsradar kliniske og pasientrettede systemer</h1>
          <p className="hero-copy">
            {data.metadata.total} saker fra {data.metadata.sourceFile}. Velg eiergruppe for å se
            boardet og eksporter visningen til presentasjon.
          </p>
        </div>
        <div className="action-panel">
          <button type="button" onClick={exportSelectedPng} disabled={isExporting}>
            Eksporter valgt PNG
          </button>
          <button type="button" onClick={exportAllPng} disabled={isExporting}>
            Eksporter alle PNG
          </button>
          <button type="button" className="primary-action" onClick={exportPowerPoint} disabled={isExporting}>
            Lag PowerPoint
          </button>
        </div>
      </header>

      <section className="summary-strip" aria-label="Total fasefordeling">
        {totals.map((phase) => (
          <div className={`summary-card ${phase.tone}`} key={phase.label}>
            <span>{phase.label}</span>
            <strong>{phase.count}</strong>
          </div>
        ))}
      </section>

      <BoardTabs groups={groups} activeGroupName={activeGroupName} onSelect={handleSelectGroup} />

      {piOptions.length > 0 && (
        <div className="pi-filter-bar">
          <div className="pi-filter-dropdown" ref={piDropdownRef}>
            <button
              type="button"
              className={`pi-filter-trigger ${selectedPIs.size > 0 ? "active" : ""}`}
              onClick={() => setPiFilterOpen((o) => !o)}
            >
              <span>
                {selectedPIs.size === 0 ? "Filtrer på PI" : `${selectedPIs.size} PI valgt`}
              </span>
              <strong>{filteredActiveCount}</strong>
            </button>
            {piFilterOpen && (
              <div className="pi-filter-panel">
                {selectedPIs.size > 0 && (
                  <button
                    type="button"
                    className="pi-filter-action"
                    onClick={() => setSelectedPIs(new Set())}
                  >
                    Nullstill filter
                  </button>
                )}
                {piOptions.map((pi) => {
                  const count = (activeGroup?.items ?? []).filter((i) => i.pi === pi).length;
                  return (
                    <label key={pi} className="pi-filter-option">
                      <input
                        type="checkbox"
                        checked={selectedPIs.has(pi)}
                        onChange={() => togglePI(pi)}
                      />
                      <span>{pi}</span>
                      <strong>{count}</strong>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      )}

      {activeGroup && (
        <PhaseBoard
          group={activeGroup}
          metadata={data.metadata}
          refCallback={(node) => setBoardRef(activeGroup.name, node)}
          selectedPIs={selectedPIs}
          visible
        />
      )}

      <div className="export-stage" aria-hidden="true">
        {groups
          .filter((group) => group.name !== activeGroupName)
          .map((group) => (
            <PhaseBoard
              key={group.name}
              group={group}
              metadata={data.metadata}
              refCallback={(node) => setBoardRef(group.name, node)}
              selectedPIs={selectedPIs}
            />
          ))}
      </div>
    </main>
  );
}

function BoardTabs({ groups, activeGroupName, onSelect }) {
  return (
    <nav className="board-tabs" aria-label="Assignment groups">
      {groups.map((group) => (
        <button
          type="button"
          key={group.name}
          className={group.name === activeGroupName ? "active" : ""}
          onClick={() => onSelect(group.name)}
        >
          <span>{group.name}</span>
          <strong>{group.items.length}</strong>
        </button>
      ))}
    </nav>
  );
}

function PhaseBoard({ group, metadata, refCallback, visible = false, selectedPIs = new Set() }) {
  const filteredItems = applyPiFilter(group.items, selectedPIs);
  const groupedByPhase = groupBy(filteredItems, "derivedPhase");
  const phaseCounts = countByPhase(filteredItems);
  const unclassified = groupedByPhase.Uklassifisert ?? [];
  const isFiltered = filteredItems.length !== group.items.length;

  return (
    <section className={`board-card ${visible ? "visible-board" : ""}`} ref={refCallback}>
      <div className="board-header">
        <div>
          <p className="eyebrow">Assignment group</p>
          <h2>{group.name}</h2>
        </div>
        <div className="board-meta">
          <span>
            {isFiltered
              ? `${filteredItems.length} av ${group.items.length} saker`
              : `${group.items.length} saker`}
          </span>
          <span>{metadata.sourceFile}</span>
        </div>
      </div>

      <div className="mini-distribution">
        {phaseCounts.map((phase) => (
          <span className={phase.tone} key={phase.label}>
            {phase.label}: <strong>{phase.count}</strong>
          </span>
        ))}
      </div>

      <div className="phase-grid">
        {PHASES.map((phase) => (
          <PhaseColumn
            key={phase.label}
            phase={phase}
            items={(groupedByPhase[phase.label] ?? []).sort((a, b) =>
              a.number.localeCompare(b.number, "nb")
            )}
          />
        ))}
      </div>

      {unclassified.length > 0 && (
        <div className="unclassified-note">
          {unclassified.length} sak(er) mangler kjent fase og er ikke plassert i hovedkolonnene.
        </div>
      )}
    </section>
  );
}

function PhaseColumn({ phase, items }) {
  return (
    <section className={`phase-column ${phase.tone}`}>
      <header>
        <span>{phase.label}</span>
        <strong>{items.length}</strong>
      </header>
      <div className="card-stack">
        {items.length === 0 ? (
          <p className="empty-column">Ingen saker</p>
        ) : (
          items.map((item) => <RequestCard key={item.id} item={item} />)
        )}
      </div>
    </section>
  );
}

function RequestCard({ item }) {
  return (
    <article className="request-card">
      <div className="request-topline">
        <strong>{item.number || "Uten nummer"}</strong>
        <span>{item.status || "Uten status"}</span>
      </div>
      {item.pi && <span className="pi-badge">{item.pi}</span>}
      <h3>{item.title || "Uten tittel"}</h3>
    </article>
  );
}

function paginateGroup(group) {
  const phaseItems = PHASES.map((phase) =>
    group.items
      .filter((item) => item.derivedPhase === phase.label)
      .sort((a, b) => a.number.localeCompare(b.number, "nb"))
  );
  const maxLen = Math.max(...phaseItems.map((items) => items.length), 0);
  const pageCount = Math.max(1, Math.ceil(maxLen / CARDS_PER_COLUMN));
  const pages = Array.from({ length: pageCount }, (_, p) =>
    phaseItems.map((items) => items.slice(p * CARDS_PER_COLUMN, (p + 1) * CARDS_PER_COLUMN))
  );
  return { pages, phaseItems, pageCount };
}

function renderPhaseColumn(slide, pptx, phase, items, x, startY, colW, colH, palette, opts) {
  const { pageCount, totalInPhase, rangeStart, rangeEnd, hasMoreBelow } = opts;

  slide.addShape(pptx.ShapeType.roundRect, {
    x, y: startY, w: colW, h: colH,
    rectRadius: 0.08,
    fill: { color: palette.fill, transparency: 8 },
    line: { color: "B8C2B0", transparency: 25 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x, y: startY, w: colW, h: 0.08,
    fill: { color: palette.accent },
    line: { color: palette.accent },
  });
  slide.addText(phase.label, {
    x: x + 0.08, y: startY + 0.13, w: colW - 0.72, h: 0.18,
    fontFace: "Aptos", fontSize: 7.5, bold: true, color: "1B2C20",
    breakLine: false, fit: "shrink",
  });

  const showRange = pageCount > 1 && totalInPhase > 0 && items.length > 0;
  const badgeText = showRange
    ? `${rangeStart + 1}–${rangeEnd} / ${totalInPhase}`
    : String(totalInPhase);
  const badgeW = showRange ? 0.6 : 0.24;
  slide.addText(badgeText, {
    x: x + colW - badgeW - 0.06,
    y: startY + 0.11,
    w: badgeW,
    h: 0.2,
    align: "right",
    fontFace: "Aptos",
    fontSize: showRange ? 6 : 8,
    bold: true,
    color: "1B2C20",
  });

  const cardsStartY = startY + 0.47;

  if (items.length === 0) {
    const emptyText = pageCount > 1 && totalInPhase > 0 ? "Ingen flere saker" : "Ingen saker";
    slide.addText(emptyText, {
      x: x + 0.1, y: cardsStartY, w: colW - 0.2, h: 0.22,
      fontFace: "Aptos", italic: true, fontSize: 7, color: "6D776C",
    });
    return;
  }

  items.forEach((item, itemIndex) => {
    const y = cardsStartY + itemIndex * (CARD_H + CARD_GAP);
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.07, y, w: colW - 0.14, h: CARD_H,
      rectRadius: 0.035,
      fill: { color: "FBFCF7", transparency: 0 },
      line: { color: "FFFFFF", transparency: 100 },
    });
    const cardText = item.pi
      ? `${item.number} · ${item.pi} · ${item.title}`
      : `${item.number} · ${item.title}`;
    slide.addText(cardText, {
      x: x + 0.12, y: y + 0.025, w: colW - 0.24, h: Math.max(CARD_H - 0.03, 0.12),
      fontFace: "Aptos", fontSize: CARD_FONT, color: "18241C",
      fit: "shrink", margin: 0, breakLine: false,
    });
  });

  if (hasMoreBelow) {
    const remaining = totalInPhase - rangeEnd;
    const indicatorY = cardsStartY + items.length * (CARD_H + CARD_GAP) + 0.02;
    slide.addText(`↓ +${remaining} på neste side`, {
      x: x + 0.07, y: indicatorY, w: colW - 0.14, h: 0.18,
      fontFace: "Aptos", fontSize: 6, color: palette.accent,
      italic: true, align: "center",
    });
  }
}

function addGroupSlides(pptx, group, metadata, piFilter = []) {
  const startX = 0.28;
  const startY = 1.15;
  const gap = 0.08;
  const colW = (12.78 - gap * 5) / 6;
  const colH = 5.95;

  const { pages, phaseItems, pageCount } = paginateGroup(group);

  pages.forEach((pagePhaseItems, pageIndex) => {
    const slide = pptx.addSlide();
    slide.background = { color: "EDF2E7" };

    slide.addText("Utviklingsradar", {
      x: 0.35, y: 0.18, w: 2.2, h: 0.24,
      fontFace: "Aptos", fontSize: 8, color: "586555", bold: true, charSpace: 1,
    });
    slide.addText(group.name, {
      x: 0.35, y: 0.45, w: 8.6, h: 0.45,
      fontFace: "Aptos Display", fontSize: 24, bold: true, color: "132117",
    });

    const pageIndicator = pageCount > 1 ? ` · side ${pageIndex + 1}/${pageCount}` : "";
    slide.addText(`${group.items.length} saker${pageIndicator} · ${metadata.sourceFile}`, {
      x: 9.2, y: 0.5, w: 3.6, h: 0.28,
      align: "right", fontFace: "Aptos", fontSize: 9, color: "586555",
    });

    const subtitleParts = [];
    if (pageIndex > 0) subtitleParts.push("Fortsettelse");
    if (piFilter.length > 0) subtitleParts.push(`PI: ${piFilter.join(", ")}`);
    if (subtitleParts.length > 0) {
      slide.addText(subtitleParts.join(" · "), {
        x: 0.35, y: 0.93, w: 8, h: 0.18,
        fontFace: "Aptos", fontSize: 9, color: "586555", italic: true,
      });
    }

    PHASES.forEach((phase, index) => {
      const items = pagePhaseItems[index];
      const totalInPhase = phaseItems[index].length;
      const rangeStart = pageIndex * CARDS_PER_COLUMN;
      const rangeEnd = rangeStart + items.length;
      const hasMoreBelow = rangeEnd < totalInPhase;
      const x = startX + index * (colW + gap);

      renderPhaseColumn(slide, pptx, phase, items, x, startY, colW, colH, SLIDE_COLORS[phase.tone], {
        pageIndex, pageCount, totalInPhase, rangeStart, rangeEnd, hasMoreBelow,
      });
    });
  });
}

createRoot(document.getElementById("root")).render(<App />);
