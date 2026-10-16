"use strict";

/* ---------- opmaak-helpers (Nederlands) ---------- */
const euro = new Intl.NumberFormat("nl-NL", { style: "currency", currency: "EUR" });
const geld = (n) => {
  if (typeof n !== "number") return "—";
  const positief = euro.format(Math.abs(n));
  return n < 0 ? `− ${positief}` : positief;
};
const datum = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y}`;
};
const datumLang = (iso) => {
  if (!iso) return "";
  return new Intl.DateTimeFormat("nl-NL", {
    day: "numeric", month: "long", year: "numeric",
  }).format(new Date(iso + "T12:00:00"));
};
const datumKort = (iso) => {
  if (!iso) return "";
  return new Intl.DateTimeFormat("nl-NL", {
    day: "numeric", month: "short",
  }).format(new Date(iso + "T12:00:00"));
};

const ICONEN = {
  herstel: '<path d="M4.5 9a8 8 0 1 1 1.3 7.8"/><path d="M4.5 4.5V9H9"/>',
  chevron: '<path d="m9 7 5 5-5 5"/>',
  check: '<path d="m5.5 12.5 4 4 9-9"/>',
  waarschuwing: '<path d="M12 4 3.5 19h17L12 4Z"/><path d="M12 9v4m0 3h.01"/>',
  zoek: '<circle cx="11" cy="11" r="6.5"/><path d="m16 16 4 4"/>',
  bon: '<path d="M7 3h10v18l-2-1.4L13 21l-2-1.4L9 21l-2-1.4V3Z"/><path d="M10 8h4m-4 4h4"/>',
};

function icoon(naam, klasse = "") {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  if (klasse) svg.setAttribute("class", klasse);
  svg.innerHTML = ICONEN[naam] || "";
  return svg;
}

/* Mini DOM-helper: el("div", {class:"x"}, [kind, ...]) */
function el(tag, props = {}, kinderen = []) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") n.className = v;
    else if (k === "text") n.textContent = v;
    else if (k === "html") n.innerHTML = v;
    else if (k.startsWith("aria-") || k === "role" || k === "scope" || k === "id" || k === "tabindex" || k === "hidden" || k === "colspan")
      { if (v !== false && v !== null) n.setAttribute(k, v === true ? "" : v); }
    else if (k.startsWith("data-")) n.setAttribute(k, v);
    else n[k] = v;
  }
  for (const kind of [].concat(kinderen)) {
    if (kind == null) continue;
    n.append(kind.nodeType ? kind : document.createTextNode(String(kind)));
  }
  return n;
}

function bedragCel(n, richting) {
  const klasse = richting === "inkomst" ? "bedrag-plus" : richting === "uitgave" ? "bedrag-min" : "";
  return el("td", { class: "getal " + klasse, text: geld(n) });
}

/* Generieke tabel uit kolomdefinities + rijen. */
function bouwTabel(kolommen, rijen, opties = {}) {
  const thead = el("thead", {}, el("tr", {}, kolommen.map((k) =>
    el("th", { scope: "col", class: k.getal ? "getal" : "" }, k.label))));
  const tbody = el("tbody", {}, rijen.map((r) =>
    el("tr", {}, kolommen.map((k) => (k.cel ? k.cel(r) :
      el("td", { class: k.getal ? "getal" : "" }, k.render ? k.render(r[k.key], r) : r[k.key]))))));
  const kinderen = [thead, tbody];
  if (opties.voet) kinderen.push(el("tfoot", {}, el("tr", {}, opties.voet)));
  const tabel = el("table", {}, kinderen);
  if (opties.caption) tabel.prepend(el("caption", { text: opties.caption }));
  return el("div", { class: "tabel-wikkel" }, tabel);
}

function uitleg(tekst) {
  return el("p", { class: "panel-uitleg", text: tekst });
}
function leegMelding(tekst) {
  return el("p", { class: "melding-leeg", text: tekst });
}

/* ---------- paneel-renderers ---------- */
const RICHTING_PIL = {
  inkomst: ["pil pil-inkomst", "Inkomst"],
  uitgave: ["pil pil-uitgave", "Uitgave"],
};

const transactieState = {
  zoek: "",
  filter: "alle",
  openId: null,
};

function getoondBedrag(r) {
  if (r.richting === "uitgave" && r.bedrag > 0) return -r.bedrag;
  return r.bedrag;
}

function maakBtwSelect(r, onthoudGetter = () => false) {
  const sel = el("select", {
    class: "btw-select", "data-id": r.id,
    "aria-label": `Btw-percentage voor ${r.naam}, nu ${r.btw_pct}%`,
  });
  const opties = [21, 9, 0];
  if (!opties.includes(r.btw_pct)) opties.push(r.btw_pct);
  opties.sort((a, b) => b - a);
  for (const p of opties) {
    const o = el("option", { value: String(p), text: p + "%" });
    if (p === r.btw_pct) o.selected = true;
    sel.append(o);
  }
  sel.addEventListener("change", () => bewerkBtw(r.id, sel.value, onthoudGetter()));
  return sel;
}

function btwCel(r) {
  const td = el("td", { class: "getal tx-btw" });
  td.append(maakBtwSelect(r));
  if (r.btw_bron === "override") {
    const h = el("button", {
      class: "herstel", type: "button", title: "Herstel automatische btw",
      "aria-label": `Herstel automatische btw voor ${r.naam}`,
    }, icoon("herstel"));
    h.addEventListener("click", () => herstelBtw(r.id));
    td.append(h, el("span", { class: "visueel-verborgen", text: "Aangepast" }));
  }
  return td;
}

function btwBronTekst(bron) {
  return {
    override: "Handmatig aangepast",
    geheugen: "Leveranciersvoorkeur",
    standaard: "Automatisch toegepast",
    patroon: "Herkenningsregel",
    verlegd: "Verlegde btw",
  }[bron] || "Automatisch toegepast";
}

function transactieKenmerken(r) {
  return [
    r.is_investering ? "Investering" : null,
    r.is_verlegd ? "Verlegde btw" : null,
    r.is_kredietfactuur ? "Kredietfactuur" : null,
    r.auto_nul ? "Automatisch 0%" : null,
  ].filter(Boolean);
}

function transactieDetailRij(r, kolommen) {
  const onthoud = el("input", { type: "checkbox", id: "onthoud-" + r.id });
  const btwVeld = maakBtwSelect(r, () => onthoud.checked);
  const acties = [
    el("label", { class: "detail-onthoud" }, [
      onthoud,
      `Gebruik deze keuze voortaan voor ${r.leverancier || "deze leverancier"}`,
    ]),
  ];
  if (r.btw_bron === "override") {
    const herstel = el("button", { class: "knop knop-compact", type: "button" }, [
      icoon("herstel"),
      "Automatische waarde herstellen",
    ]);
    herstel.addEventListener("click", () => herstelBtw(r.id));
    acties.push(herstel);
  }
  const kenmerken = transactieKenmerken(r);
  return el("tr", { class: "detail-rij", id: "detail-" + r.id }, el("td", { colspan: kolommen }, [
    el("div", { class: "transactie-detail" }, [
      el("div", { class: "detail-kop" }, [
        el("strong", { text: "Transactiedetails" }),
        el("span", { text: datumLang(r.datum) }),
      ]),
      el("dl", { class: "detail-grid" }, [
        el("div", {}, [el("dt", { text: "Tegenrekening" }), el("dd", { text: r.tegenrekening || "Niet beschikbaar" })]),
        el("div", {}, [el("dt", { text: "Classificatie" }), el("dd", { text: RICHTING_PIL[r.richting]?.[1] || r.richting })]),
        el("div", {}, [el("dt", { text: "Exclusief btw" }), el("dd", { text: geld(r.excl) })]),
        el("div", {}, [el("dt", { text: "Btw-instelling" }), el("dd", {}, btwVeld)]),
        el("div", {}, [el("dt", { text: "Herkomst btw" }), el("dd", { text: btwBronTekst(r.btw_bron) })]),
        el("div", { class: "detail-breed" }, [
          el("dt", { text: "Mededeling" }),
          el("dd", { text: r.mededelingen || "Niet beschikbaar" }),
        ]),
        el("div", { class: "detail-breed" }, [
          el("dt", { text: "Kenmerken" }),
          el("dd", { text: kenmerken.length ? kenmerken.join(" · ") : "Geen bijzonderheden" }),
        ]),
      ]),
      el("div", { class: "detail-acties" }, acties),
    ]),
  ]));
}

function bouwTransactieTabel(rijen) {
  if (!rijen.length) return leegMelding("Geen transacties gevonden met deze filters.");
  const labels = ["Datum", "Omschrijving", "Leverancier", "Factuurnummer", "Soort", "Bedrag incl.", "Btw", "Btw-bedrag"];
  const thead = el("thead", {}, el("tr", {}, labels.map((label, i) =>
    el("th", { scope: "col", class: i >= 5 ? "getal" : "" }, label))));
  const tbody = el("tbody");
  rijen.forEach((r) => {
    const open = transactieState.openId === r.id;
    const openKnop = el("button", {
      class: "transactie-open", type: "button",
      "aria-expanded": open ? "true" : "false",
      "aria-controls": open ? "detail-" + r.id : null,
      "data-detail-id": r.id,
    }, [
      el("span", { class: "tx-primair", text: r.naam }),
      el("span", { class: "tx-secundair", text: r.mededelingen || r.factuur || "Geen aanvullende omschrijving" }),
      icoon("chevron", "tx-chevron"),
    ]);
    openKnop.addEventListener("click", () => {
      transactieState.openId = open ? null : r.id;
      toonTab("transacties");
      requestAnimationFrame(() => {
        document.querySelector(`.transactie-open[data-detail-id="${r.id}"]`)?.focus();
      });
    });
    const soort = el("td", { class: "tx-soort", text: RICHTING_PIL[r.richting]?.[1] || r.richting });
    if (r.is_kredietfactuur) soort.append(el("span", { class: "tx-markering", text: "Krediet" }));
    const hoofd = el("tr", { class: "transactie-hoofd", "data-id": r.id }, [
      el("td", { class: "tx-datum", text: datum(r.datum) }),
      el("td", { class: "tx-omschrijving" }, openKnop),
      el("td", { class: "tx-leverancier", text: r.leverancier || "—" }),
      el("td", { class: "tx-factuur", text: r.factuur || "—" }),
      soort,
      el("td", { class: "getal tx-bedrag", text: geld(getoondBedrag(r)) }),
      btwCel(r),
      el("td", { class: "getal tx-btw-bedrag", text: geld(r.btw) }),
    ]);
    tbody.append(hoofd);
    if (open) tbody.append(transactieDetailRij(r, labels.length));
  });
  const tabel = el("table", { class: "transactie-tabel" }, [thead, tbody]);
  return el("div", { class: "tabel-wikkel transactie-wikkel" }, tabel);
}

function bouwNietIngedeeld(rijen) {
  if (!rijen.length) return leegMelding("Alle bankregels zijn ingedeeld.");
  return el("div", { class: "indelingslijst" }, rijen.map((r) => {
    const inkomst = el("button", { class: "knop knop-compact", type: "button", text: "Inkomst" });
    const uitgave = el("button", { class: "knop knop-compact", type: "button", text: "Uitgave" });
    inkomst.addEventListener("click", () => classificeerRichting(r.id, "inkomst"));
    uitgave.addEventListener("click", () => classificeerRichting(r.id, "uitgave"));
    return el("article", { class: "indelingsregel" }, [
      el("div", { class: "indelingsdatum", text: datum(r.datum) }),
      el("div", { class: "indelingsomschrijving" }, [
        el("strong", { text: r.naam }),
        el("span", { text: r.mededelingen || "Geen aanvullende omschrijving" }),
      ]),
      el("strong", { class: "indelingsbedrag", text: geld(r.bedrag) }),
      el("div", { class: "indelingsacties" }, [inkomst, uitgave]),
    ]);
  }));
}

function paneelTransacties(d) {
  const nietIngedeeld = d.niet_ingedeeld || [];
  if (transactieState.filter === "niet_ingedeeld" && !nietIngedeeld.length) {
    transactieState.filter = "alle";
  }
  const zoek = el("input", {
    type: "search", class: "transactie-zoek",
    placeholder: "Zoek op naam, leverancier of factuurnummer",
    value: transactieState.zoek,
    "aria-label": "Transacties zoeken",
  });
  const filter = el("select", { class: "transactie-filter", "aria-label": "Transacties filteren" });
  [
    ["alle", "Alle transacties"],
    ["inkomst", "Inkomsten"],
    ["uitgave", "Uitgaven"],
    ["bijzonder", "Bijzonderheden"],
    ["aangepast", "Aangepast"],
    ...(nietIngedeeld.length ? [["niet_ingedeeld", `Niet ingedeeld (${nietIngedeeld.length})`]] : []),
  ].forEach(([waarde, label]) => {
    const optie = el("option", { value: waarde, text: label });
    optie.selected = transactieState.filter === waarde;
    filter.append(optie);
  });
  const teller = el("span", { class: "transactie-teller", role: "status" });
  const reset = el("button", { class: "tekstknop", type: "button", text: "Wis filters" });
  const resultaat = el("div", { class: "transactie-resultaat" });

  const vernieuw = () => {
    transactieState.zoek = zoek.value;
    transactieState.filter = filter.value;
    const term = transactieState.zoek.trim().toLocaleLowerCase("nl-NL");
    if (transactieState.filter === "niet_ingedeeld") {
      const rijen = nietIngedeeld.filter((r) =>
        !term || [r.naam, r.factuur, r.mededelingen, r.tegenrekening, r.datum]
          .some((v) => String(v || "").toLocaleLowerCase("nl-NL").includes(term)));
      teller.textContent = `${rijen.length} niet ingedeeld`;
      resultaat.replaceChildren(bouwNietIngedeeld(rijen));
    } else {
      const rijen = d.transacties.filter((r) => {
        const tekstMatch = !term || [r.naam, r.leverancier, r.factuur, r.mededelingen, r.tegenrekening, r.datum]
          .some((v) => String(v || "").toLocaleLowerCase("nl-NL").includes(term));
        const filterMatch =
          transactieState.filter === "alle"
          || r.richting === transactieState.filter
          || (transactieState.filter === "bijzonder" && transactieKenmerken(r).length > 0)
          || (transactieState.filter === "aangepast" && r.btw_bron === "override");
        return tekstMatch && filterMatch;
      });
      teller.textContent = `${rijen.length} van ${d.transacties.length} transacties`;
      resultaat.replaceChildren(bouwTransactieTabel(rijen));
    }
    reset.hidden = !term && transactieState.filter === "alle";
  };
  zoek.addEventListener("input", vernieuw);
  filter.addEventListener("change", vernieuw);
  reset.addEventListener("click", () => {
    zoek.value = "";
    filter.value = "alle";
    vernieuw();
    zoek.focus();
  });
  vernieuw();
  return el("div", { class: "transactie-paneel" }, [
    el("div", { class: "transactie-toolbar" }, [
      el("label", { class: "zoekveld" }, [icoon("zoek"), zoek]),
      filter,
      teller,
      reset,
    ]),
    el("p", { class: "table-note", text: "Open een regel voor broninformatie en leveranciersspecifieke btw-instellingen." }),
    resultaat,
  ]);
}

function totaalTabel(titel, incl, excl, btw) {
  return bouwTabel(
    [{ label: titel }, { label: "Bedrag", getal: true }],
    [
      { a: "Incl. btw", b: incl },
      { a: "Excl. btw", b: excl },
      { a: "Btw", b: btw },
    ].map((r) => ({ a: r.a, b: r.b })),
    {}
  );
}

function maakActies(d) {
  const acties = [];
  const periode = d.periode || {};
  const nietIngedeeld = d.niet_ingedeeld || [];
  if (!periode.volledig_kwartaal) {
    acties.push({
      niveau: "waarschuwing",
      titel: "Periode niet als volledig kwartaal vastgesteld",
      context: "Controleer de bronperiode voordat je de aangifte overneemt.",
      actie: "Bekijk transacties",
      tab: "transacties",
    });
  }
  if (nietIngedeeld.length) {
    acties.push({
      niveau: "waarschuwing",
      titel: `${nietIngedeeld.length} ${nietIngedeeld.length === 1 ? "bankregel is" : "bankregels zijn"} niet ingedeeld`,
      context: "Deze regels tellen nog niet mee in de berekening.",
      actie: "Indelen",
      tab: "transacties",
      filter: "niet_ingedeeld",
    });
  }
  if (d.buiten_periode.length) {
    acties.push({
      niveau: "waarschuwing",
      titel: `${d.buiten_periode.length} ${d.buiten_periode.length === 1 ? "factuur valt" : "facturen vallen"} buiten de periode`,
      context: "Deze facturen tellen niet mee in de aangifte.",
      actie: "Bekijken",
      tab: "overzicht",
      doel: "buiten-periode-lijst",
    });
  }
  const buitenKwartaal = d.buiten_kwartaal || [];
  if (buitenKwartaal.length) {
    acties.push({
      niveau: "neutraal",
      titel: `${buitenKwartaal.length} ${buitenKwartaal.length === 1 ? "boeking valt" : "boekingen vallen"} na het kwartaaleinde`,
      context: "Niet meegeteld; deze gaan mee in de aangifte van het volgende kwartaal.",
      actie: "Bekijken",
      tab: "overzicht",
      doel: "buiten-kwartaal-lijst",
    });
  }
  if (!bonnenData) {
    acties.push({
      niveau: "neutraal",
      titel: "Bonnen nog niet gecontroleerd",
      context: "Koppel bewijsstukken om de voorbelasting te onderbouwen.",
      actie: "Controleer bonnen",
      tab: "bonnen",
    });
  } else if (bonnenData.tellingen.zonder_bon || bonnenData.tellingen.onzeker || bonnenData.tellingen.geen_match) {
    const open = bonnenData.tellingen.zonder_bon + bonnenData.tellingen.onzeker + bonnenData.tellingen.geen_match;
    acties.push({
      niveau: "risico",
      titel: `${open} ${open === 1 ? "boncontrole staat" : "boncontroles staan"} open`,
      context: `${geld(bonnenData.btw_risico)} voorbelasting is nog niet volledig onderbouwd.`,
      actie: "Controleer bonnen",
      tab: "bonnen",
    });
  }
  if (d.tellingen.verlegd) {
    acties.push({
      niveau: "waarschuwing",
      titel: `${d.tellingen.verlegd} ${d.tellingen.verlegd === 1 ? "post met" : "posten met"} verlegde btw`,
      context: "Controleer rubriek 1e en 4b voordat je indient.",
      actie: "Bekijk aangifte",
      tab: "aangifte",
    });
  }
  if (d.tellingen.investeringen) {
    acties.push({
      niveau: "neutraal",
      titel: `${d.tellingen.investeringen} mogelijke ${d.tellingen.investeringen === 1 ? "investering" : "investeringen"}`,
      context: "Beoordeel of deze uitgaven moeten worden afgeschreven.",
      actie: "Bekijken",
      tab: "investeringen",
    });
  }
  if (d.leveranciers.afwijkingen.length) {
    acties.push({
      niveau: "neutraal",
      titel: `${d.leveranciers.afwijkingen.length} opvallende ${d.leveranciers.afwijkingen.length === 1 ? "betaling" : "betalingen"}`,
      context: "Deze bedragen wijken af van het gebruikelijke patroon.",
      actie: "Bekijken",
      tab: "leveranciers",
    });
  }
  return acties;
}

function openActie(item) {
  if (item.filter) {
    transactieState.filter = item.filter;
    transactieState.zoek = "";
  }
  toonTab(item.tab);
  requestAnimationFrame(() => {
    const doel = item.doel ? document.getElementById(item.doel) : document.getElementById("panel-" + item.tab);
    doel?.scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "start" });
    doel?.focus?.({ preventScroll: true });
  });
}

function actieRij(item) {
  const knop = el("button", { class: "actie-link", type: "button", text: item.actie });
  knop.addEventListener("click", () => openActie(item));
  return el("li", { class: "actie-rij -" + item.niveau }, [
    el("span", { class: "actie-punt", "aria-hidden": "true" }),
    el("div", { class: "actie-tekst" }, [
      el("strong", { text: item.titel }),
      el("span", { text: item.context }),
    ]),
    knop,
  ]);
}

function resultaatRij(label, bedrag, klasse = "") {
  return el("div", { class: "resultaat-rij " + klasse }, [
    el("dt", { text: label }),
    el("dd", { text: geld(bedrag) }),
  ]);
}

function paneelOverzicht(d) {
  const t = d.totalen;
  const a = d.aangifte;
  const saldoTerug = a.btw_5c < 0;
  const acties = maakActies(d);
  const financieel = el("section", { class: "financieel-resultaat", "aria-labelledby": "financieel-titel" }, [
    el("h3", { id: "financieel-titel", text: "Financieel resultaat" }),
    el("dl", { class: "resultaat-lijst" }, [
      resultaatRij("Omzet excl. btw", t.inkomsten_excl),
      resultaatRij("Zakelijke kosten excl. btw", -Math.abs(t.uitgaven_excl)),
      resultaatRij("Winst vóór inkomstenbelasting", d.winst, "-totaal"),
    ]),
    el("section", { class: "aangifte-positie", "aria-labelledby": "positie-titel" }, [
      el("h3", { id: "positie-titel", text: "Aangiftepositie" }),
      el("dl", { class: "resultaat-lijst" }, [
        resultaatRij("Verschuldigde btw · 5a", a.btw_5a),
        resultaatRij("Voorbelasting · 5b", -Math.abs(a.btw_5b)),
        resultaatRij(saldoTerug ? "Terug te vragen · 5c" : "Af te dragen · 5c", Math.abs(a.btw_5c), "-totaal"),
      ]),
      el("p", { class: "deadline-regel", text:
        a.deadline_betrouwbaar
          ? (a.btw_5c > 0
              ? "Aangifte en betaling " + a.deadline_tekst + "."
              : "Dien de aangifte " + a.deadline_tekst + " in.")
          : "Controleer de uiterste datum in Mijn Belastingdienst Zakelijk." }),
    ]),
  ]);
  const actiecentrum = el("aside", { class: "actiecentrum", "aria-labelledby": "actiecentrum-titel" }, [
    el("header", { class: "actiecentrum-kop" }, [
      el("h3", { id: "actiecentrum-titel", text: "Te controleren" }),
      el("span", { text: acties.length ? `${acties.length} open ${acties.length === 1 ? "punt" : "punten"}` : "Geen open punten" }),
    ]),
    acties.length
      ? el("ul", { class: "actielijst" }, acties.map(actieRij))
      : el("div", { class: "acties-klaar" }, [
          el("strong", { text: "Geen open controles" }),
          el("p", { text: "De beschikbare gegevens zijn klaar om over te nemen." }),
        ]),
  ]);
  const kinderen = [el("div", { class: "overzicht-layout" }, [financieel, actiecentrum])];
  if (d.buiten_periode.length) {
    kinderen.push(el("section", {
      class: "overzicht-detail", id: "buiten-periode-lijst", tabindex: "-1",
      "aria-labelledby": "buiten-periode-titel",
    }, [
      el("div", { class: "sectie-kop" }, [
        el("h3", { id: "buiten-periode-titel", text: "Buiten de periode" }),
        el("p", { text: "Deze facturen zijn niet in de berekening opgenomen." }),
      ]),
      bouwTabel(
        [
          { key: "datum", label: "Datum", render: datum },
          { key: "naam", label: "Naam" },
          { key: "factuur", label: "Factuur" },
          { key: "bedrag", label: "Bedrag", getal: true, render: geld },
        ],
        d.buiten_periode
      ),
    ]));
  }
  const buitenKwartaal = d.buiten_kwartaal || [];
  if (buitenKwartaal.length) {
    kinderen.push(el("section", {
      class: "overzicht-detail", id: "buiten-kwartaal-lijst", tabindex: "-1",
      "aria-labelledby": "buiten-kwartaal-titel",
    }, [
      el("div", { class: "sectie-kop" }, [
        el("h3", { id: "buiten-kwartaal-titel", text: "Na het kwartaaleinde" }),
        el("p", { text: "Het bankbestand loopt een paar dagen door in het volgende kwartaal. Deze boekingen tellen hier niet mee; ze horen in de volgende aangifte." }),
      ]),
      bouwTabel(
        [
          { key: "datum", label: "Datum", render: datum },
          { key: "naam", label: "Naam" },
          { key: "factuur", label: "Omschrijving" },
          { key: "bedrag", label: "Bedrag", getal: true, render: geld },
        ],
        buitenKwartaal
      ),
    ]));
  }
  return el("div", { class: "overzicht-paneel" }, kinderen);
}

function paneelAangifte(d) {
  const a = d.aangifte;
  const kinderen = [uitleg("De cijfers die je overtikt in Mijn Belastingdienst Zakelijk. " + a.periode_tekst + ".")];

  if (a.kor_actief) {
    kinderen.push(el("div", { class: "aangifte-notitie" }, [
      el("h2", { text: "KOR actief" }),
      el("p", { text: "Onder de kleineondernemersregeling doe je geen btw-aangifte. De rubrieken hieronder zijn puur informatief." }),
    ]));
  }

  const rubrieken = [
    { label: "1a  Leveringen/diensten belast met 21%", omzet: a.omzet_1a, btw: a.btw_1a },
  ];
  if (a.heeft_1b) rubrieken.push({ label: "1b  Leveringen/diensten belast met 9%", omzet: a.omzet_1b, btw: a.btw_1b });
  if (a.heeft_1e) rubrieken.push({ label: "1e  Omzet met btw verlegd naar afnemer (0%)", omzet: a.verlegd_verkoop_excl, btw: 0 });
  if (a.heeft_overige) rubrieken.push({ label: "Overige omzet 0% / anders (controleer)", omzet: a.overige_excl, btw: a.overige_btw || 0 });
  if (a.heeft_4b) rubrieken.push({ label: "4b  Btw naar mij verlegd (inkoop onderaanneming)", omzet: a.verlegd_inkoop_excl, btw: a.verlegd_inkoop_btw });
  rubrieken.push({ label: "5a  Verschuldigde omzetbelasting", omzet: null, btw: a.btw_5a, vet: true });
  rubrieken.push({ label: "5b  Voorbelasting (terug te vragen)", omzet: null, btw: a.btw_5b, vet: true });
  rubrieken.push({ label: "5c  Te betalen (+) / terug te vragen (−)", omzet: null, btw: a.btw_5c, vet: true });

  kinderen.push(bouwTabel(
    [
      { key: "label", label: "Rubriek", cel: (r) => el("td", {}, r.vet ? el("strong", { text: r.label }) : r.label) },
      { key: "omzet", label: "Omzet (excl. btw)", getal: true, render: (v) => (v == null ? "" : geld(v)) },
      { key: "btw", label: "Btw", getal: true, cel: (r) => el("td", { class: "getal" }, r.vet ? el("strong", { text: geld(r.btw) }) : geld(r.btw)) },
    ],
    rubrieken
  ));

  if (a.heeft_1e || a.heeft_4b) {
    kinderen.push(el("p", { class: "melding-leeg", text: "Let op: verlegde btw is fiscaal maatwerk — controleer 1e/4b met je boekhouder." }));
  }

  const teBetalen = a.reserveren > 0;
  kinderen.push(el("section", { class: "aangifte-uitkomst", "aria-label": "Uitkomst aangifte" }, [
    el("span", { class: "aangifte-label", text: "Uitkomst" }),
    el("div", { class: "aangifte-uitkomst-rij" }, [
      el("div", {}, [
        el("h2", { text: teBetalen ? "Af te dragen" : "Terug te vragen" }),
        el("p", { text: teBetalen ? "Reserveer dit bedrag voor de aangifte." : "Dit is het saldo volgens het bankbestand." }),
      ]),
      el("strong", { class: "aangifte-saldo", text: geld(teBetalen ? a.reserveren : Math.abs(a.btw_5c)) }),
    ]),
    el("p", { class: "aangifte-deadline", text: a.deadline_betrouwbaar
      ? (teBetalen
          ? "Aangifte en betaling " + a.deadline_tekst + "."
          : "Dien de aangifte " + a.deadline_tekst + " in.")
      : "Controleer de uiterste datum in Mijn Belastingdienst Zakelijk." }),
  ]));

  kinderen.push(el("section", { class: "jaarcontext", "aria-label": "KOR-context" }, [
    el("div", {}, [
      el("span", { text: "Omzet in dit bankbestand" }),
      el("strong", { text: geld(a.jaaromzet_tot_nu) }),
    ]),
    el("p", { text: "De KOR-grens van € 20.000 geldt voor het volledige kalenderjaar. Neem voor die beoordeling ook omzet buiten dit bestand mee." }),
  ]));

  return el("div", { class: "aangifte-paneel" }, kinderen);
}

function richtingKolommen() {
  return [
    { key: "datum", label: "Datum", render: datum },
    { key: "naam", label: "Naam" },
    { key: "leverancier", label: "Leverancier" },
    { key: "factuur", label: "Factuur" },
    { key: "btw_pct", label: "Btw %", getal: true, render: (v) => v + "%" },
    { key: "bedrag", label: "Bedrag incl.", getal: true, cel: (r) => bedragCel(r.bedrag, r.richting) },
    { key: "excl", label: "Excl. btw", getal: true, render: geld },
    { key: "btw", label: "Btw", getal: true, render: geld },
  ];
}

function paneelRichting(rijen, titel) {
  if (!rijen.length) return leegMelding(`Geen ${titel} in deze periode.`);
  const som = (v) => rijen.reduce((a, r) => a + r[v], 0);
  const voet = [
    el("td", { colspan: 5 }, el("strong", { text: "Totaal" })),
    el("td", { class: "getal" }, el("strong", { text: geld(som("bedrag")) })),
    el("td", { class: "getal" }, el("strong", { text: geld(som("excl")) })),
    el("td", { class: "getal" }, el("strong", { text: geld(som("btw")) })),
  ];
  return bouwTabel(richtingKolommen(), rijen, { caption: `${rijen.length} ${titel}`, voet });
}

function paneelInvesteringen(d) {
  const kinderen = [uitleg(
    `Grote uitgaven vanaf ${geld(d.investering_drempel)} excl. btw — kandidaten om te controleren. ` +
    "Een bedrijfsmiddel dat jaren meegaat schrijf je af; materialen zijn gewone kosten. Overleg met je boekhouder.")];
  if (!d.investeringen.length) {
    kinderen.push(leegMelding("Geen investeringen boven de drempel in deze periode."));
    return el("div", { class: "investeringen-paneel" }, kinderen);
  }
  const totExcl = d.investeringen.reduce((a, r) => a + r.excl, 0);
  kinderen.push(bouwTabel(
    [
      { key: "datum", label: "Datum", render: datum },
      { key: "leverancier", label: "Leverancier" },
      { key: "factuur", label: "Factuur" },
      { key: "bedrag", label: "Bedrag incl.", getal: true, render: geld },
      { key: "excl", label: "Excl. btw", getal: true, render: geld },
      { key: "btw", label: "Btw", getal: true, render: geld },
    ],
    d.investeringen,
    { voet: [
        el("td", { colspan: 4 }, el("strong", { text: "Totaal (excl. btw)" })),
        el("td", { class: "getal" }, el("strong", { text: geld(totExcl) })),
        el("td", {}),
      ] }
  ));
  kinderen.push(el("aside", { class: "fiscale-notitie" }, [
    el("strong", { text: "Fiscale context" }),
    el("p", { text:
      "De KIA wordt bepaald over het totale kwalificerende investeringsbedrag in het kalenderjaar. " +
      "Gebruik deze selectie als controlelijst en laat de definitieve kwalificatie door je boekhouder beoordelen." }),
  ]));
  return el("div", { class: "investeringen-paneel" }, kinderen);
}

function paneelLeveranciers(d) {
  const lv = d.leveranciers;
  const kinderen = [
    uitleg("Je uitgaven opgeteld per leverancier. Opvallende bedragen wijken sterk af van wat je normaal bij die leverancier betaalt — even checken."),
    bouwTabel(
      [
        { key: "leverancier", label: "Leverancier" },
        { key: "aantal", label: "Aantal", getal: true },
        { key: "totaal", label: "Totaal incl.", getal: true, render: geld },
        { key: "gemiddeld", label: "Gemiddeld", getal: true, render: geld },
        { key: "eerste", label: "Eerste", render: datum },
        { key: "laatste", label: "Laatste", render: datum },
      ],
      lv.rijen
    ),
  ];
  kinderen.push(el("h2", { text: `Opvallende bedragen (${lv.afwijkingen.length})` }));
  if (lv.afwijkingen.length) {
    kinderen.push(bouwTabel(
      [
        { key: "datum", label: "Datum", render: datum },
        { key: "leverancier", label: "Leverancier" },
        { key: "bedrag", label: "Bedrag", getal: true, render: geld },
        { key: "mediaan", label: "Typisch (mediaan)", getal: true, render: geld },
      ],
      lv.afwijkingen
    ));
  } else {
    kinderen.push(leegMelding("Geen opvallende bedragen gevonden."));
  }
  return el("div", { class: "leveranciers-paneel" }, kinderen);
}

let bonnenData = null;

function bonnenSectie(titel, klasse, rijen, kolommen) {
  const kinderen = [el("h2", { text: `${titel} (${rijen.length})` })];
  if (rijen.length) kinderen.push(bouwTabel(kolommen, rijen));
  else kinderen.push(leegMelding("(geen)"));
  return el("div", { class: "bonnen-sectie " + (klasse || "") }, kinderen);
}

function paneelBonnen() {
  const kiesKnop = el("button", { class: "knop knop-primair", type: "button",
    text: bonnenData ? "Opnieuw controleren…" : "Kies map met bonnen…" });
  kiesKnop.addEventListener("click", controleerBonnen);

  const kop = el("div", {}, [
    uitleg("Kies een map met pdf's of foto's. Kwartaal leest de bestanden, koppelt ze aan uitgaven en zet onzekere matches voor je klaar."),
    el("div", { class: "toolbar" }, kiesKnop),
  ]);

  if (!bonnenData) {
    kop.append(el("section", { class: "lege-werkstroom", "aria-label": "Status bonnencontrole" }, [
      el("span", { class: "werkstroom-label", text: "Status" }),
      el("h2", { text: "Nog geen bonnen gecontroleerd" }),
      el("p", { text: "Er zijn nog geen bewijsstukken aan dit kwartaal gekoppeld." }),
    ]));
    return kop;
  }

  const b = bonnenData;
  const heeftRisico = b.btw_risico > 0;
  const risicoBlok = el("section", {
    class: "aangifte-blok risico-kaart " + (heeftRisico ? "-slecht" : "-goed"),
    "aria-label": "Controle op uitgaven zonder bon",
  }, [
    el("div", { class: "risico-kop" }, [
      el("div", { class: "risico-identiteit" }, [
        el("span", { class: "risico-status" }, [
          el("span", { class: "risico-punt", "aria-hidden": "true" }),
          heeftRisico ? "Controle nodig" : "In orde",
        ]),
        el("h2", { text: "Uitgaven zonder bon" }),
      ]),
      el("strong", { class: "risico-bedrag", text: geld(b.btw_risico) }),
    ]),
    el("p", { class: "risico-samenvatting", text:
      "Voorbelasting die is teruggevraagd zonder gekoppelde factuur." }),
    el("p", { class: "risico-notitie", text: heeftRisico
      ? "Koppel de ontbrekende facturen vóór je rubriek 5b indient; bij controle kan de Belastingdienst dit bedrag corrigeren."
      : "Alle teruggevraagde voorbelasting is onderbouwd met een gekoppelde factuur." }),
  ]);

  const tel = b.tellingen;
  const kerncijfers = bouwTabel(
    [{ key: "cat", label: "Categorie" }, { key: "n", label: "Aantal", getal: true }],
    [
      { cat: "Zeker gematcht", n: tel.zeker },
      { cat: "Controleer (onzeker)", n: tel.onzeker },
      { cat: "Geen match", n: tel.geen_match },
      { cat: "Onleesbaar", n: tel.onleesbaar },
      { cat: "Uitgaven zonder bon", n: tel.zonder_bon },
    ],
    { caption: `${b.totaal_bonnen} bonnen verwerkt` }
  );

  const secties = [
    bonnenSectie("Zeker gematcht", "-goed", b.zeker, [
      { key: "bon", label: "Bon" }, { key: "herkend", label: "Herkend nummer" },
      { key: "tx_datum", label: "Transactiedatum", render: datum }, { key: "tx_naam", label: "Naam" },
      { key: "bedrag", label: "Bedrag", getal: true, render: geld },
    ]),
    bonnenSectie("Controleer — onzekere match", "-let-op", b.onzeker, [
      { key: "bon", label: "Bon" }, { key: "bon_bedrag", label: "Bedrag (bon)", getal: true, render: (v) => (v == null ? "" : geld(v)) },
      { key: "bon_datum", label: "Datum (bon)", render: datum }, { key: "tx_naam", label: "Mogelijke naam" },
      { key: "tx_bedrag", label: "Mogelijk bedrag", getal: true, render: geld },
    ]),
    bonnenSectie("Geen match gevonden", "-slecht", b.geen_match, [
      { key: "bon", label: "Bon" }, { key: "kandidaten", label: "Kandidaat-nummers" },
      { key: "bedrag", label: "Bedrag (bon)", getal: true, render: (v) => (v == null ? "" : geld(v)) },
      { key: "datum", label: "Datum (bon)", render: datum },
    ]),
    bonnenSectie("Onleesbaar", "", b.onleesbaar, [{ key: "bon", label: "Bon" }]),
    bonnenSectie("Uitgaven zonder bon", "-slecht", b.zonder_bon, [
      { key: "datum", label: "Datum", render: datum }, { key: "naam", label: "Naam" },
      { key: "bedrag", label: "Bedrag", getal: true, render: geld },
    ]),
  ];

  return el("div", { class: "bonnen-paneel" }, [kop, risicoBlok, kerncijfers, ...secties]);
}

async function controleerBonnen() {
  if (!heeftApi()) { toonToast("Niet beschikbaar", "Bonnencontrole werkt alleen in de desktop-app.", "fout"); return; }
  status("Bonnen lezen en koppelen…");
  try {
    const r = await window.pywebview.api.kies_bonnen_map();
    if (r.geannuleerd) { status(""); return; }
    if (!r.ok) {
      toonToast("Bonnencontrole mislukt", r.fout || "Controleer de gekozen map.", "fout");
      status("");
      return;
    }
    bonnenData = r.data;
    bouwTablist();
    toonTab("bonnen");
    status("");
    toonToast("Bonnen gecontroleerd", `${bonnenData.totaal_bonnen} bestanden zijn verwerkt.`, "ok");
  } catch (fout) {
    toonToast("Bonnencontrole mislukt", "De map kon niet worden verwerkt.", "fout");
    status("");
  }
}

/* ---------- tabs ---------- */
const TABS = [
  { key: "overzicht", groep: "werk", label: "Overzicht", omschrijving: "Resultaat, aangiftepositie en open controles.", render: paneelOverzicht },
  { key: "transacties", groep: "werk", label: "Transacties", omschrijving: "Zoek bankregels, controleer details en pas btw toe.", render: paneelTransacties,
    telling: (d) => (d.niet_ingedeeld || []).length },
  { key: "bonnen", groep: "werk", label: "Bonnen", omschrijving: "Koppel bewijsstukken aan de bijbehorende uitgaven.", render: paneelBonnen,
    telling: () => bonnenData ? bonnenData.tellingen.zonder_bon + bonnenData.tellingen.onzeker + bonnenData.tellingen.geen_match : 0 },
  { key: "aangifte", groep: "werk", label: "Btw-aangifte", omschrijving: "De rubrieken voor Mijn Belastingdienst Zakelijk.", render: paneelAangifte },
  { key: "inkomsten", groep: "rapporten", label: "Inkomsten", omschrijving: "Alle opgenomen inkomsten in deze periode.", render: (d) => paneelRichting(d.inkomsten, "inkomsten") },
  { key: "uitgaven", groep: "rapporten", label: "Uitgaven", omschrijving: "Alle opgenomen zakelijke uitgaven.", render: (d) => paneelRichting(d.uitgaven, "uitgaven") },
  { key: "investeringen", groep: "rapporten", label: "Investeringen", omschrijving: "Uitgaven die mogelijk als bedrijfsmiddel gelden.", render: paneelInvesteringen,
    telling: (d) => d.investeringen.length },
  { key: "leveranciers", groep: "rapporten", label: "Leveranciers", omschrijving: "Uitgavenpatronen en opvallende bedragen.", render: paneelLeveranciers,
    telling: (d) => d.leveranciers.afwijkingen.length },
];

let huidigeData = null;
let actieveTab = "overzicht";
let importConcept = null;
let toastTimer = null;
let commandTerugFocus = null;

function toonTab(key) {
  actieveTab = key;
  document.querySelectorAll("#tablist .tab").forEach((btn) => {
    const gekozen = btn.dataset.key === key;
    btn.setAttribute("aria-selected", gekozen ? "true" : "false");
    btn.tabIndex = gekozen ? 0 : -1;
    if (gekozen) btn.setAttribute("aria-controls", "panel-" + key);
    else btn.removeAttribute("aria-controls");
  });
  const meerKnop = document.getElementById("meer-knop");
  meerKnop?.setAttribute("aria-pressed",
    TABS.find((t) => t.key === key)?.groep === "rapporten" ? "true" : "false");
  const houder = document.getElementById("tab-inhoud");
  houder.innerHTML = "";
  const tab = TABS.find((t) => t.key === key);
  const paneelKop = el("header", { class: "paneel-kop" }, el("div", {}, [
    el("h2", { text: tab.label }),
    el("p", { text: tab.omschrijving }),
  ]));
  const paneel = el("div", {
    class: "tabpanel", role: "tabpanel", id: "panel-" + key,
    "aria-labelledby": "tab-" + key, tabindex: "0",
  }, [paneelKop, tab.render(huidigeData)]);
  houder.append(paneel);
}

function bouwTablist() {
  const lijst = document.getElementById("tablist");
  lijst.innerHTML = "";
  let laatsteGroep = null;
  TABS.forEach((t) => {
    if (t.groep !== laatsteGroep) {
      laatsteGroep = t.groep;
      lijst.append(el("p", {
        class: "nav-groep", role: "presentation",
        text: t.groep === "werk" ? "Werk" : "Rapporten",
      }));
    }
    const aantal = t.telling ? t.telling(huidigeData) : 0;
    const badge = aantal ? el("span", { class: "badge", text: aantal }) : null;
    const btn = el("button", {
      class: "tab", type: "button", role: "tab", id: "tab-" + t.key,
      "aria-controls": t.key === actieveTab ? "panel-" + t.key : null,
      "aria-selected": t.key === actieveTab ? "true" : "false",
      tabindex: t.key === actieveTab ? "0" : "-1",
      "data-key": t.key,
      "data-group": t.groep,
    }, [t.label, badge]);
    btn.addEventListener("click", () => toonTab(t.key));
    lijst.append(btn);
  });
  const meerKnop = document.getElementById("meer-knop");
  meerKnop?.setAttribute("aria-pressed",
    TABS.find((t) => t.key === actieveTab)?.groep === "rapporten" ? "true" : "false");
  bouwMeerLijst();
  // Toetsenbord: pijltjes/Home/End tussen tabs (WAI-ARIA tabs-patroon).
  lijst.onkeydown = (e) => {
    const keys = TABS.map((t) => t.key);
    let i = keys.indexOf(actieveTab);
    if (e.key === "ArrowRight") i = (i + 1) % keys.length;
    else if (e.key === "ArrowLeft") i = (i - 1 + keys.length) % keys.length;
    else if (e.key === "Home") i = 0;
    else if (e.key === "End") i = keys.length - 1;
    else return;
    e.preventDefault();
    toonTab(keys[i]);
    document.getElementById("tab-" + keys[i]).focus();
  };
}

/* ---------- status + render ---------- */
function status(tekst, soort) {
  const s = document.getElementById("status");
  s.textContent = tekst || "";
  s.className = "status" + (soort ? " -" + soort : "");
}

function periodeVan(data) {
  if (data.periode) return data.periode;
  const datums = (data.transacties || []).map((r) => r.datum).filter(Boolean).sort();
  return {
    label: "Bankperiode",
    start: datums[0] || "",
    eind: datums.at(-1) || "",
    volledig_kwartaal: false,
    deadline: null,
  };
}

function periodeBereik(data) {
  const p = periodeVan(data);
  if (!p.start || !p.eind) return "Periode niet beschikbaar";
  return `${datumLang(p.start)} – ${datumLang(p.eind)}`;
}

function werkDossierContextBij(data) {
  const p = periodeVan(data);
  document.getElementById("periode-label").textContent = p.label;
  document.getElementById("periode-bereik").textContent = periodeBereik(data);
  const bedrijfsnaam = document.getElementById("bedrijfsnaam");
  bedrijfsnaam.textContent = data.bedrijfsnaam || "";
  bedrijfsnaam.hidden = !data.bedrijfsnaam;
  document.getElementById("bestandsregel").textContent = data.bestandsnaam;
  document.getElementById("dossier-context").hidden = false;
  document.title = `Kwartaal — ${p.label}`;
}

function zetHoofdscherm(scherm) {
  document.getElementById("leeg").hidden = scherm !== "leeg";
  document.getElementById("import-review").hidden = scherm !== "import";
  document.getElementById("werkblad").hidden = scherm !== "werkblad";
  if (scherm !== "werkblad") document.getElementById("dossier-context").hidden = true;
}

function werkHeaderActiesBij() {
  const kies = document.getElementById("knop-kies");
  const kiesTekst = kies.querySelector("span");
  const excel = document.getElementById("knop-excel");
  const zoeken = document.getElementById("knop-zoeken");
  if (huidigeData) {
    kies.classList.remove("knop-primair");
    kiesTekst.textContent = "Nieuw bestand";
    excel.classList.add("knop-primair");
    excel.hidden = false;
    excel.disabled = false;
    zoeken.hidden = false;
  } else {
    kies.classList.add("knop-primair");
    kiesTekst.textContent = "Open bankbestand";
    excel.classList.remove("knop-primair");
    excel.hidden = true;
    excel.disabled = true;
    zoeken.hidden = true;
  }
}

function renderModel(data) {
  huidigeData = data;
  zetHoofdscherm("werkblad");
  werkDossierContextBij(data);
  werkHeaderActiesBij();
  bouwTablist();
  toonTab(actieveTab);
  status("");
}

function toonImportReview(data) {
  importConcept = data;
  const p = periodeVan(data);
  const controle = data.importcontrole || {
    bronregels: data.transacties.length,
    verwerkt: data.tellingen.verwerkt,
    buiten_beschouwing: data.tellingen.prive + data.tellingen.belastingdienst + data.tellingen.uitgesloten,
    aandachtspunten: (data.niet_ingedeeld || []).length + data.buiten_periode.length + (p.volledig_kwartaal ? 0 : 1),
  };
  zetHoofdscherm("import");
  document.getElementById("import-periode-label").textContent = p.label;
  document.getElementById("import-periode-bereik").textContent = periodeBereik(data);
  document.getElementById("import-bestand").textContent = data.bestandsnaam;
  document.getElementById("import-verwerkt").textContent = `${controle.verwerkt} transacties`;
  document.getElementById("import-uitgesloten").textContent = `${controle.buiten_beschouwing} regels`;
  document.getElementById("import-aandacht").textContent = `${controle.aandachtspunten} ${controle.aandachtspunten === 1 ? "punt" : "punten"}`;
  document.getElementById("import-bronregels").textContent = controle.bronregels;
  document.getElementById("import-prive").textContent = data.tellingen.prive;
  document.getElementById("import-belastingdienst").textContent = data.tellingen.belastingdienst;
  document.getElementById("import-uitgesloten-detail").textContent = data.tellingen.uitgesloten;
  document.getElementById("import-niet-ingedeeld").textContent = (data.niet_ingedeeld || []).length;
  document.getElementById("import-buiten-periode").textContent = data.buiten_periode.length;
  document.getElementById("import-buiten-kwartaal").textContent = (data.buiten_kwartaal || []).length;
  const statusBlok = document.getElementById("import-status");
  const statusTekst = document.getElementById("import-status-tekst");
  statusBlok.classList.toggle("-waarschuwing", !p.volledig_kwartaal);
  statusTekst.textContent = p.volledig_kwartaal ? "Bestand gecontroleerd" : "Periode controleren";
  document.getElementById("import-oordeel").textContent = p.volledig_kwartaal
    ? `${p.label} is herkend. Gebruik dit kwartaal en rond daarna de open controles af.`
    : "De bronperiode is niet exact als één kwartaal vastgesteld. Controleer de periode voordat je aangiftebedragen overneemt.";
  document.getElementById("import-annuleren").textContent =
    huidigeData ? "Terug naar huidig kwartaal" : "Annuleren";
  document.title = "Kwartaal — import controleren";
  status("");
  document.getElementById("import-review-titel").focus?.();
}

/* ---------- pywebview-brug ---------- */
function heeftApi() {
  return window.pywebview && window.pywebview.api;
}

async function kiesBestand() {
  if (!heeftApi()) { toonToast("Niet beschikbaar", "Open de desktop-app om een ING-bestand te kiezen.", "fout"); return; }
  const knoppen = [document.getElementById("knop-kies"), document.getElementById("knop-kies-leeg"), document.getElementById("import-ander")];
  knoppen.forEach((knop) => { if (knop) knop.disabled = true; });
  status("Bankbestand controleren…");
  try {
    const r = await window.pywebview.api.kies_csv();
    if (r.geannuleerd) { status(""); return; }
    if (!r.ok) {
      toonToast("Dit bestand kan niet worden gelezen", r.fout || "Kies een CSV-export uit Mijn ING.", "fout");
      status("");
      return;
    }
    toonImportReview(r.data);
  } catch (fout) {
    toonToast("Bestand niet geopend", "De desktop-app kon het bestand niet controleren.", "fout");
    status("");
  } finally {
    knoppen.forEach((knop) => { if (knop) knop.disabled = false; });
  }
}

async function bevestigImport() {
  if (!heeftApi() || !importConcept) return;
  const knop = document.getElementById("import-bevestigen");
  knop.disabled = true;
  knop.textContent = "Kwartaal openen…";
  try {
    const r = await window.pywebview.api.bevestig_csv();
    if (!r.ok) {
      toonToast("Import niet geopend", r.fout || "Probeer het bestand opnieuw.", "fout");
      return;
    }
    bonnenData = null;
    actieveTab = "overzicht";
    transactieState.zoek = "";
    transactieState.filter = "alle";
    transactieState.openId = null;
    importConcept = null;
    renderModel(r.data);
    toonToast("Kwartaal geopend", `${r.data.tellingen.verwerkt} transacties zijn opgenomen.`, "ok");
  } catch (fout) {
    toonToast("Import niet geopend", "Probeer het bestand opnieuw te kiezen.", "fout");
  } finally {
    knop.disabled = false;
    knop.textContent = "Open kwartaaloverzicht";
  }
}

async function annuleerImport() {
  if (heeftApi()) await window.pywebview.api.annuleer_csv();
  importConcept = null;
  if (huidigeData) renderModel(huidigeData);
  else {
    zetHoofdscherm("leeg");
    werkHeaderActiesBij();
    document.title = "Kwartaal — btw uit je bankbestand";
  }
}

function naEdit(data, id, melding) {
  bonnenData = null;
  renderModel(data); // herbouwt de werkbladen uit het verse model
  const sel = document.querySelector(`.btw-select[data-id="${id}"]`);
  if (sel) sel.focus();
  toonToast("Wijziging opgeslagen", melding, "ok");
}

async function bewerkBtw(id, pct, onthoud = false) {
  if (!heeftApi()) { toonToast("Niet beschikbaar", "Bewerken werkt alleen in de desktop-app.", "fout"); return; }
  status("Bezig met bijwerken…");
  const r = await window.pywebview.api.zet_btw(id, parseInt(pct, 10), onthoud);
  if (!r.ok) { status(r.fout || "Bewerken mislukt.", "fout"); return; }
  naEdit(r.data, id, `Btw aangepast naar ${pct}%.` + (onthoud ? " Deze keuze wordt voor de leverancier onthouden." : ""));
}

async function herstelBtw(id) {
  if (!heeftApi()) return;
  status("Bezig met herstellen…");
  const r = await window.pywebview.api.herstel_btw(id);
  if (!r.ok) { status(r.fout || "Herstellen mislukt.", "fout"); return; }
  naEdit(r.data, id, "Btw hersteld naar de automatische waarde.");
}

async function classificeerRichting(id, richting) {
  if (!heeftApi()) { toonToast("Niet beschikbaar", "Indelen werkt alleen in de desktop-app.", "fout"); return; }
  status("Bankregel indelen…");
  const r = await window.pywebview.api.zet_richting(id, richting);
  if (!r.ok) {
    toonToast("Indelen mislukt", r.fout || "Probeer het opnieuw.", "fout");
    status("");
    return;
  }
  bonnenData = null;
  renderModel(r.data);
  toonToast("Bankregel ingedeeld", `De regel telt nu mee als ${richting}.`, "ok");
}

function exportBestandsnaam(data) {
  const basis = String(data.bestandsnaam || "kwartaal").replace(/\.csv$/i, "");
  return `${basis}_btw_overzicht.xlsx`;
}

function openExportDialoog() {
  if (!huidigeData) return;
  document.getElementById("export-periode").textContent =
    `${periodeVan(huidigeData).label} · ${periodeBereik(huidigeData)}`;
  document.getElementById("export-bestandsnaam").textContent = exportBestandsnaam(huidigeData);
  const onderdelen = [
    ["Excel-werkboek", "Overzicht, aangifte, transacties en rapporten"],
    ["Boekhoud-CSV", "Voor import in boekhoudsoftware, wanneer ingeschakeld"],
  ];
  if (bonnenData) onderdelen.push(["Bonnencontrole", "De laatst uitgevoerde bonnenanalyse"]);
  const lijst = document.getElementById("export-onderdelen");
  lijst.replaceChildren(...onderdelen.map(([titel, tekst]) => el("li", {}, [
    icoon("check"),
    el("div", {}, [el("strong", { text: titel }), el("span", { text: tekst })]),
  ])));
  document.getElementById("export-dialoog").showModal();
}

async function downloadExcel() {
  if (!heeftApi()) { toonToast("Niet beschikbaar", "Exporteren werkt alleen in de desktop-app.", "fout"); return; }
  const knop = document.getElementById("export-bevestigen");
  knop.disabled = true;
  knop.textContent = "Exporteren…";
  status("Bezig met opslaan…");
  try {
    const r = await window.pywebview.api.download_excel();
    if (r.geannuleerd) { status(""); return; }
    if (!r.ok) {
      toonToast("Exporteren mislukt", r.fout || "Kies een andere locatie.", "fout");
      status("");
      return;
    }
    document.getElementById("export-dialoog").close();
    status("");
    toonToast("Export opgeslagen", r.csv ? `${r.pad} · boekhoud-CSV aangemaakt` : r.pad, "ok");
  } catch (fout) {
    status("");
    toonToast(
      "Exporteren mislukt",
      fout?.message || "Probeer het opnieuw of kies een andere locatie.",
      "fout",
    );
  } finally {
    knop.disabled = false;
    knop.textContent = "Kies locatie en exporteer";
  }
}

function toonToast(titel, detail = "", soort = "") {
  const toast = document.getElementById("toast");
  clearTimeout(toastTimer);
  document.getElementById("toast-titel").textContent = titel;
  document.getElementById("toast-detail").textContent = detail;
  toast.className = "toast" + (soort ? " -" + soort : "");
  toast.hidden = false;
  toastTimer = setTimeout(() => { toast.hidden = true; }, 5200);
}

function bouwMeerLijst() {
  const lijst = document.getElementById("meer-lijst");
  if (!lijst || !huidigeData) return;
  lijst.replaceChildren(...TABS.filter((t) => t.groep === "rapporten").map((t) => {
    const aantal = t.telling ? t.telling(huidigeData) : 0;
    const knop = el("button", {
      class: "meer-item", type: "button",
      "aria-current": t.key === actieveTab ? "page" : null,
    }, [
      el("span", { text: t.label }),
      aantal ? el("span", { class: "badge", text: aantal }) : null,
    ]);
    knop.addEventListener("click", () => {
      document.getElementById("meer-menu").close();
      toonTab(t.key);
    });
    return knop;
  }));
}

function commandKnop(titel, context, actie) {
  const knop = el("button", { class: "command-item", type: "button" }, [
    el("span", { class: "command-item-titel", text: titel }),
    context ? el("span", { class: "command-item-context", text: context }) : null,
  ]);
  knop.addEventListener("click", actie);
  return knop;
}

function commandGroep(titel, knoppen) {
  if (!knoppen.length) return null;
  return el("section", { class: "command-groep" }, [
    el("h3", { text: titel }),
    ...knoppen,
  ]);
}

function sluitCommandPalette(herstelFocus = false) {
  const dialoog = document.getElementById("command-palette");
  if (dialoog.open) dialoog.close();
  if (herstelFocus) {
    const doel = commandTerugFocus;
    requestAnimationFrame(() => doel?.isConnected && doel.focus());
  }
}

function renderCommandResultaten(zoekterm = "") {
  const houder = document.getElementById("command-results");
  const term = zoekterm.trim().toLocaleLowerCase("nl-NL");
  const nav = huidigeData ? TABS
    .filter((t) => !term || t.label.toLocaleLowerCase("nl-NL").includes(term))
    .slice(0, term ? 8 : 4)
    .map((t) => commandKnop(t.label, t.groep === "werk" ? "Werk" : "Rapport", () => {
      sluitCommandPalette();
      toonTab(t.key);
    })) : [];
  const acties = [
    commandKnop("Nieuw bankbestand", "Importeer en controleer een ING-export", () => {
      sluitCommandPalette();
      kiesBestand();
    }),
    ...(huidigeData ? [
      commandKnop("Kwartaal exporteren", periodeVan(huidigeData).label, () => {
        sluitCommandPalette();
        openExportDialoog();
      }),
      commandKnop("Bonnen controleren", "Koppel facturen en bonnen", () => {
        sluitCommandPalette();
        toonTab("bonnen");
      }),
    ] : []),
  ].filter((knop) => !term || knop.textContent.toLocaleLowerCase("nl-NL").includes(term));
  const transacties = huidigeData && term ? huidigeData.transacties
    .filter((r) => [r.naam, r.leverancier, r.factuur, r.mededelingen]
      .some((v) => String(v || "").toLocaleLowerCase("nl-NL").includes(term)))
    .slice(0, 6)
    .map((r) => commandKnop(
      r.leverancier || r.naam,
      `${datum(r.datum)} · ${geld(getoondBedrag(r))} · ${r.factuur || "geen factuurnummer"}`,
      () => {
        transactieState.zoek = r.leverancier || r.naam;
        transactieState.filter = "alle";
        transactieState.openId = r.id;
        sluitCommandPalette();
        toonTab("transacties");
      },
    )) : [];
  const groepen = [
    commandGroep("Ga naar", nav),
    commandGroep("Acties", acties),
    commandGroep("Transacties", transacties),
  ].filter(Boolean);
  houder.replaceChildren(...groepen);
  if (!groepen.length) {
    houder.append(el("p", { class: "command-leeg", text: "Geen resultaten gevonden." }));
  }
}

function openCommandPalette() {
  const dialoog = document.getElementById("command-palette");
  const zoek = document.getElementById("command-search");
  if (!dialoog.open) {
    commandTerugFocus = document.activeElement;
    dialoog.showModal();
  }
  zoek.value = "";
  renderCommandResultaten();
  requestAnimationFrame(() => zoek.focus());
}

/* ---------- toegankelijkheid: tekstgrootte, thema, uitleg ---------- */
function zetTekstgrootte(schaal) {
  document.documentElement.style.setProperty("--font", parseFloat(schaal) + "rem");
  document.querySelectorAll("[data-scale]").forEach((b) =>
    b.setAttribute("aria-pressed", b.dataset.scale === String(schaal) ? "true" : "false"));
  try { localStorage.setItem("tekstgrootte", schaal); } catch (e) {}
}

function zetThema(thema) {
  if (thema) document.documentElement.setAttribute("data-theme", thema);
  else document.documentElement.removeAttribute("data-theme");
  try { localStorage.setItem("thema", thema || ""); } catch (e) {}
  werkThemaKnopBij();
}

function themaIsDonker() {
  const gekozen = document.documentElement.getAttribute("data-theme");
  return gekozen ? gekozen === "dark" : matchMedia("(prefers-color-scheme: dark)").matches;
}

function werkThemaKnopBij() {
  const knop = document.getElementById("knop-thema");
  if (knop) knop.textContent = themaIsDonker() ? "Licht thema gebruiken" : "Donker thema gebruiken";
}

function initA11y() {
  const instellingen = document.querySelector(".instellingen");
  const instellingenKnop = document.getElementById("knop-instellingen");
  const instellingenMenu = document.getElementById("instellingen-menu");
  const zetInstellingenOpen = (open, verplaatsFocus = false) => {
    if (!instellingenMenu || !instellingenKnop) return;
    instellingenMenu.hidden = !open;
    instellingenKnop.setAttribute("aria-expanded", open ? "true" : "false");
    if (open && verplaatsFocus) {
      requestAnimationFrame(() => instellingenMenu.querySelector("button")?.focus());
    }
  };

  instellingenKnop?.addEventListener("click", () =>
    zetInstellingenOpen(instellingenMenu.hidden, instellingenMenu.hidden));
  document.addEventListener("pointerdown", (e) => {
    if (instellingen && !instellingen.contains(e.target)) zetInstellingenOpen(false);
  });
  document.addEventListener("keydown", (e) => {
    if (instellingenMenu && e.key === "Escape" && !instellingenMenu.hidden) {
      zetInstellingenOpen(false);
      instellingenKnop.focus();
    }
  });

  document.querySelectorAll("[data-scale]").forEach((b) =>
    b.addEventListener("click", () => zetTekstgrootte(b.dataset.scale)));
  document.getElementById("knop-thema").addEventListener("click", () =>
    zetThema(themaIsDonker() ? "light" : "dark"));
  const help = document.getElementById("knop-help");
  const helpSluiten = document.getElementById("knop-help-sluiten");
  const glos = document.getElementById("glossary");
  help.addEventListener("click", () => {
    zetInstellingenOpen(false);
    glos.showModal();
    help.setAttribute("aria-expanded", "true");
  });
  helpSluiten?.addEventListener("click", () => {
    glos.close();
  });
  glos.addEventListener("close", () => {
    help.setAttribute("aria-expanded", "false");
    (instellingenKnop || help).focus();
  });
  glos.addEventListener("click", (e) => {
    if (e.target === glos) glos.close();
  });
  // Voorkeuren terugzetten
  try {
    const g = localStorage.getItem("tekstgrootte");
    if (g) zetTekstgrootte(g);
    const t = localStorage.getItem("thema");
    if (t) zetThema(t);
  } catch (e) {}
  werkThemaKnopBij();
}

function init() {
  document.getElementById("knop-kies").addEventListener("click", kiesBestand);
  document.getElementById("knop-kies-leeg")?.addEventListener("click", kiesBestand);
  document.getElementById("knop-excel").addEventListener("click", openExportDialoog);
  document.getElementById("knop-zoeken").addEventListener("click", openCommandPalette);
  document.getElementById("import-bevestigen").addEventListener("click", bevestigImport);
  document.getElementById("import-ander").addEventListener("click", kiesBestand);
  document.getElementById("import-annuleren").addEventListener("click", annuleerImport);

  const exportDialoog = document.getElementById("export-dialoog");
  document.getElementById("export-form").addEventListener("submit", (e) => {
    e.preventDefault();
    downloadExcel();
  });
  document.getElementById("export-annuleren").addEventListener("click", () => exportDialoog.close());
  document.getElementById("export-sluiten").addEventListener("click", () => exportDialoog.close());
  exportDialoog.addEventListener("click", (e) => { if (e.target === exportDialoog) exportDialoog.close(); });

  const commandDialoog = document.getElementById("command-palette");
  const commandZoek = document.getElementById("command-search");
  commandZoek.addEventListener("input", () => renderCommandResultaten(commandZoek.value));
  commandZoek.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      document.querySelector("#command-results .command-item")?.focus();
    }
  });
  document.getElementById("command-results").addEventListener("keydown", (e) => {
    if (!["ArrowDown", "ArrowUp"].includes(e.key)) return;
    const knoppen = [...document.querySelectorAll("#command-results .command-item")];
    const huidig = knoppen.indexOf(document.activeElement);
    const volgende = e.key === "ArrowDown"
      ? Math.min(knoppen.length - 1, huidig + 1)
      : Math.max(0, huidig - 1);
    if (knoppen[volgende]) {
      e.preventDefault();
      knoppen[volgende].focus();
    }
  });
  commandDialoog.addEventListener("cancel", (e) => {
    e.preventDefault();
    sluitCommandPalette(true);
  });
  commandDialoog.addEventListener("click", (e) => {
    if (e.target === commandDialoog) sluitCommandPalette(true);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && commandDialoog.open) {
      e.preventDefault();
      sluitCommandPalette(true);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLocaleLowerCase("nl-NL") === "k") {
      e.preventDefault();
      openCommandPalette();
    }
  });

  const meerMenu = document.getElementById("meer-menu");
  document.getElementById("meer-knop").addEventListener("click", () => {
    bouwMeerLijst();
    meerMenu.showModal();
  });
  document.getElementById("meer-sluiten").addEventListener("click", () => meerMenu.close());
  meerMenu.addEventListener("click", (e) => { if (e.target === meerMenu) meerMenu.close(); });

  initA11y();
  werkHeaderActiesBij();
  // Browser-preview zonder Python: toon meegegeven demo-data als die er is.
  if (!heeftApi() && window.__DEMO_MODEL__) {
    renderModel(window.__DEMO_MODEL__);
    toonToast("Voorbeeldgegevens", "Dit is een lokale ontwerpweergave.", "ok");
  }
}

window.addEventListener("pywebviewready", () => {
  if (window.pywebview.api) {
    window.pywebview.api.huidige_data().then((r) => { if (r.ok) renderModel(r.data); });
  }
});
document.addEventListener("DOMContentLoaded", init);
