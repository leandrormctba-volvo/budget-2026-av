/* Budget 2026 — Associação Volvo
   Lê os dados exportados do "Painel Gerencial.xlsx" (assets/data.js) e monta a apresentação. */
(function () {
  "use strict";

  const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  const MESES_LONGOS = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
  const COR = { previsto: "#8fc4d1", realizado: "#0c5868", bom: "#157d93", ruim: "#d95926", grade: "#e8ecef", texto: "#5d6476" };

  /* ---------- 1. Dados ---------- */
  const brutos = window.BUDGET_RAW.map(r => r.value || r);
  const linhas = [];
  brutos.slice(2).forEach(v => {
    const nome = String(v[0] || "").trim();
    if (!nome) return;
    const p = [], r = [];
    for (let i = 0; i < 12; i++) { p.push(+v[1 + 3 * i] || 0); r.push(+v[2 + 3 * i] || 0); }
    const m = /^([\d.]+)\s*-\s*(.*)$/.exec(nome);
    const cod = m ? m[1] : null;
    const partes = cod ? cod.split(".") : [];
    linhas.push({
      cod, nome: m ? m[2].trim() : nome, nivel: partes.length,
      raiz: partes[0], grupo: partes.length >= 2 ? partes[0] + "." + partes[1] : null,
      tipo: partes[0] === "1" ? "rec" : "desp", p, r
    });
  });

  const porCod = {};
  linhas.forEach(l => { if (l.cod) porCod[l.cod] = l; });
  const grupos = linhas.filter(l => l.nivel === 2);
  grupos.forEach(g => { g.itens = linhas.filter(l => l.nivel === 3 && l.grupo === g.cod); });

  const somaLinha = (a, b) => ({ p: a.p.map((x, i) => x + b.p[i]), r: a.r.map((x, i) => x + b.r[i]) });
  const REC = porCod["1"];
  const DESP = Object.assign({ nome: "Despesas", tipo: "desp" }, somaLinha(porCod["2"], porCod["6"] || { p: Array(12).fill(0), r: Array(12).fill(0) }));
  const RES = { nome: "Resultado", tipo: "rec", p: REC.p.map((x, i) => x - DESP.p[i]), r: REC.r.map((x, i) => x - DESP.r[i]) };

  // último mês com realizado lançado
  let ultimoMes = 0;
  for (let i = 0; i < 12; i++) if (REC.r[i] !== 0 || DESP.r[i] !== 0) ultimoMes = i;

  /* ---------- 2. Período ---------- */
  let periodo = "ano"; // "ano" = previsto do ano todo | "ytd" = previsto só até o último mês lançado
  const nMeses = () => (periodo === "ano" ? 12 : ultimoMes + 1);
  const soma = (arr, n) => arr.slice(0, n).reduce((a, b) => a + b, 0);
  function totais(l) {
    const P = soma(l.p, nMeses());
    const R = soma(l.r, 12);
    const S = l.tipo === "rec" ? R - P : P - R; // positivo = favorável
    return { P, R, S, exec: P ? R / P : null };
  }
  const saldoMes = (l, i) => (l.tipo === "rec" ? l.r[i] - l.p[i] : l.p[i] - l.r[i]);

  /* ---------- 3. Formatação ---------- */
  const fmtBRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
  const fmtBRL2 = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2 });
  function curto(v) {
    const a = Math.abs(v), s = v < 0 ? "−" : "";
    if (a >= 1e6) return s + "R$ " + (a / 1e6).toLocaleString("pt-BR", { maximumFractionDigits: 2 }) + " mi";
    if (a >= 1e3) return s + "R$ " + (a / 1e3).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + " mil";
    return s + "R$ " + a.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
  }
  const inteiro = v => fmtBRL.format(v).replace("-", "−");
  const pct = v => (v == null ? "—" : (v * 100).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + "%");
  const sinal = v => (v > 0 ? "+" : "");
  const classe = v => (v > 0.5 ? "pos" : v < -0.5 ? "neg" : "");
  const MINUSC = new Set(["de", "da", "do", "das", "dos", "e", "com", "s/", "a", "o"]);
  function titulo(s) {
    if (s !== s.toUpperCase()) return s;
    return s.toLowerCase().split(" ").map((w, i) => (i && MINUSC.has(w) ? w : w.charAt(0).toUpperCase() + w.slice(1))).join(" ").replace(/\bAv\b/, "AV");
  }
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const semAcento = s => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();

  /* ---------- 4. Gráficos ---------- */
  Chart.defaults.font.family = "Montserrat, system-ui, sans-serif";
  Chart.defaults.font.size = 11.5;
  Chart.defaults.color = COR.texto;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = "#06313a";
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.callbacks.label = ctx => " " + ctx.dataset.label + ": " + fmtBRL2.format(ctx.parsed.y ?? ctx.parsed.x ?? 0);
  const graficos = {};
  const eixoValor = { grid: { color: COR.grade }, border: { display: false }, ticks: { callback: v => curto(v).replace("R$ ", "") } };
  const eixoCat = { grid: { display: false }, border: { color: COR.grade } };

  function desenhar(id, config) {
    if (graficos[id]) graficos[id].destroy();
    graficos[id] = new Chart(document.getElementById(id), config);
  }
  const mesLabels = () => MESES.map((m, i) => (i === ultimoMes && ultimoMes < 11 ? m + "*" : m));
  const realizadoMeses = arr => arr.map((v, i) => (i > ultimoMes ? null : v));

  function grafMensal(id, l, opcoes = {}) {
    desenhar(id, {
      type: "bar",
      data: {
        labels: mesLabels(),
        datasets: [
          { label: "Previsto", data: l.p, backgroundColor: COR.previsto, borderRadius: 4, maxBarThickness: 28 },
          { label: "Realizado", data: realizadoMeses(l.r), backgroundColor: COR.realizado, borderRadius: 4, maxBarThickness: 28 }
        ].concat(opcoes.extra || [])
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: { legend: { position: "top", align: "end" } },
        scales: { x: eixoCat, y: Object.assign({ beginAtZero: true }, eixoValor), ...(opcoes.escalas || {}) }
      }
    });
  }

  /* ---------- 5. Slides ---------- */
  function renderVisaoGeral() {
    const blocos = [
      { l: REC, nome: "Receitas", rotSaldo: "Saldo" },
      { l: DESP, nome: "Despesas", rotSaldo: "Economia" },
      { l: RES, nome: "Resultado", rotSaldo: "Saldo", destaque: true }
    ];
    document.getElementById("kpis").innerHTML = blocos.map(b => {
      const t = totais(b.l);
      const larg = t.exec == null ? 0 : Math.max(0, Math.min(100, t.exec * 100));
      return `<div class="kpi ${b.destaque ? "destaque" : ""}">
        <div class="kpi-nome">${b.nome}</div>
        <div class="kpi-linhas">
          <div><div class="kpi-rot">Previsto</div><div class="kpi-val">${curto(t.P)}</div></div>
          <div><div class="kpi-rot">Realizado</div><div class="kpi-val">${curto(t.R)}</div></div>
          <div><div class="kpi-rot">${b.rotSaldo}</div><div class="kpi-val ${classe(t.S)}">${sinal(t.S)}${curto(t.S)}</div></div>
        </div>
        <div class="kpi-barra"><span style="width:${larg}%"></span></div>
        <div class="kpi-exec">Execução: <b>${pct(t.exec)}</b> do previsto</div>
      </div>`;
    }).join("");
    grafMensal("gRecMes", REC);
    grafMensal("gDespMes", DESP);
  }

  function renderResultado() {
    let acP = 0, acR = 0;
    const acumP = RES.p.map(v => (acP += v));
    const acumR = RES.r.map((v, i) => (i > ultimoMes ? null : (acR += v)));
    grafMensal("gResultado", RES, {
      extra: [
        { type: "line", label: "Acumulado previsto", data: acumP, borderColor: COR.previsto, borderDash: [5, 4], pointRadius: 0, tension: .3, yAxisID: "y2" },
        { type: "line", label: "Acumulado realizado", data: acumR, borderColor: COR.bom, backgroundColor: COR.bom, pointRadius: 3, tension: .3, yAxisID: "y2" }
      ],
      escalas: { y: Object.assign({}, eixoValor, { beginAtZero: true }), y2: Object.assign({}, eixoValor, { position: "right", grid: { display: false } }) }
    });
    const lin = (rot, f, cls = "") => `<tr class="${cls}"><td>${rot}</td>${MESES.map((m, i) => `<td>${f(i)}</td>`).join("")}</tr>`;
    const fut = i => i > ultimoMes;
    document.getElementById("tabResultado").innerHTML = `<table><thead><tr><th></th>${mesLabels().map(m => `<th>${m}</th>`).join("")}</tr></thead><tbody>
      ${lin("Receitas realizadas", i => (fut(i) ? "—" : curto(REC.r[i])))}
      ${lin("Despesas realizadas", i => (fut(i) ? "—" : curto(DESP.r[i])))}
      ${lin("Resultado previsto", i => curto(RES.p[i]))}
      ${lin("Resultado realizado", i => (fut(i) ? "—" : `<span class="${classe(RES.r[i])}">${curto(RES.r[i])}</span>`), "total")}
    </tbody></table>`;
  }

  function renderGrupos(tipo, idGraf, idTab) {
    const gs = grupos.filter(g => g.tipo === tipo).map(g => ({ g, t: totais(g) })).filter(x => x.t.P || x.t.R)
      .sort((a, b) => Math.max(b.t.P, b.t.R) - Math.max(a.t.P, a.t.R));
    desenhar(idGraf, {
      type: "bar",
      data: {
        labels: gs.map(x => titulo(x.g.nome)),
        datasets: [
          { label: "Previsto", data: gs.map(x => x.t.P), backgroundColor: COR.previsto, borderRadius: 4, maxBarThickness: 16 },
          { label: "Realizado", data: gs.map(x => x.t.R), backgroundColor: COR.realizado, borderRadius: 4, maxBarThickness: 16 }
        ]
      },
      options: {
        indexAxis: "y", responsive: true, maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false, axis: "y" },
        plugins: { legend: { position: "top", align: "end" } },
        scales: { x: Object.assign({ beginAtZero: true }, eixoValor), y: Object.assign({}, eixoCat, { ticks: { autoSkip: false, callback(v) { const s = this.getLabelForValue(v); return s.length > 26 ? s.slice(0, 25) + "…" : s; } } }) },
        onClick: (e, els) => { if (els.length) abrirDetalhe(gs[els[0].index].g); },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? "pointer" : "default"; }
      }
    });
    const tot = totais(tipo === "rec" ? REC : DESP);
    document.getElementById(idTab).innerHTML = `<table><thead><tr><th>Grupo</th><th>Previsto</th><th>Realizado</th><th>${tipo === "rec" ? "Saldo" : "Economia"}</th><th>Execução</th></tr></thead><tbody>
      ${gs.map(x => linhaTabela(x.g, x.t)).join("")}
      <tr class="total"><td>Total</td><td>${curto(tot.P)}</td><td>${curto(tot.R)}</td><td class="${classe(tot.S)}">${sinal(tot.S)}${curto(tot.S)}</td><td>${pct(tot.exec)}</td></tr>
    </tbody></table>`;
    document.querySelectorAll("#" + idTab + " tr[data-cod]").forEach(tr => tr.addEventListener("click", () => abrirDetalhe(porCod[tr.dataset.cod])));
  }
  function linhaTabela(l, t) {
    const larg = t.exec == null ? (t.R ? 100 : 0) : Math.min(100, t.exec * 100);
    const acima = l.tipo === "desp" ? t.R > t.P : false;
    return `<tr class="clicavel" data-cod="${l.cod}"><td>${esc(titulo(l.nome))}</td><td>${curto(t.P)}</td><td>${curto(t.R)}</td>
      <td class="${classe(t.S)}">${sinal(t.S)}${curto(t.S)}</td>
      <td><span class="mini-barra"><span class="${acima ? "acima" : ""}" style="width:${larg}%"></span></span>${pct(t.exec)}</td></tr>`;
  }

  let aba = "desp";
  function cartaoAtiv(l, ehItem) {
    const t = totais(l);
    const larg = t.exec == null ? (t.R ? 100 : 0) : Math.min(100, t.exec * 100);
    const acima = l.tipo === "desp" && t.R > t.P + 0.5;
    const pai = ehItem && porCod[l.grupo] ? titulo(porCod[l.grupo].nome) + " · " : "";
    return `<button class="ativ ${ehItem ? "item" : ""}" data-cod="${l.cod}">
      <div class="ativ-topo"><div><div class="ativ-cod">${esc(pai)}${l.cod}</div><div class="ativ-nome">${esc(titulo(l.nome))}</div></div><span class="ativ-seta">→</span></div>
      <div class="ativ-nums"><span>Previsto<b>${curto(t.P)}</b></span><span style="text-align:right">Realizado<b>${curto(t.R)}</b></span></div>
      <div class="ativ-barra-exec"><span class="${acima ? "acima" : ""}" style="width:${larg}%"></span></div>
      <div class="ativ-rodape"><span>${pct(t.exec)} executado</span><span class="${classe(t.S)}">${sinal(t.S)}${curto(t.S)}</span></div>
    </button>`;
  }
  function renderAtividades() {
    const q = semAcento(document.getElementById("busca").value.trim());
    const grade = document.getElementById("ativGrade");
    let html;
    if (q) {
      const achados = linhas.filter(l => l.nivel >= 2 && (semAcento(l.nome).includes(q) || l.cod.startsWith(q)) && (soma(l.p, 12) || soma(l.r, 12)));
      html = achados.length ? achados.map(l => cartaoAtiv(l, l.nivel === 3)).join("") : `<div class="vazio">Nenhuma atividade encontrada para “${esc(q)}”.</div>`;
    } else {
      html = grupos.filter(g => g.tipo === aba && (soma(g.p, 12) || soma(g.r, 12)))
        .sort((a, b) => Math.max(soma(b.p, 12), soma(b.r, 12)) - Math.max(soma(a.p, 12), soma(a.r, 12)))
        .map(g => cartaoAtiv(g, false)).join("");
    }
    grade.innerHTML = html;
    grade.querySelectorAll(".ativ").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));
  }

  function renderDestaques() {
    const itens = linhas.filter(l => l.nivel === 3 && l.tipo === "desp").map(l => ({ l, t: totais(l) }));
    const ruins = itens.filter(x => x.t.S < -0.5).sort((a, b) => a.t.S - b.t.S).slice(0, 6);
    const bons = itens.filter(x => x.t.S > 0.5).sort((a, b) => b.t.S - a.t.S).slice(0, 6);
    const max = Math.max(...ruins.map(x => -x.t.S), ...bons.map(x => x.t.S), 1);
    const item = (x, cor) => `<button class="dest" data-cod="${x.l.cod}">
      <span class="dest-nome">${esc(titulo(x.l.nome))}</span><span class="dest-val ${classe(x.t.S)}">${sinal(x.t.S)}${curto(x.t.S)}</span>
      <span class="dest-sub">Previsto ${curto(x.t.P)} · Realizado ${curto(x.t.R)}</span><span></span>
      <span class="dest-barra"><span style="width:${(Math.abs(x.t.S) / max) * 100}%;background:${cor}"></span></span></button>`;
    document.getElementById("destRuim").innerHTML = ruins.map(x => item(x, COR.ruim)).join("") || `<div class="vazio">Nenhum item acima do previsto.</div>`;
    document.getElementById("destBom").innerHTML = bons.map(x => item(x, COR.bom)).join("");
    document.querySelectorAll(".dest").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));
  }

  /* ---------- 6. Detalhe (modal) ---------- */
  const modal = document.getElementById("modal");
  let atual = null;
  function abrirDetalhe(l) {
    if (!l) return;
    atual = l;
    const grupo = l.nivel === 3 ? porCod[l.grupo] : l;
    document.getElementById("mGrupo").textContent = (l.tipo === "rec" ? "Receitas" : "Despesas") + (l.nivel === 3 ? " · " + titulo(grupo.nome) : "");
    document.getElementById("mTitulo").textContent = (l.nivel === 3 ? l.cod + " · " : "") + titulo(l.nome);

    const filhos = (grupo.itens || []).filter(i => soma(i.p, 12) || soma(i.r, 12));
    document.getElementById("mChips").innerHTML = filhos.length
      ? `<button data-cod="${grupo.cod}" class="${l === grupo ? "ativo" : ""}">Total do grupo</button>` + filhos.map(i => `<button data-cod="${i.cod}" class="${l === i ? "ativo" : ""}">${esc(titulo(i.nome))}</button>`).join("")
      : "";
    document.querySelectorAll("#mChips button").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));

    const t = totais(l);
    const rotS = l.tipo === "rec" ? "Saldo" : "Economia";
    document.getElementById("mKpis").innerHTML = [
      ["Previsto", curto(t.P), ""], ["Realizado", curto(t.R), ""], [rotS, sinal(t.S) + curto(t.S), classe(t.S)], ["Execução", pct(t.exec), ""]
    ].map(k => `<div class="kpi"><div class="kpi-rot">${k[0]}</div><div class="kpi-val ${k[2]}">${k[1]}</div></div>`).join("");

    grafMensal("gModal", l);
    let acS = 0;
    const linhasT = MESES.map((m, i) => {
      const fut = i > ultimoMes;
      const s = saldoMes(l, i);
      if (!fut) acS += s;
      return `<tr class="${fut ? "futuro" : ""}"><td>${MESES_LONGOS[i].replace(/^./, c => c.toUpperCase())}${i === ultimoMes && ultimoMes < 11 ? " *" : ""}</td>
        <td>${inteiro(l.p[i])}</td><td>${fut ? "—" : inteiro(l.r[i])}</td>
        <td class="${fut ? "" : classe(s)}">${fut ? "—" : sinal(s) + inteiro(s)}</td>
        <td class="${fut ? "" : classe(acS)}">${fut ? "—" : sinal(acS) + inteiro(acS)}</td>
        <td>${fut || !l.p[i] ? "—" : pct(l.r[i] / l.p[i])}</td></tr>`;
    }).join("");
    document.getElementById("mTabela").innerHTML = `<table><thead><tr><th>Mês</th><th>Previsto</th><th>Realizado</th><th>${rotS} no mês</th><th>${rotS} acum.</th><th>Execução</th></tr></thead>
      <tbody>${linhasT}<tr class="total"><td>Total ${periodo === "ano" ? "2026" : "acumulado"}</td><td>${inteiro(t.P)}</td><td>${inteiro(t.R)}</td><td class="${classe(t.S)}">${sinal(t.S)}${inteiro(t.S)}</td><td></td><td>${pct(t.exec)}</td></tr></tbody></table>`;

    modal.classList.add("aberto");
    modal.setAttribute("aria-hidden", "false");
  }
  function fecharDetalhe() { modal.classList.remove("aberto"); modal.setAttribute("aria-hidden", "true"); atual = null; }
  document.getElementById("mFechar").addEventListener("click", fecharDetalhe);
  modal.addEventListener("click", e => { if (e.target === modal) fecharDetalhe(); });

  /* ---------- 7. Navegação ---------- */
  const slides = [...document.querySelectorAll(".slide")];
  const pontos = document.getElementById("pontos");
  let idx = 0;
  pontos.innerHTML = slides.map((s, i) => `<button aria-label="${s.dataset.titulo}" title="${s.dataset.titulo}" data-i="${i}"></button>`).join("");
  pontos.querySelectorAll("button").forEach(b => b.addEventListener("click", () => ir(+b.dataset.i)));
  function ir(n) {
    idx = Math.max(0, Math.min(slides.length - 1, n));
    slides.forEach((s, i) => { s.classList.toggle("ativo", i === idx); s.classList.toggle("antes", i < idx); });
    pontos.querySelectorAll("button").forEach((b, i) => b.classList.toggle("ativo", i === idx));
    document.body.classList.toggle("na-capa", slides[idx].classList.contains("capa"));
    document.getElementById("ant").disabled = idx === 0;
    document.getElementById("prox").disabled = idx === slides.length - 1;
    document.getElementById("contador").textContent = `${idx + 1} / ${slides.length} · ${slides[idx].dataset.titulo}`;
    slides[idx].scrollTop = 0;
    // re-anima os gráficos do slide que entrou
    slides[idx].querySelectorAll("canvas").forEach(c => { const g = graficos[c.id]; if (g) { g.reset(); g.update(); } });
    try { history.replaceState(null, "", "#" + (idx + 1)); } catch (e) {}
  }
  document.getElementById("ant").addEventListener("click", () => ir(idx - 1));
  document.getElementById("prox").addEventListener("click", () => ir(idx + 1));
  document.addEventListener("keydown", e => {
    if (e.target.tagName === "INPUT") { if (e.key === "Escape") e.target.blur(); return; }
    if (e.key === "Escape" && atual) return fecharDetalhe();
    if (atual) return;
    if (["ArrowRight", "PageDown", " "].includes(e.key)) { e.preventDefault(); ir(idx + 1); }
    if (["ArrowLeft", "PageUp"].includes(e.key)) { e.preventDefault(); ir(idx - 1); }
    if (e.key === "Home") ir(0);
    if (e.key === "End") ir(slides.length - 1);
    if (e.key.toLowerCase() === "f") telaCheia();
  });
  let toqueX = null;
  document.getElementById("palco").addEventListener("touchstart", e => { toqueX = e.touches[0].clientX; }, { passive: true });
  document.getElementById("palco").addEventListener("touchend", e => {
    if (toqueX == null) return;
    const dx = e.changedTouches[0].clientX - toqueX; toqueX = null;
    if (Math.abs(dx) > 70) ir(idx + (dx < 0 ? 1 : -1));
  });
  function telaCheia() {
    try { if (!document.fullscreenElement) document.documentElement.requestFullscreen(); else document.exitFullscreen(); } catch (e) {}
  }
  document.getElementById("btnTela").addEventListener("click", telaCheia);

  /* ---------- 8. Controles ---------- */
  document.querySelectorAll(".periodo button").forEach(b => b.addEventListener("click", () => {
    periodo = b.dataset.periodo;
    document.querySelectorAll(".periodo button").forEach(x => x.classList.toggle("ativo", x === b));
    renderTudo();
  }));
  document.querySelectorAll(".abas button").forEach(b => b.addEventListener("click", () => {
    aba = b.dataset.aba;
    document.querySelectorAll(".abas button").forEach(x => x.classList.toggle("ativo", x === b));
    document.getElementById("busca").value = "";
    renderAtividades();
  }));
  document.getElementById("busca").addEventListener("input", renderAtividades);

  function renderTudo() {
    const mesNome = MESES_LONGOS[ultimoMes];
    document.getElementById("btnYtd").textContent = "Até " + MESES[ultimoMes];
    document.getElementById("notaPeriodo").innerHTML = (periodo === "ano"
      ? `Previsto do <b>ano inteiro</b> comparado ao realizado lançado até ${mesNome}.`
      : `Previsto e realizado <b>acumulados de janeiro a ${mesNome}</b>.`)
      + (ultimoMes < 11 ? ` * ${MESES[ultimoMes]}: mês com lançamentos parciais.` : "");
    renderVisaoGeral();
    renderResultado();
    renderGrupos("rec", "gRecGrupo", "tabRecGrupo");
    renderGrupos("desp", "gDespGrupo", "tabDespGrupo");
    renderAtividades();
    renderDestaques();
    if (atual) abrirDetalhe(atual);
  }

  document.getElementById("capaData").textContent = "Realizado lançado até " + MESES_LONGOS[ultimoMes] + " de 2026";
  const inicial = parseInt(location.hash.slice(1), 10);
  ir(isNaN(inicial) ? 0 : inicial - 1);
  // espera a fonte Montserrat carregar para os gráficos medirem os rótulos corretamente
  const iniciar = () => { renderTudo(); ir(idx); };
  (document.fonts && document.fonts.ready ? Promise.race([document.fonts.ready, new Promise(r => setTimeout(r, 2500))]) : Promise.resolve()).then(iniciar);
})();
