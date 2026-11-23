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

  // ordena códigos numericamente: 1.1, 1.2, …, 1.12 (e não 1.1, 1.12, 1.2)
  const compCod = (a, b) => {
    const x = a.cod.split(".").map(Number), y = b.cod.split(".").map(Number);
    for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d; }
    return 0;
  };
  const porCod = {};
  linhas.forEach(l => { if (l.cod) porCod[l.cod] = l; });
  const grupos = linhas.filter(l => l.nivel === 2).sort(compCod);
  grupos.forEach(g => { g.itens = linhas.filter(l => l.nivel === 3 && l.grupo === g.cod).sort(compCod); });

  const REC = porCod["1"];               // linha "1 - RECEITAS" da planilha
  const DESP = porCod["2"];              // linha "2 - DESPESAS" da planilha
  const EXTRA = porCod["6"];             // linha "6 - DESPESAS EXTRAORDINÁRIAS"
  const RES = {                          // linha "Resultado" = receitas − despesas − extraordinárias
    nome: "Resultado", tipo: "rec",
    p: REC.p.map((x, i) => x - DESP.p[i] - (EXTRA ? EXTRA.p[i] : 0)),
    r: REC.r.map((x, i) => x - DESP.r[i] - (EXTRA ? EXTRA.r[i] : 0))
  };
  const vivo = l => l.p.some(Boolean) || l.r.some(Boolean);

  // último mês com realizado lançado
  let ultimoMes = 0;
  for (let i = 0; i < 12; i++) if (REC.r[i] !== 0 || DESP.r[i] !== 0) ultimoMes = i;
  const parcial = ultimoMes < 11; // o último mês lançado ainda está em andamento

  /* ---------- 2. Período ---------- */
  // periodo = "ano" (ano inteiro) ou o índice do mês (0 = Jan … 11 = Dez): acumulado de janeiro até esse mês
  let periodo = "ano";
  const nMeses = () => (periodo === "ano" ? 12 : periodo + 1);
  const rotuloPeriodo = () => (periodo === "ano" ? "2026" : periodo === 0 ? "Jan" : "Jan–" + MESES[periodo]);
  const soma = (arr, n) => arr.slice(0, n).reduce((a, b) => a + b, 0);
  const saldoMes = (l, i) => (l.tipo === "rec" ? l.r[i] - l.p[i] : l.p[i] - l.r[i]); // positivo = favorável
  function totais(l, n = nMeses()) {
    const P = soma(l.p, n), R = soma(l.r, n);
    const S = l.tipo === "rec" ? R - P : P - R;
    return { P, R, S, exec: P ? R / P : null };
  }

  /* ---------- 3. Formatação ---------- */
  const fmtBRL2 = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2 });
  const exato = v => fmtBRL2.format(v).replace("-", "−");
  function curto(v) {
    const a = Math.abs(v), s = v < 0 ? "−" : "";
    if (a >= 1e6) return s + "R$ " + (a / 1e6).toLocaleString("pt-BR", { maximumFractionDigits: 2 }) + " mi";
    if (a >= 1e3) return s + "R$ " + (a / 1e3).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + " mil";
    return s + "R$ " + a.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
  }
  const pct = v => (v == null || !isFinite(v) ? "—" : (v * 100).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + "%");
  const sinal = v => (v > 0.005 ? "+" : "");
  const classe = v => (v > 0.5 ? "pos" : v < -0.5 ? "neg" : "");
  const MINUSC = new Set(["de", "da", "do", "das", "dos", "e", "com", "s/", "a", "o"]);
  function titulo(s) {
    if (s !== s.toUpperCase()) return s;
    return s.toLowerCase().split(" ").map((w, i) => (i && MINUSC.has(w) ? w : w.charAt(0).toUpperCase() + w.slice(1))).join(" ").replace(/\bAv\b/, "AV");
  }
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const semAcento = s => s.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  const nomeMes = i => MESES_LONGOS[i].replace(/^./, c => c.toUpperCase());
  const faixa = (a, b) => (a === b ? MESES[a] : MESES[a] + "–" + MESES[b]);

  /* ---------- 4. Gráficos ---------- */
  Chart.defaults.font.family = "Montserrat, system-ui, sans-serif";
  Chart.defaults.font.size = 11.5;
  Chart.defaults.color = COR.texto;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.boxWidth = 8;
  Chart.defaults.plugins.tooltip.backgroundColor = "#06313a";
  Chart.defaults.plugins.tooltip.padding = 12;
  Chart.defaults.plugins.tooltip.titleFont = { weight: "700", size: 12.5 };
  Chart.defaults.plugins.tooltip.bodyFont = { size: 12.5 };
  // valor exato (com centavos) ao passar o mouse — funciona em barras verticais e horizontais
  Chart.defaults.plugins.tooltip.callbacks.label = ctx => {
    const horizontal = ctx.chart.options.indexAxis === "y";
    const v = horizontal ? ctx.parsed.x : ctx.parsed.y;
    return v == null ? null : " " + ctx.dataset.label + ": " + exato(v);
  };
  const graficos = {};
  const eixoValor = { grid: { color: COR.grade }, border: { display: false }, ticks: { callback: v => curto(v).replace("R$ ", "") } };
  const eixoCat = { grid: { display: false }, border: { color: COR.grade } };

  function desenhar(id, config) {
    if (graficos[id]) graficos[id].destroy();
    graficos[id] = new Chart(document.getElementById(id), config);
  }
  const mesLabels = () => MESES.map((m, i) => (i === ultimoMes && parcial ? m + "*" : m));
  const realizadoMeses = arr => arr.map((v, i) => (i > ultimoMes ? null : v));

  // tooltip extra: saldo do mês
  function rodapeSaldo(l) {
    return items => {
      const i = items[0].dataIndex;
      if (i > ultimoMes) return "Ainda não lançado";
      const s = saldoMes(l, i);
      return "Saldo: " + sinal(s) + exato(s) + (l.p[i] ? "  ·  " + pct(l.r[i] / l.p[i]) + " executado" : "");
    };
  }
  function grafMensal(id, l) {
    desenhar(id, {
      type: "bar",
      data: {
        labels: mesLabels(),
        datasets: [
          { label: "Previsto", data: l.p, backgroundColor: COR.previsto, borderRadius: 4, maxBarThickness: 28 },
          { label: "Realizado", data: realizadoMeses(l.r), backgroundColor: COR.realizado, borderRadius: 4, maxBarThickness: 28 }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: { legend: { position: "top", align: "end" }, tooltip: { callbacks: { footer: rodapeSaldo(l) } } },
        scales: { x: eixoCat, y: Object.assign({ beginAtZero: true }, eixoValor) }
      }
    });
  }

  /* ---------- 5. Visão geral ---------- */
  function cartaoKpi(nome, l, opc = {}) {
    const t = totais(l);
    const larg = t.exec == null ? 0 : Math.max(0, Math.min(100, t.exec * 100));
    return `<div class="kpi ${opc.destaque ? "destaque" : ""}">
      <div class="kpi-nome">${nome}</div>
      <div class="kpi-linhas">
        <div><div class="kpi-rot">Previsto</div><div class="kpi-val" title="${exato(t.P)}">${curto(t.P)}</div></div>
        <div><div class="kpi-rot">Realizado</div><div class="kpi-val" title="${exato(t.R)}">${curto(t.R)}</div></div>
        <div><div class="kpi-rot">Saldo</div><div class="kpi-val ${classe(t.S)}" title="${exato(t.S)}">${sinal(t.S)}${curto(t.S)}</div></div>
      </div>
      <div class="kpi-barra"><span style="width:${larg}%"></span></div>
      <div class="kpi-exec">Execução: <b>${pct(t.exec)}</b> do previsto${opc.obs ? " · " + opc.obs : ""}</div>
    </div>`;
  }
  function renderVisaoGeral() {
    const te = EXTRA ? totais(EXTRA) : null;
    document.getElementById("kpis").innerHTML =
      cartaoKpi("1 · Receitas", REC) +
      cartaoKpi("2 · Despesas", DESP, { obs: "saldo positivo = economia" }) +
      cartaoKpi("Resultado", RES, { destaque: true, obs: te ? "já desconta extraordinárias (" + curto(te.R) + ")" : "" });
    grafMensal("gRecMesVG", REC);
    grafMensal("gDespMesVG", DESP);
  }

  /* ---------- 6. Receitas / Despesas mês a mês ---------- */
  function renderMensal(l, pre) {
    const t = totais(l);
    const rotPer = "Total " + rotuloPeriodo();
    document.getElementById(pre + "Kpis").innerHTML = `
      <div class="kpi"><div class="kpi-nome">Previsto · ${rotPer}</div><div class="kpi-val grande">${exato(t.P)}</div></div>
      <div class="kpi"><div class="kpi-nome">Realizado · ${rotPer}</div><div class="kpi-val grande">${exato(t.R)}</div><div class="kpi-exec">${pct(t.exec)} do previsto</div></div>
      <div class="kpi destaque"><div class="kpi-nome">Saldo · ${rotPer}</div><div class="kpi-val grande ${classe(t.S)}">${sinal(t.S)}${exato(t.S)}</div>
        <div class="kpi-exec">${l.tipo === "rec" ? (t.S < 0 ? "entrou menos do que o previsto" : "entrou mais do que o previsto") : (t.S >= 0 ? "economia em relação ao previsto" : "gastamos mais do que o previsto")}</div></div>`;
    grafMensal("g" + pre.charAt(0).toUpperCase() + pre.slice(1) + "Mes", l);

    const n = nMeses();
    const corpo = MESES.map((m, i) => {
      const fora = i >= n;
      const naoLancado = i > ultimoMes;
      const s = saldoMes(l, i);
      const marca = naoLancado ? ' <span class="selo">não lançado</span>' : i === ultimoMes && parcial ? ' <span class="selo">parcial</span>' : "";
      return `<tr class="${fora ? "futuro" : naoLancado ? "pendente" : ""}"><td>${nomeMes(i)}${marca}</td>
        <td>${exato(l.p[i])}</td><td>${exato(l.r[i])}</td><td class="${fora ? "" : classe(s)}">${sinal(s)}${exato(s)}</td>
        <td>${l.p[i] ? pct(l.r[i] / l.p[i]) : "—"}</td></tr>`;
    }).join("");
    document.getElementById(pre + "Tab").innerHTML = `<table><thead><tr><th>Mês</th><th>Previsto</th><th>Realizado</th><th>Saldo</th><th>Exec.</th></tr></thead>
      <tbody>${corpo}<tr class="total"><td>${rotPer}</td><td>${exato(t.P)}</td><td>${exato(t.R)}</td><td class="${classe(t.S)}">${sinal(t.S)}${exato(t.S)}</td><td>${pct(t.exec)}</td></tr></tbody></table>`;

    document.getElementById(pre + "Ins").innerHTML = insights(l, t);
  }

  function insights(l, t) {
    const rec = l.tipo === "rec";
    const li = [];
    // (a) explicação por mês
    if (periodo === "ano" && ultimoMes < 11) {
      let futuro = 0;
      for (let i = ultimoMes + 1; i < 12; i++) futuro += saldoMes(l, i);
      li.push(`<b>${faixa(ultimoMes + 1, 11)} ainda não ${ultimoMes === 10 ? "foi lançado" : "foram lançados"}:</b> ${sinal(futuro)}${curto(futuro)} do saldo (${pct(futuro / t.S)} do total) é só o ${rec ? "valor previsto que ainda vai entrar" : "valor previsto que ainda vai ser gasto"}.`);
    }
    const fim = nMeses() - 1; // último mês incluído no período escolhido
    if (parcial && fim >= ultimoMes) {
      const s = saldoMes(l, ultimoMes);
      li.push(`<b>${nomeMes(ultimoMes)} está com lançamentos parciais:</b> ${sinal(s)}${curto(s)} até agora (${pct(l.p[ultimoMes] ? l.r[ultimoMes] / l.p[ultimoMes] : null)} do previsto do mês).`);
    }
    const fechados = parcial ? ultimoMes : ultimoMes + 1;
    if (periodo !== "ano" && fim < 11) {
      li.push(`<b>Período analisado: ${faixa(0, fim)}.</b> ${faixa(fim + 1, 11)} não ${fim === 10 ? "entra" : "entram"} nesta comparação.`);
    } else if (fechados > 0) {
      const tf = totais(l, fechados);
      li.push(`<b>Nos meses fechados (${faixa(0, fechados - 1)}):</b> previsto ${curto(tf.P)}, realizado ${curto(tf.R)} → saldo de <span class="${classe(tf.S)}">${sinal(tf.S)}${curto(tf.S)}</span> (${pct(tf.exec)} executado).`);
    }
    // (b) explicação por item
    const itens = linhas.filter(x => x.nivel === 3 && x.raiz === (rec ? "1" : "2") && vivo(x)).map(x => ({ x, t: totais(x) }));
    const neg = itens.filter(i => i.t.S < -0.5).sort((a, b) => a.t.S - b.t.S).slice(0, 4);
    const pos = itens.filter(i => i.t.S > 0.5).sort((a, b) => b.t.S - a.t.S).slice(0, 3);
    const fmtLista = arr => arr.map(i => `${esc(titulo(i.x.nome))} <span class="${classe(i.t.S)}">(${sinal(i.t.S)}${curto(i.t.S)})</span>`).join(", ");
    if (neg.length) li.push(`<b>${rec ? "Itens que mais puxaram o saldo para baixo" : "Itens com maior realizado acima do previsto"}:</b> ${fmtLista(neg)}.`);
    if (pos.length) li.push(`<b>${rec ? "Itens que compensaram (entraram acima do previsto)" : "Maiores economias"}:</b> ${fmtLista(pos)}.`);

    const tituloIns = rec
      ? `Por que o saldo das receitas está em ${sinal(t.S)}${curto(t.S)}?`
      : `Por que o saldo das despesas está em ${sinal(t.S)}${curto(t.S)}?`;
    return `<h3>💡 ${tituloIns}</h3><ul>${li.map(x => `<li>${x}</li>`).join("")}</ul>`;
  }

  /* ---------- 7. Grupos detalhados (1.1, 1.1.1 … / 2.1, 2.1.1 …) ---------- */
  const buscaPainel = {}; // texto pesquisado em cada painel (mantido ao trocar o período)
  function renderGrupos(tipo, idGraf, idTab, idPainel) {
    const raizes = tipo === "rec" ? ["1"] : ["2", "6"];
    const gs = grupos.filter(g => raizes.includes(g.raiz) && vivo(g)).map(g => ({ g, t: totais(g) }));
    let ativoIdx = 0;
    let filtro = new Set(gs.map((x, i) => i)); // índices que batem com a pesquisa

    const corBarra = (cor, apagada) => ctx => (ctx.type !== "data" || filtro.has(ctx.dataIndex) ? cor : apagada);
    desenhar(idGraf, {
      type: "bar",
      data: {
        labels: gs.map(x => x.g.cod + " · " + titulo(x.g.nome)),
        datasets: [
          { label: "Previsto", data: gs.map(x => x.t.P), backgroundColor: corBarra(COR.previsto, "#e3eef1"), borderRadius: 4, maxBarThickness: 16 },
          { label: "Realizado", data: gs.map(x => x.t.R), backgroundColor: corBarra(COR.realizado, "#c9d6da"), borderRadius: 4, maxBarThickness: 16 }
        ]
      },
      options: {
        indexAxis: "y", responsive: true, maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false, axis: "y" },
        plugins: {
          legend: { position: "top", align: "end" },
          tooltip: { enabled: false } // a legenda abre no painel ao lado
        },
        scales: { x: Object.assign({ beginAtZero: true }, eixoValor), y: Object.assign({}, eixoCat, { ticks: { autoSkip: false, callback(v) { const s = this.getLabelForValue(v); const max = window.innerWidth < 640 && !imprimindo ? 18 : 30; return s.length > max ? s.slice(0, max - 1) + "…" : s; } } }) },
        onClick: (e, els) => { if (els.length) abrirDetalhe(gs[els[0].index].g); },
        onHover: (e, els) => {
          e.native.target.style.cursor = els.length ? "pointer" : "default";
          if (els.length && els[0].index !== ativoIdx) marcar(els[0].index, true);
        }
      }
    });
    const grafico = graficos[idGraf];

    /* painel lateral: legenda de todos os grupos, com pesquisa */
    const painel = document.getElementById(idPainel);
    const linhaItem = i => {
      const t = totais(i);
      return `<button class="leg-item" data-cod="${i.cod}"><span class="leg-item-nome"><span class="cod">${i.cod}</span> ${esc(titulo(i.nome))}</span>
        <span class="leg-item-val">${exato(t.R)} <small>de ${exato(t.P)}</small></span><span class="leg-item-saldo ${classe(t.S)}">${sinal(t.S)}${curto(t.S)}</span></button>`;
    };
    painel.innerHTML = `
      <div class="painel-cab">
        <h3>Legenda do gráfico</h3>
        <input type="search" class="painel-busca" placeholder="Pesquisar grupo ou item… ex.: ${tipo === "rec" ? "associado" : "manutenção"}" autocomplete="off">
      </div>
      <div class="painel-lista">
        ${gs.map((x, i) => {
          const larg = x.t.exec == null ? (x.t.R ? 100 : 0) : Math.min(100, x.t.exec * 100);
          const acima = tipo === "desp" && x.t.R > x.t.P + 0.5;
          return `<div class="leg" data-i="${i}" tabindex="0" role="button" aria-label="${esc(x.g.cod + " " + titulo(x.g.nome))}">
            <div class="leg-topo"><span class="cod">${x.g.cod}</span><span class="leg-nome">${esc(titulo(x.g.nome))}</span><span class="leg-saldo ${classe(x.t.S)}">${sinal(x.t.S)}${curto(x.t.S)}</span></div>
            <div class="leg-vals">
              <span><i class="sw sw-prev"></i>Previsto <b>${exato(x.t.P)}</b></span>
              <span><i class="sw sw-real"></i>Realizado <b>${exato(x.t.R)}</b></span>
            </div>
            <div class="leg-exec"><span class="mini-barra"><span class="${acima ? "acima" : ""}" style="width:${larg}%"></span></span>${pct(x.t.exec)} executado · saldo <b class="${classe(x.t.S)}">${sinal(x.t.S)}${exato(x.t.S)}</b></div>
            <div class="leg-itens">${x.g.itens.filter(vivo).map(linhaItem).join("")}</div>
          </div>`;
        }).join("")}
        <div class="vazio painel-vazio" hidden>Nada encontrado.</div>
      </div>`;
    const lista = painel.querySelector(".painel-lista");
    const rows = [...painel.querySelectorAll(".leg")];

    function marcar(i, rolar) {
      ativoIdx = i;
      rows.forEach((r, k) => r.classList.toggle("ativo", k === i));
      grafico.setActiveElements([{ datasetIndex: 0, index: i }, { datasetIndex: 1, index: i }]);
      grafico.update("none");
      if (rolar) {
        const r = rows[i];
        if (r.offsetTop < lista.scrollTop || r.offsetTop + r.offsetHeight > lista.scrollTop + lista.clientHeight)
          lista.scrollTo({ top: r.offsetTop - 6, behavior: "smooth" });
      }
    }
    rows.forEach((r, i) => {
      r.addEventListener("mouseenter", () => marcar(i, false));
      r.addEventListener("focus", () => marcar(i, false));
      r.addEventListener("click", e => { if (!e.target.closest(".leg-item")) abrirDetalhe(gs[i].g); });
      r.addEventListener("keydown", e => { if (e.key === "Enter") abrirDetalhe(gs[i].g); });
    });
    painel.querySelectorAll(".leg-item").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));

    const busca = painel.querySelector(".painel-busca");
    function filtrar() {
      const q = semAcento(busca.value.trim());
      buscaPainel[idPainel] = busca.value;
      filtro = new Set();
      rows.forEach((r, i) => {
        const g = gs[i].g;
        const okGrupo = !q || semAcento(g.nome).includes(q) || g.cod.startsWith(q);
        let okItem = false;
        r.querySelectorAll(".leg-item").forEach(b => {
          const it = porCod[b.dataset.cod];
          const ok = !!q && (semAcento(it.nome).includes(q) || it.cod.startsWith(q));
          b.classList.toggle("achado", ok);
          okItem = okItem || ok;
        });
        const mostra = okGrupo || okItem;
        r.hidden = !mostra;
        r.classList.toggle("aberto-busca", !!q && okItem);
        if (mostra) filtro.add(i);
      });
      painel.querySelector(".painel-vazio").hidden = filtro.size > 0;
      grafico.update("none");
      if (filtro.size && !filtro.has(ativoIdx)) marcar([...filtro][0], true);
    }
    busca.addEventListener("input", filtrar);
    busca.addEventListener("keydown", e => { if (e.key === "Escape") { busca.value = ""; filtrar(); } });
    if (buscaPainel[idPainel]) { busca.value = buscaPainel[idPainel]; filtrar(); }
    marcar(filtro.has(0) ? 0 : [...filtro][0] || 0, false);

    let html = "";
    raizes.forEach(raiz => {
      const raizLinha = porCod[raiz];
      if (!raizLinha || !vivo(raizLinha)) return;
      grupos.filter(g => g.raiz === raiz && vivo(g)).forEach(g => {
        html += linhaHier(g, "grupo");
        g.itens.filter(vivo).forEach(i => { html += linhaHier(i, "item"); });
      });
      const tt = totais(raizLinha);
      html += `<tr class="total"><td>Total ${raiz} · ${esc(titulo(raizLinha.nome))}</td><td>${exato(tt.P)}</td><td>${exato(tt.R)}</td><td class="${classe(tt.S)}">${sinal(tt.S)}${exato(tt.S)}</td><td>${pct(tt.exec)}</td></tr>`;
    });
    document.getElementById(idTab).innerHTML = `<table class="hier"><thead><tr><th>Conta</th><th>Previsto</th><th>Realizado</th><th>Saldo</th><th>Execução</th></tr></thead><tbody>${html}</tbody></table>`;
    document.querySelectorAll("#" + idTab + " tr[data-cod]").forEach(tr => tr.addEventListener("click", () => abrirDetalhe(porCod[tr.dataset.cod])));
  }
  function linhaHier(l, cls) {
    const t = totais(l);
    const larg = t.exec == null ? (t.R ? 100 : 0) : Math.min(100, t.exec * 100);
    const acima = l.tipo === "desp" ? t.R > t.P + 0.5 : false;
    return `<tr class="clicavel ${cls}" data-cod="${l.cod}"><td><span class="cod">${l.cod}</span> ${esc(titulo(l.nome))}</td>
      <td>${exato(t.P)}</td><td>${exato(t.R)}</td><td class="${classe(t.S)}">${sinal(t.S)}${exato(t.S)}</td>
      <td><span class="mini-barra"><span class="${acima ? "acima" : ""}" style="width:${larg}%"></span></span>${pct(t.exec)}</td></tr>`;
  }

  /* ---------- 8. Atividades ---------- */
  let aba = "desp";
  function cartaoAtiv(l, ehItem) {
    const t = totais(l);
    const larg = t.exec == null ? (t.R ? 100 : 0) : Math.min(100, t.exec * 100);
    const acima = l.tipo === "desp" && t.R > t.P + 0.5;
    const pai = ehItem && porCod[l.grupo] ? titulo(porCod[l.grupo].nome) + " · " : "";
    return `<button class="ativ ${ehItem ? "item" : ""}" data-cod="${l.cod}">
      <div class="ativ-topo"><div><div class="ativ-cod">${esc(pai)}${l.cod}</div><div class="ativ-nome">${esc(titulo(l.nome))}</div></div><span class="ativ-seta">→</span></div>
      <div class="ativ-nums"><span>Previsto<b title="${exato(t.P)}">${curto(t.P)}</b></span><span style="text-align:right">Realizado<b title="${exato(t.R)}">${curto(t.R)}</b></span></div>
      <div class="ativ-barra-exec"><span class="${acima ? "acima" : ""}" style="width:${larg}%"></span></div>
      <div class="ativ-rodape"><span>${pct(t.exec)} executado</span><span class="${classe(t.S)}" title="${exato(t.S)}">${sinal(t.S)}${curto(t.S)}</span></div>
    </button>`;
  }
  function renderAtividades() {
    const q = semAcento(document.getElementById("busca").value.trim());
    const grade = document.getElementById("ativGrade");
    let html;
    if (q) {
      const achados = linhas.filter(l => l.nivel >= 2 && (semAcento(l.nome).includes(q) || l.cod.startsWith(q)) && vivo(l));
      html = achados.length ? achados.map(l => cartaoAtiv(l, l.nivel === 3)).join("") : `<div class="vazio">Nenhuma atividade encontrada para “${esc(q)}”.</div>`;
    } else {
      html = grupos.filter(g => g.tipo === aba && vivo(g))
        .sort((a, b) => Math.max(soma(b.p, 12), soma(b.r, 12)) - Math.max(soma(a.p, 12), soma(a.r, 12)))
        .map(g => cartaoAtiv(g, false)).join("");
    }
    grade.innerHTML = html;
    grade.querySelectorAll(".ativ").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));
  }

  /* ---------- 9. Destaques e itens com realizado acima do previsto ---------- */
  // itens de despesa: nível 3, mais os grupos sem subitens (ex.: 6.1 Despesas Extraordinárias)
  const itensDespesa = () => linhas.filter(l => l.tipo === "desp" && vivo(l) && (l.nivel === 3 || (l.nivel === 2 && !(l.itens || []).length)));

  function renderDestaques() {
    const itens = itensDespesa().map(l => ({ l, t: totais(l) }));
    const ruins = itens.filter(x => x.t.S < -0.5).sort((a, b) => a.t.S - b.t.S).slice(0, 6);
    const bons = itens.filter(x => x.t.S > 0.5).sort((a, b) => b.t.S - a.t.S).slice(0, 6);
    const max = Math.max(...ruins.map(x => -x.t.S), ...bons.map(x => x.t.S), 1);
    const item = (x, cor) => `<button class="dest" data-cod="${x.l.cod}">
      <span class="dest-nome">${esc(titulo(x.l.nome))}</span><span class="dest-val ${classe(x.t.S)}" title="${exato(x.t.S)}">${sinal(x.t.S)}${curto(x.t.S)}</span>
      <span class="dest-sub">Previsto ${exato(x.t.P)} · Realizado ${exato(x.t.R)}</span><span></span>
      <span class="dest-barra"><span style="width:${(Math.abs(x.t.S) / max) * 100}%;background:${cor}"></span></span></button>`;
    document.getElementById("destRuim").innerHTML = ruins.map(x => item(x, COR.ruim)).join("") || `<div class="vazio">Nenhum item acima do previsto.</div>`;
    document.getElementById("destBom").innerHTML = bons.map(x => item(x, COR.bom)).join("");
    document.querySelectorAll(".dest").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));
  }

  function renderEstouros() {
    const recPrev = totais(REC).P;
    const lista = itensDespesa().map(l => ({ l, t: totais(l) })).filter(x => x.t.R > x.t.P + 0.5)
      .map(x => Object.assign(x, { est: x.t.R - x.t.P })).sort((a, b) => b.est - a.est);
    const total = lista.reduce((a, x) => a + x.est, 0);
    const semPrev = lista.filter(x => !x.t.P);
    const somaP = lista.reduce((a, x) => a + x.t.P, 0), somaR = lista.reduce((a, x) => a + x.t.R, 0);
    document.getElementById("estKpis").innerHTML = `
      <div class="kpi"><div class="kpi-nome">Itens com realizado acima do previsto</div><div class="kpi-val grande">${lista.length}</div><div class="kpi-exec">sendo ${semPrev.length} sem valor orçado no período</div></div>
      <div class="kpi"><div class="kpi-nome">Valor realizado acima do previsto</div><div class="kpi-val grande neg">${exato(total)}</div><div class="kpi-exec">soma da diferença (realizado − previsto) de cada item</div></div>
      <div class="kpi destaque"><div class="kpi-nome">Impacto sobre a receita prevista</div><div class="kpi-val grande">${pct(total / recPrev)}</div><div class="kpi-exec">da receita prevista (${curto(recPrev)}) foi absorvida por esses desvios</div></div>`;
    const max = Math.max(...lista.map(x => x.est), 1);
    document.getElementById("estTab").innerHTML = `<table><thead><tr><th>Item</th><th>Previsto</th><th>Realizado</th><th>Desvio (R$)</th><th>Desvio (%)</th><th>% da receita prevista</th></tr></thead><tbody>
      ${lista.map(x => `<tr class="clicavel" data-cod="${x.l.cod}"><td><span class="cod">${x.l.cod}</span> ${esc(titulo(x.l.nome))}<div class="sub">${esc(titulo((porCod[x.l.grupo] && x.l.nivel === 3 ? porCod[x.l.grupo] : porCod[x.l.raiz]).nome))}</div></td>
        <td>${exato(x.t.P)}</td><td>${exato(x.t.R)}</td>
        <td class="neg"><span class="mini-barra"><span class="acima" style="width:${(x.est / max) * 100}%"></span></span>${exato(x.est)}</td>
        <td>${x.t.P ? "+" + pct(x.est / x.t.P) : '<span class="selo">não orçado</span>'}</td>
        <td>${pct(x.est / recPrev)}</td></tr>`).join("")}
      <tr class="total"><td>Total</td><td>${exato(somaP)}</td><td>${exato(somaR)}</td><td class="neg">${exato(total)}</td><td>${somaP ? "+" + pct(total / somaP) : ""}</td><td>${pct(total / recPrev)}</td></tr>
    </tbody></table>`;
    document.querySelectorAll("#estTab tr[data-cod]").forEach(tr => tr.addEventListener("click", () => abrirDetalhe(porCod[tr.dataset.cod])));
  }

  /* ---------- 10. Detalhe (modal) ---------- */
  const modal = document.getElementById("modal");
  let atual = null;
  function abrirDetalhe(l) {
    if (!l) return;
    atual = l;
    const grupo = l.nivel === 3 ? porCod[l.grupo] : l;
    document.getElementById("mGrupo").textContent = (l.tipo === "rec" ? "Receitas" : "Despesas") + (l.nivel === 3 ? " · " + titulo(grupo.nome) : "");
    document.getElementById("mTitulo").textContent = l.cod + " · " + titulo(l.nome);

    const filhos = (grupo.itens || []).filter(vivo);
    document.getElementById("mChips").innerHTML = filhos.length
      ? `<button data-cod="${grupo.cod}" class="${l === grupo ? "ativo" : ""}">Total do grupo</button>` + filhos.map(i => `<button data-cod="${i.cod}" class="${l === i ? "ativo" : ""}">${esc(titulo(i.nome))}</button>`).join("")
      : "";
    document.querySelectorAll("#mChips button").forEach(b => b.addEventListener("click", () => abrirDetalhe(porCod[b.dataset.cod])));

    const t = totais(l);
    document.getElementById("mKpis").innerHTML = [
      ["Previsto", exato(t.P), ""], ["Realizado", exato(t.R), ""], ["Saldo", sinal(t.S) + exato(t.S), classe(t.S)], ["Execução", pct(t.exec), ""]
    ].map(k => `<div class="kpi"><div class="kpi-rot">${k[0]}</div><div class="kpi-val ${k[2]}">${k[1]}</div></div>`).join("");

    grafMensal("gModal", l);
    let acS = 0;
    const linhasT = MESES.map((m, i) => {
      const fut = i > ultimoMes;
      const s = saldoMes(l, i);
      if (!fut) acS += s;
      return `<tr class="${fut ? "futuro" : ""}"><td>${nomeMes(i)}${i === ultimoMes && parcial ? ' <span class="selo">parcial</span>' : ""}</td>
        <td>${exato(l.p[i])}</td><td>${fut ? "—" : exato(l.r[i])}</td>
        <td class="${fut ? "" : classe(s)}">${fut ? "—" : sinal(s) + exato(s)}</td>
        <td class="${fut ? "" : classe(acS)}">${fut ? "—" : sinal(acS) + exato(acS)}</td>
        <td>${fut || !l.p[i] ? "—" : pct(l.r[i] / l.p[i])}</td></tr>`;
    }).join("");
    document.getElementById("mTabela").innerHTML = `<table><thead><tr><th>Mês</th><th>Previsto</th><th>Realizado</th><th>Saldo no mês</th><th>Saldo acum.</th><th>Execução</th></tr></thead>
      <tbody>${linhasT}<tr class="total"><td>Total ${rotuloPeriodo()}</td><td>${exato(t.P)}</td><td>${exato(t.R)}</td><td class="${classe(t.S)}">${sinal(t.S)}${exato(t.S)}</td><td></td><td>${pct(t.exec)}</td></tr></tbody></table>`;

    modal.classList.add("aberto");
    modal.setAttribute("aria-hidden", "false");
  }
  function fecharDetalhe() { modal.classList.remove("aberto"); modal.setAttribute("aria-hidden", "true"); atual = null; }
  document.getElementById("mFechar").addEventListener("click", fecharDetalhe);
  modal.addEventListener("click", e => { if (e.target === modal) fecharDetalhe(); });

  /* ---------- 11. Navegação ---------- */
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
    slides[idx].querySelectorAll("canvas").forEach(c => { const g = graficos[c.id]; if (g) { g.reset(); g.update(); } });
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
  document.getElementById("capaTela").addEventListener("click", telaCheia);
  document.addEventListener("fullscreenchange", () => document.body.classList.toggle("em-tela-cheia", !!document.fullscreenElement));

  /* ---------- 12. Controles ---------- */
  // seletor de período: "Ano 2026" + um botão por mês (no celular, uma lista suspensa)
  const caixaMeses = document.getElementById("periodoMeses");
  const selMes = document.getElementById("selMes");
  caixaMeses.innerHTML = MESES.map((m, i) => {
    const semDados = i > ultimoMes;
    const dica = semDados ? `${nomeMes(i)}: ainda sem lançamentos`
      : `Acumulado de janeiro a ${MESES_LONGOS[i]}` + (i === ultimoMes && parcial ? " (mês ainda em lançamento)" : "");
    return `<button data-mes="${i}" title="${dica}" ${semDados ? "disabled" : ""}>${m}${i === ultimoMes && parcial ? "*" : ""}</button>`;
  }).join("");
  selMes.innerHTML = `<option value="">Mês…</option>` + MESES.map((m, i) =>
    `<option value="${i}" ${i > ultimoMes ? "disabled" : ""}>${m}${i === ultimoMes && parcial ? " (parcial)" : ""}</option>`).join("");
  function escolherPeriodo(p) {
    periodo = p;
    document.querySelector('[data-periodo="ano"]').classList.toggle("ativo", p === "ano");
    caixaMeses.querySelectorAll("button").forEach(b => b.classList.toggle("ativo", +b.dataset.mes === p));
    selMes.value = p === "ano" ? "" : String(p);
    selMes.classList.toggle("ativo", p !== "ano");
    renderTudo();
  }
  document.querySelector('[data-periodo="ano"]').addEventListener("click", () => escolherPeriodo("ano"));
  caixaMeses.querySelectorAll("button").forEach(b => b.addEventListener("click", () => escolherPeriodo(+b.dataset.mes)));
  selMes.addEventListener("change", () => escolherPeriodo(selMes.value === "" ? "ano" : +selMes.value));
  document.querySelectorAll(".abas button").forEach(b => b.addEventListener("click", () => {
    aba = b.dataset.aba;
    document.querySelectorAll(".abas button").forEach(x => x.classList.toggle("ativo", x === b));
    document.getElementById("busca").value = "";
    renderAtividades();
  }));
  document.getElementById("busca").addEventListener("input", renderAtividades);

  function renderTudo() {
    renderMensal(REC, "rec");
    renderMensal(DESP, "desp");
    renderGrupos("rec", "gRecGrupo", "tabRecGrupo", "painelRec");
    renderGrupos("desp", "gDespGrupo", "tabDespGrupo", "painelDesp");
    renderAtividades();
    renderDestaques();
    renderEstouros();
    if (atual) abrirDetalhe(atual);
  }

  /* ---------- 13. PDF / impressão ---------- */
  // monta todos os slides em sequência, com largura fixa de página A4 paisagem, e redesenha os gráficos nesse tamanho
  let imprimindo = false;
  function prepararImpressao() {
    if (imprimindo) return;
    imprimindo = true;
    if (atual) fecharDetalhe();
    document.body.classList.add("modo-impressao");
    Object.values(graficos).forEach(g => { g.options.animation = false; g.options.devicePixelRatio = 2; g.resize(); g.update("none"); });
  }
  function finalizarImpressao() {
    if (!imprimindo) return;
    imprimindo = false;
    document.body.classList.remove("modo-impressao");
    renderTudo(); ir(idx);
  }
  window.addEventListener("beforeprint", prepararImpressao);
  window.addEventListener("afterprint", finalizarImpressao);
  document.getElementById("btnPdf").addEventListener("click", () => {
    prepararImpressao();
    setTimeout(() => { window.print(); setTimeout(finalizarImpressao, 500); }, 350);
  });

  // ao abrir ou atualizar a página, sempre começa pela capa
  try { history.replaceState(null, "", location.pathname + location.search); } catch (e) {}
  window.irParaSlide = n => ir(n - 1);
  ir(0);
  // espera a fonte Montserrat carregar para os gráficos medirem os rótulos corretamente
  const iniciar = () => { renderTudo(); ir(idx); if (/[?&]pdf=1/.test(location.search)) prepararImpressao(); };
  (document.fonts && document.fonts.ready ? Promise.race([document.fonts.ready, new Promise(r => setTimeout(r, 2500))]) : Promise.resolve()).then(iniciar);
})();
