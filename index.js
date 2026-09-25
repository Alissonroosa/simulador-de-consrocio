/* =====================================================================
   Simulador de Consórcios – Sicredi
   Web part estática (index.html + index.js + style.css + data/*.csv)

   Fluxo: tela de dados → Simular → tela de resultados (Editar / Nova).
   Cada simulação pode ser registrada numa lista do SharePoint (REST) e/ou
   enviada a um fluxo do Power Automate – ver "Registro das simulações".

   Convenções de cálculo
   - Todo o plano é controlado em % do valor do crédito (fundo comum 100%
     + taxa de administração + fundo de reserva). O valor em R$ de cada
     parcela é o % do mês × crédito atualizado pelo índice do plano.
   - O crédito é reajustado a cada 12 meses pelo índice acumulado dos
     12 meses anteriores. Após a contemplação o saldo devedor continua
     sendo reajustado.
   ===================================================================== */
(function () {
  'use strict';

  const root = document.getElementById('scs-app');
  if (!root) return;

  // ------------------------------------------------------------------ Configuração
  const FILES = {
    planos: 'planos.csv',
    lances: 'historico_lances.csv',
    indices: 'indexadores.csv',
    financiamento: 'financiamento.csv',
    investimentos: 'produtos_investimento.csv',
    creditoLance: 'produtos_credito_lance.csv',
    parametros: 'parametros.csv'
  };
  const REQUIRED = ['planos', 'indices'];

  const SEGMENTOS = { IMOVEL: 'Imóveis', AUTO: 'Automóveis', SERVICOS: 'Serviços', PESADOS: 'Pesados', MOTO: 'Motocicletas' };
  const SEG_ORDEM = ['IMOVEL', 'AUTO', 'SERVICOS', 'PESADOS', 'MOTO'];
  const INDICE_LABEL = { INCC: 'INCC', IPCA: 'IPCA', IGPM: 'IGP-M', NENHUM: 'Sem reajuste', TR: 'TR', CDI: 'CDI', PRE: 'Prefixado' };

  // Paleta dos gráficos: cores da marca, validadas para daltonismo (todas as combinações)
  const COR = {
    s1: '#3FA110', // Verde Sicredi
    s2: '#0A82B4', // Azul (cor de apoio, tom escurecido p/ contraste)
    s3: '#C8327D', // Magenta (cor de apoio, tom ajustado)
    s4: '#146E37', // Verde escuro
    ink: '#323C32',
    ink2: '#5A645A',
    muted: '#7D877D',
    grid: '#E8EEE2',
    axis: '#C9D5BE'
  };
  const MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const NIVEIS = {
    CONSERVADOR: { nome: 'Conservador', param: 'confianca_lance_conservador_pct', def: 90 },
    MODERADO: { nome: 'Moderado', param: 'confianca_lance_moderado_pct', def: 70 },
    ARROJADO: { nome: 'Arrojado', param: 'confianca_lance_arrojado_pct', def: 50 }
  };

  const DB = {};
  let S = null;            // último resultado calculado
  let activeTab = 'resumo';
  let editandoId = null;   // simulação de origem quando o usuário clica em "Editar"
  const CHARTS = {};       // id do canvas -> especificação

  const $ = (s) => root.querySelector(s);
  const $$ = (s) => Array.from(root.querySelectorAll(s));

  // ------------------------------------------------------------------ Utilitários
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function parseNum(v) {
    if (v == null) return NaN;
    if (typeof v === 'number') return v;
    let s = String(v).trim().replace(/R\$|%|\s| /g, '');
    if (!s) return NaN;
    if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
    return parseFloat(s);
  }
  function num(v, def) {
    const n = parseNum(v);
    return Number.isFinite(n) ? n : (def === undefined ? 0 : def);
  }
  function parseMoney(v) {
    const s = String(v == null ? '' : v).trim().replace(/R\$|\s| /g, '');
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) return parseFloat(s.replace(/\./g, ''));
    return num(s, 0);
  }
  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const sum = (a) => a.reduce((s, x) => s + (Number.isFinite(x) ? x : 0), 0);
  const mean = (a) => (a.length ? sum(a) / a.length : NaN);
  const median = (a) => {
    const b = a.filter(Number.isFinite).sort((x, y) => x - y);
    if (!b.length) return NaN;
    const m = Math.floor(b.length / 2);
    return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2;
  };

  const nfBRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
  const fmtBRL = (v) => (Number.isFinite(v) ? nfBRL.format(v) : '—');
  const fmtNum = (v, d = 0) => (Number.isFinite(v) ? v.toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d }) : '—');
  const fmtPct = (v, d = 2) => (Number.isFinite(v) ? fmtNum(v, d) + '%' : '—');
  function fmtCompact(v) {
    const a = Math.abs(v);
    if (a >= 1e6) return 'R$ ' + fmtNum(v / 1e6, a >= 1e7 ? 0 : 1) + ' mi';
    if (a >= 1e4) return 'R$ ' + fmtNum(v / 1e3, 0) + ' mil';
    return 'R$ ' + fmtNum(v, 0);
  }

  function parseDate(v) {
    const s = String(v || '').trim();
    let m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/))) return new Date(+m[1], +m[2] - 1, +(m[3] || 1));
    if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return new Date(+m[3], +m[2] - 1, +m[1]);
    if ((m = s.match(/^(\d{1,2})\/(\d{4})$/))) return new Date(+m[2], +m[1] - 1, 1);
    return null;
  }
  const addMonths = (d, n) => new Date(d.getFullYear(), d.getMonth() + n, 1);
  const ym = (d) => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
  const fmtMes = (d) => MESES[d.getMonth()] + '/' + String(d.getFullYear()).slice(2);
  const fmtMesLongo = (d) => MESES[d.getMonth()] + '/' + d.getFullYear();
  const annual = (r) => (Number.isFinite(r) ? (Math.pow(1 + r, 12) - 1) * 100 : NaN);
  const monthly = (aa) => Math.pow(1 + aa / 100, 1 / 12) - 1;

  function P(key, def) {
    const v = DB.params ? DB.params[key] : undefined;
    if (v == null || v === '') return def;
    return typeof def === 'number' ? num(v, def) : v;
  }
  const segLabel = (s) => SEGMENTOS[s] || (s ? s.charAt(0) + s.slice(1).toLowerCase() : '');

  // Estatística
  function normCdf(x) {
    const t = 1 / (1 + 0.3275911 * Math.abs(x) / Math.SQRT2);
    const y = 1 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.exp(-x * x / 2);
    return x >= 0 ? (1 + y) / 2 : (1 - y) / 2;
  }
  function invNorm(p) {
    p = clamp(p, 1e-6, 1 - 1e-6);
    const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
    const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
    const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
    const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
    const pl = 0.02425;
    let q, r;
    if (p < pl) {
      q = Math.sqrt(-2 * Math.log(p));
      return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    }
    if (p > 1 - pl) {
      q = Math.sqrt(-2 * Math.log(1 - p));
      return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
    }
    q = p - 0.5; r = q * q;
    return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  }
  // TIR mensal (null quando o fluxo não tem raiz no intervalo)
  function irr(flows) {
    const npv = (r) => flows.reduce((s, f, t) => s + f / Math.pow(1 + r, t), 0);
    let prev = -0.05, fprev = npv(prev);
    for (let r = -0.0475; r <= 0.3; r += 0.0025) {
      const f = npv(r);
      if (Number.isFinite(f) && Number.isFinite(fprev) && Math.sign(f) !== Math.sign(fprev)) {
        let a = prev, b = r, fa = fprev;
        for (let k = 0; k < 80; k++) {
          const m = (a + b) / 2, fm = npv(m);
          if (Math.sign(fm) === Math.sign(fa)) { a = m; fa = fm; } else b = m;
        }
        return (a + b) / 2;
      }
      prev = r; fprev = f;
    }
    return null;
  }
  const pricePmt = (pv, i, n) => (n <= 0 ? 0 : i > 0 ? pv * i / (1 - Math.pow(1 + i, -n)) : pv / n);

  // ------------------------------------------------------------------ CSV
  function decode(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch (e) { return new TextDecoder('windows-1252').decode(buf); } // CSV salvo pelo Excel em ANSI
  }
  function normKey(k) {
    return String(k).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
  }
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    const nl = text.search(/\r?\n/);
    const first = nl >= 0 ? text.slice(0, nl) : text;
    const delim = (first.match(/;/g) || []).length >= (first.match(/,/g) || []).length ? ';' : ',';
    const rows = [];
    let row = [], field = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(field); field = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += c;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    const header = (rows.shift() || []).map(normKey);
    return rows
      .filter((r) => r.some((x) => String(x).trim() !== ''))
      .map((r) => { const o = {}; header.forEach((h, i) => { o[h] = (r[i] || '').trim(); }); return o; });
  }
  async function fetchCSV(base, file) {
    let res;
    const get = () => fetch(base + file, { credentials: 'same-origin', cache: 'no-cache' });
    for (let tentativa = 1; ; tentativa++) {
      try { res = await get(); break; }
      catch (e) {
        if (tentativa >= 3) throw new Error(file + ' (' + e.message + ')');
        await new Promise((r) => setTimeout(r, 400 * tentativa));
      }
    }
    if (!res.ok) throw new Error(file + ' (HTTP ' + res.status + ')');
    return parseCSV(decode(await res.arrayBuffer()));
  }
  async function loadAll() {
    let base = root.dataset.path || 'data/';
    if (!/\/$/.test(base)) base += '/';
    const keys = Object.keys(FILES);
    const results = await Promise.allSettled(keys.map((k) => fetchCSV(base, FILES[k])));
    const raw = {}, errors = [];
    results.forEach((r, i) => {
      if (r.status === 'fulfilled') raw[keys[i]] = r.value;
      else { raw[keys[i]] = []; errors.push({ key: keys[i], msg: r.reason && r.reason.message ? r.reason.message : String(r.reason) }); }
    });
    return { raw, errors, base };
  }

  function normalize(raw) {
    DB.params = {};
    raw.parametros.forEach((r) => { if (r.chave) DB.params[r.chave.trim()] = r.valor; });

    DB.planos = raw.planos.map((r) => ({
      codigo: r.codigo_plano || r.codigo || (r.segmento + '-' + r.prazo_meses),
      segmento: (r.segmento || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z]/g, ''),
      descricao: r.descricao || '',
      prazo: Math.round(num(r.prazo_meses, 0)),
      TA: num(r.taxa_adm_pct),
      FR: num(r.fundo_reserva_pct),
      indice: (r.indice_reajuste || 'IPCA').toUpperCase().replace(/[^A-Z]/g, ''),
      seguro: num(r.seguro_mensal_pct),
      adesao: num(r.taxa_adesao_pct),
      creditoMin: num(r.credito_min, 0),
      creditoMax: num(r.credito_max, 0),
      embutidoMax: num(r.lance_embutido_max_pct, 0),
      lanceFixo: num(r.lance_fixo_pct, 0),
      permiteReduzida: /^(s|sim|1|true|y)/i.test(r.permite_parcela_reduzida || ''),
      reducaoPct: num(r.reducao_parcela_pct, 50),
      baseLance: /saldo/i.test(r.base_lance || '') ? 'SALDO' : 'CREDITO',
      ativo: !/^(n|nao|não|0|false)/i.test(r.ativo || 'S')
    })).filter((x) => x.ativo && x.segmento && x.prazo > 0);

    DB.lances = raw.lances.map((r) => ({
      segmento: (r.segmento || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z]/g, ''),
      prazoPlano: Math.round(num(r.prazo_plano, 0)),
      grupo: r.grupo || '',
      assembleia: num(r.assembleia),
      data: parseDate(r.data_assembleia),
      aptas: num(r.cotas_aptas),
      sorteio: num(r.contemplados_sorteio),
      livre: num(r.contemplados_lance_livre),
      fixo: num(r.contemplados_lance_fixo),
      ofertasLivre: num(r.ofertas_lance_livre),
      ofertasFixo: num(r.ofertas_lance_fixo),
      menor: parseNum(r.menor_lance_livre_pct),
      maior: parseNum(r.maior_lance_livre_pct),
      media: parseNum(r.media_lance_livre_pct)
    })).filter((h) => h.segmento && h.data);

    DB.idx = new Map();
    raw.indices.forEach((r) => {
      const d = parseDate(r.competencia);
      if (!d) return;
      DB.idx.set(ym(d), {
        tipo: (r.tipo || '').toUpperCase(),
        IPCA: parseNum(r.ipca_pct), INCC: parseNum(r.incc_pct), IGPM: parseNum(r.igpm_pct),
        CDI: parseNum(r.cdi_pct), TR: parseNum(r.tr_pct)
      });
    });
    DB.idxKeys = Array.from(DB.idx.keys()).sort();
    DB.idxFallback = {};
    ['IPCA', 'INCC', 'IGPM', 'CDI', 'TR'].forEach((k) => {
      const vals = DB.idxKeys.map((key) => DB.idx.get(key)[k]).filter(Number.isFinite);
      DB.idxFallback[k] = {
        first: vals.length ? mean(vals.slice(0, 12)) / 100 : 0,
        last: vals.length ? mean(vals.slice(-12)) / 100 : 0
      };
    });

    DB.fin = raw.financiamento.map((r) => ({
      codigo: r.codigo || r.modalidade,
      segmento: (r.segmento || '').toUpperCase(),
      modalidade: r.modalidade || r.codigo,
      sistema: /price/i.test(r.sistema_padrao || '') ? 'PRICE' : 'SAC',
      taxaAA: num(r.taxa_juros_aa_pct),
      indexador: (r.indexador || 'PRE').toUpperCase().replace(/[^A-Z]/g, ''),
      mip: num(r.seguro_mip_mensal_pct),
      dfi: num(r.seguro_dfi_mensal_pct),
      tarifaMensal: num(r.tarifa_mensal),
      tarifaContratacao: num(r.tarifa_contratacao),
      iof: num(r.iof_pct),
      entradaMin: num(r.entrada_min_pct),
      prazoMax: num(r.prazo_max_meses, 360)
    }));
    DB.inv = raw.investimentos.map((r) => ({
      codigo: r.codigo || r.produto,
      produto: r.produto || r.codigo,
      indexador: (r.indexador || 'CDI').toUpperCase().replace(/[^A-Z]/g, ''),
      pctIndexador: num(r.percentual_indexador, 100),
      adicionalAA: num(r.taxa_adicional_aa_pct, 0),
      isento: /^(s|sim|1|true|y)/i.test(r.isento_ir || ''),
      liquidez: r.liquidez || '',
      minimo: num(r.aplicacao_minima, 0)
    }));
    DB.cred = raw.creditoLance.map((r) => ({
      codigo: r.codigo || r.produto,
      produto: r.produto || r.codigo,
      taxaAM: num(r.taxa_am_pct),
      prazoMax: num(r.prazo_max_meses, 60),
      limitePct: num(r.limite_pct_credito, 100),
      iof: num(r.iof_pct, 0),
      garantia: r.garantia || '',
      segmentos: String(r.segmentos || 'TODOS').toUpperCase().split(/[;,|]/).map((s) => s.trim()).filter(Boolean)
    }));
  }

  // ------------------------------------------------------------------ Indexadores
  function idxRate(key, date) {
    if (!key || key === 'NENHUM' || key === 'PRE') return 0;
    const r = DB.idx.get(ym(date));
    if (r && Number.isFinite(r[key])) return r[key] / 100;
    const fb = DB.idxFallback[key];
    if (!fb) return 0;
    return DB.idxKeys.length && ym(date) < DB.idxKeys[0] ? fb.first : fb.last;
  }
  function acc12(key, date) {
    let f = 1;
    for (let k = 12; k >= 1; k--) f *= 1 + idxRate(key, addMonths(date, -k));
    return f - 1;
  }

  // ------------------------------------------------------------------ Motor do consórcio
  function isAniversario(p, date, m) {
    return p.mesReajuste >= 1 && p.mesReajuste <= 12 ? date.getMonth() + 1 === p.mesReajuste : (m - 1) % 12 === 0;
  }

  /**
   * Simula o plano mês a mês.
   * p: credito, prazo, inicio, indice, TA, FR, seguro, adesao, mesReajuste, baseLance,
   *    tipoParcela, reducaoPct, reducaoBase (TOTAL|FC), reducaoAte (CONTEMPLACAO|METADE),
   *    contemplacaoMes, lanceProprioPct, lanceEmbutidoPct, abatimento (PRAZO|PARCELA)
   */
  function simulate(p) {
    const N = Math.max(1, Math.round(p.prazo));
    const ad = clamp(p.adesao || 0, 0, p.TA);
    const unit = { fc: 100 / N, ta: (p.TA - ad) / N, fr: p.FR / N };
    const rem = { fc: 100, ta: p.TA - ad, fr: p.FR };
    const c = p.contemplacaoMes >= 1 && p.contemplacaoMes <= N ? Math.round(p.contemplacaoMes) : null;
    const reduz = p.tipoParcela === 'REDUZIDA' && p.reducaoPct > 0;
    let fimReducao = 0;
    if (reduz) {
      // A redução sempre termina na contemplação; a opção METADE a limita à metade do plano.
      const ateCont = c || N;
      fimReducao = p.reducaoAte === 'METADE' ? Math.min(ateCont, Math.floor(N / 2)) : ateCont;
      fimReducao = Math.min(fimReducao, N - 1);
    }
    const fatorRed = 1 - clamp(p.reducaoPct || 0, 0, 100) / 100;

    let credit = p.credito, Nend = N, fixed = null, lance = null, acum = 0;
    const rows = [];
    const tot = { parcelas: 0, fc: 0, ta: 0, fr: 0, seguro: 0, adesao: 0 };

    for (let m = 1; m <= Nend; m++) {
      const date = addMonths(p.inicio, m - 1);
      let reaj = 0;
      if (m > 1 && isAniversario(p, date, m)) { reaj = acc12(p.indice, date); credit *= 1 + reaj; }
      const remTot = rem.fc + rem.ta + rem.fr;
      if (remTot <= 1e-9) { Nend = m - 1; break; }
      const saldoIni = remTot / 100 * credit;

      let pay;
      if (m === Nend) pay = { fc: rem.fc, ta: rem.ta, fr: rem.fr };
      else if (fixed) pay = { fc: Math.min(fixed.fc, rem.fc), ta: Math.min(fixed.ta, rem.ta), fr: Math.min(fixed.fr, rem.fr) };
      else if (reduz && m <= fimReducao) {
        const soFC = p.reducaoBase === 'FC';
        pay = { fc: unit.fc * fatorRed, ta: soFC ? unit.ta : unit.ta * fatorRed, fr: soFC ? unit.fr : unit.fr * fatorRed };
      } else {
        const left = Nend - m + 1; // dilui o saldo (inclui recomposição da parcela reduzida)
        pay = { fc: rem.fc / left, ta: rem.ta / left, fr: rem.fr / left };
      }

      const vFC = pay.fc / 100 * credit, vTA = pay.ta / 100 * credit, vFR = pay.fr / 100 * credit;
      const seguro = saldoIni * (p.seguro || 0) / 100;
      const adesao = m === 1 ? ad / 100 * credit : 0;
      const parcela = vFC + vTA + vFR + seguro + adesao;
      rem.fc -= pay.fc; rem.ta -= pay.ta; rem.fr -= pay.fr;
      acum += parcela;
      tot.parcelas += parcela; tot.fc += vFC; tot.ta += vTA + adesao; tot.fr += vFR; tot.seguro += seguro; tot.adesao += adesao;

      const remAfter = Math.max(0, rem.fc + rem.ta + rem.fr);
      const row = {
        m, date, credit, reaj, fc: vFC, ta: vTA + adesao, fr: vFR, seguro, parcela,
        pctMes: pay.fc + pay.ta + pay.fr + (m === 1 ? ad : 0),
        reduzida: reduz && m <= fimReducao,
        lance: 0, saldoPct: remAfter, saldo: remAfter / 100 * credit, acum, fase: 'pre'
      };

      if (c && m === c) {
        const remBefore = remAfter;
        const lancePct = Math.max(0, p.lanceProprioPct || 0) + Math.max(0, p.lanceEmbutidoPct || 0);
        let abat = p.baseLance === 'SALDO' ? lancePct / 100 * remBefore : lancePct;
        abat = Math.min(abat, remBefore);
        const valor = abat / 100 * credit;
        const shareEmb = lancePct > 0 ? Math.max(0, p.lanceEmbutidoPct || 0) / lancePct : 0;
        const valorEmb = valor * shareEmb;
        const monthsLeft = N - m;
        const pctMesSemLance = monthsLeft > 0 ? remBefore / monthsLeft : 0;
        const k = remBefore > 0 ? (remBefore - abat) / remBefore : 0;
        rem.fc *= k; rem.ta *= k; rem.fr *= k;
        const remLeft = remBefore - abat;
        let mesesRestantes = monthsLeft;
        if (remLeft <= 1e-9) { Nend = m; mesesRestantes = 0; }
        else if (p.abatimento === 'PRAZO' && monthsLeft > 0 && abat > 0) {
          const nm = Math.max(1, Math.ceil(remLeft / pctMesSemLance - 1e-6));
          Nend = m + nm; mesesRestantes = nm;
          const f = pctMesSemLance / remLeft;
          fixed = { fc: rem.fc * f, ta: rem.ta * f, fr: rem.fr * f };
        }
        lance = {
          mes: m, date, credito: credit, pct: lancePct, abatPct: abat, valor, valorEmb,
          valorProp: valor - valorEmb, creditoLiquido: credit - valorEmb,
          saldoAntes: remBefore / 100 * credit, saldoDepois: remLeft / 100 * credit,
          parcelaAntes: parcela, mesesRestantes, prazoRestanteOriginal: monthsLeft
        };
        row.lance = valor; row.saldoPct = remLeft; row.saldo = remLeft / 100 * credit;
      }
      rows.push(row);
    }
    rows.forEach((r) => { r.fase = !c || r.m < c ? 'pre' : r.m === c ? 'cont' : 'pos'; });
    if (lance) {
      const next = rows[lance.mes];
      lance.parcelaDepois = next ? next.parcela : 0;
    }
    const lanceProp = lance ? lance.valorProp : 0;
    const desembolso = tot.parcelas + lanceProp;

    // TIR: parcelas negativas, crédito líquido positivo no mês da contemplação
    let cet = null;
    if (lance) {
      const flows = rows.map((r) => -r.parcela);
      flows[lance.mes - 1] += lance.creditoLiquido - lance.valorProp;
      cet = irr(flows);
    }
    return {
      p, N, Nend: rows.length, rows, tot, lance, c, fimReducao,
      desembolso, cetAA: cet == null ? NaN : annual(cet)
    };
  }

  // ------------------------------------------------------------------ Contemplação (modelo estatístico)
  // Usa o histórico de assembleias do segmento (e do mesmo prazo, quando houver volume suficiente)
  // dentro de uma janela de meses. A tendência do menor lance vencedor é medida no tempo (p.p. por mês).
  const monthsDiff = (a, b) => (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());

  function contemplationModel(pl, inicio) {
    const W = Math.max(3, P('janela_historico_meses', 24));
    const minPlano = P('minimo_assembleias_plano', 12);
    const janela = (rows) => {
      if (!rows.length) return { rows, ref: null };
      const last = rows.reduce((m, h) => (h.data > m ? h.data : m), rows[0].data);
      const ref = new Date(last.getFullYear(), last.getMonth(), 1);
      const cut = addMonths(ref, -(W - 1));
      return { rows: rows.filter((h) => h.data >= cut), ref };
    };
    const doSegmento = DB.lances.filter((h) => h.segmento === pl.segmento);
    let w = janela(doSegmento.filter((h) => h.prazoPlano === pl.prazo));
    let fonte = 'plano';
    if (w.rows.length < minPlano) { w = janela(doSegmento); fonte = w.rows.length ? 'segmento' : 'nenhum'; }
    const recent = w.rows;
    const ref = w.ref || addMonths(inicio, -1);

    const pts = recent.filter((h) => Number.isFinite(h.menor)).map((h) => ({ x: monthsDiff(ref, h.data), y: h.menor }));
    let a = 30, b = 0, sigma = 6;
    if (pts.length >= 2) {
      const mx = mean(pts.map((q) => q.x)), my = mean(pts.map((q) => q.y));
      let sxy = 0, sxx = 0;
      pts.forEach((q) => { sxy += (q.x - mx) * (q.y - my); sxx += (q.x - mx) ** 2; });
      b = sxx > 0 ? sxy / sxx : 0;
      a = my - b * mx;
      const res = pts.map((q) => q.y - (a + b * q.x));
      sigma = Math.max(1.5, Math.sqrt(sum(res.map((r) => r * r)) / Math.max(1, pts.length - 2)));
    } else if (pts.length === 1) { a = pts[0].y; }
    const ys = pts.map((q) => q.y);
    const ult = recent.filter((h) => monthsDiff(ref, h.data) === 0 && Number.isFinite(h.menor));
    const aptas = sum(recent.map((h) => h.aptas));
    const ofertasFixo = sum(recent.map((h) => h.ofertasFixo));

    return {
      fonte, recent, ref, inicio, a, b, sigma, janelaMeses: W,
      horizonte: P('horizonte_tendencia_lance_meses', 12),
      minY: ys.length ? Math.min.apply(null, ys) : 5,
      maxY: ys.length ? Math.max.apply(null, ys) : 60,
      medianaY: median(ys),
      ultimoMes: ult.length ? mean(ult.map((h) => h.menor)) : NaN,
      ultimoN: ult.length,
      r0: aptas > 0 ? sum(recent.map((h) => h.sorteio)) / aptas : 0.001, // chance mensal de sorteio por cota apta
      pFixo: ofertasFixo > 0 ? clamp(sum(recent.map((h) => h.fixo)) / ofertasFixo, 0, 1) : 0
    };
  }
  // tendência do menor lance x meses após o último mês do histórico (estabiliza após o horizonte)
  function trendAt(M, x) {
    return clamp(M.a + M.b * Math.min(x, M.horizonte), Math.max(0.5, M.minY * 0.6), Math.min(100, M.maxY * 1.15));
  }
  // menor lance vencedor projetado (% da base) na assembleia do mês i da simulação
  const vProj = (M, i) => trendAt(M, monthsDiff(M.ref, addMonths(M.inicio, i - 1)));
  const confPct = (nivel) => P(NIVEIS[nivel].param, NIVEIS[nivel].def);
  function lanceRec(M, i, nivel) {
    return clamp(vProj(M, i) + invNorm(confPct(nivel) / 100) * M.sigma, 0, 100);
  }
  function projectContemplation(M, pl, prazo, lancePct, lanceInicio) {
    const out = [];
    let surv = 1;
    const N = Math.max(1, Math.round(prazo));
    for (let i = 1; i <= N; i++) {
      // cotas aptas diminuem até o fim do grupo, então a chance de sorteio por cota cresce
      const ps = Math.min(1, M.r0 / Math.max(1 / N, 1 - (i - 1) / N));
      const v = vProj(M, i);
      const ofertando = lancePct > 0 && i >= lanceInicio;
      const probL = ofertando ? normCdf((lancePct - v) / M.sigma) : 0;
      const pf = ofertando && pl.lanceFixo > 0 && lancePct >= pl.lanceFixo ? M.pFixo : 0;
      let p = 1 - (1 - ps) * (1 - probL) * (1 - pf);
      if (i === N) p = 1;
      surv *= 1 - p;
      out.push({ i, ps, pl: probL, pf, p, cum: 1 - surv, v });
    }
    return out;
  }
  function mesAlvo(proj) {
    const alvo = P('probabilidade_alvo_contemplacao_pct', 50) / 100;
    const hit = proj.find((x) => x.cum >= alvo);
    return hit ? hit.i : proj.length;
  }

  // ------------------------------------------------------------------ Financiamento
  function financing(o) {
    const entrada = o.valorBem * o.entradaPct / 100;
    const principal = Math.max(0, o.valorBem - entrada);
    const iof = principal * o.iofPct / 100;
    let saldo = principal + iof; // IOF financiado
    const i = monthly(o.taxaAA);
    const rows = [];
    for (let m = 1; m <= o.prazo; m++) {
      const date = addMonths(o.inicio, m);
      const corr = o.indexador === 'TR' || o.indexador === 'IPCA' ? idxRate(o.indexador, date) : 0;
      saldo *= 1 + corr;
      const juros = saldo * i;
      const left = o.prazo - m + 1;
      let amort, pmt;
      if (o.sistema === 'SAC') { amort = saldo / left; pmt = amort + juros; }
      else { pmt = pricePmt(saldo, i, left); amort = pmt - juros; }
      const seg = saldo * o.mip / 100 + o.valorBem * o.dfi / 100;
      const parcela = pmt + seg + o.tarifaMensal;
      saldo = Math.max(0, saldo - amort);
      rows.push({ m, date, juros, amort, seg, parcela, saldo });
    }
    const totParcelas = sum(rows.map((r) => r.parcela));
    const flows = [principal - o.tarifaContratacao].concat(rows.map((r) => -r.parcela));
    const cet = irr(flows);
    return {
      entrada, principal, iof, rows, totParcelas,
      desembolso: entrada + o.tarifaContratacao + totParcelas,
      juros: sum(rows.map((r) => r.juros)),
      cetAA: cet == null ? NaN : annual(cet)
    };
  }
  // Valor presente de uma série de desembolsos mensais (índice 0 = mês da 1ª parcela)
  function pvSerie(values, inicio) {
    const pct = P('percentual_cdi_desconto_vpl', 100) / 100;
    let df = 1, pv = 0;
    values.forEach((v, t) => { pv += v / df; df *= 1 + idxRate('CDI', addMonths(inicio, t)) * pct; });
    return pv;
  }

  // ------------------------------------------------------------------ Investimento
  function invRate(prod, date) {
    const adic = monthly(prod.adicionalAA || 0);
    const k = prod.pctIndexador / 100;
    switch (prod.indexador) {
      case 'CDI': return idxRate('CDI', date) * k + adic;
      case 'IPCA': return (1 + idxRate('IPCA', date) * k) * (1 + adic) - 1;
      case 'TR': return (1 + idxRate('TR', date) * k) * (1 + adic) - 1;
      default: return adic;
    }
  }
  function aliquotaIR(meses, isento) {
    if (isento) return 0;
    if (meses <= 6) return 0.225;
    if (meses <= 12) return 0.2;
    if (meses <= 24) return 0.175;
    return 0.15;
  }

  // ------------------------------------------------------------------ Gráficos (canvas nativo)
  function niceTicks(min, max, count) {
    const range = max - min || Math.abs(max) || 1;
    const raw = range / Math.max(1, count - 1);
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const norm = raw / mag;
    const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
    const t = [];
    for (let v = lo; v <= hi + step * 0.5; v += step) t.push(Math.abs(v) < step * 1e-9 ? 0 : v);
    return t.length >= 2 ? t : [lo, lo + step];
  }

  function drawChart(canvas, spec, opt) {
    opt = opt || {};
    const W = opt.width || canvas.clientWidth;
    const H = opt.height || canvas.clientHeight || 300;
    if (!W || !spec || !spec.labels || !spec.labels.length) return;
    const dpr = opt.dpr || window.devicePixelRatio || 1;
    canvas.width = Math.round(W * dpr);
    canvas.height = Math.round(H * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, W, H);
    const F = (w, s) => w + ' ' + s + 'px Nunito, "Segoe UI", Arial, sans-serif';
    const n = spec.labels.length;
    const series = spec.series.filter((s) => s && s.data);
    const fmtY = spec.fmtY || fmtCompact;
    const isBar = spec.type === 'bar' || spec.type === 'stack';

    // Legenda (sempre que houver 2+ séries)
    let top = 10;
    if (series.length >= 2) {
      ctx.font = F(600, 12);
      let x = 4, y = 14;
      series.forEach((s) => {
        const w = 16 + ctx.measureText(s.name).width + 18;
        if (x + w > W - 4) { x = 4; y += 20; }
        ctx.fillStyle = s.color;
        if (s.dash) { ctx.fillRect(x, y - 5, 4, 3); ctx.fillRect(x + 6, y - 5, 4, 3); }
        else roundRect(ctx, x, y - 9, 10, 10, [0, 3, 3, 3]);
        ctx.fillStyle = COR.ink2;
        ctx.textBaseline = 'alphabetic';
        ctx.textAlign = 'left';
        ctx.fillText(s.name, x + 16, y);
        x += w;
      });
      top = y + 16;
    }

    // Escala Y
    let ymin = 0, ymax = 0;
    if (spec.type === 'stack') {
      for (let i = 0; i < n; i++) ymax = Math.max(ymax, sum(series.map((s) => Math.max(0, s.data[i] || 0))));
    } else {
      series.forEach((s) => s.data.forEach((v) => { if (Number.isFinite(v)) { ymax = Math.max(ymax, v); ymin = Math.min(ymin, v); } }));
    }
    if (spec.yMax != null) ymax = spec.yMax;
    if (ymax <= ymin) ymax = ymin + 1;
    const ticks = niceTicks(ymin, ymax, 5);
    ymin = ticks[0]; ymax = ticks[ticks.length - 1];
    ctx.font = F(400, 11);
    const yLabW = Math.max.apply(null, ticks.map((t) => ctx.measureText(fmtY(t)).width));
    const L = Math.ceil(yLabW) + 14, R = 14, B = 26, T = top + (spec.markers && spec.markers.length ? 16 : 0);
    const pw = W - L - R, ph = H - T - B;
    const Y = (v) => T + ph - (v - ymin) / (ymax - ymin) * ph;
    const X = (i) => (isBar ? L + (i + 0.5) * pw / n : L + (n > 1 ? i * pw / (n - 1) : pw / 2));

    // Grade e eixo
    ctx.lineWidth = 1;
    ticks.forEach((t) => {
      const y = Math.round(Y(t)) + 0.5;
      ctx.strokeStyle = t === 0 ? COR.axis : COR.grid;
      ctx.beginPath(); ctx.moveTo(L, y); ctx.lineTo(L + pw, y); ctx.stroke();
      ctx.fillStyle = COR.muted; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(fmtY(t), L - 6, y);
    });
    const maxLabels = Math.max(2, Math.floor(pw / 66));
    const step = Math.max(1, Math.ceil(n / maxLabels));
    ctx.textBaseline = 'top'; ctx.fillStyle = COR.muted;
    for (let i = 0; i < n; i += step) {
      ctx.textAlign = i === 0 && !isBar ? 'left' : 'center';
      ctx.fillText(spec.labels[i], X(i), T + ph + 8);
    }

    // Marcadores verticais (ex.: contemplação)
    (spec.markers || []).forEach((mk) => {
      if (mk.index == null || mk.index < 0 || mk.index >= n) return;
      const x = Math.round(X(mk.index)) + 0.5;
      ctx.save();
      ctx.setLineDash([4, 4]); ctx.strokeStyle = COR.ink2; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(x, T - 4); ctx.lineTo(x, T + ph); ctx.stroke();
      ctx.restore();
      ctx.font = F(700, 11); ctx.fillStyle = COR.ink; ctx.textBaseline = 'bottom';
      const tw = ctx.measureText(mk.label).width;
      ctx.textAlign = x + tw / 2 > L + pw ? 'right' : x - tw / 2 < L ? 'left' : 'center';
      ctx.fillText(mk.label, x, T - 5);
      ctx.font = F(400, 11);
    });

    // Barras
    if (isBar) {
      const bw = pw / n;
      const w = Math.max(1, bw * (n > 80 ? 0.86 : 0.7));
      const gap = w >= 4 ? 1 : 0;
      for (let i = 0; i < n; i++) {
        let acc = 0;
        const cx = X(i) - w / 2;
        const visible = series.filter((s) => (s.data[i] || 0) > 0);
        visible.forEach((s, si) => {
          const v = s.data[i] || 0;
          const y0 = Y(acc), y1 = Y(acc + v);
          const h = Math.max(0, y0 - y1 - (si > 0 ? gap : 0));
          ctx.fillStyle = s.color;
          const topSeg = si === visible.length - 1;
          if (topSeg && w >= 8) roundRect(ctx, cx, y1, w, h, [3, 3, 0, 0]);
          else ctx.fillRect(cx, y1, w, h);
          acc += spec.type === 'stack' ? v : 0;
        });
      }
    } else {
      // Linhas
      series.forEach((s) => {
        ctx.save();
        ctx.strokeStyle = s.color; ctx.lineWidth = s.width || 2; ctx.lineJoin = 'round'; ctx.lineCap = 'round';
        if (s.dash) ctx.setLineDash(s.dash);
        ctx.beginPath();
        let started = false;
        s.data.forEach((v, i) => {
          if (!Number.isFinite(v)) { started = false; return; }
          const x = X(i), y = Y(v);
          if (!started) { ctx.moveTo(x, y); started = true; } else ctx.lineTo(x, y);
        });
        ctx.stroke();
        ctx.restore();
      });
    }

    // Camada de hover
    if (opt.hover != null && opt.hover >= 0 && opt.hover < n) {
      const x = Math.round(X(opt.hover)) + 0.5;
      ctx.strokeStyle = 'rgba(50,60,50,0.35)'; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(x, T); ctx.lineTo(x, T + ph); ctx.stroke();
      if (!isBar) {
        series.forEach((s) => {
          const v = s.data[opt.hover];
          if (!Number.isFinite(v)) return;
          ctx.beginPath(); ctx.arc(x, Y(v), 4.5, 0, Math.PI * 2);
          ctx.fillStyle = s.color; ctx.fill();
          ctx.lineWidth = 2; ctx.strokeStyle = '#FFFFFF'; ctx.stroke();
        });
      }
    }
    canvas._layout = { L, pw, n, isBar, T, ph };
  }
  function roundRect(ctx, x, y, w, h, r) {
    const [tl, tr, br, bl] = r.map((v) => Math.min(v, w / 2, h / 2));
    ctx.beginPath();
    ctx.moveTo(x + tl, y);
    ctx.lineTo(x + w - tr, y); ctx.quadraticCurveTo(x + w, y, x + w, y + tr);
    ctx.lineTo(x + w, y + h - br); ctx.quadraticCurveTo(x + w, y + h, x + w - br, y + h);
    ctx.lineTo(x + bl, y + h); ctx.quadraticCurveTo(x, y + h, x, y + h - bl);
    ctx.lineTo(x, y + tl); ctx.quadraticCurveTo(x, y, x + tl, y);
    ctx.closePath(); ctx.fill();
  }

  function setChart(id, spec) { CHARTS[id] = spec; }
  function renderChart(id) {
    const canvas = document.getElementById(id);
    if (!canvas || !CHARTS[id] || !canvas.offsetParent) return;
    drawChart(canvas, CHARTS[id]);
    if (!canvas._hoverBound) bindHover(canvas, id);
  }
  function renderActiveCharts() {
    $$('.scs-tabpanel[data-panel="' + activeTab + '"] canvas').forEach((c) => renderChart(c.id));
  }
  function bindHover(canvas, id) {
    canvas._hoverBound = true;
    const tip = canvas.parentElement.querySelector('.scs-tip');
    const hide = () => { tip.hidden = true; drawChart(canvas, CHARTS[id]); };
    canvas.addEventListener('mouseleave', hide);
    canvas.addEventListener('mousemove', (ev) => {
      const spec = CHARTS[id], lay = canvas._layout;
      if (!spec || !lay) return;
      const rect = canvas.getBoundingClientRect();
      const x = ev.clientX - rect.left;
      if (x < lay.L - 10 || x > lay.L + lay.pw + 10) { hide(); return; }
      const i = lay.isBar ? Math.floor((x - lay.L) / (lay.pw / lay.n)) : Math.round((x - lay.L) / (lay.pw / Math.max(1, lay.n - 1)));
      const idx = clamp(i, 0, lay.n - 1);
      drawChart(canvas, spec, { hover: idx });
      const fmt = spec.fmtTip || spec.fmtY || fmtBRL;
      const rows = spec.series.filter((s) => s && s.data && Number.isFinite(s.data[idx]))
        .map((s) => '<div class="scs-tip__row"><i style="background:' + s.color + '"></i><span>' + esc(s.name) + '</span><b>' + esc(fmt(s.data[idx])) + '</b></div>');
      if (spec.type === 'stack' && spec.series.length > 1) {
        rows.push('<div class="scs-tip__row"><i></i><span>Total</span><b>' + esc(fmt(sum(spec.series.map((s) => s.data[idx] || 0)))) + '</b></div>');
      }
      const title = spec.tipTitle ? spec.tipTitle(idx) : spec.labels[idx];
      tip.innerHTML = '<div class="scs-tip__title">' + esc(title) + '</div>' + rows.join('');
      tip.hidden = false;
      const tw = tip.offsetWidth, cw = canvas.clientWidth;
      let left = x + 14;
      if (left + tw > cw) left = x - tw - 14;
      tip.style.left = Math.max(0, left) + 'px';
      tip.style.top = '8px';
    });
  }

  // ------------------------------------------------------------------ HTML helpers
  function kvTable(rows) {
    return '<table class="scs-kv"><tbody>' + rows.map((r) => '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td></tr>').join('') + '</tbody></table>';
  }
  function dataTable(head, body, opt) {
    opt = opt || {};
    const th = head.map((h) => (h && typeof h === 'object' ? '<th class="' + h.cls + '">' + esc(h.v) + '</th>' : '<th>' + esc(h) + '</th>')).join('');
    const tr = body.map((r, i) => {
      const cls = opt.rowClass ? opt.rowClass(i) : '';
      return '<tr' + (cls ? ' class="' + cls + '"' : '') + '>' + r.map((c) => {
        if (c && typeof c === 'object') return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' + (c.html != null ? c.html : esc(c.v)) + '</td>';
        return '<td>' + esc(c) + '</td>';
      }).join('') + '</tr>';
    }).join('');
    const t = '<table class="scs-table"><thead><tr>' + th + '</tr></thead><tbody>' + tr + '</tbody></table>';
    return opt.noWrap ? t : '<div class="scs-table-wrap">' + t + '</div>';
  }
  function kpi(label, value, sub, hero) {
    return '<div class="scs-kpi' + (hero ? ' scs-kpi--hero' : '') + '"><div class="scs-kpi__label">' + esc(label) +
      '</div><div class="scs-kpi__value">' + esc(value) + '</div>' + (sub ? '<div class="scs-kpi__sub">' + esc(sub) + '</div>' : '') + '</div>';
  }
  const cellText = (c) => (c && typeof c === 'object' ? (c.v != null ? String(c.v) : String(c.html || '').replace(/<[^>]+>/g, '')) : String(c));

  // ------------------------------------------------------------------ Tela de dados
  const val = (sel) => { const el = $(sel); return el ? el.value : ''; };
  const segAtual = () => ($('input[name="in-segmento"]:checked') || {}).value || '';
  const prazosDe = (seg) => Array.from(new Set(DB.planos.filter((x) => x.segmento === seg).map((x) => x.prazo))).sort((a, b) => a - b);
  function faixaDe(list) {
    return {
      min: Math.min.apply(null, list.map((x) => x.creditoMin || 0)),
      max: Math.max.apply(null, list.map((x) => x.creditoMax || Infinity))
    };
  }
  // Plano = segmento + prazo; quando houver mais de uma faixa de crédito, escolhe a faixa do valor informado
  function resolvePlano(seg, prazo, credito) {
    const list = DB.planos.filter((x) => x.segmento === seg && x.prazo === prazo);
    if (!list.length) return { plano: null, list };
    const f = faixaDe(list);
    const dentro = list.find((x) => (!x.creditoMin || credito >= x.creditoMin) && (!x.creditoMax || credito <= x.creditoMax));
    if (dentro) return { plano: dentro, dentro: true, list, min: f.min, max: f.max };
    let best = list[0], bd = Infinity;
    list.forEach((x) => {
      const d = credito < x.creditoMin ? x.creditoMin - credito : credito - (x.creditoMax || Infinity);
      if (d < bd) { bd = d; best = x; }
    });
    return { plano: best, dentro: false, list, min: f.min, max: f.max };
  }

  function readInputs() {
    const seg = segAtual();
    const credito = Math.max(0, parseMoney(val('#in-credito')));
    const r = resolvePlano(seg, Math.round(num(val('#in-prazo'), 0)), credito);
    const pl = r.plano;
    if (!pl) return null;
    const ini = parseDate(val('#in-inicio')) || addMonths(new Date(), 1);
    return {
      pl, faixa: r,
      cliente: val('#in-cliente').trim(),
      consultor: val('#in-consultor').trim(),
      credito,
      prazo: pl.prazo,
      inicio: new Date(ini.getFullYear(), ini.getMonth(), 1),
      indice: pl.indice,
      tipoParcela: ($('input[name="in-tipo-parcela"]:checked') || {}).value === 'REDUZIDA' && pl.permiteReduzida ? 'REDUZIDA' : 'INTEGRAL',
      reducaoPct: clamp(num(val('#in-reducao-pct'), pl.reducaoPct || 50), 0, 90),
      reducaoBase: val('#in-reducao-base') || 'TOTAL',
      reducaoAte: val('#in-reducao-ate') || 'METADE',
      lanceProprioPct: clamp(num(val('#in-lance-proprio'), 0), 0, 100),
      lanceEmbutidoPct: clamp(num(val('#in-lance-embutido'), 0), 0, pl.embutidoMax),
      lanceInicio: clamp(Math.round(num(val('#in-lance-inicio'), 1)), 1, pl.prazo),
      contempModo: val('#in-contemp-modo') || 'PROJECAO',
      contempMes: clamp(Math.round(num(val('#in-contemp-mes'), 12)), 1, pl.prazo),
      abatimento: val('#in-abatimento') || 'PRAZO',
      TA: pl.TA, FR: pl.FR, seguro: pl.seguro, adesao: pl.adesao,
      mesReajuste: 0, baseLance: pl.baseLance
    };
  }

  function fillSegments() {
    const all = Array.from(new Set(DB.planos.map((x) => x.segmento)));
    const segs = SEG_ORDEM.filter((s) => all.includes(s)).concat(all.filter((s) => !SEG_ORDEM.includes(s)));
    $('#in-segmento').innerHTML = segs.map((s, i) => {
      const pz = prazosDe(s);
      const txt = pz.length === 1 ? pz[0] + ' meses' : pz.length + ' prazos · ' + pz[0] + ' a ' + pz[pz.length - 1] + ' meses';
      return '<label class="scs-segcard"><input type="radio" name="in-segmento" value="' + esc(s) + '"' + (i === 0 ? ' checked' : '') +
        '><span><b>' + esc(segLabel(s)) + '</b><small>' + esc(txt) + '</small></span></label>';
    }).join('');
  }
  function fillPrazos() {
    const pz = prazosDe(segAtual());
    const cur = Math.round(num(val('#in-prazo'), 0));
    $('#in-prazo').innerHTML = pz.map((x) => '<option value="' + x + '">' + x + ' meses</option>').join('');
    if (pz.includes(cur)) $('#in-prazo').value = String(cur);
  }
  function creditoPadrao(seg) {
    const f = faixaDe(DB.planos.filter((x) => x.segmento === seg));
    const max = Number.isFinite(f.max) ? f.max : f.min * 4;
    return Math.round((f.min + (max - f.min) * 0.25) / 1000) * 1000;
  }
  // aplica os limites do plano escolhido aos campos (embutido, parcela reduzida, meses)
  function applyPlanDefaults(segMudou) {
    const seg = segAtual();
    if (segMudou) {
      const f = faixaDe(DB.planos.filter((x) => x.segmento === seg));
      const cred = parseMoney(val('#in-credito'));
      if (!(cred >= f.min && cred <= f.max)) $('#in-credito').value = fmtNum(creditoPadrao(seg), 2);
    }
    const r = resolvePlano(seg, Math.round(num(val('#in-prazo'), 0)), parseMoney(val('#in-credito')));
    const pl = r.plano;
    if (!pl) return;
    if (segMudou) $('#in-reducao-pct').value = pl.reducaoPct || 50;
    $('#in-lance-embutido').max = pl.embutidoMax;
    $('#in-contemp-mes').max = pl.prazo;
    $('#in-lance-inicio').max = pl.prazo;
    const redRadio = $('input[name="in-tipo-parcela"][value="REDUZIDA"]');
    redRadio.disabled = !pl.permiteReduzida;
    if (!pl.permiteReduzida) $('input[name="in-tipo-parcela"][value="INTEGRAL"]').checked = true;
    toggleBoxes();
    updateForm();
  }
  function toggleBoxes() {
    $('#box-reduzida').hidden = ($('input[name="in-tipo-parcela"]:checked') || {}).value !== 'REDUZIDA';
    $('#box-contemp-mes').hidden = val('#in-contemp-modo') !== 'MANUAL';
  }
  function resetForm(manterConsultor) {
    const consultor = val('#in-consultor');
    $('#scs-form').reset();
    if (manterConsultor) $('#in-consultor').value = consultor;
    const first = $('input[name="in-segmento"]');
    if (first) first.checked = true;
    fillPrazos();
    $('#in-credito').value = fmtNum(creditoPadrao(segAtual()), 2);
    $('#in-inicio').value = ym(addMonths(new Date(), 1));
    $('#form-error').textContent = '';
    applyPlanDefaults(true);
  }

  function updateForm() {
    const seg = segAtual();
    const credito = parseMoney(val('#in-credito'));
    const r = resolvePlano(seg, Math.round(num(val('#in-prazo'), 0)), credito);
    const pl = r.plano;
    const hc = $('#hint-credito');
    if (!pl) { hc.textContent = ''; $('#plano-info').innerHTML = ''; $('#form-preview').innerHTML = ''; return; }

    hc.textContent = (r.dentro ? 'Faixa deste prazo: ' : 'Valor fora da faixa deste prazo: ') + fmtBRL(r.min) + ' a ' + fmtBRL(r.max);
    hc.classList.toggle('is-warn', !r.dentro);

    const info = [
      ['Taxa de administração', fmtPct(pl.TA)],
      ['Fundo de reserva', fmtPct(pl.FR)],
      ['Índice de reajuste', INDICE_LABEL[pl.indice] || pl.indice],
      ['Seguro prestamista', fmtPct(pl.seguro, 3) + ' a.m.'],
      pl.adesao ? ['Taxa de adesão', fmtPct(pl.adesao)] : null,
      ['Lance embutido máx.', fmtPct(pl.embutidoMax, 0)],
      pl.lanceFixo ? ['Lance fixo', fmtPct(pl.lanceFixo, 0)] : null
    ].filter(Boolean);
    $('#plano-info').innerHTML = '<span class="scs-plan-info__title">' + esc(pl.codigo + (pl.descricao ? ' – ' + pl.descricao : '')) + '</span>' +
      info.map((x) => '<span>' + esc(x[0]) + '</span><b>' + esc(x[1]) + '</b>').join('') +
      (r.list.length > 1 ? '<span class="scs-plan-info__title">Neste prazo a taxa varia conforme a faixa de crédito.</span>' : '');

    const embDig = num(val('#in-lance-embutido'), 0);
    const hl = $('#hint-lance');
    hl.textContent = 'Percentuais sobre o ' + (pl.baseLance === 'SALDO' ? 'saldo devedor' : 'crédito') + '. Embutido máx.: ' + fmtPct(pl.embutidoMax, 0) +
      (embDig > pl.embutidoMax ? ' – valor será limitado ao máximo.' : '');
    hl.classList.toggle('is-warn', embDig > pl.embutidoMax);

    const reduz = ($('input[name="in-tipo-parcela"]:checked') || {}).value === 'REDUZIDA';
    $('#hint-reduzida').textContent = !pl.permiteReduzida ? 'Este plano não permite parcela reduzida.'
      : reduz ? 'Após o fim da redução, a diferença não paga é diluída nas parcelas restantes.' : '';

    const p = readInputs();
    if (!p || !(p.credito > 0)) { $('#form-preview').innerHTML = ''; return; }
    const s = simulate(Object.assign({}, p, { contemplacaoMes: null }));
    const item = (l, v, main) => '<div' + (main ? ' class="is-main"' : '') + '><span>' + esc(l) + '</span><b>' + esc(v) + '</b></div>';
    $('#form-preview').innerHTML = item('1ª parcela estimada' + (p.tipoParcela === 'REDUZIDA' ? ' (reduzida)' : ''), fmtBRL(s.rows[0].parcela), true) +
      item('Crédito', fmtBRL(p.credito)) + item('Prazo', p.prazo + ' meses') + item('Taxa adm. + fundo de reserva', fmtPct(p.TA + p.FR));
  }

  function validate() {
    const seg = segAtual();
    if (!seg) return 'Escolha o segmento.';
    const credito = parseMoney(val('#in-credito'));
    if (!(credito > 0)) return 'Informe o valor do crédito.';
    const r = resolvePlano(seg, Math.round(num(val('#in-prazo'), 0)), credito);
    if (!r.plano) return 'Não há plano cadastrado para este segmento e prazo.';
    if (!r.dentro) return 'O valor do crédito para ' + segLabel(seg) + ' em ' + r.plano.prazo + ' meses deve estar entre ' + fmtBRL(r.min) + ' e ' + fmtBRL(r.max) + '.';
    const lance = num(val('#in-lance-proprio'), 0) + Math.min(num(val('#in-lance-embutido'), 0), r.plano.embutidoMax);
    if (lance > 100) return 'O lance total não pode passar de 100%.';
    if (val('#in-contemp-modo') === 'MANUAL' && num(val('#in-contemp-mes'), 0) > r.plano.prazo) return 'O mês da contemplação deve ser no máximo ' + r.plano.prazo + '.';
    return '';
  }

  // Produtos das abas de resultado dependem do segmento
  function fillProducts(pl) {
    const fin = DB.fin.filter((f) => f.segmento === pl.segmento);
    const finList = fin.length ? fin : DB.fin;
    const prevFin = val('#fin-modalidade');
    $('#fin-modalidade').innerHTML = finList.map((f) => '<option value="' + esc(f.codigo) + '">' + esc(f.modalidade) + '</option>').join('');
    if (finList.some((f) => f.codigo === prevFin)) $('#fin-modalidade').value = prevFin;
    applyFinDefaults(pl);

    const prevInv = val('#inv-produto');
    $('#inv-produto').innerHTML = DB.inv.map((x) => '<option value="' + esc(x.codigo) + '">' + esc(x.produto + ' (' + invDescricao(x) + ')') + '</option>').join('');
    if (DB.inv.some((x) => x.codigo === prevInv)) $('#inv-produto').value = prevInv;

    const creds = DB.cred.filter((x) => x.segmentos.includes('TODOS') || x.segmentos.includes(pl.segmento));
    const prevCred = val('#cred-produto');
    $('#cred-produto').innerHTML = creds.map((x) => '<option value="' + esc(x.codigo) + '">' + esc(x.produto + ' – ' + fmtPct(x.taxaAM) + ' a.m.') + '</option>').join('');
    if (creds.some((x) => x.codigo === prevCred)) $('#cred-produto').value = prevCred;
    $('#cred-mes').max = pl.prazo;
  }
  function invDescricao(x) {
    const idx = INDICE_LABEL[x.indexador] || x.indexador;
    if (x.indexador === 'PRE') return fmtPct(x.adicionalAA) + ' a.a.';
    if (x.indexador === 'CDI') return fmtNum(x.pctIndexador, 0) + '% CDI' + (x.adicionalAA ? ' + ' + fmtPct(x.adicionalAA) : '');
    return idx + (x.adicionalAA ? ' + ' + fmtPct(x.adicionalAA) + ' a.a.' : '');
  }
  function applyFinDefaults(pl) {
    const f = DB.fin.find((x) => x.codigo === val('#fin-modalidade'));
    if (!f) return;
    $('#fin-sistema').value = f.sistema;
    $('#fin-taxa').value = f.taxaAA;
    $('#fin-entrada').value = f.entradaMin;
    const n = pl ? pl.prazo : f.prazoMax;
    $('#fin-prazo').value = Math.min(f.prazoMax, Math.max(1, n));
    $('#fin-prazo').max = f.prazoMax;
  }

  // ------------------------------------------------------------------ Telas
  function showScreen(name) {
    $$('.scs-screen').forEach((s) => { s.hidden = s.dataset.screen !== name; });
    const top = root.getBoundingClientRect().top + window.pageYOffset;
    if (window.pageYOffset > top) window.scrollTo(0, top);
    if (name === 'result') renderActiveCharts();
  }
  function novoId() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return 'SIM-' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + Math.random().toString(36).slice(2, 7).toUpperCase();
  }
  function simular() {
    const err = validate();
    $('#form-error').textContent = err;
    if (err) return;
    const p = readInputs();
    fillProducts(p.pl);
    recalc();
    S.id = novoId();
    S.origem = editandoId || '';
    S.criadoEm = new Date();
    editandoId = null;
    renderResultBar();
    activateTab('resumo');
    showScreen('result');
    registrarSimulacao();
  }
  function renderResultBar() {
    const { p, sim } = S;
    $('#res-id').textContent = S.id;
    $('#res-titulo').textContent = p.cliente || 'Simulação';
    $('#res-sub').textContent = segLabel(p.pl.segmento) + ' · ' + p.prazo + ' meses · ' + fmtBRL(p.credito) + ' · 1ª parcela ' + fmtBRL(sim.rows[0].parcela);
  }

  // ------------------------------------------------------------------ Registro das simulações
  // Lista do SharePoint: grava via REST no mesmo site (usuário já autenticado; "Criado por" = usuário).
  // Power Automate: gatilho "Quando um item é criado" nessa lista. Alternativa: fluxo com gatilho HTTP (registro_flow_url).
  const SP_JSON = 'application/json;odata=nometadata';
  let spFieldsCache = null;
  function registroCfg() {
    return {
      lista: String(root.dataset.lista || P('registro_lista_sharepoint', '')).trim(),
      site: String(root.dataset.site || P('registro_site_url', '')).trim().replace(/\/+$/, ''),
      flow: String(root.dataset.flow || P('registro_flow_url', '')).trim()
    };
  }
  function siteUrl(cfg) {
    if (cfg.site) return cfg.site;
    const m = location.pathname.match(/^\/(sites|teams)\/[^/]+/i);
    return location.origin + (m ? m[0] : '');
  }
  async function spJson(url, opts) {
    const res = await fetch(url, Object.assign({ credentials: 'include' }, opts));
    if (!res.ok) {
      let msg = 'HTTP ' + res.status;
      try { const j = await res.json(); msg += ' – ' + ((j['odata.error'] && j['odata.error'].message.value) || (j.error && j.error.message && (j.error.message.value || j.error.message)) || ''); } catch (e) { /* sem corpo */ }
      throw new Error(msg);
    }
    return res.status === 204 ? {} : res.json();
  }
  async function spFields(site, lista) {
    const url = site + "/_api/web/lists/getbytitle('" + encodeURIComponent(lista.replace(/'/g, "''")) + "')/fields?$select=InternalName,TypeAsString,ReadOnlyField&$filter=Hidden eq false";
    const j = await spJson(url, { headers: { Accept: SP_JSON } });
    const map = {};
    (j.value || []).forEach((f) => { map[f.InternalName] = { type: f.TypeAsString, ro: f.ReadOnlyField }; });
    return map;
  }
  function coerce(v, type) {
    if (v == null || v === '' || (typeof v === 'number' && !Number.isFinite(v))) return null;
    switch (type) {
      case 'Number': case 'Currency': case 'Integer': { const n = typeof v === 'number' ? v : parseNum(v); return Number.isFinite(n) ? n : null; }
      case 'DateTime': return v instanceof Date ? v.toISOString() : String(v);
      case 'Boolean': return !!v;
      case 'Text': return (v instanceof Date ? v.toLocaleString('pt-BR') : String(v)).slice(0, 255);
      default: return v instanceof Date ? v.toISOString() : String(v);
    }
  }
  async function registrar(rec, cfg) {
    const out = {};
    if (cfg.lista) {
      try {
        const site = siteUrl(cfg);
        const fields = spFieldsCache || (spFieldsCache = await spFields(site, cfg.lista));
        const body = {};
        const ignorados = [];
        Object.keys(rec).forEach((k) => {
          if (!fields[k] || fields[k].ro) { ignorados.push(k); return; }
          const v = coerce(rec[k], fields[k].type);
          if (v != null) body[k] = v;
        });
        if (ignorados.length) console.info('[Simulador] Colunas não encontradas na lista (ignoradas):', ignorados.join(', '));
        const ctx = await spJson(site + '/_api/contextinfo', { method: 'POST', headers: { Accept: SP_JSON } });
        const item = await spJson(site + "/_api/web/lists/getbytitle('" + encodeURIComponent(cfg.lista.replace(/'/g, "''")) + "')/items", {
          method: 'POST',
          headers: { Accept: SP_JSON, 'Content-Type': SP_JSON, 'X-RequestDigest': ctx.FormDigestValue },
          body: JSON.stringify(body)
        });
        out.lista = { ok: true, id: item.Id || item.ID };
      } catch (e) {
        spFieldsCache = null;
        out.lista = { ok: false, msg: 'Lista: ' + e.message };
      }
    }
    if (cfg.flow) {
      try {
        let usuario = null;
        try { usuario = await spJson(siteUrl(cfg) + '/_api/web/currentuser?$select=Title,Email', { headers: { Accept: SP_JSON } }); } catch (e) { /* fora do SharePoint */ }
        const payload = Object.assign({}, rec, { UsuarioNome: usuario ? usuario.Title : '', UsuarioEmail: usuario ? usuario.Email : '' });
        const res = await fetch(cfg.flow, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        out.flow = { ok: true };
      } catch (e) {
        out.flow = { ok: false, msg: 'Power Automate: ' + e.message };
      }
    }
    return out;
  }
  const r2 = (v) => (Number.isFinite(v) ? Math.round(v * 100) / 100 : null);
  function buildRecord() {
    const { p, sim, proj, cMes, M } = S;
    const L = sim.lance;
    const pos = L && sim.rows[L.mes] ? sim.rows[L.mes].parcela : null;
    const chance12 = proj[Math.min(12, proj.length) - 1].cum * 100;
    const recMod = lanceRec(M, 1, 'MODERADO');
    const entradas = {
      segmento: p.pl.segmento, plano: p.pl.codigo, prazo: p.prazo, credito: p.credito, inicio: ym(p.inicio),
      tipoParcela: p.tipoParcela, reducaoPct: p.reducaoPct, reducaoBase: p.reducaoBase, reducaoAte: p.reducaoAte,
      lanceProprioPct: p.lanceProprioPct, lanceEmbutidoPct: p.lanceEmbutidoPct, lanceInicio: p.lanceInicio,
      contempModo: p.contempModo, contempMes: p.contempMes, abatimento: p.abatimento
    };
    const resumo = {
      parcela1: r2(sim.rows[0].parcela), parcelaPos: r2(pos), totalDesembolsado: r2(sim.desembolso), custoEfetivoAA: r2(sim.cetAA),
      mesContemplacao: cMes, chance12m: r2(chance12), lanceRecomendadoModeradoPct: r2(recMod),
      lance: L ? r2(L.valor) : 0, creditoLiquido: L ? r2(L.creditoLiquido) : null, parcelasPagas: sim.rows.length
    };
    return {
      Title: S.id,
      IdOrigem: S.origem || '',
      Associado: p.cliente,
      Consultor: p.consultor,
      Segmento: segLabel(p.pl.segmento),
      Plano: p.pl.codigo,
      PrazoMeses: p.prazo,
      ValorCredito: r2(p.credito),
      TaxaAdmPct: p.TA,
      FundoReservaPct: p.FR,
      Indice: INDICE_LABEL[p.indice] || p.indice,
      TipoParcela: p.tipoParcela === 'REDUZIDA' ? 'Reduzida' : 'Integral',
      ReducaoPct: p.tipoParcela === 'REDUZIDA' ? p.reducaoPct : 0,
      LanceProprioPct: p.lanceProprioPct,
      LanceEmbutidoPct: p.lanceEmbutidoPct,
      ModoContemplacao: p.contempModo === 'MANUAL' ? 'Manual' : 'Projeção',
      MesContemplacao: cMes,
      Parcela1: resumo.parcela1,
      ParcelaPosContemplacao: resumo.parcelaPos,
      TotalDesembolsado: resumo.totalDesembolsado,
      CustoEfetivoAA: resumo.custoEfetivoAA,
      ChanceContemplacao12m: resumo.chance12m,
      LanceRecomendadoPct: resumo.lanceRecomendadoModeradoPct,
      DataSimulacao: S.criadoEm,
      DadosJson: JSON.stringify({ versao: 2, id: S.id, origem: S.origem || null, entradas, resumo })
    };
  }
  function setChip(txt, cls, title) {
    const c = $('#res-registro');
    if (!txt) { c.hidden = true; return; }
    c.hidden = false;
    c.textContent = txt;
    c.className = 'scs-chip' + (cls ? ' ' + cls : '');
    c.title = title || '';
  }
  async function registrarSimulacao() {
    const cfg = registroCfg();
    if (!cfg.lista && !cfg.flow) { setChip(''); return; }
    const id = S.id;
    setChip('Registrando…', 'is-wait');
    const res = await registrar(buildRecord(), cfg);
    if (!S || S.id !== id) return;
    S.registro = res;
    const falhas = [res.lista, res.flow].filter((x) => x && !x.ok);
    if (falhas.length) {
      console.warn('[Simulador] Falha ao registrar a simulação:', falhas.map((f) => f.msg).join(' | '));
      setChip('Não registrada', 'is-warn', falhas.map((f) => f.msg).join('\n'));
    } else {
      setChip('Registrada' + (res.lista && res.lista.id ? ' · item ' + res.lista.id : ''), '', 'Simulação gravada');
    }
  }

  // ------------------------------------------------------------------ Cálculo principal
  function baseValue(p, base, j) {
    const r = base.rows[Math.min(base.rows.length, Math.max(1, j)) - 1];
    if (!r) return p.credito;
    return p.baseLance === 'SALDO' ? r.saldo + r.fc + r.ta + r.fr : r.credit;
  }

  function recalc() {
    const p = readInputs();
    if (!p) return;
    const M = contemplationModel(p.pl, p.inicio);
    const lanceTotal = p.lanceProprioPct + p.lanceEmbutidoPct;
    const proj = projectContemplation(M, p.pl, p.prazo, lanceTotal, p.lanceInicio);
    const projMes = mesAlvo(proj);
    const cMes = p.contempModo === 'MANUAL' ? p.contempMes : projMes;

    const sim = simulate(Object.assign({}, p, { contemplacaoMes: cMes }));
    const base = simulate(Object.assign({}, p, { tipoParcela: 'INTEGRAL', contemplacaoMes: null }));

    const keep = S ? { id: S.id, origem: S.origem, criadoEm: S.criadoEm, registro: S.registro } : {};
    S = Object.assign(keep, { p, M, proj, projMes, cMes, sim, base, report: {} });
    renderResumo();
    renderReduzida();
    renderContemplacao();
    renderProjecao();
    renderPos();
    renderFinanciamento();
    renderCombinados();
    renderActiveCharts();
  }

  const labelsMeses = (rows) => rows.map((r) => fmtMes(r.date));
  const tipMes = (rows) => (i) => 'Mês ' + rows[i].m + ' · ' + fmtMesLongo(rows[i].date);

  // ---------------------------------------------------------- Simulação geral
  function renderResumo() {
    const { p, sim, proj, cMes } = S;
    const r1 = sim.rows[0];
    const L = sim.lance;
    const pos = L && sim.rows[L.mes] ? sim.rows[L.mes].parcela : NaN;
    const chanceC = proj[cMes - 1] ? proj[cMes - 1].cum * 100 : NaN;
    const ultimo = sim.rows[sim.rows.length - 1];

    $('#resumo-kpis').innerHTML = [
      kpi('1ª parcela', fmtBRL(r1.parcela), p.tipoParcela === 'REDUZIDA' ? 'Reduzida em ' + fmtPct(p.reducaoPct, 0) : 'Parcela integral', true),
      kpi('Crédito', fmtBRL(p.credito), 'Plano ' + p.pl.codigo + ' · ' + segLabel(p.pl.segmento)),
      kpi('Prazo', p.prazo + ' meses', 'Término previsto: ' + fmtMesLongo(ultimo.date)),
      kpi('Contemplação considerada', 'Mês ' + cMes, fmtMesLongo(addMonths(p.inicio, cMes - 1)) + (p.contempModo === 'MANUAL' ? ' · informada' : ' · ' + fmtPct(chanceC, 0) + ' de chance acumulada')),
      kpi('Parcela pós-contemplação', fmtBRL(pos), L ? (p.abatimento === 'PRAZO' ? 'Lance reduz o prazo' : 'Lance reduz a parcela') : ''),
      kpi('Total desembolsado', fmtBRL(sim.desembolso), 'Parcelas + lance com recursos próprios')
    ].join('');

    const semReaj = p.credito * (100 + p.TA + p.FR) / 100;
    const comp = [
      ['Crédito contratado', fmtBRL(p.credito)],
      ['Fundo comum (100%)', fmtBRL(p.credito)],
      ['Taxa de administração (' + fmtPct(p.TA) + ')', fmtBRL(p.credito * p.TA / 100)],
      p.adesao > 0 ? ['   dos quais taxa de adesão na 1ª parcela (' + fmtPct(p.adesao) + ')', fmtBRL(p.credito * p.adesao / 100)] : null,
      ['Fundo de reserva (' + fmtPct(p.FR) + ')', fmtBRL(p.credito * p.FR / 100)],
      ['Total do plano sem reajuste', fmtBRL(semReaj)],
      ['Seguro prestamista (' + fmtPct(p.seguro, 3) + ' a.m. s/ saldo)', fmtBRL(sim.tot.seguro)],
      ['Reajuste do crédito', (INDICE_LABEL[p.indice] || p.indice) + (p.mesReajuste ? ' – em ' + MESES[p.mesReajuste - 1] : ' – a cada 12 meses')],
      ['Total de parcelas projetado (com reajustes)', fmtBRL(sim.tot.parcelas)],
      ['Custo efetivo (TIR, com reajustes)', Number.isFinite(sim.cetAA) ? fmtPct(sim.cetAA) + ' a.a.' : '—']
    ].filter(Boolean);
    $('#resumo-composicao').innerHTML = kvTable(comp);

    const tl = [];
    tl.push({ t: fmtMesLongo(p.inicio), d: '1ª parcela: ' + fmtBRL(r1.parcela), key: true });
    const reaj = sim.rows.find((r) => r.reaj > 0);
    if (reaj) tl.push({ t: fmtMesLongo(reaj.date), d: 'Primeiro reajuste: +' + fmtPct(reaj.reaj * 100) + ' (' + (INDICE_LABEL[p.indice] || p.indice) + ')' });
    if (p.tipoParcela === 'REDUZIDA' && sim.fimReducao > 0) {
      const rr = sim.rows[sim.fimReducao];
      tl.push({ t: 'Até ' + fmtMesLongo(sim.rows[sim.fimReducao - 1].date), d: 'Parcela reduzida' + (rr ? '; depois ' + fmtBRL(rr.parcela) : '') });
    }
    if (L) {
      tl.push({ t: fmtMesLongo(L.date), d: 'Contemplação (mês ' + L.mes + ')' + (L.valor > 0 ? ' com lance de ' + fmtBRL(L.valor) : ''), key: true });
      tl.push({ t: 'Crédito disponível', d: fmtBRL(L.creditoLiquido) + (L.valorEmb > 0 ? ' (descontado o lance embutido)' : '') });
    }
    tl.push({ t: fmtMesLongo(ultimo.date), d: 'Última parcela (mês ' + ultimo.m + ')', key: true });
    $('#resumo-timeline').innerHTML = '<ul class="scs-timeline">' + tl.map((x) => '<li' + (x.key ? ' class="is-key"' : '') + '><b>' + esc(x.t) + '</b><span>' + esc(x.d) + '</span></li>').join('') + '</ul>';

    setChart('ch-resumo', {
      type: 'stack',
      labels: labelsMeses(sim.rows),
      tipTitle: tipMes(sim.rows),
      series: [
        { name: 'Antes da contemplação', color: COR.s1, data: sim.rows.map((r) => (r.fase !== 'pos' ? r.parcela : 0)) },
        { name: 'Após a contemplação', color: COR.s2, data: sim.rows.map((r) => (r.fase === 'pos' ? r.parcela : 0)) }
      ],
      markers: L ? [{ index: L.mes - 1, label: 'Contemplação' }] : []
    });
    S.report.resumo = { comp, tl: tl.map((x) => [x.t, x.d]) };
  }

  // ---------------------------------------------------------- Parcela reduzida
  function renderReduzida() {
    const { p, cMes } = S;
    const g = p.pl;
    if (!g.permiteReduzida) {
      $('#reduzida-nota').className = 'scs-note is-warn';
      $('#reduzida-nota').textContent = 'O plano ' + g.codigo + ' não oferece parcela reduzida.';
      $('#reduzida-tabela').innerHTML = '';
      setChart('ch-reduzida', null);
      S.report.reduzida = null;
      return;
    }
    const red = { tipoParcela: 'REDUZIDA', reducaoPct: p.tipoParcela === 'REDUZIDA' ? p.reducaoPct : (g.reducaoPct || 50) };
    const integ = simulate(Object.assign({}, p, { tipoParcela: 'INTEGRAL', contemplacaoMes: cMes }));
    const rMeta = simulate(Object.assign({}, p, red, { reducaoAte: 'METADE', contemplacaoMes: cMes }));
    const rCont = simulate(Object.assign({}, p, red, { reducaoAte: 'CONTEMPLACAO', contemplacaoMes: cMes }));
    const sel = p.reducaoAte === 'CONTEMPLACAO' ? rCont : rMeta;
    const pctR = red.reducaoPct;

    $('#reduzida-nota').className = 'scs-note';
    $('#reduzida-nota').textContent = 'Redução de ' + fmtPct(pctR, 0) + ' sobre ' + (p.reducaoBase === 'FC' ? 'o fundo comum' : 'a parcela total') +
      '. Com contemplação no mês ' + cMes + ', a redução vale ' + (sel.fimReducao ? 'até o mês ' + sel.fimReducao : '—') +
      '; a diferença não paga é diluída nas parcelas seguintes.';

    const linha = (s) => {
      const depois = s.rows[s.fimReducao] || null;
      const posC = s.lance && s.rows[s.lance.mes] ? s.rows[s.lance.mes].parcela : NaN;
      return [s.rows[0].parcela, s.fimReducao ? 'Meses 1 a ' + s.fimReducao : '—', depois ? depois.parcela : NaN, posC, s.tot.parcelas, s.desembolso];
    };
    const li = [integ.rows[0].parcela, '—', NaN, integ.lance && integ.rows[integ.lance.mes] ? integ.rows[integ.lance.mes].parcela : NaN, integ.tot.parcelas, integ.desembolso];
    const lm = linha(rMeta), lc = linha(rCont);
    const nomes = ['Parcela inicial', 'Período com redução', 'Parcela ao fim da redução', 'Parcela pós-contemplação', 'Total de parcelas', 'Total desembolsado'];
    const f = (v) => (typeof v === 'string' ? v : fmtBRL(v));
    const body = nomes.map((n, i) => [n, f(li[i]), f(lm[i]), f(lc[i])]);
    const head = ['', 'Integral', 'Reduzida até metade/contemplação', 'Reduzida até contemplação'];
    $('#reduzida-tabela').innerHTML = dataTable(head, body) +
      '<p class="scs-caption">Economia mensal no período reduzido: <strong>' + fmtBRL(integ.rows[0].parcela - sel.rows[0].parcela) +
      '</strong>. Custo adicional total da opção reduzida (seguro sobre saldo maior e reajustes): <strong>' + fmtBRL(sel.tot.parcelas - integ.tot.parcelas) + '</strong>.</p>';

    const n = Math.max(integ.rows.length, sel.rows.length);
    const rowsRef = (integ.rows.length >= sel.rows.length ? integ : sel).rows;
    setChart('ch-reduzida', {
      type: 'line',
      labels: labelsMeses(rowsRef),
      tipTitle: tipMes(rowsRef),
      series: [
        { name: 'Integral', color: COR.s2, data: Array.from({ length: n }, (_, i) => (integ.rows[i] ? integ.rows[i].parcela : NaN)) },
        { name: 'Reduzida (' + (p.reducaoAte === 'CONTEMPLACAO' ? 'até contemplação' : 'até metade/contemplação') + ')', color: COR.s1, data: Array.from({ length: n }, (_, i) => (sel.rows[i] ? sel.rows[i].parcela : NaN)) }
      ],
      markers: [{ index: cMes - 1, label: 'Contemplação' }].concat(sel.fimReducao && sel.fimReducao !== cMes ? [{ index: sel.fimReducao - 1, label: 'Fim da redução' }] : [])
    });
    S.report.reduzida = { head, body };
  }

  // ---------------------------------------------------------- Contemplação e lance
  function renderContemplacao() {
    const { p, M, proj, projMes } = S;
    const pl = p.pl;
    const lanceTotal = p.lanceProprioPct + p.lanceEmbutidoPct;
    const at = (i) => (proj[Math.min(proj.length, i) - 1] || { cum: NaN }).cum * 100;
    const alvo = P('probabilidade_alvo_contemplacao_pct', 50);
    const fonteTxt = M.fonte === 'plano'
      ? M.recent.length + ' assembleias de grupos de ' + segLabel(pl.segmento) + ' com prazo de ' + pl.prazo + ' meses (últimos ' + M.janelaMeses + ' meses)'
      : M.fonte === 'segmento'
        ? M.recent.length + ' assembleias de grupos de ' + segLabel(pl.segmento) + ', todos os prazos (últimos ' + M.janelaMeses + ' meses)'
        : 'parâmetros padrão – não há histórico de lances cadastrado para o segmento';

    $('#contemp-kpis').innerHTML = [
      kpi('Contemplação projetada', 'Mês ' + projMes, fmtMesLongo(addMonths(p.inicio, projMes - 1)) + ' · ' + fmtPct(alvo, 0) + ' de chance acumulada', true),
      kpi('Chance na 1ª assembleia', fmtPct(proj[0].p * 100, 1), lanceTotal > 0 ? 'Com lance de ' + fmtPct(lanceTotal, 1) : 'Somente sorteio'),
      kpi('Chance em 12 meses', fmtPct(at(12), 0), 'Em 24 meses: ' + fmtPct(at(24), 0)),
      kpi('Menor lance vencedor', fmtPct(M.ultimoMes, 2), M.ultimoN ? 'Média de ' + M.ultimoN + ' assembleias em ' + fmtMesLongo(M.ref) : 'Sem histórico'),
      kpi('Mediana do menor lance', fmtPct(M.medianaY, 2), 'Últimos ' + M.janelaMeses + ' meses')
    ].join('');

    // Curvas de probabilidade acumulada
    const recNext = lanceRec(M, 1, 'MODERADO');
    const soSorteio = projectContemplation(M, pl, p.prazo, 0, 1);
    const comRec = projectContemplation(M, pl, p.prazo, recNext, 1);
    const series = [
      { name: 'Somente sorteio', color: COR.s2, data: soSorteio.map((x) => x.cum * 100) },
      { name: 'Lance recomendado (' + fmtPct(recNext, 1) + ')', color: COR.s3, data: comRec.map((x) => x.cum * 100) }
    ];
    if (lanceTotal > 0) series.unshift({ name: 'Lance informado (' + fmtPct(lanceTotal, 1) + ')', color: COR.s1, data: proj.map((x) => x.cum * 100) });
    setChart('ch-contemp', {
      type: 'line', yMax: 100,
      labels: proj.map((x) => fmtMes(addMonths(p.inicio, x.i - 1))),
      tipTitle: (i) => 'Mês ' + (i + 1) + ' · ' + fmtMesLongo(addMonths(p.inicio, i)),
      fmtY: (v) => fmtNum(v, 0) + '%',
      series,
      markers: [{ index: projMes - 1, label: 'Projeção ' + fmtPct(alvo, 0) }]
    });

    $('#contemp-metodo').textContent = 'Modelo: chance de sorteio = contemplados por sorteio ÷ cotas aptas no histórico (' + fmtPct(M.r0 * 100, 3) +
      ' ao mês por cota, crescendo à medida que o grupo contempla); chance por lance = probabilidade de o lance superar o menor lance vencedor projetado ' +
      '(tendência de ' + (M.b >= 0 ? '+' : '') + fmtNum(M.b, 2) + ' p.p. ao mês por ' + M.horizonte + ' meses, desvio de ' + fmtNum(M.sigma, 1) + ' p.p.)' +
      (pl.lanceFixo && M.pFixo ? '; lance fixo de ' + fmtPct(pl.lanceFixo, 0) + ': ' + fmtPct(M.pFixo * 100, 0) + ' dos ofertantes contemplados' : '') +
      '. Base: ' + fonteTxt + '.';

    // Recomendação
    const baseTxt = p.baseLance === 'SALDO' ? 'saldo devedor' : 'crédito';
    const bv1 = baseValue(p, S.base, 1);
    const cards = ['CONSERVADOR', 'MODERADO', 'ARROJADO'].map((nv) => {
      const pct = lanceRec(M, 1, nv);
      const emb = Math.min(pl.embutidoMax, pct);
      return { nv, pct, valor: pct / 100 * bv1, emb, prop: Math.max(0, pct - emb) / 100 * bv1 };
    });
    const reco = '<div class="scs-reco">' + cards.map((c) =>
      '<div class="scs-reco__item' + (c.nv === 'MODERADO' ? ' is-main' : '') + '">' +
      '<div class="scs-reco__name">' + NIVEIS[c.nv].nome + ' · ' + fmtPct(confPct(c.nv), 0) + ' de chance</div>' +
      '<div class="scs-reco__pct">' + fmtPct(c.pct, 2) + '</div>' +
      '<div class="scs-reco__val">' + fmtBRL(c.valor) + '</div>' +
      '<div class="scs-reco__sub">Até ' + fmtPct(c.emb, 0) + ' pode ser embutido (' + fmtBRL(c.emb / 100 * bv1) + '). Recursos próprios: <strong>' + fmtBRL(c.prop) + '</strong></div>' +
      '</div>').join('') + '</div>';

    const horiz = [1, 3, 6, 12, 24].filter((h) => h <= p.prazo);
    const head = ['Assembleia', 'Data', 'Menor lance projetado', 'Conservador', 'Moderado', 'Arrojado', 'Moderado em R$'];
    const body = horiz.map((h) => {
      const mod = lanceRec(M, h, 'MODERADO');
      return [h === 1 ? '1ª (mês 1)' : 'Mês ' + h, fmtMes(addMonths(p.inicio, h - 1)), fmtPct(vProj(M, h)),
        fmtPct(lanceRec(M, h, 'CONSERVADOR')), fmtPct(mod), fmtPct(lanceRec(M, h, 'ARROJADO')), fmtBRL(mod / 100 * baseValue(p, S.base, h))];
    });
    let infoLance = '';
    if (lanceTotal > 0) {
      const pi = proj[p.lanceInicio - 1];
      infoLance = '<div class="scs-callout">Seu lance de <strong>' + fmtPct(lanceTotal, 2) + '</strong> (' + fmtBRL(lanceTotal / 100 * baseValue(p, S.base, p.lanceInicio)) +
        ') tem cerca de <strong>' + fmtPct(pi ? pi.pl * 100 : NaN, 0) + '</strong> de chance de superar o menor lance livre na assembleia do mês ' + p.lanceInicio +
        (pl.lanceFixo > 0 && lanceTotal >= pl.lanceFixo ? ' e participa do lance fixo de ' + fmtPct(pl.lanceFixo, 0) : '') +
        '. Chance total de contemplação nessa assembleia, incluindo sorteio: <strong>' + fmtPct(pi ? pi.p * 100 : NaN, 0) + '</strong>.</div>';
    }
    $('#lance-recomendacao').innerHTML = reco + infoLance + dataTable(head, body) +
      '<p class="scs-caption">Percentuais sobre o ' + baseTxt + '. A tendência do menor lance é projetada por ' + M.horizonte + ' meses e depois mantida estável.</p>';

    // Histórico mensal (média das assembleias do mês)
    const byM = new Map();
    M.recent.forEach((h) => { const k = ym(h.data); if (!byM.has(k)) byM.set(k, []); byM.get(k).push(h); });
    const keys = Array.from(byM.keys()).sort();
    if (keys.length) {
      const fin = (arr) => arr.filter(Number.isFinite);
      const agg = keys.map((k) => {
        const hs = byM.get(k);
        const menores = fin(hs.map((h) => h.menor));
        return {
          d: parseDate(k), n: hs.length, menor: mean(menores), minMenor: menores.length ? Math.min.apply(null, menores) : NaN,
          media: mean(fin(hs.map((h) => h.media))), maior: mean(fin(hs.map((h) => h.maior))),
          aptas: sum(hs.map((h) => h.aptas)), sorteio: sum(hs.map((h) => h.sorteio)), livre: sum(hs.map((h) => h.livre)), fixo: sum(hs.map((h) => h.fixo))
        };
      });
      const extra = Math.min(6, M.horizonte);
      const labels = agg.map((x) => fmtMes(x.d));
      for (let k = 1; k <= extra; k++) labels.push(fmtMes(addMonths(M.ref, k)) + '*');
      const pad = (arr) => arr.concat(Array(extra).fill(NaN));
      const tend = agg.map((x) => trendAt(M, monthsDiff(M.ref, x.d)));
      for (let k = 1; k <= extra; k++) tend.push(trendAt(M, k));
      setChart('ch-lances', {
        type: 'line',
        labels,
        fmtY: (v) => fmtNum(v, 0) + '%',
        fmtTip: (v) => fmtPct(v, 2),
        series: [
          { name: 'Menor lance vencedor', color: COR.s1, data: pad(agg.map((x) => x.menor)) },
          { name: 'Média dos lances', color: COR.s2, data: pad(agg.map((x) => x.media)) },
          { name: 'Maior lance', color: COR.s3, data: pad(agg.map((x) => x.maior)) },
          { name: 'Tendência projetada', color: COR.s4, dash: [6, 4], data: tend }
        ]
      });
      $('#lance-historico').innerHTML = dataTable(
        ['Mês', 'Assembleias', 'Cotas aptas', 'Sorteio', 'Lance livre', 'Lance fixo', 'Menor lance (média)', 'Menor lance (mínimo)', 'Média dos lances', 'Maior lance (média)'],
        agg.slice().reverse().map((x) => [fmtMesLongo(x.d), x.n, fmtNum(x.aptas, 0), x.sorteio, x.livre, x.fixo, fmtPct(x.menor), fmtPct(x.minMenor), fmtPct(x.media), fmtPct(x.maior)])
      ) + '<p class="scs-caption">Base: ' + esc(fonteTxt) + '.</p>';
    } else {
      setChart('ch-lances', null);
      $('#lance-historico').innerHTML = '<p class="scs-caption">Sem histórico de lances cadastrado para este segmento.</p>';
    }
    S.report.contemplacao = {
      kv: [
        ['Contemplação projetada', 'Mês ' + projMes + ' (' + fmtMesLongo(addMonths(p.inicio, projMes - 1)) + ')'],
        ['Chance na 1ª assembleia', fmtPct(proj[0].p * 100, 1)],
        ['Chance acumulada em 12 / 24 meses', fmtPct(at(12), 0) + ' / ' + fmtPct(at(24), 0)],
        ['Menor lance vencedor (último mês)', fmtPct(M.ultimoMes)],
        ['Base estatística', fonteTxt]
      ],
      cards: cards.map((c) => [NIVEIS[c.nv].nome + ' (' + fmtPct(confPct(c.nv), 0) + ')', fmtPct(c.pct), fmtBRL(c.valor), fmtBRL(c.prop)]),
      head, body
    };
  }

  // ---------------------------------------------------------- Projeção de parcelas
  function renderProjecao() {
    const { sim } = S;
    const rows = sim.rows;
    setChart('ch-projecao', {
      type: 'stack',
      labels: labelsMeses(rows),
      tipTitle: tipMes(rows),
      series: [
        { name: 'Fundo comum', color: COR.s1, data: rows.map((r) => r.fc) },
        { name: 'Taxa de administração', color: COR.s2, data: rows.map((r) => r.ta) },
        { name: 'Fundo de reserva', color: COR.s3, data: rows.map((r) => r.fr) },
        { name: 'Seguro', color: COR.s4, data: rows.map((r) => r.seguro) }
      ],
      markers: sim.lance ? [{ index: sim.lance.mes - 1, label: 'Contemplação' }] : []
    });
    setChart('ch-saldo', {
      type: 'line',
      labels: labelsMeses(rows),
      tipTitle: tipMes(rows),
      series: [
        { name: 'Crédito atualizado', color: COR.s1, data: rows.map((r) => r.credit) },
        { name: 'Saldo devedor', color: COR.s2, data: rows.map((r) => r.saldo) },
        { name: 'Total pago acumulado', color: COR.s3, data: rows.map((r) => r.acum) }
      ],
      markers: sim.lance ? [{ index: sim.lance.mes - 1, label: 'Contemplação' }] : []
    });
    const head = projHead();
    const body = rows.map(projRow);
    body.push([{ v: 'Total' }, '', '', '', '', fmtBRL(sim.tot.fc), fmtBRL(sim.tot.ta), fmtBRL(sim.tot.fr), fmtBRL(sim.tot.seguro), fmtBRL(sim.tot.parcelas), fmtBRL(sim.lance ? sim.lance.valor : 0), '', '']);
    $('#projecao-tabela').innerHTML = dataTable(head, body, {
      noWrap: true,
      rowClass: (i) => (i === body.length - 1 ? 'is-total' : rows[i].fase === 'cont' ? 'is-mark' : '')
    });
  }
  const projHead = () => ['Mês', 'Vencimento', { v: 'Fase', cls: 'l' }, 'Crédito atualizado', 'Reajuste', 'Fundo comum', 'Taxa adm.', 'Fundo reserva', 'Seguro', 'Parcela', 'Lance', 'Saldo devedor', 'Pago acumulado'];
  function projRow(r) {
    const fase = r.fase === 'cont' ? '<span class="scs-tag scs-tag--cont">Contemplação</span>'
      : r.fase === 'pos' ? '<span class="scs-tag scs-tag--pos">Pós</span>'
        : '<span class="scs-tag">' + (r.reduzida ? 'Reduzida' : 'Pré') + '</span>';
    const faseTxt = r.fase === 'cont' ? 'Contemplação' : r.fase === 'pos' ? 'Pós' : r.reduzida ? 'Reduzida' : 'Pré';
    return [r.m, fmtMes(r.date), { html: fase, v: faseTxt, cls: 'l' }, fmtBRL(r.credit), r.reaj ? fmtPct(r.reaj * 100) : '', fmtBRL(r.fc), fmtBRL(r.ta), fmtBRL(r.fr),
      fmtBRL(r.seguro), fmtBRL(r.parcela), r.lance ? fmtBRL(r.lance) : '', fmtBRL(r.saldo), fmtBRL(r.acum)];
  }

  // ---------------------------------------------------------- Pós-contemplação
  function renderPos() {
    const { p, sim, cMes } = S;
    const L = sim.lance;
    if (!L) {
      $('#pos-kpis').innerHTML = '';
      $('#pos-comparativo').innerHTML = '<p class="scs-caption">Defina um mês de contemplação para ver as parcelas pós-contemplação.</p>';
      setChart('ch-pos', null);
      S.report.pos = null;
      return;
    }
    $('#pos-kpis').innerHTML = [
      kpi('Crédito na contemplação', fmtBRL(L.credito), 'Mês ' + L.mes + ' · ' + fmtMesLongo(L.date), true),
      kpi('Lance', fmtBRL(L.valor), L.valor ? 'Próprio ' + fmtBRL(L.valorProp) + ' · embutido ' + fmtBRL(L.valorEmb) : 'Contemplação por sorteio'),
      kpi('Crédito líquido disponível', fmtBRL(L.creditoLiquido), L.valorEmb ? 'Crédito menos lance embutido' : 'Valor integral'),
      kpi('Saldo devedor', fmtBRL(L.saldoDepois), 'Antes do lance: ' + fmtBRL(L.saldoAntes)),
      kpi('Parcela', fmtBRL(L.parcelaDepois), 'Antes: ' + fmtBRL(L.parcelaAntes)),
      kpi('Parcelas restantes', String(L.mesesRestantes), 'Prazo original restante: ' + L.prazoRestanteOriginal)
    ].join('');

    const sPrazo = p.abatimento === 'PRAZO' ? sim : simulate(Object.assign({}, p, { contemplacaoMes: cMes, abatimento: 'PRAZO' }));
    const sParc = p.abatimento === 'PARCELA' ? sim : simulate(Object.assign({}, p, { contemplacaoMes: cMes, abatimento: 'PARCELA' }));
    const info = (s) => {
      const l = s.lance;
      const pos = s.rows.filter((r) => r.fase === 'pos');
      return [fmtBRL(l.parcelaDepois), String(l.mesesRestantes), fmtMesLongo(s.rows[s.rows.length - 1].date), fmtBRL(sum(pos.map((r) => r.parcela))), fmtBRL(s.desembolso)];
    };
    const a = info(sPrazo), b = info(sParc);
    const nomes = ['1ª parcela após a contemplação', 'Parcelas restantes', 'Última parcela', 'Total pago após a contemplação', 'Total desembolsado no plano'];
    const head = ['', 'Reduzir prazo' + (p.abatimento === 'PRAZO' ? ' (selecionado)' : ''), 'Reduzir parcela' + (p.abatimento === 'PARCELA' ? ' (selecionado)' : '')];
    const body = nomes.map((n, i) => [n, a[i], b[i]]);
    $('#pos-comparativo').innerHTML = dataTable(head, body) + (L.valor ? '' : '<p class="scs-caption">Sem lance informado, as duas opções resultam no mesmo plano. Informe um lance no painel para comparar.</p>');

    const n = Math.max(sPrazo.rows.length, sParc.rows.length);
    const ref = sPrazo.rows.length >= sParc.rows.length ? sPrazo.rows : sParc.rows;
    setChart('ch-pos', {
      type: 'line',
      labels: labelsMeses(ref),
      tipTitle: tipMes(ref),
      series: [
        { name: 'Lance reduz o prazo', color: COR.s1, data: Array.from({ length: n }, (_, i) => (sPrazo.rows[i] ? sPrazo.rows[i].parcela : NaN)) },
        { name: 'Lance reduz a parcela', color: COR.s2, data: Array.from({ length: n }, (_, i) => (sParc.rows[i] ? sParc.rows[i].parcela : NaN)) }
      ],
      markers: [{ index: L.mes - 1, label: 'Contemplação' }]
    });
    S.report.pos = {
      kv: [
        ['Mês da contemplação', L.mes + ' (' + fmtMesLongo(L.date) + ')'],
        ['Crédito na contemplação', fmtBRL(L.credito)],
        ['Lance total (' + fmtPct(L.pct) + ')', fmtBRL(L.valor)],
        ['   recursos próprios', fmtBRL(L.valorProp)],
        ['   embutido', fmtBRL(L.valorEmb)],
        ['Crédito líquido disponível', fmtBRL(L.creditoLiquido)],
        ['Saldo devedor antes / depois do lance', fmtBRL(L.saldoAntes) + ' / ' + fmtBRL(L.saldoDepois)],
        ['Parcela antes / depois', fmtBRL(L.parcelaAntes) + ' / ' + fmtBRL(L.parcelaDepois)]
      ],
      head, body
    };
  }

  // ---------------------------------------------------------- Comparativo com financiamento
  function renderFinanciamento() {
    const { p, sim } = S;
    const f = DB.fin.find((x) => x.codigo === val('#fin-modalidade'));
    if (!f) {
      $('#fin-tabela').innerHTML = '<p class="scs-caption">Cadastre modalidades em financiamento.csv para habilitar o comparativo.</p>';
      $('#fin-conclusao').textContent = '';
      setChart('ch-fin', null);
      S.report.fin = null; S.fin = null;
      return;
    }
    const o = {
      valorBem: p.credito,
      entradaPct: clamp(num(val('#fin-entrada'), f.entradaMin), 0, 95),
      prazo: Math.max(1, Math.round(num(val('#fin-prazo'), f.prazoMax))),
      taxaAA: num(val('#fin-taxa'), f.taxaAA),
      sistema: val('#fin-sistema') || f.sistema,
      indexador: f.indexador, mip: f.mip, dfi: f.dfi, tarifaMensal: f.tarifaMensal,
      tarifaContratacao: f.tarifaContratacao, iofPct: f.iof, inicio: p.inicio
    };
    const fin = financing(o);
    S.fin = fin;

    // Desembolsos por mês (índice 0 = mês da 1ª parcela do consórcio / contratação do financiamento)
    const T = Math.max(sim.rows.length, fin.rows.length + 1);
    const outC = Array.from({ length: T }, (_, t) => (sim.rows[t] ? sim.rows[t].parcela : 0) + (sim.lance && t === sim.lance.mes - 1 ? sim.lance.valorProp : 0));
    const outF = Array.from({ length: T }, (_, t) => (t === 0 ? fin.entrada + o.tarifaContratacao : fin.rows[t - 1] ? fin.rows[t - 1].parcela : 0));
    const pvC = pvSerie(outC, p.inicio), pvF = pvSerie(outF, p.inicio);
    const L = sim.lance;
    const mediaC = mean(sim.rows.map((r) => r.parcela)), mediaF = mean(fin.rows.map((r) => r.parcela));
    const dif = fin.desembolso - sim.desembolso;
    const head = ['', 'Consórcio', o.sistema + ' – ' + f.modalidade];
    const body = [
      ['Acesso ao bem', L ? 'Mês ' + L.mes + ' (' + fmtMesLongo(L.date) + ')' : 'Na contemplação', 'Imediato'],
      ['Entrada', fmtBRL(0), fmtBRL(fin.entrada)],
      ['Valor contratado / financiado', fmtBRL(p.credito), fmtBRL(fin.principal) + (fin.iof ? ' + IOF ' + fmtBRL(fin.iof) : '')],
      ['Taxa', 'Adm. ' + fmtPct(p.TA) + ' + FR ' + fmtPct(p.FR) + ' (total)', fmtPct(o.taxaAA) + ' a.a.' + (f.indexador !== 'PRE' ? ' + ' + (INDICE_LABEL[f.indexador] || f.indexador) : '')],
      ['1ª parcela', fmtBRL(sim.rows[0].parcela), fmtBRL(fin.rows[0].parcela)],
      ['Parcela média', fmtBRL(mediaC), fmtBRL(mediaF)],
      ['Última parcela', fmtBRL(sim.rows[sim.rows.length - 1].parcela), fmtBRL(fin.rows[fin.rows.length - 1].parcela)],
      ['Prazo', sim.rows.length + ' meses', o.prazo + ' meses'],
      ['Juros / taxas totais', fmtBRL(sim.tot.ta + sim.tot.fr + sim.tot.seguro), fmtBRL(fin.juros + sum(fin.rows.map((r) => r.seg)) + o.tarifaContratacao + fin.iof)],
      ['Total desembolsado', fmtBRL(sim.desembolso), fmtBRL(fin.desembolso)],
      ['Valor presente dos desembolsos (' + fmtNum(P('percentual_cdi_desconto_vpl', 100), 0) + '% CDI)', fmtBRL(pvC), fmtBRL(pvF)],
      ['Custo efetivo (TIR)', Number.isFinite(sim.cetAA) ? fmtPct(sim.cetAA) + ' a.a.' : '—', Number.isFinite(fin.cetAA) ? fmtPct(fin.cetAA) + ' a.a.' : '—']
    ];
    $('#fin-tabela').innerHTML = dataTable(head, body);
    $('#fin-conclusao').innerHTML = dif > 0
      ? 'O consórcio desembolsa <strong>' + esc(fmtBRL(dif)) + '</strong> a menos que o financiamento (' + esc(fmtPct(dif / fin.desembolso * 100, 1)) +
        '). Em troca, o bem é adquirido na contemplação' + (L ? ' (mês ' + L.mes + ')' : '') + ', enquanto no financiamento o acesso é imediato.'
      : 'Nesta configuração o financiamento desembolsa ' + esc(fmtBRL(-dif)) + ' a menos que o consórcio. Revise prazo, taxa e entrada.';

    const labels = Array.from({ length: T }, (_, t) => fmtMes(addMonths(p.inicio, t)));
    let aC = 0, aF = 0;
    setChart('ch-fin', {
      type: 'line',
      labels,
      tipTitle: (i) => 'Mês ' + (i + 1) + ' · ' + fmtMesLongo(addMonths(p.inicio, i)),
      series: [
        { name: 'Consórcio', color: COR.s1, data: outC.map((v) => (aC += v)) },
        { name: 'Financiamento', color: COR.s2, data: outF.map((v) => (aF += v)) }
      ],
      markers: L ? [{ index: L.mes - 1, label: 'Contemplação' }] : []
    });
    S.report.fin = { head, body, conclusao: $('#fin-conclusao').textContent };
  }

  // ---------------------------------------------------------- Produtos combinados
  function renderCombinados() {
    renderInvestimento();
    renderCreditoLance();
  }

  function renderInvestimento() {
    const { p, M, base } = S;
    const prod = DB.inv.find((x) => x.codigo === val('#inv-produto')) || DB.inv[0];
    if (!prod) {
      $('#inv-resultado').innerHTML = '<p class="scs-caption">Cadastre produtos em produtos_investimento.csv.</p>';
      setChart('ch-inv', null); S.report.inv = null;
      return;
    }
    const inicial = parseMoney(val('#inv-inicial'));
    const mensal = parseMoney(val('#inv-mensal'));
    const meta = val('#inv-meta') || 'MODERADO';
    const emb = p.lanceEmbutidoPct;
    const lots = [];
    const rows = [];
    let reached = null;
    for (let j = 1; j <= p.prazo; j++) {
      const date = addMonths(p.inicio, j - 1);
      if (j > 1) { const r = invRate(prod, date); lots.forEach((l) => { l.v *= 1 + r; }); }
      if (j === 1 && inicial > 0) lots.push({ p: inicial, v: inicial, j });
      if (mensal > 0) lots.push({ p: mensal, v: mensal, j });
      const bruto = sum(lots.map((l) => l.v));
      const aportado = sum(lots.map((l) => l.p));
      const liquido = sum(lots.map((l) => l.v - aliquotaIR(j - l.j, prod.isento) * Math.max(0, l.v - l.p)));
      const pctLance = meta === 'INFORMADO' ? p.lanceProprioPct + emb : lanceRec(M, j, meta);
      const pctProprio = Math.max(0, pctLance - Math.min(emb, pctLance));
      const alvo = pctProprio / 100 * baseValue(p, base, j);
      rows.push({ j, date, bruto, liquido, aportado, alvo, pctLance, pctProprio });
      if (reached == null && liquido >= alvo - 0.005 && (aportado > 0 || alvo <= 0)) reached = j;
    }
    S.inv = { prod, rows, reached };

    // Sugestão de aporte: diferença de parcela para o financiamento
    const sugest = S.fin ? S.fin.rows[0].parcela - S.sim.rows[0].parcela : NaN;
    let html = '';
    if (Number.isFinite(sugest) && sugest > 0) {
      html += '<p class="scs-caption">Sugestão: investir a diferença entre a parcela do financiamento e a do consórcio – <strong>' + fmtBRL(sugest) +
        '</strong>/mês. <button type="button" class="scs-btn scs-btn--ghost" data-action="usar-sugestao" data-valor="' + sugest.toFixed(2) + '">Usar este valor</button></p>';
    }
    const semSorteio = projectContemplation(M, p.pl, p.prazo, 0, 1);
    if (reached) {
      const r = rows[reached - 1];
      const chanceSorteioAntes = reached > 1 ? semSorteio[reached - 2].cum * 100 : 0;
      html += '<div class="scs-callout">Aplicando ' + fmtBRL(mensal) + '/mês' + (inicial ? ' + ' + fmtBRL(inicial) + ' inicial' : '') + ' em <strong>' + esc(prod.produto) +
        '</strong>, o associado reúne <strong>' + fmtBRL(r.alvo) + '</strong> (lance de ' + fmtPct(r.pctLance) + (emb ? ', com ' + fmtPct(Math.min(emb, r.pctLance)) + ' embutido' : '') +
        ') no <strong>mês ' + reached + ' (' + fmtMesLongo(r.date) + ')</strong>. Até lá, a chance de contemplação só por sorteio é de ' + fmtPct(chanceSorteioAntes, 0) + '. ' +
        '<button type="button" class="scs-btn scs-btn--primary" data-action="aplicar-inv">Usar na simulação</button></div>';
    } else {
      html += '<div class="scs-callout">Com esses aportes a meta de lance não é atingida dentro do prazo. Aumente o aporte mensal ou use lance embutido.</div>';
    }
    const marcos = [6, 12, 24, 36, 60].filter((k) => k <= rows.length);
    const head = ['Mês', 'Data', 'Aportado', 'Saldo bruto', 'Saldo líquido de IR', 'Lance necessário (próprio)', 'Situação'];
    const body = marcos.map((k) => {
      const r = rows[k - 1];
      return [k, fmtMes(r.date), fmtBRL(r.aportado), fmtBRL(r.bruto), fmtBRL(r.liquido), fmtBRL(r.alvo), r.liquido >= r.alvo ? { v: 'Atingido', cls: 'pos' } : { v: 'Faltam ' + fmtBRL(r.alvo - r.liquido), cls: 'neg' }];
    });
    $('#inv-resultado').innerHTML = html + dataTable(head, body);

    const lim = Math.min(rows.length, Math.max(24, (reached || 0) + 12));
    const rr = rows.slice(0, lim);
    setChart('ch-inv', {
      type: 'line',
      labels: rr.map((r) => fmtMes(r.date)),
      tipTitle: (i) => 'Mês ' + (i + 1) + ' · ' + fmtMesLongo(rr[i].date),
      series: [
        { name: 'Saldo líquido do investimento', color: COR.s1, data: rr.map((r) => r.liquido) },
        { name: 'Lance necessário (recursos próprios)', color: COR.s3, data: rr.map((r) => r.alvo) },
        { name: 'Total aportado', color: COR.s2, dash: [5, 4], data: rr.map((r) => r.aportado) }
      ],
      markers: reached && reached <= lim ? [{ index: reached - 1, label: 'Meta atingida' }] : []
    });
    S.report.inv = { texto: $('#inv-resultado .scs-callout') ? $('#inv-resultado .scs-callout').textContent.replace('Usar na simulação', '').trim() : '', head, body, produto: prod.produto + ' (' + invDescricao(prod) + ')' };
  }

  function renderCreditoLance() {
    const { p, M, base, sim } = S;
    const prod = DB.cred.find((x) => x.codigo === val('#cred-produto'));
    if (!prod) {
      $('#cred-resultado').innerHTML = '<p class="scs-caption">Cadastre produtos em produtos_credito_lance.csv.</p>';
      setChart('ch-cred', null); S.report.cred = null;
      return;
    }
    const nivel = val('#cred-nivel') || 'MODERADO';
    const k = clamp(Math.round(num(val('#cred-mes'), 1)), 1, p.prazo);
    const recursos = parseMoney(val('#cred-recursos'));
    const prazoDig = Math.max(1, Math.round(num(val('#cred-prazo'), 24)));
    const prazo = Math.min(prazoDig, prod.prazoMax);
    const pct = lanceRec(M, k, nivel);
    const embUsed = Math.min(p.lanceEmbutidoPct, pct);
    const bv = baseValue(p, base, k);
    const lanceValor = pct / 100 * bv;
    const embValor = embUsed / 100 * bv;
    const proprio = lanceValor - embValor;
    const valor = Math.max(0, proprio - recursos);
    const limite = prod.limitePct / 100 * base.rows[k - 1].credit;
    const iof = valor * prod.iof / 100;
    const i = prod.taxaAM / 100;
    const pmt = pricePmt(valor + iof, i, prazo);
    const totalCred = pmt * prazo;
    const cetCred = valor > 0 ? irr([valor].concat(Array(prazo).fill(-pmt))) : null;

    const simC = simulate(Object.assign({}, p, { contemplacaoMes: k, lanceProprioPct: pct - embUsed, lanceEmbutidoPct: embUsed }));
    const L = simC.lance;
    const antecipa = S.cMes - k;
    const semLance = projectContemplation(M, p.pl, p.prazo, 0, 1);
    const chanceAntes = k > 1 ? semLance[k - 2].cum : 0;
    const chanceK = chanceAntes + (1 - chanceAntes) * confPct(nivel) / 100;

    const avisos = [];
    if (valor > limite) avisos.push('O valor necessário (' + fmtBRL(valor) + ') excede o limite do produto (' + fmtPct(prod.limitePct, 0) + ' do crédito = ' + fmtBRL(limite) + ').');
    if (prazoDig > prod.prazoMax) avisos.push('Prazo limitado a ' + prod.prazoMax + ' meses para este produto.');

    const head = ['Item', 'Valor'];
    const body = [
      ['Lance ' + NIVEIS[nivel].nome.toLowerCase() + ' no mês ' + k, fmtPct(pct) + ' = ' + fmtBRL(lanceValor)],
      ['Lance embutido usado', fmtPct(embUsed) + ' = ' + fmtBRL(embValor)],
      ['Recursos próprios disponíveis', fmtBRL(Math.min(recursos, proprio))],
      ['Crédito para lance contratado', fmtBRL(valor) + (iof ? ' + IOF ' + fmtBRL(iof) : '')],
      ['Parcela do crédito (' + prazo + 'x, ' + fmtPct(prod.taxaAM) + ' a.m.)', fmtBRL(pmt)],
      ['Juros + IOF do crédito', fmtBRL(totalCred - valor)],
      ['CET do crédito', Number.isFinite(annual(cetCred)) ? fmtPct(annual(cetCred)) + ' a.a.' : '—'],
      ['Parcela do consórcio após a contemplação', L ? fmtBRL(L.parcelaDepois) : '—'],
      ['Desembolso mensal enquanto houver as duas parcelas', L ? fmtBRL(L.parcelaDepois + pmt) : '—'],
      ['Crédito líquido na contemplação', L ? fmtBRL(L.creditoLiquido) : '—']
    ];
    let html = avisos.length ? '<div class="scs-note is-warn">' + avisos.map(esc).join('<br>') + '</div>' : '';
    html += '<div class="scs-callout">Ofertando o lance ' + NIVEIS[nivel].nome.toLowerCase() + ' no <strong>mês ' + k + '</strong>, a chance estimada de estar contemplado até essa assembleia é de <strong>' +
      fmtPct(chanceK * 100, 0) + '</strong>' + (antecipa > 0 ? ', antecipando a contemplação em cerca de <strong>' + antecipa + ' meses</strong> em relação ao cenário atual (mês ' + S.cMes + ').' : '.') + '</div>';
    html += dataTable(head, body);
    $('#cred-resultado').innerHTML = html;

    const n = Math.max(simC.rows.length, sim.rows.length);
    const ref = simC.rows.length >= sim.rows.length ? simC.rows : sim.rows;
    setChart('ch-cred', {
      type: 'line',
      labels: labelsMeses(ref),
      tipTitle: tipMes(ref),
      series: [
        { name: 'Consórcio + crédito p/ lance', color: COR.s1, data: Array.from({ length: n }, (_, j) => (simC.rows[j] ? simC.rows[j].parcela + (j + 1 > k && j + 1 <= k + prazo ? pmt : 0) : NaN)) },
        { name: 'Somente consórcio (mesmo lance)', color: COR.s2, data: Array.from({ length: n }, (_, j) => (simC.rows[j] ? simC.rows[j].parcela : NaN)) },
        { name: 'Cenário atual do painel', color: COR.s3, dash: [5, 4], data: Array.from({ length: n }, (_, j) => (sim.rows[j] ? sim.rows[j].parcela : NaN)) }
      ],
      markers: [{ index: k - 1, label: 'Lance' }]
    });
    S.report.cred = { produto: prod.produto, head, body, texto: $('#cred-resultado .scs-callout').textContent, avisos };
  }

  // ------------------------------------------------------------------ Exportações
  function downloadBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  function exportCSV() {
    if (!S) return;
    const head = projHead();
    const lines = [head.map(cellText).join(';')].concat(S.sim.rows.map((r) => projRow(r).map((c) => cellText(c).replace(/R\$\s?| /g, '').trim()).join(';')));
    downloadBlob(new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'projecao-parcelas-' + S.id + '.csv');
  }

  function withAllPanels(fn) {
    const panels = $$('.scs-tabpanel');
    const prev = panels.map((pn) => pn.hidden);
    panels.forEach((pn) => { pn.hidden = false; });
    Object.keys(CHARTS).forEach(renderChart);
    const restore = () => { panels.forEach((pn, i) => { pn.hidden = prev[i]; }); renderActiveCharts(); };
    fn(restore);
  }
  function doPrint() {
    withAllPanels((restore) => {
      const after = () => { window.removeEventListener('afterprint', after); restore(); };
      window.addEventListener('afterprint', after);
      window.print();
      setTimeout(() => { if (document.hasFocus()) after(); }, 1500);
    });
  }

  function loadLogo() {
    return new Promise((resolve) => {
      const img = $('.scs-logo');
      if (!img || img.hidden || !img.complete || !img.naturalWidth) return resolve(null);
      try {
        const c = document.createElement('canvas');
        c.width = img.naturalWidth; c.height = img.naturalHeight;
        c.getContext('2d').drawImage(img, 0, 0);
        resolve({ data: c.toDataURL('image/png'), w: img.naturalWidth, h: img.naturalHeight });
      } catch (e) { resolve(null); }
    });
  }
  const pdfText = (s) => String(s == null ? '' : s)
    .replace(/[  ]/g, ' ').replace(/[–—]/g, '-').replace(/≈/g, '~').replace(/≥/g, '>=').replace(/≤/g, '<=').replace(/[“”]/g, '"').replace(/…/g, '...');

  async function exportPDF() {
    if (!S) return;
    const J = window.jspdf && window.jspdf.jsPDF;
    const probe = J ? new J() : null;
    if (!J || typeof probe.autoTable !== 'function') {
      alert('A biblioteca de PDF não foi carregada (verifique o acesso a cdnjs.cloudflare.com). Será aberta a impressão – escolha "Salvar como PDF".');
      doPrint();
      return;
    }
    const btn = $('[data-action="pdf"]');
    btn.disabled = true; btn.textContent = 'Gerando…';
    try {
      const doc = new J({ unit: 'mm', format: 'a4', compress: true });
      const PW = 210, PH = 297, M = 14, CW = PW - 2 * M;
      const VERDE = [63, 161, 16], ESCURO = [20, 110, 55], INK = [50, 60, 50], CINZA = [90, 100, 90];
      const logo = await loadLogo();
      const { p, sim } = S;
      const titulo = P('titulo_app', 'Simulador de Consórcios');
      const coop = P('nome_cooperativa', '');
      const hoje = new Date().toLocaleDateString('pt-BR');
      let y = 32;

      const pageH = () => doc.internal.pageSize.getHeight();
      const ensure = (h) => { if (y + h > pageH() - 16) { doc.addPage(); y = 32; } };
      const section = (t) => {
        ensure(16);
        doc.setFillColor(...VERDE);
        doc.rect(M, y - 3.6, 3.2, 3.2, 'F');
        doc.setFont('helvetica', 'bolditalic'); doc.setFontSize(12.5); doc.setTextColor(...INK);
        doc.text(pdfText(t), M + 5.5, y);
        y += 4;
      };
      const tableOpts = (extra) => Object.assign({
        startY: y,
        margin: { left: M, right: M, top: 30, bottom: 16 },
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8.5, textColor: INK, lineColor: [225, 232, 218], lineWidth: 0.1, cellPadding: 1.6, overflow: 'linebreak' },
        headStyles: { fillColor: ESCURO, textColor: 255, fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [247, 250, 244] }
      }, extra);
      const kv = (rows) => {
        doc.autoTable(tableOpts({
          body: rows.map((r) => [pdfText(r[0]), pdfText(r[1])]),
          columnStyles: { 0: { cellWidth: CW * 0.58 }, 1: { halign: 'right', fontStyle: 'bold' } }
        }));
        y = doc.lastAutoTable.finalY + 7;
      };
      const table = (head, body, extra) => {
        doc.autoTable(tableOpts(Object.assign({
          head: [head.map((h) => pdfText(cellText(h)))],
          body: body.map((r) => r.map((c) => pdfText(cellText(c)))),
          columnStyles: { 0: { halign: 'left' } },
          didParseCell: (d) => { if (d.column.index > 0) d.cell.styles.halign = 'right'; }
        }, extra || {})));
        y = doc.lastAutoTable.finalY + 7;
      };
      const paragraph = (t, size) => {
        doc.setFont('helvetica', 'normal'); doc.setFontSize(size || 9); doc.setTextColor(...CINZA);
        const lines = doc.splitTextToSize(pdfText(t), CW);
        ensure(lines.length * 4.2 + 2);
        doc.text(lines, M, y);
        y += lines.length * 4.2 + 3;
      };
      const chart = (id) => {
        const spec = CHARTS[id];
        if (!spec) return;
        const cv = document.createElement('canvas');
        drawChart(cv, spec, { width: 1000, height: 380, dpr: 1.6 });
        const h = CW * 380 / 1000;
        ensure(h + 4);
        doc.addImage(cv.toDataURL('image/jpeg', 0.9), 'JPEG', M, y, CW, h, undefined, 'FAST');
        y += h + 6;
      };

      // Capa / identificação
      section('Dados da simulação');
      kv([
        ['Simulação nº', S.id + (S.origem ? ' (editada a partir de ' + S.origem + ')' : '')],
        ['Associado(a)', p.cliente || '-'],
        ['Consultor(a)', p.consultor || '-'],
        ['Segmento / plano', segLabel(p.pl.segmento) + ' / ' + p.pl.codigo + (p.pl.descricao ? ' - ' + p.pl.descricao : '')],
        ['Taxas do plano', 'Adm. ' + fmtPct(p.TA) + ' · Fundo de reserva ' + fmtPct(p.FR) + ' · ' + (INDICE_LABEL[p.indice] || p.indice)],
        ['Crédito', fmtBRL(p.credito)],
        ['Prazo', p.prazo + ' meses (1ª parcela em ' + fmtMesLongo(p.inicio) + ')'],
        ['Tipo de parcela', p.tipoParcela === 'REDUZIDA' ? 'Reduzida em ' + fmtPct(p.reducaoPct, 0) + (p.reducaoAte === 'METADE' ? ' até a contemplação ou metade do plano' : ' até a contemplação') : 'Integral'],
        ['Lance', fmtPct(p.lanceProprioPct) + ' recursos próprios + ' + fmtPct(p.lanceEmbutidoPct) + ' embutido'],
        ['Contemplação considerada', 'Mês ' + S.cMes + ' (' + (p.contempModo === 'MANUAL' ? 'informada' : 'projeção estatística') + ')']
      ]);

      section('Simulação geral');
      kv([
        ['1ª parcela', fmtBRL(sim.rows[0].parcela)],
        ['Parcela pós-contemplação', sim.lance ? fmtBRL(sim.lance.parcelaDepois) : '-'],
        ['Total desembolsado', fmtBRL(sim.desembolso)]
      ].concat(S.report.resumo.comp));
      chart('ch-resumo');

      if (S.report.reduzida) {
        section('Parcela reduzida');
        table(S.report.reduzida.head, S.report.reduzida.body);
        chart('ch-reduzida');
      }

      section('Composição da parcela ao longo do plano');
      chart('ch-projecao');

      section('Projeção de contemplação');
      kv(S.report.contemplacao.kv);
      chart('ch-contemp');

      section('Recomendação de lance (próxima assembleia)');
      table(['Cenário', '% do lance', 'Valor do lance', 'Recursos próprios'], S.report.contemplacao.cards);
      table(S.report.contemplacao.head, S.report.contemplacao.body);
      if (CHARTS['ch-lances']) chart('ch-lances');

      if (S.report.pos) {
        section('Parcelas pós-contemplação');
        kv(S.report.pos.kv);
        table(S.report.pos.head, S.report.pos.body);
        chart('ch-pos');
      }

      if (S.report.fin) {
        section('Comparativo com financiamento');
        table(S.report.fin.head, S.report.fin.body);
        paragraph(S.report.fin.conclusao);
        chart('ch-fin');
      }

      if (S.report.inv) {
        section('Investimento para formar o lance');
        paragraph('Produto: ' + S.report.inv.produto + '. ' + S.report.inv.texto);
        table(S.report.inv.head, S.report.inv.body);
        chart('ch-inv');
      }
      if (S.report.cred) {
        section('Crédito para lance - ' + S.report.cred.produto);
        if (S.report.cred.avisos.length) paragraph(S.report.cred.avisos.join(' '));
        paragraph(S.report.cred.texto);
        table(S.report.cred.head, S.report.cred.body);
        chart('ch-cred');
      }

      doc.addPage('a4', 'l'); y = 32;
      section('Projeção de parcelas mês a mês');
      table(projHead(), sim.rows.map(projRow), {
        styles: { font: 'helvetica', fontSize: 6.6, textColor: INK, lineColor: [225, 232, 218], lineWidth: 0.1, cellPadding: 1 },
        headStyles: { fillColor: ESCURO, textColor: 255, fontStyle: 'bold', fontSize: 6.6 }
      });

      section('Aviso legal');
      paragraph(P('aviso_legal', ''), 8.5);

      // Cabeçalho e rodapé em todas as páginas
      const pages = doc.getNumberOfPages();
      for (let i = 1; i <= pages; i++) {
        doc.setPage(i);
        const PW = doc.internal.pageSize.getWidth(), PH = doc.internal.pageSize.getHeight();
        doc.setFillColor(...VERDE);
        doc.rect(0, 0, PW, 22, 'F');
        let tx = M;
        if (logo) {
          const lh = 9, lw = Math.min(40, lh * logo.w / logo.h);
          doc.setFillColor(255, 255, 255);
          doc.rect(M, 4.5, lw + 6, lh + 4, 'F');
          doc.addImage(logo.data, 'PNG', M + 3, 6.5, lw, lh);
          tx = M + lw + 12;
        }
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bolditalic'); doc.setFontSize(15);
        doc.text(pdfText(titulo), tx, 11.5);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
        doc.text(pdfText((coop ? coop + ' · ' : '') + 'Simulação ' + (S.id || '') + ' · ' + hoje), tx, 17);
        doc.setDrawColor(...VERDE); doc.setLineWidth(0.4);
        doc.line(M, PH - 11, PW - M, PH - 11);
        doc.setTextColor(...CINZA); doc.setFontSize(7.5);
        doc.text(pdfText('Simulação ilustrativa, sujeita às regras do grupo. Não constitui proposta.'), M, PH - 6.5);
        doc.text('Página ' + i + ' de ' + pages, PW - M, PH - 6.5, { align: 'right' });
      }
      const nome = (p.cliente || 'associado').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-zA-Z0-9]+/g, '-').toLowerCase();
      doc.save('simulacao-consorcio-' + nome + '-' + ym(new Date()) + '.pdf');
    } catch (e) {
      console.error(e);
      alert('Não foi possível gerar o PDF: ' + e.message);
    } finally {
      btn.disabled = false; btn.textContent = 'Exportar PDF';
      renderActiveCharts();
    }
  }

  // ------------------------------------------------------------------ Eventos
  let timer = null, formTimer = null;
  const schedule = () => { clearTimeout(timer); timer = setTimeout(recalc, 250); };
  const scheduleForm = () => { clearTimeout(formTimer); formTimer = setTimeout(updateForm, 200); };

  function activateTab(tab) {
    activeTab = tab;
    $$('.scs-tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === tab)));
    $$('.scs-tabpanel').forEach((pn) => { pn.hidden = pn.dataset.panel !== tab; });
    renderActiveCharts();
  }
  const naTelaDeDados = (el) => !!el.closest('[data-screen="form"]');

  function bindEvents() {
    $('#scs-form').addEventListener('submit', (ev) => { ev.preventDefault(); simular(); });
    root.addEventListener('change', (ev) => {
      const t = ev.target;
      if (naTelaDeDados(t)) {
        if (t.name === 'in-segmento') { fillPrazos(); applyPlanDefaults(true); }
        else if (t.id === 'in-prazo' || t.id === 'in-credito') applyPlanDefaults(false);
        else { toggleBoxes(); scheduleForm(); }
        $('#form-error').textContent = '';
        return;
      }
      if (!S) return;
      if (t.id === 'fin-modalidade') applyFinDefaults(S.p.pl);
      schedule();
    });
    root.addEventListener('input', (ev) => {
      if (!ev.target.matches('input')) return;
      if (naTelaDeDados(ev.target)) scheduleForm();
      else if (S) schedule();
    });
    root.addEventListener('focusout', (ev) => {
      if (ev.target.classList && ev.target.classList.contains('scs-money')) ev.target.value = fmtNum(parseMoney(ev.target.value), 2);
    });
    root.addEventListener('click', (ev) => {
      const tabBtn = ev.target.closest('.scs-tab');
      if (tabBtn) { activateTab(tabBtn.dataset.tab); return; }
      const act = ev.target.closest('[data-action]');
      if (!act) return;
      const a = act.dataset.action;
      if (a === 'pdf') exportPDF();
      else if (a === 'print') doPrint();
      else if (a === 'csv') exportCSV();
      else if (a === 'editar') { editandoId = S ? S.id : null; showScreen('form'); updateForm(); }
      else if (a === 'nova') { editandoId = null; resetForm(true); showScreen('form'); }
      else if (a === 'limpar') resetForm(false);
      else if (a === 'usar-sugestao') { $('#inv-mensal').value = fmtNum(num(act.dataset.valor), 2); recalc(); }
      else if (a === 'aplicar-inv' && S && S.inv && S.inv.reached) {
        // leva a estratégia de investimento para a simulação principal (nova simulação derivada desta)
        const r = S.inv.rows[S.inv.reached - 1];
        $('#in-contemp-modo').value = 'MANUAL';
        $('#in-contemp-mes').value = S.inv.reached;
        $('#in-lance-inicio').value = S.inv.reached;
        $('#in-lance-proprio').value = Math.round(r.pctProprio * 100) / 100;
        toggleBoxes();
        editandoId = S.id;
        simular();
      }
    });
    // Navegação por teclado entre abas
    $('.scs-tabs').addEventListener('keydown', (ev) => {
      if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
      const tabs = $$('.scs-tab');
      const i = tabs.findIndex((t) => t.dataset.tab === activeTab);
      const next = tabs[(i + (ev.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length];
      next.focus(); activateTab(next.dataset.tab);
    });
    let rt = null;
    const onResize = () => { clearTimeout(rt); rt = setTimeout(renderActiveCharts, 150); };
    if (window.ResizeObserver) new ResizeObserver(onResize).observe($('[data-screen="result"]'));
    else window.addEventListener('resize', onResize);
  }

  // ------------------------------------------------------------------ Inicialização
  function showStatus(msg, isError) {
    const st = $('.scs-status');
    st.textContent = msg;
    st.classList.toggle('is-error', !!isError);
  }

  async function init() {
    const logoSrc = root.dataset.logo;
    const img = $('.scs-logo');
    if (logoSrc && img) {
      img.addEventListener('load', () => { img.hidden = false; });
      img.addEventListener('error', () => { img.closest('.scs-logo-box').hidden = true; });
      img.src = logoSrc;
    } else if (img) img.closest('.scs-logo-box').hidden = true;

    let loaded;
    try {
      loaded = await loadAll();
    } catch (e) {
      showStatus('Erro ao carregar as bases: ' + e.message, true);
      return;
    }
    const fatal = loaded.errors.filter((e) => REQUIRED.includes(e.key));
    if (fatal.length) {
      showStatus('Não foi possível carregar as bases obrigatórias a partir de "' + loaded.base + '":\n' + fatal.map((e) => '• ' + e.msg).join('\n') +
        (location.protocol === 'file:' ? '\n\nAbra a página por um servidor web (SharePoint ou "python -m http.server"): navegadores bloqueiam a leitura de CSV via file://.' : ''), true);
      return;
    }
    normalize(loaded.raw);
    if (!DB.planos.length) { showStatus('planos.csv não possui planos ativos válidos.', true); return; }

    $$('[data-bind="titulo"]').forEach((el) => { el.textContent = P('titulo_app', el.textContent); });
    const coop = P('nome_cooperativa', '');
    if (coop) $$('[data-bind="cooperativa"]').forEach((el) => { el.textContent = coop + ' · simulação, contemplação e lance'; });
    $$('[data-bind="aviso"]').forEach((el) => { el.textContent = P('aviso_legal', ''); });
    const ultIdx = DB.idxKeys.filter((k) => DB.idx.get(k).tipo !== 'PROJETADO').pop();
    $('#base-info').textContent = 'Base: ' + DB.planos.length + ' planos · ' + DB.lances.length + ' assembleias no histórico · indexadores realizados até ' +
      (ultIdx || '—') + ', projetados até ' + (DB.idxKeys[DB.idxKeys.length - 1] || '—') + '.';

    fillSegments();
    $$('.scs-money').forEach((el) => { el.value = fmtNum(parseMoney(el.value), 2); });
    resetForm(false);
    bindEvents();

    if (loaded.errors.length) showStatus('Bases opcionais não carregadas: ' + loaded.errors.map((e) => e.msg).join(', ') + '. As seções dependentes ficam desabilitadas.', true);
    else showStatus('', false);
    showScreen('form');

    // Atalho de teste: index.html#simular&tab=contemplacao
    if (/simular/.test(location.hash)) {
      simular();
      const hashTab = (location.hash.match(/tab=([a-z]+)/) || [])[1];
      if (hashTab && $('.scs-tab[data-tab="' + hashTab + '"]')) activateTab(hashTab);
    }
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(renderActiveCharts);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
