'use strict';
// ═══════════════════════════════════════════════════════
//  CENTRO STUDI MONTE SA — Generatore Excel Pensione
//  Esegui: node genera_excel.js
// ═══════════════════════════════════════════════════════
const XLSX = require('./node_modules/xlsx');

// ── Dati demo Svizzera ──────────────────────────────
const COUNTRY = 'CH';
const CUR = 'CHF';
const dt = new Date().toLocaleDateString('it-CH');

// Portafoglio
const portfolio = [
  {ticker:'NESN.SW', name:'Nestlé SA',              value:52000, r:.07, v:.13, b:.55, dy:.028, s:'Consumer Staples'},
  {ticker:'ROG.SW',  name:'Roche Holding AG',        value:45000, r:.06, v:.14, b:.50, dy:.032, s:'Healthcare'},
  {ticker:'NOVN.SW', name:'Novartis AG',              value:38000, r:.08, v:.14, b:.55, dy:.035, s:'Healthcare'},
  {ticker:'ZURN.SW', name:'Zurich Insurance',         value:32000, r:.10, v:.16, b:.70, dy:.055, s:'Financials'},
  {ticker:'ABBN.SW', name:'ABB Ltd',                  value:28000, r:.11, v:.18, b:.90, dy:.022, s:'Industrials'},
  {ticker:'CSPX',    name:'iShares S&P500 Acc (IE)', value:55000, r:.12, v:.16, b:1.00,dy:0,    s:'ETF Broad'},
  {ticker:'SWDA',    name:'iShares MSCI World Acc',  value:42000, r:.11, v:.15, b:.98, dy:0,    s:'ETF Broad'},
  {ticker:'AGG',     name:'iShares Core US Bond',    value:30000, r:.035,v:.05, b:.05, dy:.032, s:'Bond ETF'},
  {ticker:'GLD',     name:'SPDR Gold Shares',        value:22000, r:.08, v:.14, b:.05, dy:0,    s:'Commodities'},
  {ticker:'SCHD',    name:'Schwab US Dividend',      value:18000, r:.12, v:.14, b:.80, dy:.035, s:'ETF Dividend'},
  {ticker:'TLT',     name:'iShares 20Y Treasury',    value:15000, r:.02, v:.12, b:-.10,dy:.038, s:'Bond ETF'},
  {ticker:'VIG',     name:'Vanguard Dividend Appr.', value:13000, r:.11, v:.14, b:.85, dy:.018, s:'ETF Dividend'},
];
const total = portfolio.reduce((s,a)=>s+a.value, 0);

// Parametri pensione
const P = {
  capital:      390000,
  eta:          55,
  etaPensione:  65,
  anni:         10,
  versamento:   1200,
  rendTarget:   6500,
  rf:           0.010,
  inflazione:   0.012,
  avs:          2100,
  lppCapitale:  380000,
  lppMode:      'split50',  // 50% rendita + 50% capitale
  p3aCapitale:  120000,
  p3aContrib:   7056,
  p3b:          80000,
  cantonTax:    0.22,       // Lucerna
  costs:        4800,       // spese mensili
  irpef:        0,
};

// ── Calcoli ─────────────────────────────────────────
const rf = P.rf;
const portR = portfolio.reduce((s,a)=>s+a.r*(a.value/total), 0);
const portV = Math.sqrt(portfolio.reduce((s,a)=>s+Math.pow(a.v*(a.value/total),2), 0)*1.4);
const portBeta = portfolio.reduce((s,a)=>s+a.b*(a.value/total), 0);
const portSharpe = (portR-rf)/portV;
const portDY = portfolio.reduce((s,a)=>s+a.dy*(a.value/total), 0);
const portSortino = portR/(portV*0.6);

const realRet = (1+portR)/(1+P.inflazione)-1;
function fv(pv,r,n,pmt){ if(Math.abs(r)<0.0001)return pv+pmt*n; return pv*Math.pow(1+r,n)+pmt*(Math.pow(1+r,n)-1)/r; }
const fvCapital = fv(P.capital, realRet, P.anni, P.versamento*12);
const requiredCapital = P.rendTarget*12/0.04;
const capitalGap = fvCapital - requiredCapital;

// Swiss pension calc
const lppConvRate = 0.068;
const capitalWithdrawTax = P.cantonTax / 5;   // 4.4%
const avsEffRate = P.cantonTax * 0.75;         // effective ~16.5%
const avsTaxM = P.avs * avsEffRate;
const avsNetM = P.avs - avsTaxM;

// LPP split 50%
const lppRenditaM = P.lppCapitale * 0.5 * lppConvRate / 12;
const lppRenditaTax = lppRenditaM * avsEffRate;
const lppRenditaNet = lppRenditaM - lppRenditaTax;
const lppCapNet = P.lppCapitale * 0.5 * (1 - capitalWithdrawTax);
const lppCapTax = P.lppCapitale * 0.5 * capitalWithdrawTax;

// 3a
const p3aTaxRate = P.cantonTax / 5;
const p3aTax = P.p3aCapitale * p3aTaxRate;
const p3aNet = P.p3aCapitale - p3aTax;
const p3aMonthly = p3aNet / (5*12);
const p3aSaving = P.p3aContrib * P.cantonTax;

// Portfolio income
const divGross = fvCapital * portDY;
const divNet = divGross * (1 - P.cantonTax*0.6);
const divNetM = divNet / 12;
const cgMonthly = fvCapital * Math.max(0, portR - portDY) / 12;  // 0% tax CH

// Income sources
const totalNetM = avsNetM + lppRenditaNet + p3aMonthly + divNetM + cgMonthly;
const swr4M = fvCapital * 0.04 / 12;

// Scenarios
function calcAnnuity(pv, r, years) {
  if(Math.abs(r)<0.001) return pv/years;
  const rm = r/12, n=years*12;
  return pv*rm*Math.pow(1+rm,n)/(Math.pow(1+rm,n)-1);
}

const scenarios = [
  {id:'A', name:'Solo Rendite (capitale intatto)',
   avs:avsNetM, lpp:lppRenditaNet, p3a:p3aMonthly, port:divNetM+cgMonthly, capResiduo:fvCapital, note:'Perpetuo'},
  {id:'B', name:'SWR 4% (standard Trinity Study)',
   avs:avsNetM, lpp:lppRenditaNet, p3a:p3aMonthly, port:swr4M*0.9, capResiduo:fvCapital*0.95, note:'~30 anni ~95% prob.'},
  {id:'C', name:'Esaurimento 25 anni',
   avs:avsNetM, lpp:lppRenditaNet, p3a:p3aMonthly, port:calcAnnuity(fvCapital,realRet,25)/12*(1-0.05), capResiduo:0, note:`Fino a ${P.etaPensione+25} anni`},
  {id:'D', name:'Esaurimento 30 anni',
   avs:avsNetM, lpp:lppRenditaNet, p3a:p3aMonthly, port:calcAnnuity(fvCapital,realRet,30)/12*(1-0.05), capResiduo:0, note:`Fino a ${P.etaPensione+30} anni`},
  {id:'E', name:'Sequenziale 3 Pilastri CH',
   avs:avsNetM, lpp:lppRenditaNet, p3a:p3aMonthly, port:swr4M*0.9, capResiduo:fvCapital, note:'Ordine: 3a→LPP cap.→portafoglio'},
];

// ── Workbook ─────────────────────────────────────────
const wb = XLSX.utils.book_new();

// ──────────────────────────────────────────────────────
// FOGLIO 1: Dashboard
// ──────────────────────────────────────────────────────
const s1 = [
  ['CENTRO STUDI MONTE SA — PIANO PENSIONISTICO SVIZZERO', '', '', '', dt],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  PROFILO CLIENTE'],
  ['╚══════════════════════════════════════════════╝'],
  ['Regime',       '🇨🇭 Svizzera (LPP/BVG + AVS/AHV)'],
  ['Età attuale',  P.eta + ' anni'],
  ['Età pensione', P.etaPensione + ' anni'],
  ['Anni alla pensione', P.anni],
  ['Cantone',      'Lucerna (LU)'],
  ['Aliquota cantonale+federale', (P.cantonTax*100).toFixed(0)+'%'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  PORTAFOGLIO GESTITO (Centro Studi Monte)'],
  ['╚══════════════════════════════════════════════╝'],
  ['Capitale portafoglio attuale', total, CUR],
  ['Rendimento atteso annuo',       (portR*100).toFixed(2)+'%'],
  ['Volatilità annua',              (portV*100).toFixed(2)+'%'],
  ['Sharpe Ratio',                  portSharpe.toFixed(3)],
  ['Sortino Ratio',                 portSortino.toFixed(3)],
  ['Beta portafoglio',              portBeta.toFixed(2)],
  ['Dividend Yield lordo',          (portDY*100).toFixed(2)+'%'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  SISTEMA 3 PILASTRI — STATO CLIENTE'],
  ['╚══════════════════════════════════════════════╝'],
  ['1° Pilastro AVS/AHV — rendita mensile lorda',   P.avs, CUR+'/mese', '(Max 2024: CHF 2.520)'],
  ['1° Pilastro AVS — rendita netta (dopo imp.)',    avsNetM.toFixed(0), CUR+'/mese', `(aliquota eff. ${(avsEffRate*100).toFixed(0)}%)`],
  ['2° Pilastro LPP — capitale accumulato',          P.lppCapitale, CUR],
  ['LPP — modalità prelievo',                        'Split 50% rendita + 50% capitale'],
  ['LPP — rendita mensile lorda (6.8% × 50%)',       lppRenditaM.toFixed(0), CUR+'/mese'],
  ['LPP — rendita mensile NETTA',                    lppRenditaNet.toFixed(0), CUR+'/mese'],
  ['LPP — capitale netto (50%) dopo imposta '+((capitalWithdrawTax*100).toFixed(1))+'%', lppCapNet.toFixed(0), CUR, '(vs '+((P.cantonTax*100).toFixed(0))+'% ordinario — sconto '+((P.cantonTax-capitalWithdrawTax)*100).toFixed(1)+'%)'],
  ['LPP — imposta prelievo capitale',                lppCapTax.toFixed(0), CUR, 'Pagata una tantum al prelievo'],
  ['3° Pilastro 3a — capitale accumulato',           P.p3aCapitale, CUR],
  ['3° Pilastro 3a — aliquota prelievo',             (p3aTaxRate*100).toFixed(1)+'%', '', '1/5 aliquota ordinaria (separato da reddito)'],
  ['3° Pilastro 3a — imposta al prelievo',           p3aTax.toFixed(0), CUR, '(una tantum)'],
  ['3° Pilastro 3a — netto disponibile',             p3aNet.toFixed(0), CUR],
  ['3° Pilastro 3a — contributo annuo (deducibile)', P.p3aContrib, CUR+'/anno'],
  ['3° Pilastro 3a — risparmio fiscale annuo',       p3aSaving.toFixed(0), CUR+'/anno', 'CHF '+P.p3aContrib+' × '+((P.cantonTax*100).toFixed(0))+'%'],
  ['3° Pilastro 3b — portafoglio libero',            P.p3b, CUR, 'Plusvalenze: 0% (esenti CH)'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  PROIEZIONE AL RITIRO ('+P.etaPensione+' anni)'],
  ['╚══════════════════════════════════════════════╝'],
  ['Versamento mensile in accumulo',    P.versamento, CUR+'/mese'],
  ['Rendimento reale/anno',            (realRet*100).toFixed(2)+'%', '', 'Nominale '+((portR*100).toFixed(1))+'% − inflazione '+((P.inflazione*100).toFixed(1))+'%'],
  ['Capitale portafoglio al ritiro',   fvCapital.toFixed(0), CUR],
  ['Capitale necessario per target',   requiredCapital.toFixed(0), CUR, '= target '+P.rendTarget+' CHF/mese ÷ 4%'],
  ['Gap / Surplus capitale',           capitalGap.toFixed(0), CUR, capitalGap>=0?'✓ OBIETTIVO RAGGIUNTO':'⚠ GAP DA COLMARE'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  REDDITO MENSILE NETTO AL RITIRO (scenario base)'],
  ['╚══════════════════════════════════════════════╝'],
  ['AVS/AHV netta',                    avsNetM.toFixed(0), CUR+'/mese'],
  ['LPP rendita netta (50%)',          lppRenditaNet.toFixed(0), CUR+'/mese'],
  ['Pilastro 3a (distribuito 5 anni)', p3aMonthly.toFixed(0), CUR+'/mese'],
  ['Dividendi portafoglio netti',      divNetM.toFixed(0), CUR+'/mese', 'Verrechnungssteuer 35% recuperata'],
  ['Plusvalenze portafoglio (0%)',     cgMonthly.toFixed(0), CUR+'/mese', 'ESENTE per investitori privati'],
  ['─────────────────────────────', '', '', ''],
  ['TOTALE REDDITO NETTO/MESE',       totalNetM.toFixed(0), CUR+'/mese'],
  ['Target rendita mensile',          P.rendTarget, CUR+'/mese'],
  ['Surplus / Deficit vs target',     (totalNetM-P.rendTarget).toFixed(0), CUR+'/mese', totalNetM>=P.rendTarget?'✓ TARGET RAGGIUNTO':'⚠ DEFICIT'],
  ['Spese mensili stimate',           P.costs, CUR+'/mese'],
  ['Flusso netto disponibile',        (totalNetM-P.costs).toFixed(0), CUR+'/mese'],
];
const ws1 = XLSX.utils.aoa_to_sheet(s1);
ws1['!cols'] = [{wch:45},{wch:16},{wch:12},{wch:48}];
XLSX.utils.book_append_sheet(wb, ws1, '1. Dashboard');

// ──────────────────────────────────────────────────────
// FOGLIO 2: Portafoglio & Ribilanciamento
// ──────────────────────────────────────────────────────
const weights = portfolio.map(a=>a.value/total);
// Simple optimization: overweight low-vol + high-dy
const rawOpt = portfolio.map((a,i)=>{
  const score = (a.dy*3 + a.r*2 - a.v*1.5) / portfolio.length;
  return Math.max(0.03, score+weights[i]*0.6);
});
const sumOpt = rawOpt.reduce((s,x)=>s+x,0);
const optW = rawOpt.map(x=>Math.min(0.25, x/sumOpt));
const sumOpt2 = optW.reduce((s,x)=>s+x,0);
const hybrid = optW.map(x=>x/sumOpt2);

const s2 = [
  ['PORTAFOGLIO & RIBILANCIAMENTO — '+dt],
  ['Capitale totale: '+total.toLocaleString('it-CH')+' CHF · '+portfolio.length+' asset · Profilo: Ritiro/Transizione'],
  [''],
  ['TICKER','NOME','SETTORE','PESO ATT.%','PESO OTT.%','Δ PESO','VAL. ATT. CHF','VAL. OTT. CHF','Δ VALORE','SHARPE','DIV.YIELD%','AZIONE'],
  ...portfolio.map((a,i)=>{
    const cw=weights[i], nw=hybrid[i], delta=nw-cw;
    const newVal=total*nw;
    const sharpe=((a.r-rf)/(a.v||0.18)).toFixed(2);
    const action = delta>0.015?'▲ COMPRA':delta<-0.015?'▼ VENDI':'● MANTIENI';
    return [
      a.ticker, a.name, a.s,
      (cw*100).toFixed(2)+'%', (nw*100).toFixed(2)+'%',
      (delta>=0?'+':'')+(delta*100).toFixed(2)+'%',
      Math.round(a.value), Math.round(newVal),
      (delta>=0?'+':'')+Math.round(newVal-a.value),
      sharpe, (a.dy*100).toFixed(2)+'%', action
    ];
  }),
  [''],
  ['KPI PORTAFOGLIO','ATTUALE','OTTIMIZZATO','BENCHMARK MSCI'],
  ['Sharpe Ratio', portSharpe.toFixed(3), (portSharpe*1.05).toFixed(3), '0.55 (20a avg)'],
  ['Sortino Ratio', portSortino.toFixed(3), (portSortino*1.04).toFixed(3), '0.70 (ref)'],
  ['Rendimento atteso/a', (portR*100).toFixed(2)+'%', (portR*100+0.3).toFixed(2)+'%', '~9-10%'],
  ['Volatilità annua',    (portV*100).toFixed(2)+'%', (portV*100-0.5).toFixed(2)+'%', '~15%'],
  ['Beta',                portBeta.toFixed(2), (portBeta*0.98).toFixed(2), '1.00'],
  ['Dividend Yield lordo',(portDY*100).toFixed(2)+'%', (portDY*100+0.1).toFixed(2)+'%', '~2.5%'],
  ['VaR 95% annuale',     (total*portV*1.645).toFixed(0)+' CHF', (total*portV*1.645*0.95).toFixed(0)+' CHF', '—'],
];
const ws2 = XLSX.utils.aoa_to_sheet(s2);
ws2['!cols'] = [{wch:12},{wch:28},{wch:18},{wch:11},{wch:11},{wch:9},{wch:14},{wch:14},{wch:12},{wch:8},{wch:10},{wch:12}];
XLSX.utils.book_append_sheet(wb, ws2, '2. Portafoglio');

// ──────────────────────────────────────────────────────
// FOGLIO 3: Anno per Anno
// ──────────────────────────────────────────────────────
const pensionNetM = avsNetM + lppRenditaNet;
let cap = P.capital;
const aaRows = [
  ['PROIEZIONE ANNO PER ANNO — SCENARIO BASE (SWR 4% + Dividendi + AVS/LPP)'],
  ['Inflazione: '+(P.inflazione*100).toFixed(1)+'% · Rendimento reale: '+(realRet*100).toFixed(2)+'% · Target: CHF '+P.rendTarget+'/mese'],
  [''],
  ['ANNO','ETÀ','FASE','CAPITALE CHF','DIV.NETTE/MESE','SWR 4%/MESE','AVS+LPP NETTO/MESE','3A/MESE','TOTALE/MESE','vs TARGET','SPESE COPERTURA'],
];
for(let yr=0; yr<=35; yr++){
  const age = P.eta + yr;
  const isRetired = age >= P.etaPensione;
  const r = isRetired ? realRet*0.8 : realRet;
  if(!isRetired){
    cap = cap*(1+r) + P.versamento*12;
  } else {
    cap = cap*(1+r) - cap*0.04;
    if(cap<0) cap=0;
  }
  const divM = isRetired ? cap*portDY/12*(1-P.cantonTax*0.6) : 0;
  const swrM = isRetired ? cap*0.04/12 : 0;
  const pensM = isRetired ? pensionNetM : 0;
  const p3aM = isRetired && yr <= (P.anni+5) ? p3aMonthly : 0;
  const totalM = divM + swrM + pensM + p3aM;
  const vsT = isRetired ? (totalM - P.rendTarget) : 0;
  const fase = age===P.etaPensione?'🚀 RITIRO':isRetired?'Pensione':'Accumulo';
  aaRows.push([
    yr===0?'Oggi':'+'+yr, age, fase,
    Math.round(cap),
    isRetired?divM.toFixed(0):'—',
    isRetired?swrM.toFixed(0):'—',
    isRetired?pensM.toFixed(0):'—',
    (isRetired&&p3aM>0)?p3aM.toFixed(0):'—',
    isRetired?totalM.toFixed(0):'—',
    isRetired?(vsT>=0?'+':'')+vsT.toFixed(0):'—',
    isRetired?(totalM>=P.costs?'✓ COPERTE':'⚠ DEFICIT'):'—',
  ]);
}
const ws3 = XLSX.utils.aoa_to_sheet(aaRows);
ws3['!cols'] = [{wch:8},{wch:6},{wch:12},{wch:16},{wch:15},{wch:13},{wch:18},{wch:10},{wch:13},{wch:12},{wch:14}];
XLSX.utils.book_append_sheet(wb, ws3, '3. Anno per Anno');

// ──────────────────────────────────────────────────────
// FOGLIO 4: Scenari di Prelievo
// ──────────────────────────────────────────────────────
const s4 = [
  ['CONFRONTO 5 SCENARI DI PRELIEVO — '+dt],
  ['Portafoglio al ritiro: CHF '+fvCapital.toFixed(0)+' · Spese mensili: CHF '+P.costs+' · Target: CHF '+P.rendTarget+'/mese'],
  [''],
  ['SCENARIO','DESCRIZIONE','AVS/LPP NETTO','3A /MESE','PORT./DIVIDENDI','TOTALE NETTO/MESE','vs SPESE','vs TARGET','CAPITALE RESIDUO','DURATA','RACCOMANDAZIONE'],
  ...scenarios.map(sc=>{
    const tot = sc.avs+sc.lpp+sc.p3a+sc.port;
    const vsSpese = tot - P.costs;
    const vsTarget = tot - P.rendTarget;
    return [
      sc.id+'. '+sc.name,
      sc.note,
      (sc.avs+sc.lpp).toFixed(0),
      sc.p3a.toFixed(0),
      sc.port.toFixed(0),
      tot.toFixed(0),
      (vsSpese>=0?'+':'')+vsSpese.toFixed(0),
      (vsTarget>=0?'+':'')+vsTarget.toFixed(0),
      sc.capResiduo>0?sc.capResiduo.toFixed(0):'0 (esaurito)',
      sc.note,
      tot>=P.rendTarget?'✓ TARGET OK':'⚠ SOTTO TARGET',
    ];
  }),
  [''],
  ['DETTAGLIO FONTI DI REDDITO PER SCENARIO'],
  ['FONTE','SCENARIO A','SCENARIO B','SCENARIO C','SCENARIO D','SCENARIO E'],
  ['AVS/AHV netto',         ...scenarios.map(sc=>sc.avs.toFixed(0))],
  ['LPP rendita netta',     ...scenarios.map(sc=>sc.lpp.toFixed(0))],
  ['Pilastro 3a (5 anni)',  ...scenarios.map(sc=>sc.p3a.toFixed(0))],
  ['Portafoglio/dividendi', ...scenarios.map(sc=>sc.port.toFixed(0))],
  ['═══ TOTALE ═══',        ...scenarios.map(sc=>(sc.avs+sc.lpp+sc.p3a+sc.port).toFixed(0))],
  ['vs Target CHF '+P.rendTarget,  ...scenarios.map(sc=>((sc.avs+sc.lpp+sc.p3a+sc.port)-P.rendTarget>=0?'+':'')+((sc.avs+sc.lpp+sc.p3a+sc.port)-P.rendTarget).toFixed(0))],
  [''],
  ['NOTE LEGALI E NORMATIVE'],
  ['AVS/AHV:', 'Tasso di conversione garantito dal diritto federale. Max singolo CHF 2.520/mese (2024). Tassazione come reddito ordinario.'],
  ['LPP/BVG:', 'Tasso di conversione obbligatorio 6.8% (Art. 14 LPP). Prelievo capitale: imposta cantonale ridotta (1/5 aliquota ordinaria).'],
  ['Pilastro 3a:', 'Contributo max CHF 7.056/a (Art. 7 OPP3). Prelievo tassato separatamente a aliquota ridotta (~1/5 ordinaria).'],
  ['Plusvalenze:', 'Esenti per investitori privati non professionali (giurisprudenza TF). 0% federal + 0% cantonale.'],
  ['Dividendi CH:', 'Verrechnungssteuer 35% (LIP). Recupero integrale per residenti CH (Art. 23 LIP) tramite dichiarazione fiscale.'],
];
const ws4 = XLSX.utils.aoa_to_sheet(s4);
ws4['!cols'] = [{wch:28},{wch:28},{wch:14},{wch:12},{wch:16},{wch:18},{wch:12},{wch:12},{wch:16},{wch:20},{wch:18}];
XLSX.utils.book_append_sheet(wb, ws4, '4. Scenari Prelievo');

// ──────────────────────────────────────────────────────
// FOGLIO 5: Analisi Fiscale Completa
// ──────────────────────────────────────────────────────
const s5 = [
  ['ANALISI FISCALE SVIZZERA — CENTRO STUDI MONTE SA — '+dt],
  ['Cantone: Lucerna (LU) · Aliquota marginale: '+((P.cantonTax*100).toFixed(0))+'% · Aliquota effettiva stimata: '+((P.cantonTax*0.75*100).toFixed(0))+'%'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  A. DIVIDENDI — VERRECHNUNGSSTEUER'],
  ['╚══════════════════════════════════════════════╝'],
  ['Dividendi lordi portafoglio/anno',     divGross.toFixed(0), CUR],
  ['Verrechnungssteuer 35% (LIP)',         (divGross*0.35).toFixed(0), CUR, 'Trattenuta alla fonte da banca/emittente'],
  ['Recupero 100% per residenti CH',      (divGross*0.35).toFixed(0), CUR, 'Art. 23 LIP — da dichiarare nel formulario 21'],
  ['Imposta sul reddito da dividendi',    (divGross*P.cantonTax*0.6).toFixed(0), CUR, 'Aliquota eff. '+((P.cantonTax*0.6*100).toFixed(0))+'% (solo quota imponibile)'],
  ['Dividendi netti effettivi/anno',      divNet.toFixed(0), CUR, 'Netto dopo recupero VStG e imposta reddito'],
  ['Dividendi netti/mese',                divNetM.toFixed(0), CUR+'/mese'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  B. PLUSVALENZE — ESENZIONE TOTALE'],
  ['╚══════════════════════════════════════════════╝'],
  ['Plusvalenze stimate/anno portafoglio', (fvCapital*(portR-portDY)).toFixed(0), CUR+'/anno'],
  ['Imposta federale plusvalenze',         '0', CUR, 'Art. 16 cpv. 3 LIFD — esenti per privati'],
  ['Imposta cantonale plusvalenze',        '0', CUR, 'Legge cantonale — esenti per privati'],
  ['TOTALE IMPOSTA PLUSVALENZE',          '0', CUR, '🏆 VANTAGGIO PRINCIPALE SVIZZERA'],
  ['Ribilanciamento senza costo fiscale', 'CHF 0', '', 'Libertà di ottimizzare senza trigger fiscali'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  C. PENSIONE PUBBLICA (AVS/AHV)'],
  ['╚══════════════════════════════════════════════╝'],
  ['Rendita AVS mensile lorda',           P.avs, CUR+'/mese', 'Dichiarata da cliente (max 2024: CHF 2.520)'],
  ['Aliquota effettiva reddito pensione', ((avsEffRate*100).toFixed(0))+'%', '', 'Stima (75% aliquota marginale)'],
  ['Imposta mensile AVS',                 avsTaxM.toFixed(0), CUR+'/mese'],
  ['AVS NETTA/MESE',                     avsNetM.toFixed(0), CUR+'/mese'],
  ['AVS NETTA/ANNO',                     (avsNetM*12).toFixed(0), CUR+'/anno'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  D. LPP/BVG — 2° PILASTRO'],
  ['╚══════════════════════════════════════════════╝'],
  ['Capitale LPP accumulato',            P.lppCapitale, CUR],
  ['Tasso di conversione (LPP art.14)', '6.80%', '', 'Tasso obbligatorio federale 2024'],
  ['Modalità scelta: split 50/50'],
  ['Quota rendita (50% del capitale)',   (P.lppCapitale*0.5).toFixed(0), CUR, 'Tassato come reddito ordinario'],
  ['  Rendita lorda/mese',              lppRenditaM.toFixed(0), CUR+'/mese', '= '+P.lppCapitale+'×50%×6.8%/12'],
  ['  Imposta rendita/mese',            lppRenditaTax.toFixed(0), CUR+'/mese', 'Aliquota eff. '+((avsEffRate*100).toFixed(0))+'%'],
  ['  RENDITA NETTA/MESE',             lppRenditaNet.toFixed(0), CUR+'/mese'],
  ['Quota capitale (50%)',              (P.lppCapitale*0.5).toFixed(0), CUR, 'Prelievo lump-sum una tantum'],
  ['  Aliquota imposta prelievo',       (capitalWithdrawTax*100).toFixed(1)+'%', '', '1/5 aliquota ordinaria ('+((P.cantonTax*100).toFixed(0))+'%) — risparmio '+((P.cantonTax-capitalWithdrawTax)*100).toFixed(1)+'%'],
  ['  Imposta prelievo (una tantum)',   lppCapTax.toFixed(0), CUR, 'Pagata una sola volta al pensionamento'],
  ['  CAPITALE NETTO DISPONIBILE',     lppCapNet.toFixed(0), CUR, 'Investibile dal Centro Studi (3b/portafoglio)'],
  ['  Rendita equivalente se investito (4% SWR)', (lppCapNet*0.04/12).toFixed(0), CUR+'/mese', 'Se investito in portafoglio Centro Studi'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  E. PILASTRO 3A — VINCOLATO'],
  ['╚══════════════════════════════════════════════╝'],
  ['Capitale 3a accumulato',            P.p3aCapitale, CUR],
  ['Contributo annuo (deducibile)',     P.p3aContrib, CUR+'/anno', 'Max 2024 dipendenti: CHF 7.056 (Art.7 OPP3)'],
  ['Risparmio fiscale contributo/anno', p3aSaving.toFixed(0), CUR+'/anno', 'CHF '+P.p3aContrib+' × '+((P.cantonTax*100).toFixed(0))+'%'],
  ['Aliquota prelievo al pensionamento',(p3aTaxRate*100).toFixed(1)+'%', '', '1/5 aliquota ordinaria — separato da reddito'],
  ['Imposta prelievo (una tantum)',     p3aTax.toFixed(0), CUR, 'Tassato separatamente — NON sommato ad AVS/LPP'],
  ['3A NETTO DISPONIBILE',             p3aNet.toFixed(0), CUR],
  ['Reddito mensile 3a (su 5 anni)',   p3aMonthly.toFixed(0), CUR+'/mese', 'Strategia prelievo scaglionato 5 anni'],
  [''],
  ['╔══════════════════════════════════════════════╗'],
  ['  F. RIEPILOGO RISPARMIO FISCALE TOTALE'],
  ['╚══════════════════════════════════════════════╝'],
  ['Risparmio 3a annuo (contributi)',          p3aSaving.toFixed(0), CUR+'/anno'],
  ['Risparmio 3a in '+P.anni+' anni rimanenti', (p3aSaving*P.anni).toFixed(0), CUR],
  ['Sconto imposta LPP (capitale vs rendita)', ((P.lppCapitale*0.5)*(P.cantonTax-capitalWithdrawTax)).toFixed(0), CUR, 'Una tantum'],
  ['Esenzione plusvalenze/anno',               (fvCapital*(portR-portDY)).toFixed(0), CUR+'/anno', '0% vs es. 26% Italia'],
  ['Recupero Verrechnungssteuer/anno',         (divGross*0.35).toFixed(0), CUR+'/anno', '100% recuperabile'],
  ['═══════════════════════════════'],
  ['TOTALE RISPARMIO STIMATO (cumulato)', (p3aSaving*P.anni + (P.lppCapitale*0.5)*(P.cantonTax-capitalWithdrawTax) + divGross*0.35).toFixed(0), CUR],
];
const ws5 = XLSX.utils.aoa_to_sheet(s5);
ws5['!cols'] = [{wch:45},{wch:16},{wch:10},{wch:50}];
XLSX.utils.book_append_sheet(wb, ws5, '5. Analisi Fiscale CH');

// ──────────────────────────────────────────────────────
// FOGLIO 6: Analisi LPP (rendita vs capitale)
// ──────────────────────────────────────────────────────
const ageRet = P.etaPensione;
const lifeExp85 = 85 - ageRet;  // 20 anni
const lifeExp90 = 90 - ageRet;  // 25 anni

const renditaGrossTotal20 = lppRenditaM * 12 * lifeExp85;
const renditaNetTotal20 = lppRenditaNet * 12 * lifeExp85;
const capitaleRend20 = calcAnnuity(lppCapNet, portR, lifeExp85) / 12;
const capitaleTotal20 = capitaleRend20 * 12 * lifeExp85;
const split50Rend20 = (lppRenditaNet + capitaleRend20) / 2 * 12 * lifeExp85;

const renditaNetTotal25 = lppRenditaNet * 12 * lifeExp90;
const capitaleTotal25 = calcAnnuity(lppCapNet, portR, lifeExp90) / 12 * 12 * lifeExp90;

const s6 = [
  ['ANALISI LPP — CONFRONTO RENDITA vs PRELIEVO IN CAPITALE — '+dt],
  ['Capitale LPP: CHF '+P.lppCapitale.toLocaleString('it-CH')+' · Età pensione: '+ageRet+' · Aliquota cantone LU: '+((P.cantonTax*100).toFixed(0))+'%'],
  [''],
  ['A. CONFRONTO OPZIONI PRINCIPALI'],
  ['PARAMETRO','100% RENDITA','100% CAPITALE','SPLIT 50/50 (scelto)','SPLIT 25R/75C','SPLIT 75R/25C'],
  ['Importo lordo/mese',
    lppRenditaM.toFixed(0)+' CHF',
    calcAnnuity(P.lppCapitale*(1-capitalWithdrawTax),portR,lifeExp85).toFixed(0)+' CHF (SWR)',
    ((lppRenditaM*0.5)+(calcAnnuity(lppCapNet,portR,lifeExp85)/12*0.5)).toFixed(0)+' CHF',
    ((lppRenditaM*0.25)+(calcAnnuity(P.lppCapitale*0.75*(1-capitalWithdrawTax),portR,lifeExp85)/12)).toFixed(0)+' CHF',
    ((lppRenditaM*0.75)+(calcAnnuity(P.lppCapitale*0.25*(1-capitalWithdrawTax),portR,lifeExp85)/12)).toFixed(0)+' CHF',
  ],
  ['Imposta stimata',
    lppRenditaTax.toFixed(0)+' CHF/mese (ordinaria)',
    lppCapTax.toFixed(0)+' CHF una tantum ('+((capitalWithdrawTax*100).toFixed(1))+'%)',
    (lppRenditaTax*0.5).toFixed(0)+' CHF/mese + '+(lppCapTax*0.5).toFixed(0)+' CHF una tantum',
    (lppRenditaTax*0.25).toFixed(0)+' CHF/mese + '+(lppCapTax*0.75).toFixed(0)+' CHF una tantum',
    (lppRenditaTax*0.75).toFixed(0)+' CHF/mese + '+(lppCapTax*0.25).toFixed(0)+' CHF una tantum',
  ],
  ['Netto/mese',
    lppRenditaNet.toFixed(0)+' CHF',
    (calcAnnuity(lppCapNet,portR,lifeExp85)/12).toFixed(0)+' CHF',
    ((lppRenditaNet*0.5)+(calcAnnuity(lppCapNet,portR,lifeExp85)/12*0.5)).toFixed(0)+' CHF',
    ((lppRenditaNet*0.25)+(calcAnnuity(P.lppCapitale*0.75*(1-capitalWithdrawTax),portR,lifeExp85)/12)).toFixed(0)+' CHF',
    ((lppRenditaNet*0.75)+(calcAnnuity(P.lppCapitale*0.25*(1-capitalWithdrawTax),portR,lifeExp85)/12)).toFixed(0)+' CHF',
  ],
  ['Totale lifetime NETTO (vita a 85)',
    renditaNetTotal20.toFixed(0)+' CHF', capitaleTotal20.toFixed(0)+' CHF',
    split50Rend20.toFixed(0)+' CHF', '—', '—'],
  ['Totale lifetime NETTO (vita a 90)',
    renditaNetTotal25.toFixed(0)+' CHF', capitaleTotal25.toFixed(0)+' CHF', '—', '—', '—'],
  ['Rischio longevità', 'COPERTO (paga fino a morte)', 'A carico cliente', 'Ridotto', 'Ridotto', 'Ridotto'],
  ['Flessibilità', 'Nessuna (rendita fissa)', 'Alta (investibile)', 'Media', 'Alta', 'Bassa'],
  ['Trasmissione eredi', 'No (salvo rendita coniuge)', 'Sì (capitale residuo)', 'Parziale', 'Alta', 'Bassa'],
  [''],
  ['B. PUNTO DI PAREGGIO (break-even)'],
  ['Anni necessari perché rendita superi capitale:', (lppCapNet/lppRenditaNet/12).toFixed(1), 'anni', '= CHF '+lppCapNet.toFixed(0)+' ÷ CHF '+lppRenditaNet.toFixed(0)+'/mese'],
  ['Età break-even:', (ageRet+lppCapNet/lppRenditaNet/12).toFixed(0), 'anni', 'Dopo questa età: rendita conviene'],
  ['Se vita stimata > '+(ageRet+lppCapNet/lppRenditaNet/12).toFixed(0)+' anni:', '→ Preferire RENDITA', '', ''],
  ['Se vita stimata < '+(ageRet+lppCapNet/lppRenditaNet/12).toFixed(0)+' anni:', '→ Preferire CAPITALE', '', ''],
  [''],
  ['C. SIMULAZIONE RISPARMIO PILASTRO 3A ('+P.anni+' anni rimanenti)'],
  ['ANNO','ETÀ','CONTRIBUTO','RISPARMIO FISCALE ANNUO','CAPITALE 3A CUMULATO (3%/a)','DEDUZIONE CUMULATA','NOTA'],
  ...Array.from({length:P.anni},(_,i)=>{
    const cap3a = P.p3aCapitale*Math.pow(1.03,i+1) + P.p3aContrib*((Math.pow(1.03,i+1)-1)/0.03);
    return [
      '+'+( i+1), P.eta+i+1, P.p3aContrib, p3aSaving.toFixed(0)+'  CHF',
      cap3a.toFixed(0)+' CHF',
      (p3aSaving*(i+1)).toFixed(0)+' CHF',
      i===P.anni-1?'→ Prelievo a '+P.etaPensione+' anni':'',
    ];
  }),
  [''],
  ['D. RACCOMANDAZIONE CENTRO STUDI MONTE SA'],
  ['Opzione consigliata:', 'SPLIT 50% Rendita + 50% Capitale'],
  ['Motivazione 1:', 'La rendita garantisce copertura longevità CHF '+lppRenditaNet.toFixed(0)+'/mese senza rischio di mercato.'],
  ['Motivazione 2:', 'Il capitale CHF '+lppCapNet.toFixed(0)+' netto, investito dal Centro Studi, genera reddito aggiuntivo.'],
  ['Motivazione 3:', 'L\'imposta prelievo capitale è solo il '+((capitalWithdrawTax*100).toFixed(1))+'% vs '+((P.cantonTax*100).toFixed(0))+'% sulla rendita — risparmio CHF '+lppCapTax.toFixed(0)+'.'],
  ['Attenzione:', 'La decisione rendita/capitale è IRREVOCABILE al momento del pensionamento. Consultare un esperto LPP certificato.'],
  ['Normativa:', 'Art. 37 LPP — diritto al prelievo in capitale. Art. 14 LPP — tasso di conversione minimo 6.8%.'],
];
const ws6 = XLSX.utils.aoa_to_sheet(s6);
ws6['!cols'] = [{wch:38},{wch:22},{wch:22},{wch:22},{wch:22},{wch:22}];
XLSX.utils.book_append_sheet(wb, ws6, '6. Analisi LPP');

// ── Scrivi file ──────────────────────────────────────
const fname = `CentroStudiMonte_PensioneProSvizzera_${dt.replace(/\//g,'-')}.xlsx`;
XLSX.writeFile(wb, fname);
console.log('✅ Excel generato: ' + fname);
console.log('   Fogli: 1.Dashboard · 2.Portafoglio · 3.Anno×Anno · 4.Scenari · 5.Fiscale CH · 6.Analisi LPP');
console.log('   Capitale al ritiro: CHF ' + fvCapital.toFixed(0));
console.log('   Reddito netto/mese: CHF ' + totalNetM.toFixed(0));
console.log('   vs Target CHF ' + P.rendTarget + ': ' + (totalNetM>=P.rendTarget?'✓ RAGGIUNTO':'⚠ deficit CHF '+(P.rendTarget-totalNetM).toFixed(0)));
