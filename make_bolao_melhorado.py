#!/usr/bin/env python3
"""
make_bolao.py  –  Gerador automático do Bolão da COATE
======================================================
Uso:
  python3 make_bolao.py Bolao.xlsx           → gera bolao_coate.html
  python3 make_bolao.py Bolao.xlsx saida.html

A planilha deve ter colunas: DATA | POSIÇÃO | NOME | PONTOS
(ordem não importa, o script detecta pelos cabeçalhos)

Para adicionar rodadas: acrescente novas linhas na planilha com
a nova data e os novos pontos. O gráfico de Evolução mostrará
automaticamente todas as rodadas.
"""
import sys, json, os, re
import openpyxl
from datetime import datetime

# ── leitura da planilha ──────────────────────────────────────────────────────
def read_bolao(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = [row for row in ws.iter_rows(values_only=True) if any(c is not None for c in row)]

    header_idx = None
    for i, row in enumerate(rows):
        vals = [str(v).strip().upper() if v is not None else '' for v in row]
        if any(k in ' '.join(vals) for k in ['NOME','PONTOS','PTS']):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Não encontrei cabeçalho com NOME/PONTOS na planilha")

    header = [str(v).strip().upper() if v is not None else '' for v in rows[header_idx]]
    date_col = next((i for i,c in enumerate(header) if 'DATA' in c), None)
    name_col = next((i for i,c in enumerate(header) if 'NOME' in c), None)
    pts_col  = next((i for i,c in enumerate(header) if 'PONTO' in c or 'PTS' in c), None)

    if name_col is None or pts_col is None:
        raise ValueError("Planilha precisa ter colunas NOME e PONTOS")

    history = {}   # date_str -> {name: pts}
    current_date = None

    for row in rows[header_idx+1:]:
        def safe(idx):
            return row[idx] if idx is not None and idx < len(row) else None

        name = safe(name_col)
        pts  = safe(pts_col)
        dval = safe(date_col)

        if not name:
            continue

        if dval:
            if isinstance(dval, datetime):
                current_date = dval.strftime('%d/%m/%Y')
            else:
                s = str(dval).strip()
                if s:
                    current_date = s

        if current_date is None:
            current_date = datetime.now().strftime('%d/%m/%Y')

        try:
            pts_int = int(float(str(pts))) if pts is not None and str(pts).strip() != '' else 0
        except:
            pts_int = 0

        if current_date not in history:
            history[current_date] = {}
        history[current_date][str(name).strip()] = pts_int

    # Converte para lista ordenada de rodadas
    def date_key(d):
        try:
            return datetime.strptime(d, '%d/%m/%Y')
        except:
            return datetime.min

    rodadas = []
    for dt in sorted(history.keys(), key=date_key):
        players = [{'name': n, 'pts': p} for n, p in history[dt].items()]
        rodadas.append({'date': dt, 'players': players})

    return rodadas

# ── template HTML ─────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>🏇 Bolão da COATE – Copa do Mundo 2026</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Fredoka+One&family=Nunito:wght@400;600;700;800&display=swap');
  :root {
    --green:#2ecc71;--green2:#27ae60;--sky:#c9eaff;--ground:#6b3e10;
    --grass:#3db52a;--gold:#FFD700;--silver:#C0C0C0;--bronze:#CD7F32;
    --bg:#f0faf0;--text:#1a1a2e;
  }
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Nunito',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

  /* HEADER */
  header{background:linear-gradient(135deg,#0d1b2a,#1b3a5c,#0d1b2a);padding:1.8rem 1rem 1.4rem;text-align:center;position:relative;overflow:hidden}
  header::before{content:'';position:absolute;inset:0;background:radial-gradient(ellipse at 50% -20%,rgba(46,204,113,.18) 0%,transparent 65%)}
  .header-top{display:flex;align-items:center;justify-content:center;gap:.8rem;flex-wrap:wrap}
  header h1{font-family:'Fredoka One',cursive;font-size:clamp(1.8rem,6vw,3.2rem);color:#fff;letter-spacing:1px;text-shadow:0 0 24px rgba(46,204,113,.5)}
  header h1 .hi{color:#FFD700} header h1 .lo{color:#2ecc71}
  header .sub{color:rgba(255,255,255,.65);font-size:.95rem;margin-top:.3rem;font-weight:700;letter-spacing:.5px}
  .trophy-badge{font-size:2.8rem;line-height:1;filter:drop-shadow(0 0 10px rgba(255,215,0,.6));animation:float 3s ease-in-out infinite}
  @keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-6px)}}
  .stars-wrap{position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none}
  .star{position:absolute;background:white;border-radius:50%;animation:twinkle 2s infinite alternate}
  @keyframes twinkle{from{opacity:.15}to{opacity:1}}
  .wc-badge{display:inline-flex;align-items:center;gap:.4rem;background:rgba(255,215,0,.15);border:1px solid rgba(255,215,0,.4);border-radius:20px;padding:.2rem .8rem;margin-top:.5rem;font-size:.82rem;color:#FFD700;font-weight:700}

  /* TABS */
  .tabs{display:flex;justify-content:center;gap:.5rem;padding:1.2rem 1rem .4rem;flex-wrap:wrap}
  .tab-btn{font-family:'Fredoka One',cursive;font-size:1rem;padding:.45rem 1.4rem;border:2px solid #ccc;border-radius:50px;cursor:pointer;background:white;color:#666;transition:all .2s}
  .tab-btn.active{background:var(--green);border-color:var(--green);color:white;box-shadow:0 4px 14px rgba(46,204,113,.4);transform:translateY(-2px)}
  .tab-btn:hover:not(.active){border-color:var(--green);color:var(--green)}

  /* SECTIONS */
  .section{display:none}.section.active{display:block}

  /* RACE */
  .race-wrap{padding:1.2rem 1rem;max-width:960px;margin:0 auto}
  .race-title{font-family:'Fredoka One',cursive;font-size:1.35rem;text-align:center;margin-bottom:.2rem}
  .race-date{text-align:center;color:#888;font-size:.82rem;margin-bottom:1.2rem;font-weight:700}

  .track-container{background:linear-gradient(180deg,var(--sky) 0%,#dff1ff 100%);border-radius:18px;padding:1rem .5rem .3rem;border:3px solid #b8d8f0;overflow:visible;position:relative}

  .lane{display:flex;align-items:center;margin-bottom:4px;position:relative;height:60px}
  .lane-ground{position:absolute;bottom:0;left:0;right:0;height:20px;background:linear-gradient(180deg,var(--grass) 0%,var(--ground) 100%);border-radius:0 0 5px 5px}
  .lane-grass-line{position:absolute;bottom:18px;left:0;right:0;height:4px;background:var(--grass);border-radius:2px}

  .pos-badge{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:'Fredoka One',cursive;font-size:.8rem;flex-shrink:0;margin-right:5px;position:relative;z-index:10;box-shadow:0 2px 6px rgba(0,0,0,.25)}
  .pos-1{background:var(--gold);color:#333}
  .pos-2{background:var(--silver);color:#333}
  .pos-3{background:var(--bronze);color:#fff}
  .pos-other{background:#4a4a6a;color:#fff}

  .track-bar-wrap{flex:1;position:relative;height:60px}

  .horse-runner{position:absolute;bottom:20px;left:0;transition:left 1.55s cubic-bezier(.2,1.25,.35,1);z-index:5;display:flex;flex-direction:row;align-items:flex-end;gap:5px;will-change:left}
  .horse-runner::before,.horse-runner::after{content:'';position:absolute;left:5px;bottom:4px;border-radius:50%;background:rgba(181,135,79,.28);filter:blur(.4px);opacity:0;pointer-events:none}
  .horse-runner::before{width:13px;height:7px;animation:dustPuff calc(var(--pace,.48s)*2) ease-out infinite}
  .horse-runner::after{width:8px;height:5px;animation:dustPuff calc(var(--pace,.48s)*2) ease-out infinite calc(var(--pace,.48s)*.75)}
  .horse-stage{position:relative;width:92px;height:66px;flex-shrink:0}
  .horse-svg{width:92px;height:66px;overflow:visible;filter:drop-shadow(2px 3px 4px rgba(0,0,0,.30));animation:horseBob var(--pace,.48s) ease-in-out infinite;transform-origin:50% 85%;will-change:transform}
  .horse-svg .leg{transform-box:fill-box;transform-origin:50% 8%;will-change:transform}
  .horse-svg .leg-a{animation:legA var(--pace,.48s) ease-in-out infinite}
  .horse-svg .leg-b{animation:legB var(--pace,.48s) ease-in-out infinite}
  .horse-svg .tail{transform-box:fill-box;transform-origin:95% 50%;animation:tailWave calc(var(--pace,.48s)*1.35) ease-in-out infinite}
  .horse-svg .jockey{transform-box:fill-box;transform-origin:50% 100%;animation:jockeyLean calc(var(--pace,.48s)*2) ease-in-out infinite}
  .horse-svg .speed-line{animation:speedBlink calc(var(--pace,.48s)*1.6) ease-in-out infinite}
  .runner-leader .horse-svg{filter:drop-shadow(0 0 7px rgba(255,215,0,.55)) drop-shadow(2px 3px 4px rgba(0,0,0,.28))}
  .runner-leader::before,.runner-leader::after{background:rgba(255,215,0,.24)}
  .runner-tired .horse-svg{animation:tiredBob .95s ease-in-out infinite;filter:drop-shadow(2px 3px 3px rgba(0,0,0,.25)) saturate(.82)}
  .runner-tired .horse-svg .leg-a,.runner-tired .horse-svg .leg-b{animation-duration:.9s}
  .runner-last .horse-svg{animation:lastWobble 1.15s ease-in-out infinite}
  .horse-label{background:rgba(15,15,40,.88);color:white;font-size:.65rem;font-weight:800;padding:4px 8px;border-radius:8px;white-space:nowrap;pointer-events:none;border:1px solid rgba(255,255,255,.22);margin-bottom:7px;line-height:1.15;max-width:104px;overflow:hidden;text-overflow:ellipsis;box-shadow:0 2px 5px rgba(0,0,0,.18)}
  .runner-leader .horse-label{background:linear-gradient(135deg,#5b4700,#9a7400);border-color:#ffe36b;color:#fff8cf}
  .runner-tired .horse-label{background:rgba(80,56,70,.88)}

  @keyframes horseBob{0%,100%{transform:translateY(0) rotate(-1deg)}50%{transform:translateY(-3px) rotate(1deg)}}
  @keyframes tiredBob{0%,100%{transform:translateY(1px) rotate(1deg)}50%{transform:translateY(-1px) rotate(-1deg)}}
  @keyframes lastWobble{0%,100%{transform:translateY(1px) rotate(2deg)}50%{transform:translateY(-1px) rotate(-3deg)}}
  @keyframes legA{0%,100%{transform:rotate(13deg)}50%{transform:rotate(-17deg)}}
  @keyframes legB{0%,100%{transform:rotate(-16deg)}50%{transform:rotate(14deg)}}
  @keyframes tailWave{0%,100%{transform:rotate(-6deg)}50%{transform:rotate(10deg)}}
  @keyframes jockeyLean{0%,100%{transform:rotate(-2deg) translateY(0)}50%{transform:rotate(2deg) translateY(1px)}}
  @keyframes speedBlink{0%,100%{opacity:.25;transform:translateX(0)}50%{opacity:1;transform:translateX(-3px)}}
  @keyframes dustPuff{0%{opacity:0;transform:translate(8px,0) scale(.35)}25%{opacity:.65}100%{opacity:0;transform:translate(-18px,-7px) scale(1.55)}}

  .pts-label{width:40px;flex-shrink:0;text-align:right;font-family:'Fredoka One',cursive;font-size:1.1rem;color:var(--text);padding-right:4px;position:relative;z-index:10}

  .finish-pole{position:absolute;right:44px;top:0;bottom:0;width:7px;background:repeating-linear-gradient(180deg,#111 0,#111 7px,#fff 7px,#fff 14px);z-index:4}
  .finish-top{position:absolute;right:36px;top:-4px;font-size:1.4rem;z-index:5;filter:drop-shadow(0 2px 4px rgba(0,0,0,.4))}

  /* RANKING */
  .ranking-wrap{max-width:700px;margin:0 auto;padding:1.4rem 1rem}
  .section-title{font-family:'Fredoka One',cursive;font-size:1.35rem;text-align:center;margin-bottom:1rem}
  .rank-card{display:flex;align-items:center;gap:.8rem;background:white;border-radius:14px;padding:.65rem 1rem;margin-bottom:.55rem;box-shadow:0 2px 8px rgba(0,0,0,.07);transition:transform .15s,box-shadow .15s;animation:slideIn .35s ease both}
  .rank-card:hover{transform:translateX(5px);box-shadow:0 4px 16px rgba(0,0,0,.12)}
  @keyframes slideIn{from{opacity:0;transform:translateX(-18px)}to{opacity:1;transform:translateX(0)}}
  .rank-medal{font-size:1.6rem;width:2rem;text-align:center}
  .rank-num{font-family:'Fredoka One',cursive;font-size:1rem;color:#bbb;width:1.6rem;text-align:center}
  .rank-inner{flex:1;min-width:0}
  .rank-name{font-weight:800;font-size:.95rem}
  .rank-pts{font-family:'Fredoka One',cursive;font-size:1.25rem}
  .rank-pts small{font-size:.6rem;color:#aaa;font-family:'Nunito',sans-serif;font-weight:600}
  .rank-bar-bg{width:100%;height:5px;background:#eee;border-radius:5px;margin-top:.3rem;overflow:hidden}
  .rank-bar-fill{height:100%;border-radius:5px;transition:width 1s ease .3s}

  /* EVOLUTION */
  .evo-wrap{max-width:920px;margin:0 auto;padding:1.4rem 1rem}
  .chart-area{background:white;border-radius:16px;padding:1.4rem;box-shadow:0 4px 20px rgba(0,0,0,.08);overflow-x:auto}
  canvas#evoChart{width:100%!important;max-height:430px}
  .legend-grid{display:flex;flex-wrap:wrap;gap:.4rem .7rem;margin-top:1rem;justify-content:center}
  .legend-item{display:flex;align-items:center;gap:.3rem;font-size:.76rem;font-weight:700}
  .legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}

  /* FOOTER */
  footer{text-align:center;padding:1.2rem;color:#bbb;font-size:.78rem;font-weight:600}

  /* CONFETTI */
  .confetti-piece{position:fixed;width:8px;height:8px;top:-10px;animation:confettiFall linear forwards;pointer-events:none;z-index:9999;border-radius:2px}
  @keyframes confettiFall{0%{transform:translateY(0) rotate(0deg);opacity:1}100%{transform:translateY(110vh) rotate(720deg);opacity:0}}

  @media(max-width:480px){.horse-stage{width:68px;height:52px}.horse-svg{width:68px;height:52px}.lane{height:64px}.track-bar-wrap{height:64px}.horse-label{font-size:.54rem;max-width:70px;padding:3px 5px}.horse-runner{bottom:19px;gap:2px}}
</style>
</head>
<body>

<header>
  <div class="stars-wrap" id="stars"></div>
  <div class="header-top">
    <div class="trophy-badge">🏆</div>
    <div>
      <h1><span class="hi">Bolão</span> da <span class="lo">COATE</span></h1>
      <div class="sub">Copa do Mundo 2026</div>
    </div>
    <div class="trophy-badge" style="animation-delay:.8s">⚽</div>
  </div>
  <div><span class="wc-badge">🌍 FIFA World Cup 2026 &nbsp;·&nbsp; USA · CAN · MEX</span></div>
</header>

<div class="tabs">
  <button class="tab-btn active" onclick="showTab('race',this)">🏇 Corrida</button>
  <button class="tab-btn" onclick="showTab('ranking',this)">🏆 Ranking</button>
  <button class="tab-btn" onclick="showTab('evo',this)">📈 Evolução</button>
</div>

<div class="section active" id="tab-race">
  <div class="race-wrap">
    <div class="race-title">🏇 Hipódromo da COATE</div>
    <div class="race-date" id="raceDate"></div>
    <div class="track-container" id="trackContainer">
      <div id="lanes"></div>
    </div>
  </div>
</div>

<div class="section" id="tab-ranking">
  <div class="ranking-wrap">
    <div class="section-title">🏆 Classificação</div>
    <div id="rankingCards"></div>
  </div>
</div>

<div class="section" id="tab-evo">
  <div class="evo-wrap">
    <div class="section-title">📈 Histórico de Pontos</div>
    <div class="chart-area"><canvas id="evoChart"></canvas></div>
    <div class="legend-grid" id="evoLegend"></div>
    <p style="text-align:center;margin-top:.8rem;color:#aaa;font-size:.76rem;font-weight:600">
      💡 Atualize a planilha e gere novamente com <code>python3 make_bolao.py Bolao.xlsx</code>
    </p>
  </div>
</div>

<footer>🏇 Bolão da COATE · Copa do Mundo 2026 · Que vença o melhor!</footer>

<script>
// ╔══════════════════════════════════════════════════════╗
// ║  DADOS GERADOS AUTOMATICAMENTE — NÃO EDITE AQUI     ║
// ║  Para atualizar: python3 make_bolao.py Bolao.xlsx   ║
// ╚══════════════════════════════════════════════════════╝
const HISTORY = __HISTORY_JSON__;

// ── cores por participante ──────────────────────────────
const PALETTE = [
  "#e74c3c","#3498db","#2ecc71","#f39c12","#9b59b6",
  "#1abc9c","#e67e22","#e91e63","#00bcd4","#8bc34a",
  "#ff5722","#607d8b","#c0392b"
];
const ALL_NAMES = HISTORY[0].players.map(p => p.name);
const colorOf   = {};
ALL_NAMES.forEach((n,i) => colorOf[n] = PALETTE[i % PALETTE.length]);

const latest = HISTORY[HISTORY.length - 1];
const sorted  = [...latest.players].sort((a,b) => b.pts - a.pts);
const maxPts  = sorted[0].pts || 1;
const total   = sorted.length;

// ── horse SVGs ──────────────────────────────────────────
// 3 estados: leader (pomposo), mid (normal), tired (caindo)
function horseSVG(color, rank, total) {
  const leader = rank === 1;
  const podium = rank <= 3;
  const tired  = rank > total * 0.65;
  const last   = rank === total;
  const dk     = shadeColor(color, -42);
  const deep   = shadeColor(color, -72);
  const lt     = shadeColor(color, 48);
  const jc     = shadeColor(color, 68);
  const blanket = leader ? '#f6c90e' : (podium ? '#f4f1de' : shadeColor(color,-12));
  const headTransform = tired ? 'translate(1 7) rotate(12 80 18)' : '';
  const neckPath = tired
    ? 'M66 34 Q70 28 70 22 Q70 17 74 16 Q79 17 79 22 Q77 29 72 37Z'
    : 'M65 33 Q70 23 73 14 Q75 9 80 11 Q83 15 80 21 Q76 28 71 37Z';

  return `<svg viewBox="0 0 104 70" xmlns="http://www.w3.org/2000/svg" class="horse-svg" role="img" aria-label="Cavalo na posição ${rank}">
    <!-- sombra e poeira -->
    <ellipse cx="51" cy="66" rx="31" ry="3.2" fill="rgba(0,0,0,.16)"/>

    <!-- cauda -->
    <g class="tail">
      <path d="M21 34 Q9 26 8 15 Q8 9 3 6" stroke="${deep}" stroke-width="5.2" fill="none" stroke-linecap="round"/>
      <path d="M21 35 Q8 34 5 27" stroke="${jc}" stroke-width="2.4" fill="none" stroke-linecap="round" opacity=".9"/>
    </g>

    <!-- corpo -->
    <ellipse cx="49" cy="39" rx="31" ry="16" fill="${color}"/>
    <ellipse cx="49" cy="43" rx="24" ry="9" fill="${lt}" opacity=".23"/>
    <ellipse cx="22" cy="37" rx="12" ry="14" fill="${color}"/>
    <path d="${neckPath}" fill="${color}"/>

    <!-- pernas traseiras -->
    <g class="leg leg-a">
      <path d="M28 49 Q23 57 22 65" stroke="${dk}" stroke-width="6" stroke-linecap="round" fill="none"/>
      <ellipse cx="21" cy="65" rx="4.7" ry="2.2" fill="#252525"/>
    </g>
    <g class="leg leg-b">
      <path d="M39 51 Q38 59 34 65" stroke="${deep}" stroke-width="5.5" stroke-linecap="round" fill="none"/>
      <ellipse cx="34" cy="65" rx="4.5" ry="2.1" fill="#252525"/>
    </g>

    <!-- pernas dianteiras -->
    <g class="leg leg-b">
      <path d="M67 49 Q72 57 76 63" stroke="${dk}" stroke-width="6" stroke-linecap="round" fill="none"/>
      <ellipse cx="77" cy="64" rx="4.8" ry="2.2" fill="#252525"/>
    </g>
    <g class="leg leg-a">
      <path d="M58 51 Q61 59 59 66" stroke="${deep}" stroke-width="5.5" stroke-linecap="round" fill="none"/>
      <ellipse cx="59" cy="66" rx="4.5" ry="2.1" fill="#252525"/>
    </g>

    <!-- manta com número -->
    <path d="M37 32 Q50 27 63 32 L60 48 Q49 52 38 47Z" fill="${blanket}" stroke="${deep}" stroke-width="1" opacity=".96"/>
    <circle cx="50" cy="39" r="7.2" fill="${leader ? '#fff3a6' : '#ffffff'}" stroke="${deep}" stroke-width="1.2"/>
    <text x="50" y="42.3" font-size="9" font-weight="900" text-anchor="middle" fill="${deep}" font-family="Arial, sans-serif">${rank}</text>

    <!-- cabeça -->
    <g transform="${headTransform}">
      <ellipse cx="85" cy="17" rx="11" ry="8" fill="${color}" transform="rotate(-10 85 17)"/>
      <path d="M92 18 Q101 20 99 25 Q96 28 90 24Z" fill="${dk}"/>
      <ellipse cx="98" cy="22" rx="1.6" ry="1.2" fill="#111" opacity=".58"/>
      <circle cx="83" cy="14" r="2.4" fill="#161616"/>
      <circle cx="84" cy="13" r=".85" fill="#fff"/>
      <path d="M79 10 L81 3 L85 9Z" fill="${color}" stroke="${dk}" stroke-width=".8"/>
      <path d="M82 10 L87 5 L88 11Z" fill="${color}" stroke="${dk}" stroke-width=".8"/>
      <path d="M80 13 Q74 8 71 13 Q70 17 73 22" stroke="${jc}" stroke-width="4" fill="none" stroke-linecap="round"/>
      ${tired ? '<path d="M94 25 Q92 31 96 33 Q100 32 99 27Z" fill="#ef5b78"/><path d="M81 14 Q84 16 87 14" stroke="#333" stroke-width="1.4" fill="none"/>' : ''}
    </g>

    <!-- jóquei -->
    <g class="jockey">
      <path d="M43 28 Q51 22 60 28 L58 35 Q51 38 44 34Z" fill="${jc}" stroke="${deep}" stroke-width=".8"/>
      <circle cx="52" cy="19" r="6.5" fill="${shadeColor(jc,18)}"/>
      <path d="M45 18 Q52 10 59 18" fill="${leader ? '#ffd700' : deep}"/>
      <rect x="44.5" y="18" width="15" height="2.5" rx="1.2" fill="${leader ? '#c69b00' : '#292929'}"/>
      <path d="M59 28 Q67 28 75 23" stroke="${deep}" stroke-width="2" fill="none" stroke-linecap="round"/>
    </g>

    <!-- detalhes de humor/posição -->
    ${leader ? '<text x="52" y="10" font-size="12" text-anchor="middle">👑</text><path class="speed-line" d="M17 22 L4 19 M18 29 L1 29 M18 36 L5 39" stroke="#ffd700" stroke-width="2.5" stroke-linecap="round"/>' : ''}
    ${podium && !leader ? '<text x="52" y="10" font-size="9" text-anchor="middle">⭐</text>' : ''}
    ${tired ? '<text x="91" y="9" font-size="10">💧</text>' : ''}
    ${last ? '<text x="15" y="14" font-size="11">🏳️</text><text x="86" y="7" font-size="9">😮‍💨</text>' : ''}
  </svg>`;
}
function shadeColor(hex, pct) {
  hex = hex.replace('#','');
  if (hex.length === 3) hex = hex.split('').map(c=>c+c).join('');
  let r = parseInt(hex.slice(0,2),16);
  let g = parseInt(hex.slice(2,4),16);
  let b = parseInt(hex.slice(4,6),16);
  r = Math.min(255,Math.max(0,r+pct)); g=Math.min(255,Math.max(0,g+pct)); b=Math.min(255,Math.max(0,b+pct));
  return '#'+[r,g,b].map(x=>x.toString(16).padStart(2,'0')).join('');
}

// ── estrelas header ──────────────────────────────────────
const starsEl = document.getElementById('stars');
for(let i=0;i<45;i++){
  const s = document.createElement('div');
  s.className='star';
  const sz = .5+Math.random()*2;
  s.style.cssText=`left:${Math.random()*100}%;top:${Math.random()*100}%;width:${sz}px;height:${sz}px;animation-delay:${Math.random()*3}s;animation-duration:${1+Math.random()*2.5}s`;
  starsEl.appendChild(s);
}

// ── pista ────────────────────────────────────────────────
document.getElementById('raceDate').textContent = 'Atualizado em: ' + latest.date;

const lanesEl = document.getElementById('lanes');
const tc = document.getElementById('trackContainer');

// nuvens decorativas
['12%,6%','42%,3%','68%,8%'].forEach((pos,i) => {
  const c = document.createElement('div');
  const [l,t] = pos.split(',');
  const w = 70+i*25, h = 22+i*6;
  c.style.cssText=`position:absolute;left:${l};top:${t};width:${w}px;height:${h}px;background:white;border-radius:50px;opacity:.8;pointer-events:none;z-index:0`;
  tc.appendChild(c);
});

// linha de chegada
const fp = document.createElement('div'); fp.className='finish-pole'; tc.appendChild(fp);
const ft = document.createElement('div'); ft.className='finish-top'; ft.textContent='🏁'; tc.appendChild(ft);

sorted.forEach((p,i) => {
  const rank = i+1;
  const color = colorOf[p.name] || '#888';
  const lane = document.createElement('div');
  lane.className='lane';
  lane.innerHTML=`
    <div class="pos-badge pos-${rank<=3?rank:'other'}">${rank}</div>
    <div class="track-bar-wrap" id="wrap-${i}">
      <div class="lane-ground"></div>
      <div class="lane-grass-line"></div>
      <div class="horse-runner ${rank===1?'runner-leader':''} ${rank<=3?'runner-podium':''} ${rank>total*.65?'runner-tired':''} ${rank===total?'runner-last':''}" id="horse-${i}" style="left:0;--pace:${(0.36 + Math.min(rank-1,8)*0.035).toFixed(2)}s;animation-delay:${(i*.04).toFixed(2)}s" title="${p.name}: ${p.pts} pontos">
        <div class="horse-stage">${horseSVG(color,rank,total)}</div>
        <div class="horse-label">${p.name}</div>
      </div>
    </div>
    <div class="pts-label">${p.pts}</div>
  `;
  lanesEl.appendChild(lane);
});

function positionHorses() {
  sorted.forEach((p,i) => {
    const wrap  = document.getElementById('wrap-'+i);
    const horse = document.getElementById('horse-'+i);
    if (!wrap||!horse) return;
    const horseWidth = horse.offsetWidth || 150;
    const avail = Math.max(0, wrap.offsetWidth - horseWidth - 8);
    const x = Math.max(0, (maxPts>0 ? p.pts/maxPts : 0) * avail);
    horse.style.left = x + 'px';
  });
}
setTimeout(positionHorses, 220);

// ── ranking ──────────────────────────────────────────────
const medals=['🥇','🥈','🥉'];
const rankEl=document.getElementById('rankingCards');
sorted.forEach((p,i) => {
  const rank=i+1, color=colorOf[p.name]||'#888';
  const pct = maxPts>0 ? (p.pts/maxPts*100).toFixed(0) : 0;
  const card=document.createElement('div');
  card.className='rank-card';
  card.style.animationDelay=(i*.06)+'s';
  card.innerHTML=`
    <div class="rank-medal">${medals[i]||'🏇'}</div>
    <div class="rank-num">${rank}º</div>
    <div class="rank-inner">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <span class="rank-name" style="color:${color}">${p.name}</span>
        <span class="rank-pts" style="color:${color}">${p.pts} <small>pts</small></span>
      </div>
      <div class="rank-bar-bg">
        <div class="rank-bar-fill" style="width:0%;background:${color}" data-pct="${pct}"></div>
      </div>
    </div>`;
  rankEl.appendChild(card);
});
function animateBars(){document.querySelectorAll('.rank-bar-fill').forEach(b=>{b.style.width=b.dataset.pct+'%'})}

// ── gráfico de evolução ──────────────────────────────────
function buildEvoChart() {
  const labels=HISTORY.map(h=>h.date);
  const datasets=ALL_NAMES.map(name=>({
    label:name,
    data:HISTORY.map(h=>{const f=h.players.find(p=>p.name===name);return f?f.pts:null}),
    borderColor:colorOf[name],backgroundColor:colorOf[name]+'33',
    borderWidth:2.5,pointRadius:5,pointHoverRadius:8,tension:.4,fill:false
  }));
  const ctx=document.getElementById('evoChart').getContext('2d');
  if(window._ec) window._ec.destroy();
  window._ec=new Chart(ctx,{type:'line',data:{labels,datasets},options:{
    responsive:true,interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${c.parsed.y} pts`}}},
    scales:{y:{beginAtZero:true,grid:{color:'#f0f0f0'},ticks:{font:{family:'Nunito',weight:'700'}}},
            x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}}}}
  }});
  const legEl=document.getElementById('evoLegend');
  legEl.innerHTML='';
  ALL_NAMES.forEach(n=>{
    const item=document.createElement('div');
    item.className='legend-item';
    item.innerHTML=`<div class="legend-dot" style="background:${colorOf[n]}"></div>${n}`;
    legEl.appendChild(item);
  });
}

// ── troca de abas ────────────────────────────────────────
function showTab(tab, btn) {
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+tab).classList.add('active');
  btn.classList.add('active');
  if(tab==='ranking') setTimeout(animateBars,100);
  if(tab==='evo')     setTimeout(buildEvoChart,100);
  if(tab==='race') {
    sorted.forEach((p,i)=>{
      const horse=document.getElementById('horse-'+i);
      const wrap=document.getElementById('wrap-'+i);
      if(!horse||!wrap) return;
      horse.style.transition='none'; horse.style.left='0px';
      setTimeout(()=>{
        horse.style.transition='left 1.4s cubic-bezier(.34,1.4,.64,1)';
        const avail=wrap.offsetWidth-60-52;
        horse.style.left=Math.max(0,(maxPts>0?p.pts/maxPts:0)*avail)+'px';
      },120);
    });
  }
}

// ── Chart.js ─────────────────────────────────────────────
const s=document.createElement('script');
s.src='https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js';
document.head.appendChild(s);

// ── confetti ─────────────────────────────────────────────
function spawnConfetti(){
  const cols=['#e74c3c','#3498db','#2ecc71','#f39c12','#9b59b6','#FFD700','#ff5722'];
  for(let i=0;i<70;i++){
    const c=document.createElement('div'); c.className='confetti-piece';
    c.style.left=Math.random()*100+'vw';
    c.style.background=cols[Math.floor(Math.random()*cols.length)];
    c.style.animationDuration=(1.5+Math.random()*2)+'s';
    c.style.animationDelay=(Math.random()*1.2)+'s';
    c.style.transform=`rotate(${Math.random()*360}deg)`;
    c.style.borderRadius=Math.random()>.5?'50%':'2px';
    document.body.appendChild(c);
    setTimeout(()=>c.remove(),4000);
  }
}
window.addEventListener('load',()=>setTimeout(spawnConfetti,400));

</script>
</body>
</html>
"""

def generate(xlsx_path, out_path=None):
    rodadas = read_bolao(xlsx_path)
    history_json = json.dumps(rodadas, ensure_ascii=False, indent=2)
    html = HTML_TEMPLATE.replace('__HISTORY_JSON__', history_json)
    if out_path is None:
        base = os.path.splitext(os.path.basename(xlsx_path))[0]
        out_path = f"bolao_coate.html"
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"✅ Gerado: {out_path} ({len(rodadas)} rodada(s), {len(rodadas[-1]['players'])} participantes)")
    return out_path

if __name__ == '__main__':
    xlsx = sys.argv[1] if len(sys.argv) > 1 else 'Bolao.xlsx'
    out  = sys.argv[2] if len(sys.argv) > 2 else None
    generate(xlsx, out)
