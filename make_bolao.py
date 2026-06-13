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

  .track-container{background:linear-gradient(180deg,#cce8fb 0%,#dff1ff 46%,#d9f0ff 100%);border-radius:22px;padding:1.25rem .8rem .95rem;border:3px solid #b8d8f0;overflow:visible;position:relative;box-shadow:inset 0 1px 0 rgba(255,255,255,.55),0 6px 18px rgba(124,160,185,.12)}
  .track-container::before{content:"";position:absolute;inset:14px 18px auto 18px;height:32px;border-radius:18px;background:linear-gradient(180deg,rgba(255,255,255,.32),rgba(255,255,255,0));pointer-events:none}
  #lanes{padding:.25rem 0 .35rem}

  .lane{display:flex;align-items:center;margin-bottom:11px;position:relative;height:84px}
  .lane:first-child{margin-top:8px}
  .lane:last-child{margin-bottom:10px}
  .lane-ground{position:absolute;bottom:0;left:0;right:0;height:22px;background:linear-gradient(180deg,#99672d 0%,#7a4f1b 52%,#603b14 100%);border-radius:0 0 8px 8px;box-shadow:0 3px 0 rgba(255,255,255,.1) inset}
  .lane-grass-line{position:absolute;bottom:18px;left:0;right:0;height:6px;background:linear-gradient(180deg,#59e248 0%,#26c51f 100%);border-radius:3px;box-shadow:0 1px 0 rgba(255,255,255,.35) inset}

  .pos-badge{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-family:'Fredoka One',cursive;font-size:.8rem;flex-shrink:0;margin-right:5px;position:relative;z-index:10;box-shadow:0 2px 6px rgba(0,0,0,.25)}
  .pos-1{background:var(--gold);color:#333}
  .pos-2{background:var(--silver);color:#333}
  .pos-3{background:var(--bronze);color:#fff}
  .pos-other{background:#4a4a6a;color:#fff}

  .track-bar-wrap{flex:1;position:relative;height:84px;padding-right:88px}
  .track-bar-wrap::before{content:"";position:absolute;left:0;right:0;top:15px;height:1px;background:rgba(255,255,255,.34)}
  .track-bar-wrap::after{content:"";position:absolute;left:0;right:0;bottom:25px;height:1px;background:rgba(102,136,156,.08)}

  .horse-runner{position:absolute;bottom:20px;left:0;transition:left 1.55s cubic-bezier(.2,1.25,.35,1);z-index:5;display:flex;flex-direction:row;align-items:flex-end;gap:7px;will-change:left}
  .horse-runner::before,.horse-runner::after{content:'';position:absolute;left:5px;bottom:4px;border-radius:50%;background:rgba(181,135,79,.28);filter:blur(.4px);opacity:0;pointer-events:none}
  .horse-runner::before{width:13px;height:7px;animation:dustPuff calc(var(--pace,.48s)*2) ease-out infinite}
  .horse-runner::after{width:8px;height:5px;animation:dustPuff calc(var(--pace,.48s)*2) ease-out infinite calc(var(--pace,.48s)*.75)}
  .horse-stage{position:relative;width:128px;height:86px;flex-shrink:0}
  .horse-svg{width:128px;height:86px;overflow:visible;filter:drop-shadow(2px 3px 4px rgba(0,0,0,.30));animation:horseBob var(--pace,.48s) ease-in-out infinite;transform-origin:50% 85%;will-change:transform}
  .horse-svg .leg{transform-box:fill-box;transform-origin:50% 8%;will-change:transform}
  .horse-svg .leg-a{animation:legA var(--pace,.48s) ease-in-out infinite}
  .horse-svg .leg-b{animation:legB var(--pace,.48s) ease-in-out infinite}
  .horse-svg .tail{transform-box:fill-box;transform-origin:95% 50%;animation:tailWave calc(var(--pace,.48s)*1.35) ease-in-out infinite}
  .horse-svg .jockey{transform-box:fill-box;transform-origin:50% 100%;animation:jockeyLean calc(var(--pace,.48s)*2) ease-in-out infinite}
  .horse-svg .speed-line{animation:speedBlink calc(var(--pace,.48s)*1.6) ease-in-out infinite}
  .runner-leader .horse-svg{filter:drop-shadow(0 0 12px rgba(255,215,0,.55)) drop-shadow(2px 3px 4px rgba(0,0,0,.28))}
  .runner-leader::before,.runner-leader::after{background:rgba(255,215,0,.24)}
  .runner-tired .horse-svg{animation:tiredBob .95s ease-in-out infinite;filter:drop-shadow(2px 3px 3px rgba(0,0,0,.25)) saturate(.82)}
  .runner-tired .horse-svg .leg-a,.runner-tired .horse-svg .leg-b{animation-duration:.9s}
  .runner-last .horse-svg{animation:lastWobble 1.15s ease-in-out infinite}
  .horse-label{background:rgba(15,15,40,.92);color:white;font-size:.69rem;font-weight:800;padding:6px 10px;border-radius:10px;white-space:nowrap;pointer-events:none;border:1px solid rgba(255,255,255,.22);margin-bottom:11px;line-height:1.15;max-width:134px;overflow:hidden;text-overflow:ellipsis;box-shadow:0 2px 5px rgba(0,0,0,.18)}
  .runner-leader .horse-label{background:linear-gradient(135deg,#5b4700,#9a7400);border-color:#ffe36b;color:#fff8cf}
  .runner-tired .horse-label{background:rgba(80,56,70,.88)}
  .runner-podium .horse-label{box-shadow:0 0 0 2px rgba(255,255,255,.15),0 2px 8px rgba(0,0,0,.22)}
  .runner-self .horse-svg{filter:drop-shadow(0 0 10px rgba(52,152,219,.36)) drop-shadow(2px 3px 4px rgba(0,0,0,.28))}
  .runner-self .horse-label{background:linear-gradient(135deg,#0f3b61,#2563eb);border-color:#a5c8ff;color:#f5fbff;box-shadow:0 0 0 2px rgba(165,200,255,.22),0 4px 12px rgba(37,99,235,.22)}
  .runner-self .horse-svg{filter:drop-shadow(0 0 10px rgba(52,152,219,.36)) drop-shadow(2px 3px 4px rgba(0,0,0,.28))}

  @keyframes horseBob{0%,100%{transform:translateY(0) rotate(-1deg)}50%{transform:translateY(-3px) rotate(1deg)}}
  @keyframes tiredBob{0%,100%{transform:translateY(1px) rotate(1deg)}50%{transform:translateY(-1px) rotate(-1deg)}}
  @keyframes lastWobble{0%,100%{transform:translateY(1px) rotate(2deg)}50%{transform:translateY(-1px) rotate(-3deg)}}
  @keyframes legA{0%,100%{transform:rotate(13deg)}50%{transform:rotate(-17deg)}}
  @keyframes legB{0%,100%{transform:rotate(-16deg)}50%{transform:rotate(14deg)}}
  @keyframes tailWave{0%,100%{transform:rotate(-6deg)}50%{transform:rotate(10deg)}}
  @keyframes jockeyLean{0%,100%{transform:rotate(-2deg) translateY(0)}50%{transform:rotate(2deg) translateY(1px)}}
  @keyframes speedBlink{0%,100%{opacity:.25;transform:translateX(0)}50%{opacity:1;transform:translateX(-3px)}}
  @keyframes dustPuff{0%{opacity:0;transform:translate(8px,0) scale(.35)}25%{opacity:.65}100%{opacity:0;transform:translate(-18px,-7px) scale(1.55)}}

  .pts-label{width:50px;flex-shrink:0;text-align:right;font-family:'Fredoka One',cursive;font-size:1.18rem;color:var(--text);padding-right:10px;position:relative;z-index:10}

  .finish-pole{position:absolute;right:58px;top:10px;bottom:10px;width:8px;background:repeating-linear-gradient(180deg,#111 0,#111 7px,#fff 7px,#fff 14px);z-index:4;border-radius:3px;box-shadow:0 0 0 2px rgba(255,255,255,.3)}
  .finish-top{position:absolute;right:45px;top:1px;font-size:1.38rem;z-index:5;filter:drop-shadow(0 2px 4px rgba(0,0,0,.4))}

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
  .chart-area{background:white;border-radius:16px;padding:1.2rem 1.8rem 1.05rem 1rem;box-shadow:0 4px 20px rgba(0,0,0,.08);overflow-x:auto}
  canvas#evoChart,canvas#posChart{width:100%!important;max-height:450px}
  .legend-grid{display:flex;flex-wrap:wrap;gap:.4rem .7rem;margin-top:1rem;justify-content:center}
  .legend-item{display:flex;align-items:center;gap:.3rem;font-size:.76rem;font-weight:700}
  .legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
  .evo-switch{display:flex;justify-content:center;gap:.55rem;flex-wrap:wrap;margin:.2rem 0 1rem}
  .evo-btn{font-family:'Fredoka One',cursive;font-size:.95rem;padding:.45rem 1.1rem;border:2px solid #cfd8d4;border-radius:999px;cursor:pointer;background:white;color:#64706d;transition:all .2s}
  .evo-btn.active{background:var(--green);border-color:var(--green);color:#fff;box-shadow:0 4px 14px rgba(46,204,113,.25)}
  .evo-btn:hover:not(.active){border-color:var(--green);color:var(--green)}
  .evo-panel{display:none}
  .evo-panel.active{display:block}
  .chart-note{text-align:center;margin:-.2rem 0 .8rem;color:#888;font-size:.82rem;font-weight:700}



  .top-shell{max-width:1100px;margin:0 auto;padding:0 1rem}
  .hero-strip{max-width:1100px;margin:.35rem auto 0;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.7rem;padding:0 1rem}
  .hero-card{background:rgba(255,255,255,.92);border:1px solid rgba(184,216,240,.95);border-radius:16px;padding:.85rem 1rem;box-shadow:0 6px 18px rgba(20,50,80,.07)}
  .hero-label{font-size:.72rem;font-weight:800;color:#6e7a86;text-transform:uppercase;letter-spacing:.7px}
  .hero-value{font-family:'Fredoka One',cursive;font-size:1.05rem;color:#1b2f49;margin-top:.18rem}
  .hero-sub{font-size:.78rem;color:#7c8793;font-weight:700;margin-top:.1rem}
  .tabs-wrap{max-width:1100px;margin:0 auto;padding:1.1rem 1rem .3rem}
  .tabs{background:rgba(255,255,255,.65);border:1px solid rgba(184,216,240,.9);border-radius:999px;padding:.35rem;display:inline-flex;justify-content:center;gap:.35rem;flex-wrap:wrap;box-shadow:0 6px 18px rgba(20,50,80,.06)}
  .tabs-wrap{text-align:center}
  .tab-btn{min-width:142px}

  .ranking-wrap{max-width:980px;margin:0 auto;padding:1.4rem 1rem}
  .summary-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.8rem;margin-bottom:1rem}
  .summary-card{background:white;border:1px solid #e5edf3;border-radius:16px;padding:.95rem 1rem;box-shadow:0 4px 14px rgba(0,0,0,.06)}
  .summary-card .label{font-size:.72rem;font-weight:800;color:#768392;text-transform:uppercase;letter-spacing:.8px}
  .summary-card .value{font-family:'Fredoka One',cursive;font-size:1.08rem;color:#17314d;margin-top:.18rem}
  .summary-card .sub{font-size:.78rem;color:#7a8691;font-weight:700;margin-top:.1rem}
  .podium-row{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.8rem;margin-bottom:1rem}
  .podium-card{border-radius:18px;padding:1rem;background:linear-gradient(135deg,#ffffff,#f5f9ff);border:1px solid #e5edf3;box-shadow:0 6px 18px rgba(0,0,0,.06);position:relative;overflow:hidden}
  .podium-card::after{content:'';position:absolute;right:-20px;top:-20px;width:90px;height:90px;background:radial-gradient(circle,rgba(255,255,255,.55),rgba(255,255,255,0) 70%)}
  .podium-1{background:linear-gradient(135deg,#fff7cf,#fffdf0);border-color:#f1d76f}
  .podium-2{background:linear-gradient(135deg,#f1f4f8,#ffffff);border-color:#d9e0e8}
  .podium-3{background:linear-gradient(135deg,#fff0e4,#ffffff);border-color:#efc29d}
  .podium-top{display:flex;align-items:center;justify-content:space-between;gap:.6rem}
  .podium-medal{font-size:1.5rem}
  .podium-place{font-family:'Fredoka One',cursive;font-size:.95rem;color:#6a7885}
  .podium-name{font-weight:900;font-size:1rem;margin-top:.55rem;color:#132a43}
  .podium-meta{display:flex;justify-content:space-between;gap:.6rem;margin-top:.45rem;font-size:.82rem;color:#65717d;font-weight:800}
  .podium-pts{font-family:'Fredoka One',cursive;font-size:1.18rem;color:#132a43;margin-top:.25rem}
  .rank-card{display:flex;align-items:center;gap:.85rem;background:white;border-radius:16px;padding:.78rem 1rem;margin-bottom:.62rem;box-shadow:0 2px 8px rgba(0,0,0,.07);transition:transform .15s,box-shadow .15s;animation:slideIn .35s ease both;border:1px solid #edf2f7}
  .rank-card.rank-leader{border-color:#f2d77c;box-shadow:0 6px 18px rgba(242,215,124,.28)}
  .rank-card.rank-self{border-color:#a6c8ff;box-shadow:0 6px 18px rgba(80,130,255,.16)}
  .rank-card:hover{transform:translateX(5px);box-shadow:0 6px 18px rgba(0,0,0,.12)}
  .rank-head{display:flex;justify-content:space-between;align-items:flex-start;gap:.8rem}
  .rank-name-line{display:flex;align-items:center;gap:.45rem;flex-wrap:wrap}
  .rank-name{font-weight:900;font-size:1rem}
  .rank-meta{display:flex;gap:.45rem;flex-wrap:wrap;margin-top:.22rem}
  .mini-pill{font-size:.68rem;font-weight:800;color:#61707f;background:#f2f6fa;border:1px solid #e1e8ef;padding:.16rem .45rem;border-radius:999px}
  .trend{display:inline-flex;align-items:center;gap:.2rem;padding:.18rem .48rem;border-radius:999px;font-size:.72rem;font-weight:900;border:1px solid transparent}
  .trend-up{background:#ebfff3;color:#1f9d56;border-color:#bfe9d0}
  .trend-down{background:#fff0f0;color:#d94b58;border-color:#f3c3c7}
  .trend-same{background:#f3f6f9;color:#677483;border-color:#dbe4ec}
  .trend-new{background:#eef4ff;color:#3e63dd;border-color:#c9d7ff}
  .evo-wrap{max-width:980px;margin:0 auto;padding:1.4rem 1rem}
  .chart-area{position:relative}
  .chart-topline{display:flex;justify-content:space-between;gap:.8rem;align-items:center;flex-wrap:wrap;margin-bottom:.65rem}
  .chart-topline .caption{font-size:.84rem;color:#75808c;font-weight:800}
  .chart-note{margin:.1rem 0 .8rem}
  @media(max-width:900px){.hero-strip,.summary-grid,.podium-row{grid-template-columns:repeat(2,minmax(0,1fr))}}
  @media(max-width:560px){.hero-strip,.summary-grid,.podium-row{grid-template-columns:1fr}.tab-btn{min-width:unset;width:100%}.tabs{display:flex}.hero-card,.summary-card,.podium-card{padding:.8rem .9rem}}


  /* FOOTER */
  footer{text-align:center;padding:1.2rem;color:#bbb;font-size:.78rem;font-weight:600}

  /* CONFETTI */
  .confetti-piece{position:fixed;width:8px;height:8px;top:-10px;animation:confettiFall linear forwards;pointer-events:none;z-index:9999;border-radius:2px}
  @keyframes confettiFall{0%{transform:translateY(0) rotate(0deg);opacity:1}100%{transform:translateY(110vh) rotate(720deg);opacity:0}}

  @media(max-width:480px){.horse-stage{width:80px;height:58px}.horse-svg{width:80px;height:58px}.lane{height:68px}.track-bar-wrap{height:68px}.horse-label{font-size:.56rem;max-width:78px;padding:3px 6px}.horse-runner{bottom:19px;gap:3px}.chart-area{padding:1rem .8rem .8rem .8rem}}
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

<div class="hero-strip" id="heroStats"></div>

<div class="tabs-wrap">
<div class="tabs">
  <button class="tab-btn active" onclick="showTab('race',this)">🏇 Corrida</button>
  <button class="tab-btn" onclick="showTab('ranking',this)">🏆 Ranking</button>
  <button class="tab-btn" onclick="showTab('evo',this)">📈 Evolução</button>
</div>
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
    <div class="summary-grid" id="rankingSummary"></div>
    <div class="podium-row" id="rankingPodium"></div>
    <div id="rankingCards"></div>
  </div>
</div>

<div class="section" id="tab-evo">
  <div class="evo-wrap">
    <div class="section-title">📈 Evolução</div>
    <div class="evo-switch">
      <button class="evo-btn active" onclick="showEvoPanel('points',this)">🏁 Pontos</button>
      <button class="evo-btn" onclick="showEvoPanel('positions',this)">📊 Posições</button>
    </div>

    <div class="evo-panel active" id="evo-panel-points">
      <div class="chart-area">
        <div class="chart-topline"><div class="caption">Histórico de pontuação por rodada</div></div>
        <canvas id="evoChart"></canvas>
      </div>
      <div class="legend-grid" id="evoLegend"></div>
    </div>

    <div class="evo-panel" id="evo-panel-positions">
      <p class="chart-note">1º lugar aparece no topo do gráfico. Os nomes ao lado direito mostram a posição final.</p>
      <div class="chart-area">
        <div class="chart-topline"><div class="caption">Variação das posições ao longo das rodadas</div></div>
        <canvas id="posChart"></canvas>
      </div>
      <div class="legend-grid" id="posLegend"></div>
    </div>

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

const prevRound = HISTORY.length > 1 ? HISTORY[HISTORY.length - 2] : null;
function orderedRound(round){ return [...round.players].sort((a,b)=>b.pts-a.pts); }
function positionMap(round){
  if(!round) return {};
  const map = {};
  orderedRound(round).forEach((p,i)=>{ map[p.name] = i + 1; });
  return map;
}
const prevPosMap = positionMap(prevRound);
const prevPtsMap = prevRound ? Object.fromEntries(prevRound.players.map(p=>[p.name,p.pts])) : {};
function trendInfo(name, currPos){
  if(!prevRound || !(name in prevPosMap)) return {cls:'trend-new', icon:'🆕', text:'estreou'};
  const delta = prevPosMap[name] - currPos;
  if(delta > 0) return {cls:'trend-up', icon:'▲', text:`subiu ${delta}`};
  if(delta < 0) return {cls:'trend-down', icon:'▼', text:`caiu ${Math.abs(delta)}`};
  return {cls:'trend-same', icon:'•', text:'manteve'};
}
function heroCards(){
  const leader = sorted[0];
  const rodadaAtual = HISTORY.length;
  const totalPontos = sorted.reduce((acc,p)=>acc+p.pts,0);
  const top3 = sorted.slice(0,3).map(p=>shortName(p.name)).join(' · ');
  return [
    {label:'Líder', value: leader.name, sub: `${leader.pts} pontos`},
    {label:'Rodada atual', value: `${rodadaAtual}ª rodada`, sub: latest.date},
    {label:'Participantes', value: `${total}`, sub: 'no bolão'},
    {label:'Pódio', value: top3, sub: `${totalPontos} pontos somados`}
  ];
}
function renderHeroCards(){
  const el = document.getElementById('heroStats');
  if(!el) return;
  el.innerHTML = heroCards().map(card=>`<div class="hero-card"><div class="hero-label">${card.label}</div><div class="hero-value">${card.value}</div><div class="hero-sub">${card.sub}</div></div>`).join('');
}
function renderRankingSummary(){
  const leader = sorted[0];
  const vice = sorted[1] || sorted[0];
  const maiorPontuacao = Math.max(...HISTORY.flatMap(r=>r.players.map(p=>p.pts)));
  const summary = [
    {label:'Líder atual', value: leader.name, sub: `${leader.pts} pontos`},
    {label:'Diferença p/ 2º', value: `${leader.pts - vice.pts} ponto${leader.pts - vice.pts === 1 ? '' : 's'}`, sub: vice ? `sobre ${vice.name}` : 'disputa inicial'},
    {label:'Maior pontuação', value: `${maiorPontuacao} pts`, sub: 'entre todas as rodadas'},
    {label:'Última atualização', value: latest.date, sub: `${HISTORY.length} rodada${HISTORY.length > 1 ? 's' : ''}`}
  ];
  const el = document.getElementById('rankingSummary');
  if(el) el.innerHTML = summary.map(s=>`<div class="summary-card"><div class="label">${s.label}</div><div class="value">${s.value}</div><div class="sub">${s.sub}</div></div>`).join('');
}
function renderRankingPodium(){
  const medals = ['🥇','🥈','🥉'];
  const el = document.getElementById('rankingPodium');
  if(!el) return;
  el.innerHTML = sorted.slice(0,3).map((p,i)=>{
    const trend = trendInfo(p.name, i+1);
    const prevPts = prevPtsMap[p.name];
    const deltaPts = prevPts === undefined ? '' : (p.pts - prevPts > 0 ? `+${p.pts - prevPts} na rodada` : (p.pts - prevPts < 0 ? `${p.pts - prevPts} na rodada` : 'mesma pontuação'));
    return `<div class="podium-card podium-${i+1}"><div class="podium-top"><div class="podium-medal">${medals[i]}</div><div class="podium-place">${i+1}º lugar</div><div class="trend ${trend.cls}">${trend.icon} ${trend.text}</div></div><div class="podium-name">${p.name}</div><div class="podium-pts">${p.pts} pts</div><div class="podium-meta"><span>${deltaPts || 'sem comparativo'}</span><span style="color:${colorOf[p.name]}">●</span></div></div>`;
  }).join('');
}


// ── horse SVGs ──────────────────────────────────────────
// 3 estados: leader (pomposo), mid (normal), tired (caindo)
function shortName(name) {
  const map = {
    'Frutuoso Junior':'Frutuoso Jr.',
    'Mauricio Hexa':'Mauricio H.',
    'Mauro Bastos':'Mauro Bastos',
    'Wilton Bessa':'Wilton Bessa'
  };
  if (map[name]) return map[name];
  return name.length > 14 ? name.slice(0,13) + '…' : name;
}

function horseSVG(color, rank, total) {
  const leader = rank === 1;
  const podium = rank <= 3;
  const tired  = rank > total * 0.65;
  const last   = rank === total;
  const deep   = shadeColor(color, -54);
  const dark   = shadeColor(color, -34);
  const light  = shadeColor(color, 30);
  const soft   = shadeColor(color, 56);
  const mane   = shadeColor(color, -66);
  const blanket= leader ? '#ffd54a' : (podium ? '#f8f3ea' : '#fff6ec');
  const trim   = leader ? '#9a7400' : (podium ? '#64438f' : '#8f5d2f');
  const jersey = leader ? '#ffe48a' : soft;
  const headT  = tired ? 'rotate(8 104 22)' : 'rotate(-4 104 22)';

  return `<svg viewBox="0 0 136 88" xmlns="http://www.w3.org/2000/svg" class="horse-svg" role="img" aria-label="Cavalo na posição ${rank}">
    <defs>
      <linearGradient id="bodyShine-${rank}" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stop-color="${light}" stop-opacity=".82"/>
        <stop offset="1" stop-color="${color}" stop-opacity="0"/>
      </linearGradient>
    </defs>

    <ellipse cx="63" cy="82" rx="42" ry="4.3" fill="rgba(0,0,0,.14)"/>

    <g class="tail">
      <path d="M30 43 Q14 31 12 18 Q12 8 6 4" stroke="${mane}" stroke-width="7" fill="none" stroke-linecap="round"/>
      <path d="M31 45 Q14 44 8 34" stroke="${soft}" stroke-width="2.3" fill="none" stroke-linecap="round" opacity=".9"/>
    </g>

    <ellipse cx="61" cy="44" rx="36" ry="18.5" fill="${color}"/>
    <ellipse cx="42" cy="44" rx="15.6" ry="16.5" fill="${color}"/>
    <ellipse cx="64" cy="43" rx="25" ry="12" fill="url(#bodyShine-${rank})" opacity=".72"/>
    <path d="M67 33 Q77 25 88 17 Q95 13 101 17 Q104 23 99 30 Q92 38 79 44Z" fill="${color}"/>
    <path d="M67 34 Q77 25 88 18 Q95 14 100 17 Q92 26 81 43Z" fill="#fff" opacity=".08"/>
    <path d="M50 35 Q58 31 67 33" stroke="#fff" stroke-width="2.2" opacity=".12" fill="none"/>

    <g transform="${headT}">
      <ellipse cx="105" cy="23" rx="14.5" ry="9.4" fill="${color}"/>
      <path d="M114 25 Q126 27 122 34 Q116 37 108 32Z" fill="${dark}"/>
      <circle cx="101.5" cy="20.7" r="2.8" fill="#151515"/>
      <circle cx="102.3" cy="20" r=".85" fill="#fff"/>
      <ellipse cx="121.5" cy="29.5" rx="1.9" ry="1.2" fill="#111" opacity=".6"/>
      <path d="M97 18 Q90 12 87 18 Q88 25 93 29" stroke="${mane}" stroke-width="5" fill="none" stroke-linecap="round"/>
      <path d="M97 17 L99 8 L104 16Z" fill="${color}" stroke="${deep}" stroke-width="1"/>
      <path d="M102 16 L109 10 L110 18Z" fill="${color}" stroke="${deep}" stroke-width="1"/>
      <path d="M96 20 Q106 24 116 22" stroke="#2b2b2b" stroke-width="1.15" fill="none" opacity=".52"/>
      <path d="M95 17 Q103 16 112 18" stroke="${deep}" stroke-width="1" fill="none" opacity=".35"/>
      ${tired ? '<path d="M116 34 Q114 39 118 40.5 Q123 39 121 34Z" fill="#ef6d86"/>' : ''}
    </g>

    <path d="M44 35 Q57 29 74 35 L72 54 Q58 60 43 53Z" fill="${blanket}" stroke="${trim}" stroke-width="1.2"/>
    <path d="M51 37 Q59 33 69 36 L67 50 Q58 54 49 49Z" fill="${trim}" opacity=".22"/>
    <circle cx="58" cy="44" r="8.5" fill="#fffdf8" stroke="${trim}" stroke-width="1.2"/>
    <text x="58" y="47.1" font-size="9.8" font-weight="900" text-anchor="middle" fill="${trim}" font-family="Arial, sans-serif">${rank}</text>

    <g class="leg leg-a">
      <path d="M43 56 Q39 66 35 79" stroke="${deep}" stroke-width="6.5" fill="none" stroke-linecap="round"/>
      <path d="M34 79 Q40 79 43 78" stroke="#222" stroke-width="3.9" fill="none" stroke-linecap="round"/>
    </g>
    <g class="leg leg-b">
      <path d="M56 56 Q54 66 50 79" stroke="${dark}" stroke-width="5.9" fill="none" stroke-linecap="round"/>
      <path d="M49 79 Q55 79 57 78" stroke="#222" stroke-width="3.6" fill="none" stroke-linecap="round"/>
    </g>
    <g class="leg leg-b">
      <path d="M77 56 Q83 65 87 77" stroke="${deep}" stroke-width="6.2" fill="none" stroke-linecap="round"/>
      <path d="M87 77 Q92 77 95 76" stroke="#222" stroke-width="3.6" fill="none" stroke-linecap="round"/>
    </g>
    <g class="leg leg-a">
      <path d="M66 56 Q70 67 68 80" stroke="${dark}" stroke-width="5.9" fill="none" stroke-linecap="round"/>
      <path d="M68 80 Q73 80 76 79" stroke="#222" stroke-width="3.4" fill="none" stroke-linecap="round"/>
    </g>

    <g class="jockey">
      <path d="M47 28.8 Q59 20 72 30 L67 41 Q56 45 47 37Z" fill="${jersey}" stroke="${trim}" stroke-width="1"/>
      <circle cx="61" cy="18" r="6.9" fill="#f1c8a8"/>
      <path d="M53 16.8 Q61 9 69 16" fill="${leader ? '#ffd84c' : trim}"/>
      <rect x="53" y="17" width="16" height="2.9" rx="1.4" fill="${leader ? '#c7a000' : '#1e2430'}"/>
      <path d="M70 29 Q81 29 94 23" stroke="#2b2b2b" stroke-width="2.1" fill="none" stroke-linecap="round"/>
      <path d="M48 30 Q44 33 40 35" stroke="${trim}" stroke-width="1.3" fill="none" opacity=".38"/>
    </g>

    ${leader ? '<text x="60" y="9" font-size="13" text-anchor="middle">👑</text><path class="speed-line" d="M22 24 L6 21 M22 32 L3 32 M21 40 L8 43" stroke="#ffd54a" stroke-width="2.6" stroke-linecap="round"/>' : ''}
    ${podium && !leader ? '<text x="61" y="10" font-size="10">⭐</text>' : ''}
    ${last ? '<text x="15" y="16" font-size="11">🏳️</text><text x="112" y="10" font-size="9">😮‍💨</text>' : ''}
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

// linha de chegada
const fp = document.createElement('div'); fp.className='finish-pole'; tc.appendChild(fp);
const ft = document.createElement('div'); ft.className='finish-top'; ft.textContent='🏁'; tc.appendChild(ft);

sorted.forEach((p,i) => {
  const rank = i+1;
  const color = colorOf[p.name] || '#888';
  const isSelf = /ernaldo/i.test(p.name);
  const medal = rank===1 ? '🥇 ' : (rank===2 ? '🥈 ' : (rank===3 ? '🥉 ' : ''));
  const labelName = shortName(p.name);
  const lane = document.createElement('div');
  lane.className='lane';
  lane.innerHTML=`
    <div class="pos-badge pos-${rank<=3?rank:'other'}">${rank}</div>
    <div class="track-bar-wrap" id="wrap-${i}">
      <div class="lane-ground"></div>
      <div class="lane-grass-line"></div>
      <div class="horse-runner ${rank===1?'runner-leader':''} ${rank<=3?'runner-podium':''} ${rank>total*.65?'runner-tired':''} ${rank===total?'runner-last':''} ${isSelf?'runner-self':''}" id="horse-${i}" style="left:0;--pace:${(0.36 + Math.min(rank-1,8)*0.035).toFixed(2)}s;animation-delay:${(i*.04).toFixed(2)}s" title="${p.name}: ${p.pts} pontos">
        <div class="horse-stage">${horseSVG(color,rank,total)}</div>
        <div class="horse-label" title="${p.name}">${medal}${labelName}</div>
      </div>
    </div>
    <div class="pts-label">${p.pts}</div>
  `;
  lanesEl.appendChild(lane);
});

function horseLeftPx(p,i){
  const wrap=document.getElementById('wrap-'+i);
  const horse=document.getElementById('horse-'+i);
  if(!wrap||!horse) return 0;
  const horseWidth = horse.offsetWidth || 180;
  const reserve = window.innerWidth <= 480 ? 24 : 36;
  const avail = Math.max(0, wrap.offsetWidth - horseWidth - reserve);
  return Math.max(0,(maxPts>0 ? p.pts/maxPts : 0) * avail);
}
function positionHorses() {
  sorted.forEach((p,i) => {
    const horse = document.getElementById('horse-'+i);
    if (!horse) return;
    horse.style.left = horseLeftPx(p,i) + 'px';
  });
}
setTimeout(positionHorses, 220);
window.addEventListener('resize',()=>{clearTimeout(window.__reTrack);window.__reTrack=setTimeout(positionHorses,120)});

// ── ranking ──────────────────────────────────────────────
const medals=['🥇','🥈','🥉'];
renderHeroCards();
renderRankingSummary();
renderRankingPodium();
const rankEl=document.getElementById('rankingCards');
sorted.forEach((p,i) => {
  const rank=i+1, color=colorOf[p.name]||'#888';
  const pct = maxPts>0 ? (p.pts/maxPts*100).toFixed(0) : 0;
  const trend = trendInfo(p.name, rank);
  const prevPts = prevPtsMap[p.name];
  const ptsDelta = prevPts === undefined ? 'sem base anterior' : (p.pts - prevPts > 0 ? `+${p.pts - prevPts} ponto${p.pts - prevPts === 1 ? '' : 's'} na rodada` : (p.pts - prevPts < 0 ? `${p.pts - prevPts} pontos na rodada` : 'mesma pontuação'));
  const gap = sorted[0].pts - p.pts;
  const card=document.createElement('div');
  card.className='rank-card' + (rank===1 ? ' rank-leader' : '') + (/ernaldo/i.test(p.name) ? ' rank-self' : '');
  card.style.animationDelay=(i*.06)+'s';
  card.innerHTML=`
    <div class="rank-medal">${medals[i]||'🏇'}</div>
    <div class="rank-num">${rank}º</div>
    <div class="rank-inner">
      <div class="rank-head">
        <div>
          <div class="rank-name-line">
            <span class="rank-name" style="color:${color}">${p.name}</span>
            <span class="trend ${trend.cls}">${trend.icon} ${trend.text}</span>
          </div>
          <div class="rank-meta">
            <span class="mini-pill">${ptsDelta}</span>
            <span class="mini-pill">${gap===0 ? 'liderança' : `-${gap} do líder`}</span>
          </div>
        </div>
        <span class="rank-pts" style="color:${color}">${p.pts} <small>pts</small></span>
      </div>
      <div class="rank-bar-bg">
        <div class="rank-bar-fill" style="width:0%;background:${color}" data-pct="${pct}"></div>
      </div>
    </div>`;
  rankEl.appendChild(card);
});
function animateBars(){document.querySelectorAll('.rank-bar-fill').forEach(b=>{b.style.width=b.dataset.pct+'%'})}

// ── gráficos de evolução ─────────────────────────────────
const lineEndLabelsPlugin = {
  id:'lineEndLabelsPlugin',
  afterDatasetsDraw(chart, args, options) {
    if(!options || !options.enabled) return;
    const {ctx, chartArea} = chart;
    const items = [];
    chart.data.datasets.forEach((ds, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if(meta.hidden) return;
      let idx = ds.data.length - 1;
      while(idx >= 0 && (ds.data[idx] === null || ds.data[idx] === undefined)) idx--;
      if(idx < 0 || !meta.data[idx]) return;
      const point = meta.data[idx];
      items.push({
        x: point.x,
        y: point.y,
        color: ds.borderColor,
        text: options.formatter ? options.formatter(ds.label, ds.data[idx]) : ds.label
      });
    });
    items.sort((a,b)=>a.y-b.y);
    const minGap = options.minGap || 14;
    for(let i=1;i<items.length;i++){
      if(items[i].y < items[i-1].y + minGap) items[i].y = items[i-1].y + minGap;
    }
    if(items.length){
      const bottomLimit = chartArea.bottom - 4;
      if(items[items.length-1].y > bottomLimit){
        const shift = items[items.length-1].y - bottomLimit;
        items.forEach(it=>it.y -= shift);
      }
      const topLimit = chartArea.top + 4;
      if(items[0].y < topLimit){
        const shift = topLimit - items[0].y;
        items.forEach(it=>it.y += shift);
      }
    }
    ctx.save();
    ctx.font = '700 12px Nunito, sans-serif';
    ctx.textBaseline = 'middle';
    ctx.lineWidth = 1.2;
    items.forEach(item => {
      const labelX = chartArea.right + 14;
      ctx.strokeStyle = item.color;
      ctx.beginPath();
      ctx.moveTo(item.x + 4, item.y);
      ctx.lineTo(labelX - 6, item.y);
      ctx.stroke();
      ctx.fillStyle = '#ffffff';
      ctx.strokeStyle = 'rgba(255,255,255,.85)';
      ctx.lineWidth = 4;
      ctx.strokeText(item.text, labelX, item.y);
      ctx.fillStyle = item.color;
      ctx.fillText(item.text, labelX, item.y);
      ctx.lineWidth = 1.2;
    });
    ctx.restore();
  }
};

function buildEvoChart() {
  const labels=HISTORY.map(h=>h.date);
  const datasets=ALL_NAMES.map(name=>({
    label:name,
    data:HISTORY.map(h=>{const f=h.players.find(p=>p.name===name);return f?f.pts:null}),
    borderColor:colorOf[name],backgroundColor:colorOf[name]+'33',
    borderWidth:2.8,pointRadius:5,pointHoverRadius:8,tension:.4,fill:false
  }));
  const ctx=document.getElementById('evoChart').getContext('2d');
  if(window._ec) window._ec.destroy();
  window._ec=new Chart(ctx,{type:'line',data:{labels,datasets},plugins:[lineEndLabelsPlugin],options:{
    responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},
    layout:{padding:{right:110,left:8,top:4,bottom:0}},
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${c.parsed.y} pts`}},lineEndLabelsPlugin:{enabled:false}},
    scales:{y:{beginAtZero:true,grid:{color:'#f0f0f0'},ticks:{font:{family:'Nunito',weight:'700'}}},
            x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}},offset:true}}
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

function computePositions(round) {
  const ordered=[...round.players].sort((a,b)=> (b.pts-a.pts) || a.name.localeCompare(b.name));
  const posMap={};
  let prevPts=null;
  let prevRank=0;
  ordered.forEach((player,index)=>{
    const rank = player.pts===prevPts ? prevRank : index+1;
    posMap[player.name]=rank;
    prevPts=player.pts;
    prevRank=rank;
  });
  return posMap;
}

function buildPositionChart() {
  const labels=HISTORY.map(h=>h.date);
  const roundPositions=HISTORY.map(computePositions);
  const datasets=ALL_NAMES.map(name=>({
    label:name,
    data:roundPositions.map(pos=> pos[name] ?? null),
    borderColor:colorOf[name],backgroundColor:colorOf[name]+'33',
    borderWidth:2.8,pointRadius:5,pointHoverRadius:8,tension:.35,fill:false
  }));
  const ctx=document.getElementById('posChart').getContext('2d');
  if(window._pc) window._pc.destroy();
  window._pc=new Chart(ctx,{type:'line',data:{labels,datasets},plugins:[lineEndLabelsPlugin],options:{
    responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},
    layout:{padding:{right:140,left:8,top:4,bottom:0}},
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${c.parsed.y}º lugar`}},lineEndLabelsPlugin:{enabled:true, minGap:16, formatter:(label, value)=>`${label} (${value}º)`}},
    scales:{y:{reverse:true,min:1,max:ALL_NAMES.length,grid:{color:'#f0f0f0'},ticks:{stepSize:1,font:{family:'Nunito',weight:'700'},callback:v=> `${v}º`}},
            x:{grid:{display:false},ticks:{font:{family:'Nunito',weight:'700'}},offset:true}}
  }});
  const legEl=document.getElementById('posLegend');
  legEl.innerHTML='';
  ALL_NAMES.forEach(n=>{
    const item=document.createElement('div');
    item.className='legend-item';
    item.innerHTML=`<div class="legend-dot" style="background:${colorOf[n]}"></div>${n}`;
    legEl.appendChild(item);
  });
}

function showEvoPanel(panel, btn) {
  document.querySelectorAll('.evo-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.evo-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('evo-panel-'+panel).classList.add('active');
  btn.classList.add('active');
  if(panel==='points') setTimeout(buildEvoChart,80);
  if(panel==='positions') setTimeout(buildPositionChart,80);
}

// ── troca de abas ────────────────────────────────────────
function showTab(tab, btn) {
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+tab).classList.add('active');
  btn.classList.add('active');
  if(tab==='ranking') setTimeout(animateBars,100);
  if(tab==='evo')     setTimeout(()=>{buildEvoChart(); buildPositionChart();},100);
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
