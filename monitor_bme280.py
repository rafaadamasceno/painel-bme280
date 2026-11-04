#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
monitor_bme280.py
=================
Lê as leituras gravadas pelo ESP32 (BME280) na planilha do Google Sheets e gera
gráficos de temperatura, umidade, pressão e sinal Wi-Fi ao longo do tempo.

Uso rápido
----------
    python monitor_bme280.py                      # gera a página interativa (graficos/painel_interativo.html)
    python monitor_bme280.py --loop 10            # atualiza a cada 10 minutos
    python monitor_bme280.py --formato ambos      # página HTML + imagens PNG
    python monitor_bme280.py --formato png --periodos 6h 3d --individuais   # só imagens
    python monitor_bme280.py --html-periodo 24h   # janela inicial da página
    python monitor_bme280.py --csv-url "<link CSV publicado>"   # sem conta de serviço

Veja o README.md para configurar o acesso à planilha.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import matplotlib

matplotlib.use("Agg")  # sem janela: funciona em agendador/servidor
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ============================================================
# CONFIGURAÇÃO (pode ser sobrescrita pelos argumentos da linha de comando)
# ============================================================
PLANILHA = os.environ.get("BME280_PLANILHA", "COLE_AQUI_O_ID_OU_A_URL_DA_PLANILHA")
ABA = "Dados"
CREDENCIAIS = "credenciais.json"       # chave JSON da conta de serviço
PASTA_SAIDA = "graficos"
FUSO = "America/Sao_Paulo"             # fuso da planilha
PERIODOS_PADRAO = ["24h", "7d", "30d", "tudo"]
ALERTA_MINUTOS = 30                    # avisa se a última leitura for mais antiga que isso

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Ordem das colunas na planilha (mesma ordem gravada pelo Code.gs)
NOMES = ["data_hora", "temperatura", "umidade", "pressao", "rssi"]

# faixa = valores fora dela são tratados como leitura inválida e ignorados
VARIAVEIS = {
    "temperatura": {"titulo": "Temperatura", "rotulo": "Temperatura (°C)",
                    "unidade": "°C", "cor": "#d62728", "faixa": (-40, 85)},
    "umidade":     {"titulo": "Umidade relativa", "rotulo": "Umidade (%)",
                    "unidade": "%", "cor": "#1f77b4", "faixa": (0.01, 100)},
    "pressao":     {"titulo": "Pressão atmosférica", "rotulo": "Pressão (hPa)",
                    "unidade": "hPa", "cor": "#2ca02c", "faixa": (300, 1100)},
    "rssi":        {"titulo": "Sinal Wi-Fi", "rotulo": "RSSI (dBm)",
                    "unidade": "dBm", "cor": "#7f7f7f", "faixa": (-120, 0)},
}


# ============================================================
# LEITURA DOS DADOS
# ============================================================
def ler_via_api(planilha: str, aba: str, credenciais: str) -> list[list]:
    """Lê a aba usando a API do Google Sheets (conta de serviço, somente leitura)."""
    import gspread

    if "COLE_AQUI" in planilha:
        raise SystemExit("Configure o ID/URL da planilha (constante PLANILHA ou --planilha).")
    if not Path(credenciais).is_file():
        raise SystemExit(f"Arquivo de credenciais não encontrado: {credenciais}\n"
                         "Veja o README.md (passo da conta de serviço) ou use --csv-url.")

    with open(credenciais, encoding="utf-8") as f:
        email_servico = json.load(f).get("client_email", "(e-mail da conta de serviço)")

    try:
        gc = gspread.service_account(filename=credenciais, scopes=SCOPES)
        sh = gc.open_by_url(planilha) if planilha.startswith("http") else gc.open_by_key(planilha)
        ws = sh.worksheet(aba)
        # UNFORMATTED_VALUE: números e datas vêm "crus", sem depender do idioma/formatação da planilha
        return ws.get_all_values(value_render_option="UNFORMATTED_VALUE")
    except gspread.exceptions.SpreadsheetNotFound:
        raise SystemExit("Planilha não encontrada ou sem acesso. Compartilhe-a (como Leitor) com:\n"
                         f"  {email_servico}")
    except gspread.exceptions.WorksheetNotFound:
        raise SystemExit(f'A planilha não tem uma aba chamada "{aba}".')
    except gspread.exceptions.APIError as e:
        raise SystemExit(f"Erro da API do Google: {e}\n"
                         "Confirme se a Google Sheets API está ativada no projeto da conta de serviço.")


def ler_via_csv(origem: str) -> list[list]:
    """Lê o CSV publicado na web (ou um arquivo CSV local)."""
    return pd.read_csv(origem, header=None, dtype=str, keep_default_na=False).values.tolist()


def converter_datas(s: pd.Series) -> pd.Series:
    """Aceita número serial do Sheets (API) ou texto dd/mm/aaaa HH:MM:SS (CSV)."""
    num = pd.to_numeric(s, errors="coerce")
    datas = pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce").dt.round("s")
    texto = s.astype(str).str.strip().where(datas.isna())
    parsed = pd.to_datetime(texto, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    return datas.fillna(parsed)


def converter_numero(s: pd.Series) -> pd.Series:
    """Aceita número ou texto com vírgula decimal (ex.: '25,3')."""
    num = pd.to_numeric(s, errors="coerce")
    faltando = num.isna() & s.astype(str).str.strip().ne("")
    if faltando.any():
        texto = s.astype(str).str.replace(",", ".", regex=False).where(faltando)
        num = num.fillna(pd.to_numeric(texto, errors="coerce"))
    return num


def valores_para_df(valores: list[list]) -> pd.DataFrame:
    if len(valores) < 2:
        raise SystemExit("A aba não tem dados (apenas cabeçalho ou vazia).")

    df = pd.DataFrame(valores[1:]).iloc[:, :len(NOMES)]   # linha 1 = cabeçalho
    df.columns = NOMES[:df.shape[1]]
    for nome in NOMES:
        if nome not in df:
            df[nome] = np.nan                              # ex.: dados antigos sem RSSI

    df["data_hora"] = converter_datas(df["data_hora"])
    for nome in NOMES[1:]:
        df[nome] = converter_numero(df[nome])

    df = (df.dropna(subset=["data_hora"])
            .sort_values("data_hora")
            .drop_duplicates("data_hora")
            .set_index("data_hora"))

    # Leituras fora da faixa física viram NaN (mesma regra de validação do ESP32)
    for nome, cfg in VARIAVEIS.items():
        lo, hi = cfg["faixa"]
        df.loc[(df[nome] < lo) | (df[nome] > hi), nome] = np.nan

    if df.empty:
        raise SystemExit("Nenhuma linha com data/hora válida foi encontrada.")
    return df


def carregar_dados(args) -> pd.DataFrame:
    if args.csv_url:
        valores = ler_via_csv(args.csv_url)
    else:
        valores = ler_via_api(args.planilha, args.aba, args.credenciais)
    return valores_para_df(valores)


# ============================================================
# PROCESSAMENTO
# ============================================================
def parse_periodo(texto: str):
    """'24h' -> (timedelta, 'Últimas 24 h') | 'tudo' -> (None, 'Todo o período')."""
    t = texto.lower().strip()
    if t in ("tudo", "todo", "all"):
        return None, "Todo o período"
    m = re.fullmatch(r"(\d+)([hd])", t)
    if not m:
        raise SystemExit(f'Período inválido: "{texto}". Use, por exemplo: 6h, 24h, 7d, 30d ou tudo.')
    n, unidade = int(m.group(1)), m.group(2)
    if unidade == "h":
        return timedelta(hours=n), f"Últimas {n} h"
    return timedelta(days=n), f"Últimos {n} dias"


def passo_automatico(duracao: timedelta):
    """Escolhe o intervalo de agregação conforme a duração do gráfico."""
    if duracao <= timedelta(hours=6):
        return None                 # dados brutos
    if duracao <= timedelta(days=2):
        return "5min"
    if duracao <= timedelta(days=10):
        return "30min"
    if duracao <= timedelta(days=60):
        return "2h"
    return "6h"


def quebrar_lacunas(df: pd.DataFrame) -> pd.DataFrame:
    """Insere NaN nos buracos grandes para a linha não 'ligar' períodos sem dados."""
    if len(df) < 2:
        return df
    dif = df.index.to_series().diff()
    mediana = dif.median()
    limite = max(5 * mediana, pd.Timedelta(minutes=10)) if pd.notna(mediana) else pd.Timedelta(minutes=10)
    buracos = df.index[dif > limite]
    if buracos.empty:
        return df
    extras = pd.DataFrame(np.nan, index=buracos - pd.Timedelta(seconds=1), columns=df.columns)
    return pd.concat([df, extras]).sort_index()


def agregar(recorte: pd.DataFrame, passo):
    if passo is None:
        return {"media": quebrar_lacunas(recorte), "min": None, "max": None}
    r = recorte.resample(passo)
    return {"media": r.mean(), "min": r.min(), "max": r.max()}


# ============================================================
# GRÁFICOS
# ============================================================
def formatar_eixo_x(ax):
    loc = mdates.AutoDateLocator(minticks=4, maxticks=9)
    ax.xaxis.set_major_locator(loc)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(loc))


def anotar_extremos(ax, serie: pd.Series, cor: str, unidade: str):
    s = serie.dropna()
    if s.empty:
        return
    i_max, i_min = s.idxmax(), s.idxmin()
    marcados = [(i_max, "máx")] if i_max == i_min else [(i_max, "máx"), (i_min, "mín")]
    for idx, rotulo in marcados:
        valor = s.loc[idx]
        ax.plot(idx, valor, "o", ms=4, color=cor)
        deslocamento = 8 if rotulo == "máx" else -13
        ax.annotate(f"{rotulo} {valor:.1f} {unidade}", (idx, valor),
                    textcoords="offset points", xytext=(0, deslocamento),
                    ha="center", fontsize=8)


def desenhar(ax, dados: dict, nome: str, ini, fim):
    cfg = VARIAVEIS[nome]
    media = dados["media"][nome]
    # faixa mín–máx só onde faz sentido (no RSSI, valores inteiros, vira um bloco sem informação)
    com_faixa = dados["min"] is not None and nome != "rssi"

    if com_faixa:
        ax.fill_between(dados["min"].index, dados["min"][nome], dados["max"][nome],
                        color=cfg["cor"], alpha=0.18, linewidth=0, label="mín–máx")
    ax.plot(media.index, media.values, color=cfg["cor"], lw=1.4,
            label="média" if com_faixa else None)

    anotar_extremos(ax, media, cfg["cor"], cfg["unidade"])
    ax.set_ylabel(cfg["rotulo"])
    ax.set_title(cfg["titulo"], loc="left", fontsize=10, fontweight="bold")
    ax.set_xlim(ini, fim)
    ax.margins(y=0.15)
    ax.grid(alpha=0.3)
    formatar_eixo_x(ax)
    if com_faixa:
        ax.legend(loc="upper right", fontsize=8, frameon=False)


def gerar_painel(dados, ini, fim, titulo, subtitulo, caminho: Path) -> bool:
    nomes = [n for n in VARIAVEIS if dados["media"][n].notna().any()]
    if not nomes:
        return False
    fig, axes = plt.subplots(len(nomes), 1, figsize=(12, 2.9 * len(nomes) + 0.8),
                             sharex=True, squeeze=False)
    for ax, nome in zip(axes[:, 0], nomes):
        desenhar(ax, dados, nome, ini, fim)
    fig.suptitle(f"{titulo}\n{subtitulo}", fontsize=12)
    fig.tight_layout()
    fig.savefig(caminho, dpi=130)
    plt.close(fig)
    return True


def gerar_individual(dados, nome, ini, fim, titulo, caminho: Path) -> bool:
    if not dados["media"][nome].notna().any():
        return False
    fig, ax = plt.subplots(figsize=(12, 4))
    desenhar(ax, dados, nome, ini, fim)
    fig.suptitle(titulo, fontsize=11)
    fig.tight_layout()
    fig.savefig(caminho, dpi=130)
    plt.close(fig)
    return True


def imprimir_resumo(recorte: pd.DataFrame, titulo: str):
    print(f"  {titulo}: {len(recorte)} leituras")
    for nome, cfg in VARIAVEIS.items():
        s = recorte[nome].dropna()
        if s.empty:
            continue
        print(f"    {cfg['titulo']:<20} mín {s.min():8.1f} | máx {s.max():8.1f} | "
              f"média {s.mean():8.1f} | última {s.iloc[-1]:8.1f} {cfg['unidade']}")


# ============================================================
# PÁGINA HTML INTERATIVA (sem dependências: só HTML, CSS e JavaScript)
# ============================================================
TITULO_HTML = "Estação BME280"

# (casas decimais nas estatísticas do gráfico, casas decimais na dica do mouse)
_DECIMAIS = {"temperatura": (1, 2), "umidade": (1, 2), "pressao": (1, 2), "rssi": (0, 0)}

HTML_TEMPLATE = r'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITULO__</title>
<style>
:root{color-scheme:light dark;--bg:#f6f7f9;--fg:#1f2430;--mut:#667085;--grid:#e6e8ec;--panel:#fff;--bord:#d8dce3;--accent:#2563eb;--onacc:#fff;--warn:#92400e;--warnbg:#fef3c7;--warnbd:#f5d98b}
@media (prefers-color-scheme:dark){:root{--bg:#0e1116;--fg:#e7e9ee;--mut:#9aa3b2;--grid:#262b36;--panel:#151922;--bord:#2a303d;--accent:#6ea8fe;--onacc:#0b1220;--warn:#fcd34d;--warnbg:#33290a;--warnbd:#5b4a14}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 system-ui,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:1280px;margin:0 auto;padding:16px}
h1{font-size:20px;margin:0}
.sub{color:var(--mut);font-size:13px;margin-top:2px}
.aviso{display:none;margin:12px 0 0;padding:8px 12px;border-radius:8px;background:var(--warnbg);border:1px solid var(--warnbd);color:var(--warn);font-size:13px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:14px 0}
.card{background:var(--panel);border:1px solid var(--bord);border-top:3px solid var(--c);border-radius:10px;padding:10px 12px}
.card .n{font-size:12px;color:var(--mut)}
.card .v{font-size:24px;font-weight:600;line-height:1.25}
.card .v small{font-size:13px;font-weight:500;color:var(--mut);margin-left:3px}
.bar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:0 0 8px}
.bar .lbl{color:var(--mut);font-size:13px;margin-right:4px}
.bar button{font:inherit;font-size:13px;padding:4px 12px;border-radius:999px;border:1px solid var(--bord);background:var(--panel);color:var(--fg);cursor:pointer}
.bar button:hover{border-color:var(--accent)}
.bar button.on{background:var(--accent);border-color:var(--accent);color:var(--onacc)}
.bar .jan{margin-left:auto;color:var(--mut);font-size:12.5px}
.chart{position:relative;background:var(--panel);border:1px solid var(--bord);border-radius:10px;overflow:hidden}
.chart canvas{position:absolute;left:0;top:0}
#top{cursor:crosshair;touch-action:none}
#tip{position:absolute;display:none;pointer-events:none;z-index:3;background:var(--panel);border:1px solid var(--bord);border-radius:8px;padding:8px 10px;font-size:12.5px;box-shadow:0 4px 14px rgba(0,0,0,.18);white-space:nowrap}
#tip b{display:block;margin-bottom:4px}
#tip i{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}
.dica{color:var(--mut);font-size:12px;margin:8px 2px 0}
</style>
</head>
<body>
<div class="wrap">
  <h1>__TITULO__</h1>
  <div class="sub" id="sub"></div>
  <div class="aviso" id="aviso"></div>
  <div class="cards" id="cards"></div>
  <div class="bar" id="bar"><span class="lbl">Período:</span></div>
  <div class="chart" id="chart"><canvas id="base"></canvas><canvas id="top"></canvas><div id="tip"></div></div>
  <div class="dica">Passe o mouse para ver os valores &middot; arraste para dar zoom &middot; Shift + arrastar (ou setas &larr; &rarr;) para mover &middot; roda do mouse para aproximar/afastar &middot; duplo clique para ver tudo.</div>
</div>
<script>
(function(){
'use strict';
var D=__DADOS__;
var T=D.t,N=T.length,V=D.vars,nV=V.length;
var tMin=T[0],tMax=T[N-1];
var MIN_SPAN=60;
var padFull=Math.max(60,(tMax-tMin)*0.015);
var fullLo=tMin-padFull,fullHi=tMax+padFull;
var FONT='system-ui,"Segoe UI",Roboto,Arial,sans-serif';
var BTNS=[['6 h',21600],['24 h',86400],['7 d',604800],['30 d',2592000],['Tudo',null]];

function $(id){return document.getElementById(id);}
var chart=$('chart'),cBase=$('base'),cTop=$('top'),tip=$('tip'),bar=$('bar');
var gB=cBase.getContext('2d'),gT=cTop.getContext('2d');
var dpr=window.devicePixelRatio||1;
var W=0,H=0,x0=0,x1=0;
var M={l:66,r:18,t:6,b:26},PH=170,TITLE=24,GAPY=18,rowH=0;
var v0=fullLo,v1=fullHi,activeBtn=null;
var yr=[],stats=[];
var hover=null,sel=null,drag=null;

function css(n){return getComputedStyle(document.documentElement).getPropertyValue(n).trim();}
var ESCURO=false;
function tema(){try{ESCURO=window.matchMedia('(prefers-color-scheme: dark)').matches;}catch(e){}}
function clara(hex){   /* no tema escuro, clareia a cor da série para dar contraste */
  if(!ESCURO)return hex;
  var n=parseInt(hex.slice(1),16),r=n>>16,g=(n>>8)&255,b=n&255;
  return 'rgb('+Math.round(r+(255-r)*.3)+','+Math.round(g+(255-g)*.3)+','+Math.round(b+(255-b)*.3)+')';
}
function p2(n){return n<10?'0'+n:String(n);}
function fmtData(t){var d=new Date(t*1000);return p2(d.getUTCDate())+'/'+p2(d.getUTCMonth()+1)+'/'+d.getUTCFullYear()+' '+p2(d.getUTCHours())+':'+p2(d.getUTCMinutes())+':'+p2(d.getUTCSeconds());}
function num(v,dec){return v.toFixed(dec).replace('.',',');}
function lb(x){var lo=0,hi=N;while(lo<hi){var m=(lo+hi)>>1;if(T[m]<x)lo=m+1;else hi=m;}return lo;}
function nearest(t){var i=lb(t);if(i<=0)return 0;if(i>=N)return N-1;return (t-T[i-1]<=T[i]-t)?i-1:i;}
function xs(t){return x0+(t-v0)/(v1-v0)*(x1-x0);}
function tx(px){return v0+(px-x0)/(x1-x0)*(v1-v0);}
function panelY0(k){return M.t+k*rowH+TITLE;}
function ys(k,v){return panelY0(k)+PH-(v-yr[k][0])/(yr[k][1]-yr[k][0])*PH;}

/* ---------- janela visível (zoom / deslocamento) ---------- */
function setView(a,b,keepBtn){
  var full=fullHi-fullLo,span=Math.min(Math.max(b-a,MIN_SPAN),full),c=(a+b)/2;
  a=c-span/2;b=c+span/2;
  if(a<fullLo){b+=fullLo-a;a=fullLo;}
  if(b>fullHi){a-=b-fullHi;b=fullHi;}
  v0=Math.max(a,fullLo);v1=Math.min(b,fullHi);
  if(!keepBtn)activeBtn=null;
  draw();
}
function setBtn(i){
  var dur=BTNS[i][1];
  activeBtn=i;
  if(dur===null)setView(fullLo,fullHi,true);else setView(tMax-dur,tMax+dur*0.02,true);
}

/* ---------- escalas e estatísticas da janela visível ---------- */
function compute(){
  var i0=Math.max(0,lb(v0)-1),i1=Math.min(N-1,lb(v1));
  for(var k=0;k<nV;k++){
    var y=V[k].y,mn=Infinity,mx=-Infinity,s=0,c=0,i,q;
    for(i=i0;i<=i1;i++){if(T[i]<v0||T[i]>v1)continue;q=y[i];if(q===null)continue;if(q<mn)mn=q;if(q>mx)mx=q;s+=q;c++;}
    var vazio=(c===0);
    if(vazio){for(i=i0;i<=i1;i++){q=y[i];if(q===null)continue;if(q<mn)mn=q;if(q>mx)mx=q;s+=q;c++;}}
    if(c===0){mn=0;mx=1;s=0;c=1;}
    var p=(mx-mn)*0.12;if(p===0)p=Math.max(1,Math.abs(mx)*0.01);
    yr[k]=[mn-p,mx+p];
    stats[k]=vazio?null:{mn:mn,mx:mx,avg:s/c};
  }
}
function niceTicks(a,b,n){
  var raw=(b-a)/n,mag=Math.pow(10,Math.floor(Math.log10(raw))),r=raw/mag;
  var st=(r<=1?1:r<=2?2:r<=5?5:10)*mag,t=[];
  for(var v=Math.ceil(a/st-1e-9)*st;v<=b+1e-9;v+=st)t.push(v);
  return {t:t,st:st};
}
var STEPS=[10,30,60,300,600,900,1800,3600,7200,10800,21600,43200,86400,172800,604800,1209600,2592000,5184000,7776000];
function xticks(){
  var span=v1-v0,maxT=Math.max(2,Math.floor((x1-x0)/(W<600?62:100))),st=STEPS[STEPS.length-1],i,t=[];
  for(i=0;i<STEPS.length;i++){if(span/STEPS[i]<=maxT){st=STEPS[i];break;}}
  for(var v=Math.ceil(v0/st)*st;v<=v1;v+=st)t.push(v);
  return {t:t,st:st};
}
function xlabel(t,st){
  var d=new Date(t*1000),h=d.getUTCHours(),m=d.getUTCMinutes(),s=d.getUTCSeconds();
  var dia=p2(d.getUTCDate())+'/'+p2(d.getUTCMonth()+1);
  if(st>=86400||(h===0&&m===0&&s===0))return dia;
  return st<60?p2(h)+':'+p2(m)+':'+p2(s):p2(h)+':'+p2(m);
}

/* ---------- desenho ---------- */
function series(k){
  var y=V[k].y,i0=Math.max(0,lb(v0)-1),i1=Math.min(N-1,lb(v1)),i,px,py,pen=false;
  gB.beginPath();
  if(i1-i0<=(x1-x0)*2){
    for(i=i0;i<=i1;i++){
      if(y[i]===null){pen=false;continue;}
      px=xs(T[i]);py=ys(k,y[i]);
      if(pen&&T[i]-T[i-1]>D.gap)pen=false;
      if(pen)gB.lineTo(px,py);else{gB.moveTo(px,py);pen=true;}
    }
    gB.stroke();
    if(i1-i0<=300){
      gB.fillStyle=clara(V[k].cor);
      for(i=i0;i<=i1;i++){if(y[i]===null)continue;gB.beginPath();gB.arc(xs(T[i]),ys(k,y[i]),2,0,6.2832);gB.fill();}
    }
  }else{
    /* muitos pontos por pixel: faixa mín–máx (clara) + linha da média por coluna */
    var col=-1,mn=0,mx=0,sm=0,cn=0,tl=0,c,pts=[];
    var gapC=Math.max(D.gap,3*(v1-v0)/(x1-x0));   /* um buraco só conta se for maior que ~3 colunas */
    var dens=(i1-i0)/(x1-x0);                      /* leituras por pixel */
    function fecha(){
      if(col<0)return;
      gB.moveTo(col+.5,ys(k,mn));gB.lineTo(col+.5,ys(k,mx));
      if(cn>=Math.max(2,0.4*dens))pts.push([col+.5,sm/cn,tl]);   /* colunas parciais (bordas/buracos) são ruidosas */
    }
    for(i=i0;i<=i1;i++){
      if(y[i]===null)continue;
      c=Math.floor(xs(T[i]));
      if(c!==col){fecha();col=c;mn=mx=sm=y[i];cn=1;}
      else{if(y[i]<mn)mn=y[i];if(y[i]>mx)mx=y[i];sm+=y[i];cn++;}
      tl=T[i];
    }
    fecha();
    gB.globalAlpha=.3;gB.stroke();gB.globalAlpha=1;
    gB.beginPath();
    for(i=0;i<pts.length;i++){
      if(i>0&&pts[i][2]-pts[i-1][2]<=gapC)gB.lineTo(pts[i][0],ys(k,pts[i][1]));
      else gB.moveTo(pts[i][0],ys(k,pts[i][1]));
    }
    gB.stroke();
  }
}
function draw(){
  tema();compute();
  var mut=css('--mut'),grid=css('--grid'),bord=css('--bord');
  gB.setTransform(dpr,0,0,dpr,0,0);gB.clearRect(0,0,W,H);
  var xt=xticks(),k,j,y0,y1,yt,dec,s,txt,larg;
  var yLast=panelY0(nV-1)+PH;
  for(k=0;k<nV;k++){
    y0=panelY0(k);y1=y0+PH;
    gB.textBaseline='alphabetic';gB.textAlign='left';gB.font='600 13px '+FONT;gB.fillStyle=clara(V[k].cor);
    gB.fillText(V[k].rotulo,x0,y0-9);
    larg=gB.measureText(V[k].rotulo).width;
    s=stats[k];
    gB.font='12px '+FONT;gB.fillStyle=mut;gB.textAlign='right';
    if(s){
      txt='mín '+num(s.mn,V[k].ds)+'  ·  méd '+num(s.avg,V[k].ds)+'  ·  máx '+num(s.mx,V[k].ds);
      if(larg+gB.measureText(txt).width+24>x1-x0)txt='mín '+num(s.mn,V[k].ds)+' · máx '+num(s.mx,V[k].ds);
    }else txt='sem dados nesta janela';
    gB.fillText(txt,x1,y0-9);
    yt=niceTicks(yr[k][0],yr[k][1],Math.max(3,Math.round(PH/32)));
    dec=Math.max(0,Math.ceil(-Math.log10(yt.st)-1e-9));
    gB.lineWidth=1;gB.strokeStyle=grid;gB.beginPath();
    for(j=0;j<yt.t.length;j++){var yy=Math.round(ys(k,yt.t[j]))+.5;gB.moveTo(x0,yy);gB.lineTo(x1,yy);}
    for(j=0;j<xt.t.length;j++){var xx=Math.round(xs(xt.t[j]))+.5;gB.moveTo(xx,y0);gB.lineTo(xx,y1);}
    gB.stroke();
    gB.fillStyle=mut;gB.textAlign='right';gB.textBaseline='middle';gB.font='12px '+FONT;
    for(j=0;j<yt.t.length;j++)gB.fillText(num(yt.t[j],dec),x0-8,ys(k,yt.t[j]));
    gB.strokeStyle=bord;gB.strokeRect(x0+.5,y0+.5,x1-x0-1,PH-1);
    gB.save();gB.beginPath();gB.rect(x0,y0,x1-x0,PH);gB.clip();
    gB.strokeStyle=clara(V[k].cor);gB.lineWidth=1.5;gB.lineJoin='round';series(k);
    gB.restore();
  }
  gB.fillStyle=mut;gB.textAlign='center';gB.textBaseline='top';gB.font='12px '+FONT;
  for(j=0;j<xt.t.length;j++){
    var px=xs(xt.t[j]);
    if(px<x0+12||px>x1-12)continue;
    gB.fillText(xlabel(xt.t[j],xt.st),px,yLast+8);
  }
  drawTop();updBtns();
}
function drawTop(){
  gT.setTransform(dpr,0,0,dpr,0,0);gT.clearRect(0,0,W,H);
  var yTop=panelY0(0),yBot=panelY0(nV-1)+PH,k;
  if(sel){
    var a=Math.min(sel.a,sel.b),b=Math.max(sel.a,sel.b);
    gT.globalAlpha=.16;gT.fillStyle=css('--accent');gT.fillRect(a,yTop,b-a,yBot-yTop);gT.globalAlpha=1;
    gT.strokeStyle=css('--accent');gT.lineWidth=1;gT.strokeRect(a+.5,yTop+.5,b-a,yBot-yTop-1);
  }
  if(hover!==null){
    var px=xs(T[hover]);
    if(px>=x0&&px<=x1){
      gT.strokeStyle=css('--mut');gT.lineWidth=1;gT.setLineDash([4,3]);
      gT.beginPath();gT.moveTo(Math.round(px)+.5,yTop);gT.lineTo(Math.round(px)+.5,yBot);gT.stroke();gT.setLineDash([]);
      for(k=0;k<nV;k++){
        var q=V[k].y[hover];if(q===null)continue;
        gT.fillStyle=css('--panel');gT.strokeStyle=clara(V[k].cor);gT.lineWidth=2;
        gT.beginPath();gT.arc(px,ys(k,q),4,0,6.2832);gT.fill();gT.stroke();
      }
    }
  }
}
function showTip(cx,cy){
  if(hover===null){tip.style.display='none';return;}
  var h='<b>'+fmtData(T[hover])+'</b>',k,q;
  for(k=0;k<nV;k++){
    q=V[k].y[hover];
    h+='<div><i style="background:'+V[k].cor+'"></i>'+V[k].titulo+': <strong>'+(q===null?'—':num(q,V[k].dt)+' '+V[k].unidade)+'</strong></div>';
  }
  tip.innerHTML=h;tip.style.display='block';
  var r=chart.getBoundingClientRect(),w=tip.offsetWidth,hh=tip.offsetHeight;
  var lx=cx-r.left+16,ly=cy-r.top+12;
  if(lx+w>r.width-4)lx=cx-r.left-w-16;
  if(lx<4)lx=4;
  if(ly+hh>r.height-4)ly=r.height-hh-4;
  tip.style.left=lx+'px';tip.style.top=ly+'px';
}

/* ---------- interação ---------- */
function localP(e){var r=cTop.getBoundingClientRect();return {x:e.clientX-r.left,y:e.clientY-r.top};}
function inPlot(p){return p.x>=x0&&p.x<=x1&&p.y>=M.t&&p.y<=panelY0(nV-1)+PH;}
function setHover(p,e){hover=inPlot(p)?nearest(tx(p.x)):null;drawTop();showTip(e.clientX,e.clientY);}
cTop.addEventListener('pointerdown',function(e){
  if(e.pointerType==='mouse'&&e.button!==0)return;
  var p=localP(e);if(!inPlot(p))return;
  cTop.setPointerCapture(e.pointerId);
  drag={mode:(e.shiftKey||e.pointerType==='touch')?'pan':'zoom',px:p.x,v0:v0,v1:v1};
  if(drag.mode==='pan')cTop.style.cursor='grabbing';else sel={a:p.x,b:p.x};
  setHover(p,e);
});
cTop.addEventListener('pointermove',function(e){
  var p=localP(e);
  if(drag){
    if(drag.mode==='pan'){var dt=(p.x-drag.px)/(x1-x0)*(drag.v1-drag.v0);setView(drag.v0-dt,drag.v1-dt);}
    else sel.b=Math.max(x0,Math.min(x1,p.x));
  }
  setHover(p,e);
});
function endDrag(){
  if(!drag)return;
  var d=drag,s=sel;drag=null;sel=null;cTop.style.cursor='crosshair';
  if(d.mode==='zoom'&&s&&Math.abs(s.b-s.a)>6)setView(tx(Math.min(s.a,s.b)),tx(Math.max(s.a,s.b)));
  else drawTop();
}
cTop.addEventListener('pointerup',endDrag);
cTop.addEventListener('pointercancel',endDrag);
cTop.addEventListener('pointerleave',function(){if(!drag){hover=null;tip.style.display='none';drawTop();}});
cTop.addEventListener('wheel',function(e){
  var p=localP(e);
  if(p.x<x0||p.x>x1)return;
  e.preventDefault();
  if(e.shiftKey||Math.abs(e.deltaX)>Math.abs(e.deltaY)){
    var dx=(e.shiftKey?(e.deltaX||e.deltaY):e.deltaX)/(x1-x0)*(v1-v0)*0.6;
    setView(v0+dx,v1+dx);
  }else{
    var ta=tx(p.x),f=Math.exp(e.deltaY*0.0016);
    setView(ta-(ta-v0)*f,ta+(v1-ta)*f);
  }
  setHover(p,e);
},{passive:false});
cTop.addEventListener('dblclick',function(){setBtn(BTNS.length-1);});
document.addEventListener('keydown',function(e){
  var s=v1-v0,m=(v0+v1)/2;
  if(e.key==='ArrowLeft')setView(v0-s*.2,v1-s*.2);
  else if(e.key==='ArrowRight')setView(v0+s*.2,v1+s*.2);
  else if(e.key==='+'||e.key==='=')setView(m-s*.35,m+s*.35);
  else if(e.key==='-'||e.key==='_')setView(m-s*.7,m+s*.7);
  else if(e.key==='Home')setBtn(BTNS.length-1);
  else return;
  e.preventDefault();
});

/* ---------- botões, cartões e avisos ---------- */
BTNS.forEach(function(b,i){
  var el=document.createElement('button');el.textContent=b[0];el.onclick=function(){setBtn(i);};bar.appendChild(el);
});
var jan=document.createElement('span');jan.className='jan';bar.appendChild(jan);
function updBtns(){
  var bs=bar.getElementsByTagName('button'),i;
  for(i=0;i<bs.length;i++)bs[i].className=(i===activeBtn)?'on':'';
  jan.textContent=fmtData(Math.max(v0,tMin))+'  →  '+fmtData(Math.min(v1,tMax));
}
V.forEach(function(v){
  var last=null,i;
  for(i=N-1;i>=0;i--){if(v.y[i]!==null){last=v.y[i];break;}}
  var d=document.createElement('div');d.className='card';d.style.setProperty('--c',v.cor);
  d.innerHTML='<div class="n">'+v.titulo+'</div><div class="v">'+(last===null?'—':num(last,v.dt))+'<small>'+v.unidade+'</small></div>';
  $('cards').appendChild(d);
});
$('sub').textContent='Última leitura: '+fmtData(tMax)+'  ·  '+N+' pontos  ·  de '+fmtData(tMin).slice(0,10)+' a '+fmtData(tMax).slice(0,10)+(D.passo?'  ·  médias de '+Math.round(D.passo/60)+' min':'');
function aviso(){
  var agora=D.ger.naive+(Date.now()/1000-D.ger.real),atraso=agora-tMax,el=$('aviso');
  if(atraso>D.alerta*60){
    var h=atraso/3600;
    el.textContent='⚠ Sem dados novos há '+(h<1?Math.round(atraso/60)+' min':num(h,1)+' h')+' — o ESP32 pode estar offline.';
    el.style.display='block';
  }else el.style.display='none';
}
aviso();setInterval(aviso,30000);

/* ---------- tamanho, tema e início ---------- */
function resize(){
  W=chart.clientWidth;
  var small=W<600;M.l=small?52:66;PH=small?140:170;rowH=TITLE+PH+GAPY;
  H=M.t+(nV-1)*rowH+TITLE+PH+M.b;
  x0=M.l;x1=W-M.r;
  chart.style.height=H+'px';
  [cBase,cTop].forEach(function(c){c.width=Math.round(W*dpr);c.height=Math.round(H*dpr);c.style.width=W+'px';c.style.height=H+'px';});
  draw();
}
if(window.ResizeObserver)new ResizeObserver(function(){if(chart.clientWidth!==W)resize();}).observe(chart);
else window.addEventListener('resize',resize);
try{window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change',draw);}catch(e){}

var saved=null;
try{saved=JSON.parse(sessionStorage.getItem('bme_vista')||'null');sessionStorage.removeItem('bme_vista');}catch(e){}
resize();
if(saved){
  if(saved.b!==null&&saved.b!==undefined)setBtn(saved.b);else setView(saved.a,saved.c);
}else{
  var ib=-1,i;
  for(i=0;i<BTNS.length;i++){if(BTNS[i][1]===D.ini)ib=i;}
  if(ib>=0)setBtn(ib);else setView(tMax-D.ini,tMax+D.ini*0.02);
}

/* ---------- recarrega sozinha (só se ninguém estiver mexendo) ---------- */
if(D.refresh>0){
  var t0=Date.now(),ult=Date.now();
  ['pointermove','pointerdown','wheel','keydown','touchstart'].forEach(function(n){
    document.addEventListener(n,function(){ult=Date.now();},{passive:true});
  });
  setInterval(function(){
    if(Date.now()-t0>=D.refresh*60000&&Date.now()-ult>120000){
      try{sessionStorage.setItem('bme_vista',JSON.stringify({b:activeBtn,a:v0,c:v1}));}catch(e){}
      location.reload();
    }
  },20000);
}
window.painelBME={vista:function(){return [v0,v1,activeBtn];}};  /* apenas para depuração */
})();
</script>
</body>
</html>
'''


def reduzir_pontos(d: pd.DataFrame, maximo: int = 150_000):
    """Se houver pontos demais, usa médias por intervalo para a página continuar fluida."""
    if len(d) <= maximo:
        return d, None
    segundos = max(int(np.ceil((d.index.max() - d.index.min()).total_seconds() / maximo)), 60)
    return d.resample(f"{segundos}s").mean(), segundos


def gerar_html(df: pd.DataFrame, caminho: Path, agora: datetime,
               periodo_inicial: str, refresh_min: int, alerta_min: int) -> bool:
    """Gera uma página HTML única (sem internet e sem bibliotecas) com zoom e leitura ao passar o mouse."""
    nomes = [n for n in VARIAVEIS if df[n].notna().any()]
    if not nomes:
        return False
    duracao, _ = parse_periodo(periodo_inicial)
    d, passo = reduzir_pontos(df[nomes])

    # O horário da planilha é "local sem fuso": guardamos como segundos corridos e o
    # JavaScript escreve exatamente o mesmo horário, sem conversão de fuso do navegador.
    t = ((d.index - pd.Timestamp(0)) // pd.Timedelta(seconds=1)).astype("int64").tolist()
    dif = np.diff(t)
    gap = int(max(5 * np.median(dif), 600)) if len(dif) else 600   # buracos maiores quebram a linha

    variaveis = []
    for n in nomes:
        cfg = VARIAVEIS[n]
        estat, dica = _DECIMAIS[n]
        casas = 0 if n == "rssi" else 2
        y = [None if pd.isna(v) else (int(v) if casas == 0 else float(v)) for v in d[n].round(casas)]
        variaveis.append({"id": n, "titulo": cfg["titulo"], "rotulo": cfg["rotulo"],
                          "unidade": cfg["unidade"], "cor": cfg["cor"],
                          "ds": estat, "dt": dica, "y": y})

    dados = {
        "t": t, "gap": gap, "vars": variaveis, "passo": passo,
        "ini": int(duracao.total_seconds()) if duracao else None,
        "refresh": max(int(refresh_min), 0), "alerta": int(alerta_min),
        "ger": {"naive": int((agora - datetime(1970, 1, 1)).total_seconds()), "real": int(time.time())},
    }
    texto = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = HTML_TEMPLATE.replace("__TITULO__", TITULO_HTML).replace("__DADOS__", texto)

    tmp = caminho.with_name(caminho.name + ".tmp")     # grava e troca, para o navegador nunca ler pela metade
    tmp.write_text(html, encoding="utf-8")
    tmp.replace(caminho)
    return True


# ============================================================
# EXECUÇÃO
# ============================================================
def executar(args):
    saida = Path(args.saida)
    saida.mkdir(parents=True, exist_ok=True)

    df = carregar_dados(args)
    agora = datetime.now(ZoneInfo(args.fuso)).replace(tzinfo=None)
    ultima = df.index.max()
    atraso = agora - ultima

    print(f"[{agora:%d/%m/%Y %H:%M:%S}] {len(df)} leituras carregadas "
          f"({df.index.min():%d/%m/%Y %H:%M} → {ultima:%d/%m/%Y %H:%M}).")
    if atraso > timedelta(minutes=args.alerta_minutos):
        horas = atraso.total_seconds() / 3600
        print(f"  ⚠ Última leitura há {horas:.1f} h. O ESP32 pode estar offline.")

    atualizado = f"Atualizado em {agora:%d/%m/%Y %H:%M}  |  última leitura: {ultima:%d/%m/%Y %H:%M}"

    if args.formato in ("html", "ambos"):
        arq_html = saida / "painel_interativo.html"
        imprimir_resumo(df, "Todo o período")
        if gerar_html(df, arq_html, agora, args.html_periodo, args.html_refresh, args.alerta_minutos):
            print(f"    → {arq_html}")

    for p in (args.periodos if args.formato in ("png", "ambos") else []):
        duracao, rotulo = parse_periodo(p)
        ini = agora - duracao if duracao else df.index.min()
        fim = agora
        recorte = df.loc[(df.index >= ini) & (df.index <= fim)]
        if recorte.empty:
            print(f"  {rotulo}: sem dados no período, ignorado.")
            continue

        if args.passo == "auto":
            passo = passo_automatico(fim - ini)
        elif args.passo.lower() in ("nenhum", "none"):
            passo = None
        else:
            passo = args.passo

        dados = agregar(recorte, passo)
        detalhe = "dados brutos" if passo is None else f"média a cada {passo}"
        titulo = f"Estação BME280 — {rotulo}"

        imprimir_resumo(recorte, rotulo)
        arq = saida / f"painel_{p}.png"
        if gerar_painel(dados, ini, fim, titulo, f"{atualizado}  |  {detalhe}", arq):
            print(f"    → {arq}")
        if args.individuais:
            for nome in VARIAVEIS:
                arq_i = saida / f"{p}_{nome}.png"
                if gerar_individual(dados, nome, ini, fim, f"{titulo} ({detalhe})", arq_i):
                    print(f"    → {arq_i}")


def main():
    ap = argparse.ArgumentParser(description="Gráficos das leituras do BME280 gravadas no Google Sheets.")
    ap.add_argument("--planilha", default=PLANILHA, help="ID ou URL da planilha")
    ap.add_argument("--aba", default=ABA, help=f'nome da aba (padrão: "{ABA}")')
    ap.add_argument("--credenciais", default=CREDENCIAIS, help="chave JSON da conta de serviço")
    ap.add_argument("--csv-url", default=None,
                    help="link CSV da planilha publicada na web (alternativa à conta de serviço)")
    ap.add_argument("--saida", default=PASTA_SAIDA, help="pasta onde salvar as imagens")
    ap.add_argument("--periodos", nargs="+", default=PERIODOS_PADRAO,
                    help="ex.: 6h 24h 7d 30d tudo")
    ap.add_argument("--passo", default="auto",
                    help='agregação: "auto", "nenhum" ou um intervalo como 10min, 1h')
    ap.add_argument("--formato", choices=["html", "png", "ambos"], default="html",
                    help="html = página interativa (padrão) | png = imagens | ambos")
    ap.add_argument("--html-periodo", default="7d",
                    help="período mostrado ao abrir a página: 6h, 24h, 7d, 30d ou tudo (padrão: 7d)")
    ap.add_argument("--html-refresh", type=int, default=10, metavar="MINUTOS",
                    help="a página aberta recarrega sozinha a cada N min, se ninguém estiver mexendo "
                         "(0 = nunca; padrão: 10)")
    ap.add_argument("--individuais", action="store_true", help="(PNG) gera também um gráfico por variável")
    ap.add_argument("--loop", type=int, default=0, metavar="MINUTOS",
                    help="repete a cada N minutos (0 = roda uma vez)")
    ap.add_argument("--alerta-minutos", type=int, default=ALERTA_MINUTOS,
                    help="avisa se a última leitura for mais antiga que isso")
    ap.add_argument("--fuso", default=FUSO, help="fuso horário da planilha")
    args = ap.parse_args()

    if args.loop <= 0:
        executar(args)
        return

    print(f"Modo automático: atualizando a cada {args.loop} min. Ctrl+C para parar.")
    while True:
        try:
            executar(args)
        except Exception as e:  # noqa: BLE001 - em modo contínuo, um erro não deve derrubar o processo
            print(f"  ✖ Falha nesta rodada: {e}")
        time.sleep(args.loop * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(0)
