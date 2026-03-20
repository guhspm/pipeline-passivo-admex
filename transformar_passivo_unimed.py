#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Transformar Passivo p/ Layout ADMEX – v9.1 (Motor FIFO + ID Movimento)
- Inclui a coluna ID Movimento (1 = Débitos/Acertos, 5 = Informativos).
- Os acertos em folha abatem progressivamente a dívida mais antiga (FIFO).
- Formato de data ajustado nativamente para o padrão aceito pelo ADMEX (mmm/aa).
"""

import argparse
import math
import unicodedata
import os
import re
import sys
import traceback
import datetime
import pandas as pd

# -------- UI auto-suficiente (janela) --------
def pick_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None
    root = tk.Tk(); root.withdraw()
    path = filedialog.askopenfilename(
        title="Selecione o arquivo Excel Consolidado",
        filetypes=[("Planilhas Excel", "*.xlsx;*.xls")],
    )
    return path

# -------- Auxiliares --------
def normaliza(s):
    if not isinstance(s, str): return s
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").strip()

def parse_val(x):
    if x is None or (isinstance(x, float) and math.isnan(x)): return math.nan
    if isinstance(x, (int, float)): return float(x)
    s = str(x).strip()
    if s == "" or s.upper() == "AUSENTE": return math.nan
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try: return float(s)
    except Exception: return math.nan

MESES_ABREV = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12}
MESES_NOME = {1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun", 7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"}

RE_ANY_MMM  = re.compile(r"\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[\s\./_-]*(\d{2,4})\b", re.I)
RE_ANY_MMYY = re.compile(r"\b(0?[1-9]|1[0-2])[\s\./_-]+(\d{2,4})\b")
RE_PURE_MMM  = re.compile(r"^(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[\s\./_-]*(\d{2,4})$", re.I)
RE_PURE_MMYY = re.compile(r"^(0?[1-9]|1[0-2])[\s\./_-]+(\d{2,4})$")

def _to_yyyy(a):
    a = int(a); return a + 2000 if a < 100 else a

def parse_ref_any(nome_col: str):
    if not isinstance(nome_col, str): return None
    nome = normaliza(nome_col.lower())
    m = RE_ANY_MMM.search(nome)
    if m: return f"{MESES_ABREV[m.group(1)[:3]]:02d}/{_to_yyyy(m.group(2))}"
    m = RE_ANY_MMYY.search(nome)
    if m: return f"{int(m.group(1)):02d}/{_to_yyyy(m.group(2))}"
    return None

def parse_ref_pura(nome_col: str):
    if not isinstance(nome_col, str): return None
    nome = normaliza(nome_col.strip().lower())
    m = RE_PURE_MMM.fullmatch(nome)
    if m: return f"{MESES_ABREV[m.group(1)[:3]]:02d}/{_to_yyyy(m.group(2))}"
    m = RE_PURE_MMYY.fullmatch(nome)
    if m: return f"{int(m.group(1)):02d}/{_to_yyyy(m.group(2))}"
    return None

def format_date_admex(ref_mm_yyyy):
    if not ref_mm_yyyy: return ""
    mm, yyyy = ref_mm_yyyy.split("/")
    mes_str = MESES_NOME[int(mm)]
    aa = yyyy[-2:] 
    return f"{mes_str}/{aa}"

def extract_date_from_cell(raw_date):
    if pd.isna(raw_date) or raw_date == "": 
        return ""
    if isinstance(raw_date, (pd.Timestamp, datetime.datetime)):
        return format_date_admex(raw_date.strftime("%m/%Y"))
    
    s = str(raw_date).strip()
    m1 = re.search(r"(\d{2})[/.-](\d{2})[/.-](\d{4})", s)
    if m1: return format_date_admex(f"{m1.group(2)}/{m1.group(3)}")
    
    m2 = RE_ANY_MMM.search(s)
    if m2: return format_date_admex(f"{MESES_ABREV[m2.group(1)[:3].lower()]:02d}/{_to_yyyy(m2.group(2))}")
    m3 = RE_ANY_MMYY.search(s)
    if m3: return format_date_admex(f"{int(m3.group(1)):02d}/{_to_yyyy(m3.group(2))}")
    
    return ""

# -------- Núcleo --------
def run_transform(xlsx_path, outdir):
    print(f"[START] Lendo arquivo '{xlsx_path}'...")
    xl = pd.ExcelFile(xlsx_path, engine="openpyxl")
    
    all_deb_records = []
    all_acertos_records = []

    NOME_LANCAMENTO = "Débito plano UNIMEDBH"

    for sheet_name in xl.sheet_names:
        if "UNIMED" not in sheet_name.strip().upper() or "RESIDUOS" in sheet_name.strip().upper():
            continue
            
        print(f"  -> Processando aba: {sheet_name}")

        raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, engine="openpyxl")
        hdr = None
        for i in range(min(100, len(raw))):
            vals = raw.iloc[i].astype(str).str.strip().tolist()
            if any(v.upper() in ("CPF", "CPF TITULAR") for v in vals):
                hdr = i; break
        
        if hdr is None: continue

        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=hdr, engine="openpyxl")
        id_col = 'CPF' if 'CPF' in df.columns else 'CPF TITULAR'
        if id_col not in df.columns: continue

        i_saldo = None; i_saldo_liq = None
        for i, c in enumerate(df.columns):
            if isinstance(c, str) and "SALDO DEBITOS" in c.upper(): i_saldo = i
            if isinstance(c, str) and "SALDO LIQUIDO" in c.upper(): i_saldo_liq = i
        
        if i_saldo is None or i_saldo_liq is None: continue

        cols_deb    = [c for j,c in enumerate(df.columns) if j < i_saldo and parse_ref_any(c)]
        cols_folha  = [c for j,c in enumerate(df.columns) if i_saldo < j < i_saldo_liq and parse_ref_pura(c)]
        cols_valmes = [c for j,c in enumerate(df.columns) if j > i_saldo_liq and isinstance(c,str) and parse_ref_any(c)]
        
        cols_info = []
        for j, c in enumerate(df.columns):
            if isinstance(c, str) and "INFORMATIVO" in c.upper():
                date_col = df.columns[j+1] if j+1 < len(df.columns) else None
                cols_info.append((c, date_col))

        for _, row in df.iterrows():
            cpf = re.sub(r"\D", "", str(row[id_col]))
            if not cpf: continue

            # 1) MAPEAMENTO DE DÉBITOS (Obrigações)
            deb_list = []
            for col in cols_deb:
                val = parse_val(row.get(col, math.nan))
                if not math.isnan(val) and abs(val) > 0:
                    ref = parse_ref_any(col)
                    if ref:
                        deb_list.append({'ref': ref, 'valor': abs(val), 'abatido': 0.0})
                        all_deb_records.append({
                            "CPF Titular": cpf, 
                            "Referência": format_date_admex(ref), 
                            "ID Movimento": 1,
                            "Valor": round(abs(val), 2),
                            "Descrição": NOME_LANCAMENTO,
                            "Auditoria": "Débito"
                        })
            
            deb_list.sort(key=lambda x: int(x['ref'].split('/')[1]) * 100 + int(x['ref'].split('/')[0]))

            # 2) ACERTOS AVULSOS
            for col_v in cols_valmes:
                vv = parse_val(row.get(col_v, math.nan))
                if not math.isnan(vv) and abs(vv) > 0:
                    rv = parse_ref_any(col_v) 
                    if rv:
                        all_acertos_records.append({
                            "CPF Titular": cpf, 
                            "Referência": format_date_admex(rv), 
                            "ID Movimento": 1,
                            "Valor": round(-abs(vv), 2), 
                            "Descrição": NOME_LANCAMENTO,
                            "Auditoria": "Avulso"
                        })
                        for d in deb_list:
                            if d['ref'] == rv:
                                d['abatido'] += abs(vv)
                                break

            # 3) ACERTOS EM FOLHA (Motor FIFO)
            for col_f in cols_folha:
                vf = parse_val(row.get(col_f, math.nan))
                if not math.isnan(vf) and abs(vf) > 0:
                    restante = abs(vf)
                    rf_original = parse_ref_pura(col_f)
                    
                    for d in deb_list:
                        if restante <= 0: break
                        saldo_devedor = round(d['valor'] - d['abatido'], 2)
                        
                        if saldo_devedor > 0:
                            aplicar = min(saldo_devedor, restante)
                            
                            all_acertos_records.append({
                                "CPF Titular": cpf, 
                                "Referência": format_date_admex(d['ref']), 
                                "ID Movimento": 1,
                                "Valor": round(-aplicar, 2), 
                                "Descrição": NOME_LANCAMENTO,
                                "Auditoria": f"Folha (Ref. Folha {format_date_admex(rf_original)})"
                            })
                            
                            d['abatido'] += aplicar
                            restante -= aplicar
                    
                    if round(restante, 2) > 0 and rf_original:
                        all_acertos_records.append({
                            "CPF Titular": cpf, 
                            "Referência": format_date_admex(rf_original), 
                            "ID Movimento": 1,
                            "Valor": round(-restante, 2), 
                            "Descrição": NOME_LANCAMENTO,
                            "Auditoria": "Folha (Excedente sem dívida)"
                        })
            
            # 4) INFORMATIVOS DE SALDO
            for col_i, col_d in cols_info:
                vi = parse_val(row.get(col_i, math.nan))
                if not math.isnan(vi) and abs(vi) > 0:
                    raw_date = row.get(col_d) if col_d else None
                    ref_date = extract_date_from_cell(raw_date)
                    
                    all_acertos_records.append({
                        "CPF Titular": cpf, 
                        "Referência": ref_date, 
                        "ID Movimento": 5,
                        "Valor": round(-abs(vi), 2), 
                        "Descrição": NOME_LANCAMENTO,
                        "Auditoria": "INFORMATIVO"
                    })

    # Consolidação final
    df_entradas = pd.DataFrame(all_deb_records)
    if not df_entradas.empty:
        df_entradas = df_entradas.groupby(['CPF Titular','Referência','ID Movimento','Descrição', 'Auditoria'], as_index=False)['Valor'].sum()
        df_entradas['Valor'] = df_entradas['Valor'].round(2)

    df_acertos = pd.DataFrame(all_acertos_records)
    if not df_acertos.empty:
        df_acertos = df_acertos.groupby(['CPF Titular','Referência','ID Movimento','Descrição', 'Auditoria'], as_index=False)['Valor'].sum()
        df_acertos['Valor'] = df_acertos['Valor'].round(2)

    os.makedirs(outdir, exist_ok=True)
    arq_entradas = os.path.join(outdir, "Entradas_ADMEX.csv")
    arq_acertos = os.path.join(outdir, "Acertos_ADMEX.csv")

    if not df_entradas.empty:
        # Reordenando as colunas para o ID Movimento ficar bem visível
        df_entradas = df_entradas[['CPF Titular', 'Referência', 'ID Movimento', 'Valor', 'Descrição', 'Auditoria']]
        df_entradas.to_csv(arq_entradas, sep=";", index=False, decimal=",")
        print(f"\n[SUCESSO] {arq_entradas} gerado com {len(df_entradas)} linhas.")
    
    if not df_acertos.empty:
        df_acertos = df_acertos[['CPF Titular', 'Referência', 'ID Movimento', 'Valor', 'Descrição', 'Auditoria']]
        df_acertos.to_csv(arq_acertos, sep=";", index=False, decimal=",")
        print(f"[SUCESSO] {arq_acertos} gerado com {len(df_acertos)} linhas.")

def main():
    try:
        ap = argparse.ArgumentParser(description="Transformar Passivo p/ Layout ADMEX – v9.1", add_help=True)
        ap.add_argument('--xlsx', help='Excel de entrada consolidado')
        ap.add_argument('--outdir', default=None, help="Pasta de saída")
        args, _ = ap.parse_known_args()

        if not args.xlsx and len(sys.argv) == 1:
            xlsx = pick_file()
            if not xlsx:
                print("[ABORTADO] Nenhum arquivo selecionado.")
                return
            outdir = os.path.join(os.path.dirname(os.path.abspath(xlsx)), "saida")
            run_transform(xlsx, outdir)
            return

        if not args.xlsx:
            print("ERRO: O argumento --xlsx é obrigatório na linha de comando.")
            return
        
        outdir = args.outdir or os.path.join(os.path.dirname(os.path.abspath(args.xlsx)), "saida")
        run_transform(args.xlsx, outdir)
        
    except Exception as e:
        print("\n" + "="*50)
        print("[ERRO CRÍTICO] Ops! O script encontrou um problema:")
        print("="*50)
        traceback.print_exc()
        print("="*50)
    finally:
        input("\nPressione ENTER para fechar a janela...")

if __name__ == '__main__':
    main()