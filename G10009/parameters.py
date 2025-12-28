# =========================================================
# Imports
# =========================================================

from deep_translator import GoogleTranslator
import time
from concurrent.futures import ThreadPoolExecutor
import os
import pandas as pd
import numpy as np
from datetime import datetime
import re
from ftfy import fix_text
import glob
from rapidfuzz import process, fuzz


# =========================================================
# Paths e configurações
# =========================================================

path_root = (
    r"C:\Users\TAKAYK01\OneDrive - Heineken International"
    r"\PROJECTS\migracao hana\materials\files"
)

path_source = f"{path_root}\\source\\xxxxx.xlsx"
path_result = f"{path_root}\\result\\xxxxx_result.xlsx"

columns_final = [
    "column_name",
    "description_name",
    "company_code",
    "rule_id",
    "unique_id",
    "column_name_value",
]


# =========================================================
# Utilidades de limpeza
# =========================================================

def limpar_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.fillna("")
    df = df.replace("nan", "")
    df = df.astype(str)
    df = df.map(
        lambda x: " ".join(str(x).split()) if isinstance(x, str) else x
    )
    return df


def corrigir_mojibake(s):
    if not isinstance(s, str):
        return s
    try:
        return s.encode("latin1").decode("utf-8")
    except UnicodeError:
        return s.replace("Â", "")


# =========================================================
# Leitura de arquivos
# =========================================================

def ler_arquivo_aux(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df = limpar_df(df)
    df = df.map(
        lambda x: " ".join(str(x).split()) if isinstance(x, str) else x
    )
    return df


def ler_arquivo(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype={"unique_id": str})
    df = limpar_df(df)
    df = df.map(
        lambda x: " ".join(str(x).split()) if isinstance(x, str) else x
    )

    df["description"] = df["column_name_value"].str.replace(
        "Description ", "", regex=False
    )

    df[["column_name", "description_name"]] = (
        df["description"].str.split(":", expand=True)
    )

    df["description_name"] = df["description_name"].str.title()

    df = df[df["column_name_value"] != ""][columns_final].copy()

    return df


# =========================================================
# Materiais PT (de-para)
# =========================================================

def carregar_depara_materiais_pt() -> pd.DataFrame:
    file_path_mat = f"{path_root}\\source\\descricao_pt_materiais.xlsx"

    df = pd.read_excel(file_path_mat)

    df = df.rename(
        columns={
            "Material": "m
