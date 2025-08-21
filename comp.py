import json
import requests
import urllib3
import pandas as pd
import re
import time
from collections import defaultdict
from pathlib import Path

def comparar_diccionarios(aceptadas: dict, obtenidas: dict):
    coincidencias = 0
    total = 0
    
    # Copia de obtenidas para ir quitando coincidencias
    obtenidas_restantes = {k: v.copy() for k, v in obtenidas.items()}

    for clave, valores_aceptados in aceptadas.items():
        total += len(valores_aceptados)
        if clave in obtenidas_restantes:
            for valor in valores_aceptados:
                if valor in obtenidas_restantes[clave]:
                    coincidencias += 1
                    # eliminar coincidencia de obtenidas
                    obtenidas_restantes[clave].remove(valor)

    return f"{coincidencias}/{total}", obtenidas_restantes


def parse_dict(row_str):
    elements = row_str.split("],")
    result = {}
    for e in elements:
        e = e.strip().strip(",").strip("[]")
        if e:
            parts = e.split(",")
            key = parts[0].strip('"').strip("'")
            values = list(map(int, parts[1:]))
            result[key] = values
    return result

def compare_dicts_detalle(dict1, dict2,dict3):
    coincidencias = {}
    faltantes = {}
    sobrantes = {}
    adicional = {}


    total_coincidencias_valores = 0
    total_faltantes = 0
    total_sobrantes = 0

    claves_union = set(dict1.keys()).union(dict2.keys()).union(dict2.keys())

    for clave in claves_union:
        valores1 = dict1.get(clave, [])
        valores2 = dict2.get(clave, [])
        valores3 = dict3.get(clave, [])
        aceptadas , valores2 = comparar_diccionarios(valores3, valores2)



        coinciden = sorted([v for v in valores1 if v in valores2])
        faltan = sorted([v for v in valores1 if v not in valores2])
        sobran = sorted([v for v in valores2 if v not in valores1])

        if coinciden:
            coincidencias[clave] = coinciden
            total_coincidencias_valores += len(coinciden)
        if faltan:
            faltantes[clave] = faltan
            total_faltantes += len(faltan)
        if sobran:
            sobrantes[clave] = sobran
            total_sobrantes += len(sobran)

    return {
        "coincidencias": coincidencias,
        "faltantes": faltantes,
        "sobrantes": sobrantes,
        "total_coinciden": total_coincidencias_valores,
        "total_faltan": total_faltantes,
        "total_sobran": total_sobrantes,
        "Aceptadas": aceptadas
    }




def Procesar_respuestas():
    # Cargar el archivo original
    dfs = pd.read_excel("respuestas.xlsx")
    # Preparar lista para resultados nuevos
    resultados = []
   
    for idx, row in dfs.iterrows():
        dict1 = parse_dict(str(row["Expected Quotes"]))
        dict2 = parse_dict(str(row["Result Quotes"]))
        dict3 = parse_dict(str(row["Accepted Quotes"]))
        comparacion = compare_dicts_detalle(dict1, dict2,dict3)

        fila_resultado = row.to_dict()  # Copia los datos originales
        fila_resultado.update({
            "Total de documentos citados correctamente": comparacion["total_coinciden"],
            "Total de documentos faltantes": comparacion["total_faltan"],
            "Total de documentos sobrantes": comparacion["total_sobran"],
            "Detalle de documentos correctos": str(comparacion["coincidencias"]),
            "Detalle de documentos faltantes": str(comparacion["faltantes"]),
            "Detalle de documentos Sobrantes": str(comparacion["sobrantes"]),
            "aceptadas": str(comparacion["aceptadas"]),
        })

        resultados.append(fila_resultado)

    # Guardar en nuevo Excel
    pd.DataFrame(resultados).to_excel("resultados.xlsx", index=False)
    print(f"\n✅ Archivo guardado")
 
Procesar_respuestas()
