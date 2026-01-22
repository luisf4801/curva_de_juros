# -*- coding: utf-8 -*-
"""
Created on Wed Jan 21 15:09:45 2026

# FGV - Fundacao Getúlio Vargas
# EESP - Escola de Economia São Paulo
# Mestrado Academico em Economia 
@author: Luis felipe porro
contato: luisfelipe.porro@gmail.com


"""
import requests
import os
import re
# 1. Defina o caminho exato da pasta de destino
# Usamos o prefixo 'r' antes das aspas para que o Python trate as barras invertidas corretamente

dir_base=r'C:\Users\Luis\Desktop\tese\curva de juros KR'
diretorio_destino=dir_base+"\database"

# 2. Verifique se a pasta existe; se não existir, o Python irá criá-la
if not os.path.exists(diretorio_destino):
    os.makedirs(diretorio_destino)
    print(f"Pasta criada: {diretorio_destino}")
  
    
  
    

for titulo in ["LFT","LTN", "NTN-B","NTN-B_Principal","NTN-F"]:
    for i in range(2002,2026):
        print(titulo)
        print(i)
        
    # 3. URL do arquivo que deseja baixar
        url = f"https://cdn.tesouro.gov.br/sistemas-internos/apex/producao/sistemas/sistd/{i}/{titulo}_{i}.xls"
        
        # 4. Defina o nome do arquivo final combinando o diretório e o nome do ficheiro
        nome_arquivo = f"{titulo}_{i}.xls"
        caminho_completo = os.path.join(diretorio_destino, nome_arquivo)
        
        # 5. Executa o download
        try:
            print(f"A descarregar arquivo para: {caminho_completo}...")
            
            # Faz a requisição ao servidor
            resposta = requests.get(url, stream=True)
            
            # Verifica se o download foi permitido (Status 200)
            resposta.raise_for_status()
            
            # Escreve o conteúdo no arquivo local em modo binário ('wb')
            with open(caminho_completo, 'wb') as f:
                for chunk in resposta.iter_content(chunk_size=8192):
                    f.write(chunk)
                    
            print("Download concluído com sucesso!")
        
        except requests.exceptions.RequestException as e:
            print(f"Erro ao baixar o arquivo: {e}")
            

import pandas as pd
import os
import xlrd

# 1. Caminho da pasta onde os arquivos foram baixados
diretorio_database = diretorio_destino
diretorio_database_limpo=dir_base+'\database_limpa'
# Lista todos os arquivos .xls na pasta
arquivos = [f for f in os.listdir(diretorio_database) if f.endswith('.xls')]



for nome_arquivo in arquivos:
    caminho_completo = os.path.join(diretorio_database, nome_arquivo)
    
    # --- EXTRAÇÃO DO NOME DO CONTRATO (NTN-B_principal) ---
    # Remove o sufixo de ano (ex: _2006) e a extensão .xls ou .xlsx
    nome_base_contrato = re.sub(r'_\d{4}\.xls[x]?$', '', nome_arquivo)
    nome_base_contrato_data = re.sub(r'.xls[x]?$', '', nome_arquivo)

    print(f"\n--- Processando: {nome_base_contrato} ---")
    
    # Criamos um dicionário para guardar os DataFrames de cada aba deste arquivo
    dict_abas_processadas = {}
    
    try:
        excel_file = pd.ExcelFile(caminho_completo, engine='xlrd')
        abas = excel_file.sheet_names
        
        for aba in abas:
            try:
                # Leitura
                df = pd.read_excel(caminho_completo, sheet_name=aba, dtype=str, engine='xlrd')
                
                # Pega a data de vencimento (ajustado conforme sua lógica)
                vencimento = df.columns[1]
                vencimento_dt = pd.to_datetime(vencimento, dayfirst=True)

                # Ajusta cabeçalhos
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)
                
                # Coluna de vencimento (fixa para a aba)
                df["vencimento"] = vencimento_dt.date()

                # Converte a primeira coluna (Datas de referência)
                coluna_dt = pd.to_datetime(df.iloc[:, 0], dayfirst=True, errors='coerce')
                df.iloc[:, 0] = coluna_dt.dt.date

                # Cálculo dos dias corridos
                # Nota: df.iloc[:, -1] é a coluna "vencimento" que acabamos de criar
                diferenca = pd.to_datetime(df["vencimento"]) - pd.to_datetime(df.iloc[:, 0])
                df['dias corridos vencimento'] = diferenca.dt.days

                # Adiciona o nome limpo do contrato
                df["contrato"] = nome_base_contrato

                # Adiciona o DF processado ao dicionário da aba
                dict_abas_processadas[aba] = df

            except Exception as e:
                print(f"Erro na aba {aba}: {e}")

        # --- SALVAR O NOVO EXCEL ---
        # Define o nome do arquivo de saída (ex: Processado_NTN-B_principal.xlsx)
        caminho_saida = os.path.join(diretorio_database_limpo, f"Processado_{nome_base_contrato_data}.xlsx")
        
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            for nome_aba, df_final in dict_abas_processadas.items():
                df_final.to_excel(writer, sheet_name=nome_aba, index=False)
        
        print(f"Arquivo salvo com sucesso: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao abrir {nome_arquivo}: {e}")

print("\n--- Todos os arquivos foram gerados com as abas originais ---")
