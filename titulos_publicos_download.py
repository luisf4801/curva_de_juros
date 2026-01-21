# -*- coding: utf-8 -*-
"""
Created on Wed Jan 21 15:09:45 2026

@author: Luis
"""

import requests
import os

# 1. Defina o caminho exato da pasta de destino
# Usamos o prefixo 'r' antes das aspas para que o Python trate as barras invertidas corretamente
diretorio_destino = "C:\"# mude para o seu dir

# 2. Verifique se a pasta existe; se não existir, o Python irá criá-la
if not os.path.exists(diretorio_destino):
    os.makedirs(diretorio_destino)
    print(f"Pasta criada: {diretorio_destino}")
    
for titulo in ["LFT","LTN", "NTN-B","NTN-B_Principal"]:
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
