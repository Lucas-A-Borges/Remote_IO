import pandas as pd
import xml.etree.ElementTree as ET
import os
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Side, Border
from datetime import datetime
from typing import Dict, Any, List
import sys

#definições ------------------------------------------
ARQUIVO_UNITPRO = 'unitpro.xef'
TIPOS_PERMITIDOS = ['WORD', 'BOOL', 'EBOOL', 'INT']



MODELOS_CAPACIDADE = {
    "140ACI03000": 8, "140ACO02000": 4, "140ACO13000": 8,
    "140ARI03010": 8, "140DDI84100": 32, "140DAI54000": 16,
    "140DAI55300": 32, "140DAO84210": 16, "140DAI74000": 16,
    "140DDI35300": 32, "140DDO35300": 32, "BMXDDI3202K": 32,
    "BMXDDO3202K": 32
}

class Canal:
    def __init__(self, numero):
        self.numero = numero
        self.nome = ""
        self.comentario = ""

class Slot:
    def __init__(self, numero, modelo):
        self.numero = numero
        self.modelo = modelo
        self.qtd_canais = MODELOS_CAPACIDADE.get(modelo, 0)
        # Inicializa a lista de canais com base na capacidade do modelo
        self.canais = [Canal(i+1) for i in range(self.qtd_canais)]

class Drop:
    def __init__(self, numero):
        self.numero = numero
        self.slots = {} # Dicionário {numero_slot: Objeto Slot}

def gerar_matriz_plc(caminho):
    tree = ET.parse(caminho)
    root = tree.getroot()
    
    drops = {} # Dicionário {numero_drop: Objeto Drop}

    # Procurar por módulos (Quantum ou outros que sigam a mesma estrutura)
    # O XPath pode variar dependendo da estrutura completa do seu arquivo
    for module in root.findall(".//moduleQuantum"):
        try:
            # Pegar o Part Number (Modelo)
            part_item = module.find("partItem")
            modelo = part_item.get("partNumber")

            # Pegar o TopoAddress (Ex: \2.2\1.3)
            equip_info = module.find("equipInfo")
            address = equip_info.get("topoAddress")

            # Regex para extrair DROP e SLOT do endereço \2.X\1.Y
            match = re.search(r'\\2\.(\d+)\\1\.(\d+)', address)
            if match:
                num_drop = int(match.group(1))
                num_slot = int(match.group(2))

                # Adiciona Drop se não existir
                if num_drop not in drops:
                    drops[num_drop] = Drop(num_drop)
                
                # Adiciona Slot ao Drop
                drops[num_drop].slots[num_slot] = Slot(num_slot, modelo)
                
                print(f"Mapeado: Drop {num_drop}, Slot {num_slot} -> Modelo {modelo}")

        except Exception as e:
            print(f"Erro ao processar módulo: {e}")

    return drops

def ler_variaveis_unitpro(caminho_arquivo: str) -> List[Dict[str, str]]:
    """Lê todas as variáveis do unitpro.xef."""
    # ... (Implementação omitida por brevidade, assumida como funcional)
    lista_variaveis = []
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
    except (FileNotFoundError, ET.ParseError) as e:
        print(f"ERRO: Não foi possível ler ou fazer o parse do arquivo: {e}")
        return lista_variaveis

    for var_element in root.findall('.//variables'):
        nome = var_element.get('name')
        tipo = var_element.get('typeName')
        endereco = var_element.get('topologicalAddress')
        comentario_element = var_element.find('comment')
        comentario = comentario_element.text.strip() if comentario_element is not None and comentario_element.text else ""

    

        if nome and tipo in TIPOS_PERMITIDOS: # and endereco
            lista_variaveis.append({
                'nome': nome,
                'comentario': comentario,
                'endereco': endereco,
                'tipo': tipo
            })
    return lista_variaveis


#----------------------MAIN----------------------------
if __name__ == "__main__":

    # --- DEFINIÇÃO UNIVERSAL DO CAMINHO BASE ---
    # Essa lógica funciona tanto para o script .py quanto para o executável .exe (frozen)
    if getattr(sys, 'frozen', False):
        # Se estiver rodando como executável (PyInstaller), usa o caminho do binário.
        diretorio_script = os.path.dirname(sys.executable)
    else:
        # Se estiver rodando como script Python (.py), usa o caminho do arquivo de script.
        # É fundamental usar o try-except ou um método robusto para evitar erros ao ser chamado de outro diretório.
        try:
            diretorio_script = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            # Fallback caso __file__ não esteja definido (raro, mas seguro)
            diretorio_script = os.path.getcwd() 

    # --- Configuração de Caminhos ---
    caminho_unitpro = os.path.join(diretorio_script, ARQUIVO_UNITPRO)

    # 1. Leitura e Catalogação das Variáveis
    lista_variaveis_lidas = ler_variaveis_unitpro(caminho_unitpro)

    # 1. Gerar a estrutura a partir do hardware do PLC
    matriz_hardware = gerar_matriz_plc(caminho_unitpro)
    print("Matriz de Hardware Gerada:")   
        
        