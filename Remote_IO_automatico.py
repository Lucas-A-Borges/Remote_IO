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

def preencher_canais_da_matriz(caminho_arquivo, matriz_hardware):
    tree = ET.parse(caminho_arquivo)
    root = tree.getroot()

    # Itera sobre todas as variáveis no XML
    for var in root.findall(".//variables"):
        nome_var = var.get("name", "")
        
        # Regex para extrair Drop e Slot do nome (Ex: ED_DROP02_SLOT04)
        match_info = re.search(r'DROP(\d+)_SLOT(\d+)', nome_var)
        
        if match_info:
            num_drop = int(match_info.group(1))
            num_slot = int(match_info.group(2))
            
            # Verifica se esse Drop e Slot existem na nossa matriz
            if num_drop in matriz_hardware and num_slot in matriz_hardware[num_drop].slots:
                obj_slot = matriz_hardware[num_drop].slots[num_slot]
                
                # Navega na estrutura interna: DIS_CH_IN -> [n] -> VALUE -> Alias
                # Usamos .//instanceElementDesc para buscar os níveis de canais
                for ch_desc in var.findall(".//instanceElementDesc"):
                    ch_name = ch_desc.get("name", "") # Ex: "[0]", "[1]"
                    
                    if ch_name.startswith("[") and ch_name.endswith("]"):
                        try:
                            # Extrai o índice do canal: [0] -> 0
                            indice_canal = int(ch_name.strip("[]"))
                            
                            # Busca o atributo 'Alias' dentro do elemento 'VALUE'
                            value_elem = ch_desc.find(".//instanceElementDesc[@name='VALUE']")
                            if value_elem is not None:
                                attribute_alias = value_elem.find("attribute[@name='Alias']")
                                if attribute_alias is not None:
                                    tag_nome = attribute_alias.get("value")
                                    
                                    # Atribui ao objeto Canal correspondente
                                    # Nota: No seu código Canal(i+1), então indice 0 é Canal(1)
                                    if indice_canal < len(obj_slot.canais):
                                        obj_slot.canais[indice_canal].nome = tag_nome
                                        #print(f"Preenchido: Drop {num_drop} Slot {num_slot} Canal {indice_canal} -> {tag_nome}")
                        except ValueError:
                            continue

def preencher_comentarios_na_matriz(matriz_hardware, lista_variaveis_lidas):
    # 1. Criar um dicionário de busca rápida {nome_da_tag: comentario}
    # Isso evita ter que percorrer a lista inteira para cada canal (O(1) vs O(n))
    mapa_comentarios = {
        var['nome']: var['comentario'] 
        for var in lista_variaveis_lidas 
        if var['nome']
    }

    print("Iniciando preenchimento de comentários...")
    contador = 0

    # 2. Navegar na matriz de hardware
    for drop in matriz_hardware.values():
        for slot in drop.slots.values():
            for canal in slot.canais:
                # Se o canal tiver um nome (tag) atribuído
                if canal.nome:
                    # Busca o comentário no mapa que criamos
                    comentario = mapa_comentarios.get(canal.nome)
                    
                    if comentario:
                        canal.comentario = comentario
                        contador += 1

    print(f"Sucesso: {contador} comentários vinculados aos canais.")


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

    # 2. Gerar a estrutura a partir do hardware do PLC
    matriz_hardware = gerar_matriz_plc(caminho_unitpro)
 
    # 3. Preencher os nomes dos canais com base nas variáveis do arquivo
    preencher_canais_da_matriz(caminho_unitpro, matriz_hardware)

    # 4. Preencher os COMENTÁRIOS nos canais
    # Cruza os dados da matriz com a lista_variaveis_lidas
    preencher_comentarios_na_matriz(matriz_hardware, lista_variaveis_lidas)

    print("Processamento concluído.") 