import pandas as pd
import os
import xml.etree.ElementTree as ET
import messagebox

# Ler a planilha de dados
dados = pd.read_excel(r'C:\Users\Log20-2\Desktop\Projeto Lançamento Automatico\Planilhas\notas.xlsx')

# Extrair os números da coluna "numero"
numeros_planilha = set(dados['numero'].tolist())

# Diretório onde estão os arquivos XML
diretorio_xml = r'C:\Users\Log20-2\Desktop\Doctos'

# Lista para armazenar o log
log = []

# Função para extrair o número da nota fiscal do XML
def extrair_numero_nf(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        # Considerando o namespace ao buscar a tag
        numero_nf_element = root.find('.//{http://www.portalfiscal.inf.br/nfe}nNF')
        if numero_nf_element is not None:
            numero_nf = numero_nf_element.text
            return numero_nf
        else:
            print(f"Tag <nNF> não encontrada no arquivo {xml_file}.")
            return None
    except Exception as e:
        print(f"Erro ao extrair número da nota fiscal do arquivo {xml_file}: {e}")
        return None

# Iterar sobre os arquivos XML
for arquivo_xml in os.listdir(diretorio_xml):
    if arquivo_xml.endswith('.xml'):
        caminho_xml = os.path.join(diretorio_xml, arquivo_xml)
        numero_xml = extrair_numero_nf(caminho_xml)
        
        if numero_xml is not None:
            # Comparar o número do XML com os números da planilha
            if int(numero_xml) in numeros_planilha:
                log.append({'Numero_NF_Encontrado': numero_xml, 'Arquivo_XML': arquivo_xml, 'Status': 'Encontrado'})
            else:
                log.append({'Numero_NF_Encontrado': numero_xml, 'Arquivo_XML': arquivo_xml, 'Status': 'Não encontrado'})
                # Se o arquivo não foi encontrado, vamos removê-lo
                os.remove(caminho_xml)

# Criar DataFrame com o log apenas para as entradas encontradas
df_log = pd.DataFrame(log)

# Salvar o log em um arquivo Excel apenas se houver entradas no log
if not df_log.empty:
    df_log.to_excel('log_notas_encontradas.xlsx', index=False)
    messagebox.showinfo("Log de notas encontradas salvo com sucesso.","Notas Excluidas com sucesso")
else:
    print("Nenhuma nota encontrada para gerar o log.")
