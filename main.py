import pandas as pd
import os
import smtplib
import datetime
from datetime import datetime
from openpyxl import load_workbook
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time


def filtrar_vendas():

    df = pd.read_excel(r"Arquivos\Vendas.xlsx", engine="openpyxl")
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    df_dezembro = df[df['Data'].dt.month == 12]
    
    # Converte as colunas para numérico, ignorando erros
    df_dezembro['Quantidade'] = pd.to_numeric(df_dezembro['Quantidade'], errors='coerce')
    df_dezembro['Receita'] = pd.to_numeric(df_dezembro['Receita'], errors='coerce') 
    
    # Agrupa os dados por categoria, somando as quantidades e as receitas
    resumo_vendas = df_dezembro.groupby('Categoria').agg(
        Quantidade=('Quantidade', 'sum'),
        Vendas_Totais=('Receita', 'sum')
    ).reset_index()
    
    resumo_vendas['Quantidade'] = resumo_vendas['Quantidade'].apply(lambda x: f'{x:,.0f}'.replace(',', '.'))
    resumo_vendas['Vendas_Totais'] = resumo_vendas['Vendas_Totais'].apply(
        lambda x: f'R$ {x:,.2f}'.replace(',', '@').replace('.', ',').replace('@', '.')
    )
    
    resumo_vendas = resumo_vendas.sort_values(by='Categoria', ascending=True)
    print(resumo_vendas)
    
    # Define o diretório e o nome do arquivo
    diretorio = "output"
    os.makedirs(diretorio, exist_ok=True)
    nome_arquivo = f"Resumo_Vendas_{datetime.now().strftime('%Y%m%d')}.xlsx"
    caminho_completo = os.path.join(diretorio, nome_arquivo)
    
    resumo_vendas.to_excel(caminho_completo, index=False, engine="openpyxl")

    # Ajuste automático da largura das colunas
    wb = load_workbook(caminho_completo)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter  # Obtém a letra da coluna
        
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        
        ws.column_dimensions[col_letter].width = max_length + 2  # Ajusta a largura
    wb.save(caminho_completo)
    print(f"Arquivo salvo em: {caminho_completo}")        
                
def enviar_email():
    
    # Carrega variáveis do arquivo .env
    load_dotenv()

    # Obtém configurações das variáveis de ambiente
    pasta = os.getenv('PASTA_ARQUIVOS')
    email_remetente = os.getenv('EMAIL_REMETENTE')
    senha_app = os.getenv('SENHA_APP')
    
    # Verifica se todas as variáveis necessárias estão definidas
    if None in [pasta, email_remetente, senha_app]:
        print("Erro: Algumas variáveis de ambiente não estão definidas.")
        print("Verifique seu arquivo .env e as seguintes variáveis:")
        print("PASTA_ARQUIVOS, EMAIL_REMETENTE, SENHA_APP")
        return
    
    data_atual = datetime.now().strftime('%d/%m/%Y')
    
    try:
        # 1. Encontrar o arquivo mais recente na pasta
        arquivos = [f for f in os.listdir(pasta) if f.endswith('.xlsx')]
        if not arquivos:
            print(f"Nenhum arquivo .xlsx encontrado na pasta: {pasta}")
            return
        
        # Pegar o arquivo mais recente
        arquivo_mais_recente = max(
            [os.path.join(pasta, f) for f in arquivos],
            key=os.path.getmtime
        )
        nome_arquivo = os.path.basename(arquivo_mais_recente)
        
        msg = MIMEMultipart()
        msg['From'] = email_remetente
        msg['To'] = 'exemplo@empresa.com'
        msg['Subject'] = f'Resumo de Vendas - {data_atual}'
        
        corpo = """Prezado Gerente,  
Segue em anexo o resumo de vendas do mês atual.  

Atenciosamente,  
Equipe de Automação"""
        
        msg.attach(MIMEText(corpo, 'plain'))
        
        with open(arquivo_mais_recente, 'rb') as anexo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(anexo.read())
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{nome_arquivo}"'
            )
            msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(email_remetente, senha_app)
            server.send_message(msg)
            print("Email enviado com sucesso!")
            
    except Exception as e:
        print(f"Erro ao enviar email: {str(e)}")
        if hasattr(e, 'smtp_error'):
            print(f"Detalhes SMTP: {e.smtp_error.decode()}")  
                                                 
def main():
    filtrar_vendas()
    time.sleep(3)
    enviar_email()

if __name__ == "__main__":
    main()
