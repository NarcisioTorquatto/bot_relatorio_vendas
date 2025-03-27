
# Bot de Automação de Relatório de Vendas

Este script automatiza o processo de leitura de um arquivo Excel de vendas, gera um resumo com as vendas totais por categoria, salva o resumo em um novo arquivo Excel e envia o arquivo gerado por e-mail.

## Funcionalidades

1. **Leitura do arquivo Excel**:
   - O script lê o arquivo `Vendas.xlsx` localizado no diretório `./Arquivos/`.

2. **Filtragem dos dados**:
   - Filtra os dados do mês de dezembro.

3. **Geração de resumo**:
   - Agrupa os dados por categoria e calcula as somas das quantidades e vendas totais.

4. **Salvamento do resumo**:
   - O resumo gerado é salvo no diretório `./output/` com o nome `Resumo_Vendas_{DataAtual}.xlsx`.

5. **Envio de e-mail**:
   - O arquivo gerado é enviado por e-mail para o endereço `exemplo@empresa.com` com o assunto "Resumo de Vendas - {DataAtual}".
   - O corpo do e-mail inclui uma mensagem informando sobre o envio do resumo de vendas.

## Bibliotecas Requeridas

O script utiliza as seguintes bibliotecas:

- `pandas`: Para manipulação e análise de dados.
- `openpyxl`: Para leitura e escrita de arquivos Excel.
- `smtplib`: Para envio de e-mails.
- `email.mime`: Para construção do e-mail com anexos.
- `dotenv`: Para carregar variáveis de ambiente a partir de um arquivo `.env`.

Instale as bibliotecas necessárias executando:

```bash
pip install pandas openpyxl python-dotenv
```

## Configuração de Variáveis de Ambiente

O script utiliza variáveis de ambiente para armazenar informações sensíveis, como o e-mail do remetente e a senha do aplicativo. Crie um arquivo `.env` na raiz do projeto com as seguintes variáveis:

```env
EMAIL_REMETENTE: E-mail do remetente.
SENHA_APP: Senha do aplicativo gerada no Google (não a senha comum).
PASTA_ARQUIVOS: Caminho para a pasta com a Planilha.
EMAIL_DESTINATARIO: Email de Quem vai receber.
```

- **EMAIL_REMETENTE**: E-mail do remetente.
- **SENHA_APP**: Senha do aplicativo gerada no Google (não a senha comum).
- **PASTA_ARQUIVOS**: Caminho para a pasta com a Planilha.
- **EMAIL_DESTINATARIO**: Email de Quem vai receber.

## Como Usar

1. **Certifique-se de que o arquivo Excel** `Vendas.xlsx` esteja localizado no diretório `./Arquivos/`.
2. **Execute o script** para realizar as tarefas de leitura, filtragem, geração do resumo e envio do e-mail.

Exemplo de execução:

```bash
python seu_script.py
```

## Exemplo de Estrutura de Diretórios

```text
- seu_projeto/
  - Arquivos/
    - Vendas.xlsx
  - output/
  - seu_script.py
  - .env
  - README.md
```

## Detalhes da Tarefa

### 1. Leitura e Filtragem dos Dados

O arquivo `Vendas.xlsx` é lido com `pandas`, e a coluna de datas é convertida para o formato datetime. Os dados são então filtrados para incluir apenas as vendas realizadas no mês de dezembro.

### 2. Geração do Resumo

A função `groupby` do `pandas` é utilizada para agrupar os dados por categoria e somar as quantidades e valores de vendas.

### 3. Salvamento do Resumo

O resumo gerado é salvo como um arquivo Excel no diretório `./output/`, com o nome `Resumo_Vendas_{DataAtual}.xlsx`.

### 4. Envio de E-mail

O script usa `smtplib` para se conectar ao servidor SMTP do Gmail e enviar um e-mail com o resumo de vendas em anexo.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests.

## Licença

Este projeto é de código aberto e está disponível sob a licença MIT.
