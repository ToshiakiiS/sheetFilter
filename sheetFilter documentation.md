# **README: Large Spreadsheet Column Filter 📊**

Este projeto é uma ferramenta web interativa desenvolvida em **Python** utilizando **Streamlit**. O objetivo dela é solucionar um problema clássico do time: **manipular arquivos massivos (Excel ou CSV) com centenas de milhares de linhas que travam o computador ou ultrapassam os limites de memória.**

Com esta ferramenta, você faz o upload de uma planilha pesada, seleciona de forma instantânea apenas as colunas que realmente importam para a sua tarefa e exporta um arquivo limpo, otimizado e muito mais leve.

## **📌 Principais Recursos e Diferenciais de Performance**

Para garantir que o script consiga processar bases com mais de **800.000 linhas** sem estourar a memória RAM do servidor ou travar o navegador, aplicamos técnicas avançadas de otimização:

* **Varredura Otimizada (Lazy Loading):** Utiliza o `openpyxl` nos modos estruturais `read_only=True` e `write_only=True`. Isso significa que o script lê e escreve o arquivo linha por linha em fluxo (streaming), sem carregar o arquivo inteiro na memória do computador de uma vez só.  
* **Interface Instantânea (Session State Caching):** O script analisa os cabeçalhos do arquivo apenas **uma única vez** no momento do upload. Ele memoriza a estrutura de colunas no estado da sessão do Streamlit (`st.session_state`), permitindo que você marque e desmarque colunas na interface sem ter que esperar o arquivo ser relido do zero a cada clique.  
* **Barreira Física do Excel (Safety Guard):** O Excel possui um limite físico intransponível de **1.048.576 linhas**. O script monitora isso ativamente e exibe um erro amigável se você tentar gerar um arquivo `.xlsx` que estoure esse teto, sugerindo automaticamente a exportação em `.csv`.

## **🛠️ Pré-requisitos e Dependências**

Para rodar este script localmente em sua máquina, você precisará do Python instalado e das seguintes bibliotecas:

* **Streamlit:** Para renderizar a interface web.  
* **Pandas:** Para manipulação ágil das fatias de dados (especialmente útil em CSVs).  
* **Openpyxl:** Engine responsável por ler e escrever os arquivos Excel de forma ultra otimizada.

### **Como Instalar:**

Abra o seu terminal na pasta do projeto e execute o comando abaixo:

Bash

```
pip install streamlit pandas openpyxl
```

## **🚀 Como Executar a Ferramenta**

1. Salve o código fornecido em um arquivo chamado `app.py`.  
2. No seu terminal, navegue até a pasta onde salvou o arquivo e execute:

Bash

```
streamlit run app.py
```

3.   
   Uma aba no seu navegador padrão será aberta automaticamente no endereço local (geralmente `http://localhost:8501`).

## **⚙️ Fluxo de Trabalho (Passo a Passo)**

```
[ Upload do Arquivo ] ➔ [ Leitura Única dos Cabeçalhos ] ➔ [ Seleção de Colunas no UI ] ➔ [ Processamento em Lote ] ➔ [ Download Link ]
```

1.   
   **Upload da Base:** Arraste e solte ou clique para selecionar seu arquivo `.csv` ou `.xlsx` de grande porte.  
2. **Seleção de Colunas:** Uma caixa de seleção múltipla (`multiselect`) será preenchida instantaneamente com todas as colunas encontradas. Por padrão, todas vêm selecionadas. Remova as que você não vai usar na sua task.  
3. **Formato de Saída:** Escolha se deseja exportar o resultado final em **CSV** (mais rápido e sem limite de linhas) ou **Excel (.xlsx)**.  
4. **Processar:** Clique no botão **`Process & Generate Download Link`**. O script processará a filtragem em segundo plano exibindo uma animação de carregamento.  
5. **Download:** Assim que concluído, uma mensagem de sucesso exibirá o total exato de linhas salvas e liberará um botão para baixar o seu novo arquivo filtrado.

## **⚠️ Notas de Atenção para Operação**

💡 **Dica de Ouro:** Sempre que o arquivo de origem passar de 1 milhão de linhas, **obrigatoriamente escolha o formato de saída como CSV**. O formato Excel não suporta essa volumetria de linhas por limitação do próprio ecossistema da Microsoft, e o script interromperá a execução para proteger os dados.

