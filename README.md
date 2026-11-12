# 📊 Gerador de Painel de Cobrança via WhatsApp Web



[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/amisadaimodesto/gerador-cobranca-whatsapp/blob/main/gerador_cobranca_colab.ipynb)



Script e notebook Python que processam planilhas de inadimplentes (`.xls` ou `.xlsx`) e geram uma página **HTML autônoma** para gestão de cobranças, contendo:



* **Links diretos para o WhatsApp Web** com saudação automática pelo horário de Brasília, cálculo de valores e diferenciação singular/plural;

* **Controle de contatos (`Contatado?`)** com indicadores dinâmicos de quantidade, percentual e valor monetário em R$;

* **Campo de observação** persistente no navegador e botão para exportar uma cópia do HTML com as anotações gravadas no próprio arquivo;

* **Exportação para Excel (`.xlsx`)** limpa, consolidada e pronta para prestação de contas.



---



## 📁 Estrutura do Repositório



```text

gerador-cobranca-whatsapp/

├── gerador_cobranca_colab.ipynb           # Notebook interativo (Google Colab / Jupyter)

├── gerar_inadimplentes_whatsapp_colab.py  # Motor principal de processamento (Python)

├── requirements.txt                       # Dependências do projeto

├── .gitignore                             # Proteção contra envio de dados sensíveis (.xls/.xlsx/.html)

└── README.md                              # Documentação do projeto

```

---

## 🚀 Como usar

### Opção 1: Google Colab (Recomendado — 1 clique)

1. Clique no botão **Open in Colab** no topo desta página.

2. No **Passo 1**, preencha o nome do atendente e o nome da empresa ou deixe em branco para os valores padrão).

3. No **Passo 2**, execute a célula e selecione sua planilha `.xls` ou `.xlsx`. O arquivo HTML será baixado automaticamente.



### Opção 2: Execução Local (Linux / Windows / macOS)

1. Clone o repositório e ative o ambiente virtual:

   ```bash

   git clone https://github.com/amisadaimodesto/gerador-cobranca-whatsapp.git

   cd gerador-cobranca-whatsapp

   python3 -m venv .venv

   source .venv/bin/activate

   pip install -r requirements.txt

   ```

2. Execute via terminal ou abra `gerador_cobranca_colab.ipynb` no VS Code / Jupyter:

   ```bash

   python3 gerar_inadimplentes_whatsapp_colab.py caminho/da/planilha.xlsx

   ```

   *(Opcional)* Você também pode especificar o nome do arquivo de saída como segundo argumento:

   ```bash

   python3 gerar_inadimplentes_whatsapp_colab.py caminho/da/planilha.xlsx painel_saida.html

   ```

---

## 🔒 Segurança e Privacidade

O arquivo `.gitignore` já está pré-configurado para bloquear planilhas (`.xls`, `.xlsx`, `.csv`) e arquivos gerados (`.html`), garantindo que informações reais de clientes nunca sejam enviadas acidentalmente para o repositório.
