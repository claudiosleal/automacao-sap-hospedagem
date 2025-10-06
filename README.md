# Sistema Gestor de Hospedagem — Automação SAP

Ferramenta de automação com interface gráfica desenvolvida em Python para otimizar e agilizar o lançamento de despesas de hospedagem no sistema SAP. A aplicação lê os dados de uma planilha Excel e executa processos complexos de forma automatizada, reduzindo o tempo de trabalho manual e minimizando erros.

## Funcionalidades Principais

-   **Interface Gráfica Amigável:** Criada com PySide2 para facilitar a interação do usuário.
-   **Leitura de Dados Estruturados:** Utiliza uma planilha Excel como fonte de dados para as automações.
-   **Gerenciador de Senhas Seguro:** Integra-se ao `keyring` do sistema operacional para armazenar e consultar credenciais SAP de forma segura, sem a necessidade de hardcoding.
-   **Automação de Múltiplos Processos SAP:**
    -   **Requisição de Compra (RC):** Criação automática de RCs (transação `ME51N`).
    -   **Pedido de Compra (PC):** Geração de PCs a partir de RCs (transação `ME21N`).
    -   **Registro de Serviço (FRS):** Lançamento de Folhas de Registro de Serviço (transação `ML81N`).
    -   **Gestão de Documentos (GD):** Anexo e gerenciamento de documentos fiscais (transação `MLGD`).

## Pré-requisitos

Para executar este projeto, você precisará de:

1.  **Python 3.8+**
2.  **Ambiente Windows:** A automação depende da biblioteca `pywin32` para se comunicar com a API de scripting do SAP GUI.
3.  **SAP GUI for Windows:** O cliente SAP deve estar instalado e configurado.
4.  **Scripting do SAP Habilitado:** O scripting precisa estar ativado no lado do servidor SAP e no cliente local.

## Instalação

Siga os passos abaixo para configurar o ambiente de desenvolvimento:

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/claudiosleal/automacao-sap-hospedagem.git](https://github.com/claudiosleal/automacao-sap-hospedagem.git)
    cd automacao-sap-hospedagem
    ```

2.  **Crie e ative um ambiente virtual:**
    ```bash
    # Criar o ambiente
    python -m venv venv

    # Ativar no Windows
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    O projeto requer algumas bibliotecas. Instale-as usando o arquivo `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```
    *Se o arquivo `requirements.txt` ainda não existir, crie-o com o seguinte conteúdo:*
    ```
    PySide2
    keyring
    pywin32
    pandas
    openpyxl
    ```

## Como Usar

1.  **Execute a Aplicação:**
    Inicie a interface gráfica executando o arquivo `main.py`.
    ```bash
    python main.py
    ```

2.  **Selecione a Planilha:**
    Clique no botão **"Selecione"** para carregar a planilha Excel (`.xlsx` ou `.xls`) contendo os dados de hospedagem.

3.  **Configure a Senha (Primeiro Uso):**
    -   Clique no botão **"Senha"**.
    -   Preencha os campos "Usuário", "Sistema" (ex: `saplogon`) e "Senha".
    -   Clique em **"Cadastrar Senha"**. Sua credencial será salva de forma segura no gerenciador de senhas do seu sistema operacional. O usuário SAP será salvo em um arquivo `sap_user.txt` na mesma pasta da planilha para ser carregado automaticamente nas próximas vezes.

4.  **Execute um Processo:**
    Com a planilha carregada, clique no botão correspondente ao processo que deseja automatizar:
    -   **Requisição**
    -   **Pedido**
    -   **Registro de Serviço**
    -   **Gestão de Documentos**

5.  **Acompanhe o Log:**
    O campo de texto na parte inferior da janela exibirá logs em tempo real, informando sobre o progresso da automação, conexões e possíveis erros.

## Estrutura do Projeto

```
automacao-sap-hospedagem/
│
├── core/
│   ├── __init__.py
│   └── servicos.py         # Lógica de negócio e automação SAP
│
├── ui/
│   ├── __init__.py
│   ├── hospeda.ui          # Arquivo de design da interface (Qt Designer)
│   └── main_ui.py    		# Código Python gerado a partir do .ui
│
├── .gitignore
├── LICENSE
├── main.py                 # Inicialização da aplicação e lógica da UI
├── README.md
└── requirements.txt
```

## Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
