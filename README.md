![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/CustomTkinter-005177?style=for-the-badge)

# 🤖 Automação de Rescaldo de Softwares
Ferramenta desenvolvida para automatizar atualizações em planilhas de controle de rescaldos de softwares e gestão de vulnerabilidades (CVEs). A aplicação substitui o processo manual de copiar e colar informações repetitivas, inserindo dados automaticamente e gerando backups de segurança a cada execução.

## 🚀 Funcionalidades
- 🖥️ **Interface Gráfica (GUI):** Tela limpa e fácil de usar feita com `customtkinter`.
- ✅ **Validação de Dados:** Garante que todos os campos obrigatórios sejam preenchidos antes da execução.
- 📊 **Inserção de Dados Automatizada:** Expande as tabelas necessárias no Excel e preserva formatação/fórmulas originais.
- 🛡️ **Backup Inteligente:** Cria cópias de segurança do arquivo principal antes de qualquer modificação, controlando o máximo de backups armazenados simultaneamente.
- 📂 **Execução Baseada em Arquivos Externos:** Lê lista de hostnames a partir de um arquivo Excel e os aplica na planilha principal.

## 📂 Estrutura do Projeto
```text
atualizacao_rescaldo_software/
├── 📂 amostras/                          # Planilhas de teste e dados gerados
│   └── 📂 backups/                       # Cópias de segurança criadas durante a execução
├── 📂 .venv/                             # Ambiente virtual (isolamento de pacotes)
├── 📄 .env                               # Configurações de caminhos locais (não versionado)
├── 📄 .env.example                       # Modelo para configuração do .env
├── 📄 .gitignore                         # Lista de ficheiros ignorados pelo Git
├── 📄 automacao_rescaldo_softwares.ico   # Ícone oficial do projeto
├── 📄 automation_core.py                 # Lógica principal da automação
├── 📄 automation_gui.py                  # Interface gráfica (CustomTkinter)
├── 📄 gerar_amostras.py                  # Script para criar dados de teste
├── 📄 README.md                          # Documentação do projeto
└── 📄 requirements.txt                   # Lista de dependências do Python
```

## 🛠️ Como usar

### 1. 📦 Preparação
1. Certifique-se de ter o [Python](https://www.python.org/downloads/) instalado (versão 3.13 ou superior recomendada).
2. Clone o repositório e acesse o repositório do projeto:
   ```bash
   git clone https://github.com/hugopinaffo/automacao-rescaldo-software.git

   cd automacao-rescaldo-software
   ```
3. Crie o ambiente virtual e instale as dependências necessárias:
   ```bash
   python -m venv .venv
   ```
4. Ative o ambiente virtual:
   ```bash
   # Windows:
   .venv\Scripts\activate

   # Linux/Mac:
   source .venv/bin/activate
   ```
5. Instale as dependências necessárias:
      ```bash
   pip install -r requirements.txt
   ```
6. Gere planilhas de exemplo:
   ```bash
   python gerar_amostras.py
   ```
   *Isso criará uma pasta `amostras/` com um arquivo principal e uma lista de máquinas.*

7. Configure o caminho da sua planilha principal no arquivo .env (baseie-se no .env.example):
   ```
   PLANILHA_PRINCIPAL_PATH="C:\caminho\para\planilha_principal.xlsx"
   ```

### 2. ▶️ Execução
Com o ambiente virtual ativo, execute o aplicativo de interface gráfica com o seguinte comando:

```bash
python automation_gui.py
```

Na interface:
1. Preencha **Requisição**, **WO**, **Software** e **CVE**.
2. Selecione a **Planilha Principal** (caso não tenha configurado no arquivo `.env`).
3. Selecione o arquivo de **Máquinas** (`.xlsx` ou `.xlsm`) contendo a lista de hostnames a serem processados.
4. Clique em **Executar Automação**.

### 3. Distribuição (Gerando o .exe)
Se desejar gerar uma versão executável para utilizar sem a necessidade de instalar o Python em outras máquinas:
1. Instale o PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Execute o build:
   ```bash
   pyinstaller --onefile --noconsole --icon=automacao_rescaldo_softwares.ico automation_gui.py
   ```
  - `--onefile`: Compacta tudo num único ficheiro .exe.
  - `--noconsole`: Impede a abertura do terminal ao iniciar a interface.
  - `--icon=automacao_rescaldo_softwares.ico`: Define o ícone do executável.
  - `automation_gui.py`: Nome do arquivo Python a ser compilado.
3. O arquivo final será gerado na pasta `/dist`.

## ⚠️ Considerações
- O arquivo `.env` deve ser mantido privado e nunca versionado (já está ignorado no `.gitignore`).
- As tabelas e abas do arquivo principal devem aderir ao formato esperado.

## ⚖️ Licença
Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
