<p align="center">
  <!-- Logo do projeto (opcional) -->
  <img src="docs/logo.png" alt="Logo zimbra-email-automation" width="180" />
</p>

<h1 align="center">zimbra-email-automation</h1>

<p align="center">
  Automação para criação em massa de contas no Zimbra Admin
</p>

<p align="center">
  <!-- Badges -->
  <a href="https://www.python.org/">
    <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python" />
  </a>
  <a href="https://www.selenium.dev/">
    <img src="https://img.shields.io/badge/Selenium-Automation-43B02A?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium" />
  </a>
  <img src="https://img.shields.io/badge/Status-Production%20Ready-success?style=for-the-badge" alt="Status" />
  <img src="https://img.shields.io/badge/Planilha-Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Excel" />
</p>

-----------------------------------------------------------

Automação para criação em massa de contas no Zimbra Admin

## 📌 Sobre o projeto
Este projeto nasceu a partir de uma **demanda real da diretoria da empresa**, que precisava criar **dezenas de contas de e-mail corporativo** para colaboradores que ainda não possuíam acesso ao ambiente institucional.

Criar cada conta manualmente no Zimbra Admin é um processo:

- lento  
- repetitivo  
- sujeito a erros humanos  
- inviável em grande escala  

Para resolver isso, foi desenvolvida uma automação usando **Python + Selenium**, capaz de criar contas automaticamente, popular os dados dos colaboradores a partir de uma planilha Excel, vincular grupos, tratar erros e retomar a execução caso algo falhe.

---

## ✨ Funcionalidades

- ✔️ Criação automática de contas no Zimbra  
- ✔️ Leitura de usuários via planilha Excel  
- ✔️ Remoção automática de acentos (Zimbra não aceita)  
- ✔️ Associação ao grupo corporativo  
- ✔️ Senha inicial configurada e obrigatoriedade de troca  
- ✔️ Registro automático em planilha (Email criado / ERRO)  
- ✔️ Retomada automática da linha onde parou  
- ✔️ Reinício automático do navegador em caso de falha  
- ✔️ Evita duplicações ou recriações desnecessárias  
- ✔️ Totalmente tolerante a falhas temporárias do Zimbra  

---
## 🛠 Tecnologias utilizadas

- **Linguagem:** Python 3.10+
- **Automação de navegador:** Selenium WebDriver (Google Chrome)
- **Manipulação de planilhas:** pandas + openpyxl
- **Normalização de texto:** unidecode (remoção de acentos)
- **Planilha de entrada:** Microsoft Excel (`Emails.xlsx`)
- **Ambiente alvo:** Zimbra Admin (criação de contas de e-mail)

---

## 📁 Estrutura do projeto

```
GIT - AUTOMAÇÃO EMAIL/
│
├── auto.py               # Script principal da automação
├── config.py             # Configurações reais (IGNORADO no Git)
├── config_example.py     # Exemplo de configuração (sem senhas)
│
├── arquivos/
│   └── Emails.xlsx       # Planilha com os usuários (IGNORADA no Git)
│
├── .gitignore
├── README.md
└── requirements.txt
```

---

## 📄 Planilha de entrada (`arquivos/Emails.xlsx`)

A planilha deve conter:

- **Email** — nome do usuário (sem domínio)
- **1 nomes** — primeiro nome
- **Sobrenome**
- **Senha** — opcional
- **Situação** — preenchida automaticamente

### Exemplo:

| Email        | 1 nomes | Sobrenome       | Senha     | Situação     |
|--------------|---------|------------------|------------|--------------|
| joao.pereira | João    | Pereira Silva    | Unimed123  | Email criado |
| maria.sousa  | Maria   | Souza Lima       | Unimed123  | ERRO         |

---

## ⚙️ Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/seu-usuario/zimbra-email-automation.git
cd zimbra-email-automation
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

### 3. Configure o arquivo `config.py`

Copie o arquivo de exemplo:

```bash
cp config_example.py config.py
```

Edite o arquivo `config.py`:

```python
email = "seu_usuario_admin"
senha = "sua_senha_admin"
senhaemail = "senha_padrao_para_novos_usuarios"
```

> ⚠️ Importante: `config.py` está no `.gitignore` e **não será enviado ao GitHub**.

---

## ▶️ Como executar

1. Coloque `Emails.xlsx` dentro da pasta `/arquivos`  
2. Execute a automação:

```bash
python auto.py
```

### O sistema irá:

- Abrir o Zimbra Admin  
- Criar as contas automaticamente  
- Associar ao grupo corporativo  
- Registrar o status na planilha  
- Reiniciar o navegador caso o Zimbra trave  
- Continuar até que todas as contas sejam criadas  

---

## 🧠 Tolerância a erros

### ✔️ Se o botão **Novo** não aparecer  
Isso é um erro temporário do Zimbra.  
→ O sistema **reinicia o navegador** e tenta novamente, sem registrar ERRO.

### ✔️ Se ocorrer erro real na criação  
Ex.: usuário já existe, dados inválidos etc.  
→ A linha é marcada como **ERRO**  
→ O script segue para o próximo usuário.

### ✔️ Se o processo for interrompido  
Basta rodar novamente.  
Ele ignora linhas já marcadas como "Email criado" ou "ERRO".

---

## 🔐 Segurança

- Dados sensíveis **não são enviados** ao repositório  
- Planilhas com colaboradores **são ignoradas**  
- Arquivo `config.py` fica apenas no ambiente local  
- `.gitignore` garante proteção automática  

---

## 🧭 Motivação do projeto

A diretoria solicitou a criação de muitas contas corporativas rapidamente, pois num levantamento foi identificado que vários colaboradores ainda não possuíam e-mails institucionais.

A automação reduziu a criação de **horas de trabalho manual** para **poucos minutos**, garantindo:

- agilidade  
- consistência  
- redução de falhas  
- padronização  
- eficiência operacional  

---

## 📄 Licença

Sugerido:

```
MIT License
```

---

## 🎯 Melhorias futuras

- Logs detalhados em arquivo `.log`  
- Dashboard com progresso e estatísticas  
- Interface gráfica  
- Exportação de relatório final  

---

## 🙌 Contribuindo

Sugestões e melhorias são bem-vindas!  
Abra uma *issue* ou envie um *pull request* 🚀

---

# 🧾 `.gitignore` (use este conteúdo)

```gitignore
# Arquivo de configuração com senhas
config.py

# Planilhas com dados sensíveis
arquivos/*.xlsx

# Cache do Python
__pycache__/
*.pyc

# Ambientes virtuais
venv/
env/
```

---

# 📦 requirements.txt

```text
selenium
pandas
openpyxl
unidecode
```

---

# 🔐 config_example.py

```python
# Preencha este arquivo e renomeie para config.py

email = "USUARIO_DO_ZIMBRA_ADMIN"
senha = "SENHA_DO_ZIMBRA_ADMIN"
senhaemail = "SENHA_INICIAL_PARA_NOVOS_USUARIOS"
```
