from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from unidecode import unidecode
from config import email, senha, senhaemail
import time

# ========= CONFIG =========
ARQUIVO = "arquivos/Emails.xlsx"           # Caminho da planilha de usuários
ABA = "Planilha1"                          # Nome da aba do Excel
GRUPO_CORP = "corporativo@empresa..."       # Grupo de e-mail padrão a associar
URL = "https://seu-zimbra-admin..."        # URL do Zimbra Admin

TIMEOUT = 15
# ==========================

def carregar_planilha():
    df = pd.read_excel(ARQUIVO, sheet_name=ABA)
    if "Situação" not in df.columns:
        df["Situação"] = ""
    return df

def salvar_planilha(df):
    with pd.ExcelWriter(ARQUIVO, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=ABA, index=False)

def start_driver():
    options = Options()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    d = webdriver.Chrome(options=options)
    d.maximize_window()
    w = WebDriverWait(d, TIMEOUT)
    return d, w

def login(d, w):
    d.get(URL)
    w.until(EC.visibility_of_element_located((By.NAME, "ZLoginUserName"))).send_keys(email)
    w.until(EC.visibility_of_element_located((By.NAME, "ZLoginPassword"))).send_keys(senha)
    w.until(EC.element_to_be_clickable((By.ID, "ZLoginButton"))).click()

def abrir_form_novo(d, w):
    # Vai para a tela de Gerenciar e abre o menu "Novo"
    w.until(EC.element_to_be_clickable((By.ID, "zti__AppAdmin__Home__manActHV_textCell"))).click()
    w.until(EC.element_to_be_clickable((By.ID, "DWT63_settingimg"))).click()
    w.until(EC.element_to_be_clickable((By.ID, "zmi__zb_currentApp__NEW_MENU"))).click()

def criar_email_para_linha(d, w, row):
    col_user = "Email"
    col_nome = "1 nomes"
    col_sobrenome = "Sobrenome"
    col_senha = "Senha"

    usuario = str(row.get(col_user, "")).strip()
    usuario = unidecode(usuario)
    usuario = unidecode(usuario).lower().replace(" ", "")
    if not usuario:
        raise ValueError("Usuário (Email) vazio na planilha")

    nome = str(row.get(col_nome, "")).strip() if col_nome in row else ""
    sobrenome = str(row.get(col_sobrenome, "")).strip() if col_sobrenome in row else ""
    senha_para_esta_conta = (str(row.get(col_senha, "")).strip() or senhaemail)

    # Preenche dados (limpando antes)
    el = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_name_2"))); el.clear(); el.send_keys(usuario)
    el = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_givenName"))); el.clear(); el.send_keys(nome)
    el = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_sn"))); el.clear(); el.send_keys(sobrenome)
    el = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_password"))); el.clear(); el.send_keys(senha_para_esta_conta)
    el = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_confirmPassword"))); el.clear(); el.send_keys(senha_para_esta_conta)

    # Obrigar troca de senha
    w.until(EC.element_to_be_clickable((By.ID, "zdlgv__NEW_ACCT_zimbraPasswordMustChange"))).click()

    # Avança (duas vezes)
    w.until(EC.element_to_be_clickable((By.ID, "zdlg__NEW_ACCT_button12_title"))).click()
    time.sleep(1)
    w.until(EC.element_to_be_clickable((By.ID, "zdlg__NEW_ACCT_button12_title"))).click()

    # Pesquisa grupo e ENTER
    campo_pesquisa = w.until(EC.visibility_of_element_located((By.ID, "zdlgv__NEW_ACCT_query")))
    campo_pesquisa.clear()
    campo_pesquisa.send_keys("corp")
    ActionChains(d).send_keys(Keys.ENTER).perform()

    # Duplo clique no grupo
    membro_email = w.until(EC.element_to_be_clickable((By.XPATH, f"//td[text()='{GRUPO_CORP}']")))
    ActionChains(d).double_click(membro_email).perform()

    # Concluir adição do membro
    w.until(EC.element_to_be_clickable((By.ID, "zdlg__NEW_ACCT_button13_title"))).click()

def indices_pendentes(df):
    # Pendentes = Situação não é "email criado" nem "erro"
    mask = ~df["Situação"].astype(str).str.lower().isin(["email criado", "erro"])
    return df.index[mask].tolist()

# ============ EXECUÇÃO ============
# ============ EXECUÇÃO ============
df = carregar_planilha()

# 1) Verifica se há pendentes antes de abrir o navegador
if not indices_pendentes(df):
    print("Nenhum registro pendente para criar. Processo concluído.")
else:
    driver = None
    wait = None

    try:
        while True:
            pend = indices_pendentes(df)
            if not pend:
                print("Nenhum registro pendente para criar. Processo concluído.")
                break

            idx = pend[0]
            row = df.loc[idx]

            # (re)abre navegador e faz login se necessário
            if driver is None:
                driver, wait = start_driver()
                login(driver, wait)

            # 2) Tenta abrir o formulário "Novo"
            try:
                abrir_form_novo(driver, wait)
            except Exception:
                # Falha ao encontrar/abrir "Novo": NÃO marcar ERRO, apenas reiniciar e tentar de novo
                try:
                    driver.quit()
                except:
                    pass
                driver, wait = None, None
                continue  # volta ao while com a MESMA linha (idx)

            # 3) Tenta criar o e-mail para a linha atual
            try:
                criar_email_para_linha(driver, wait, row)
                df.loc[idx, "Situação"] = "Email criado"
                salvar_planilha(df)
                # Sucesso: segue para a próxima linha (a UI já volta para a área onde chamaremos abrir_form_novo)
                continue

            except Exception as e:
                # Erro na criação (qualquer outro que não seja "abrir NOVO"): marca ERRO
                print(f"Erro na linha {idx}: {e}")
                df.loc[idx, "Situação"] = "ERRO"
                salvar_planilha(df)

                # Reinicia sessão para limpar estado e seguir para o próximo pendente
                try:
                    driver.quit()
                except:
                    pass
                driver, wait = None, None
                # continua o loop; próxima iteração pegará o próximo pendente

    finally:
        try:
            if driver is not None:
                driver.quit()
        except:
            pass

    print("Processo finalizado. 'Email criado' e 'ERRO' registrados na planilha.")
