import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = 'dados_chamados.xlsx' # Certifica-te que o nome está igual ao arquivo
URL_LOGIN = 'https://chamados.somagrupo.com.br/login'
USUARIO = 'marcus.boto'  # Coloca o teu usuário aqui
SENHA = 'Soma@122025'      # Coloca a tua senha aqui

def iniciar_automacao():
    # --- 1. MENU DE ESCOLHA ---
    print("--- MENU DE OPÇÕES ---")
    print("1 - PAGAMENTOS")
    print("2 - TRANSFERÊNCIA DE VERBA")
    
    try:
        opcao = int(input("Digite o número da opção desejada: "))
    except ValueError:
        print("Erro: Digite apenas números.")
        return

    if opcao == 1:
        url_destino = "https://chamados.somagrupo.com.br/projects/compras-de-indiretos/issues/new"
        titulo_prefixo = "Pagamento Extra Feirinha Ouvidor | Novembro/2025 | "
    elif opcao == 2:
        url_destino = "https://chamados.somagrupo.com.br/projects/incremento-de-verba-capex/issues/new"
        titulo_prefixo = "TRANSF. VERBA"
    else:
        print("Opção inválida.")
        return

    # --- 2. LER O EXCEL ---
    try:
        # Lê o Excel garantindo que tudo venha como texto para evitar erros de formatação
        tabela = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
        # Removemos linhas vazias se houverem (onde a coluna Nome esteja vazia)
        tabela = tabela.dropna(subset=['Nome'])
        print(f"Excel carregado! {len(tabela)} linhas encontradas.")
    except Exception as e:
        print(f"Erro ao ler Excel: {e}")
        return

    # --- 3. INICIAR O NAVEGADOR (SEM O DRIVER MANAGER ANTIGO) ---
    print("Abrindo navegador...")
    # Apenas isto é necessário nas versões novas do Selenium:
    navegador = webdriver.Chrome() 
    navegador.maximize_window()

    # --- 4. LOGIN ---
    try:
        navegador.get(URL_LOGIN)
        time.sleep(3)
        navegador.find_element(By.ID, 'username').send_keys(USUARIO)
        navegador.find_element(By.ID, 'password').send_keys(SENHA)
        navegador.find_element(By.ID, 'login-submit').click()
        print("Login feito. Aguardando...")
        time.sleep(5)
    except Exception as e:
        print(f"Erro no Login: {e}")
        return

    # --- 5. PREENCHIMENTO DOS DADOS ---
    for index, linha in tabela.iterrows():
        print(f"Processando: {linha['Nome']}...")
        navegador.get(url_destino)
        time.sleep(3)

        try:
            # Capturando dados do Excel (baseado na tua imagem)
            nome = linha['Nome']
            valor_bruto = linha['Valor Total RPA']
            valor_liquido = linha['Valor Liquido']
            municipio = linha['Municipio']
            bairro = linha['Bairro']
            cep = linha['CEP']
            cpf = linha['CPF']
            rua = linha['Rua']
            numero = linha['N']
            complemento = linha['Complemento']
            nome_completo = linha['Nome Completo']
            naturalidade = linha['Naturalidade']
            data_nasc = linha['Data de Nascimento']
            pis = linha['PIS']
            nome_banco = linha['Nome Banco']
            agencia = linha['Agencia Banco']
            conta_bancaria = linha['Conta Bancaria']
            
            
            # PREENCHER
            
            # TÍTULO DO CHAMADO
            campo_titulo = navegador.find_element(By.ID, 'issue_subject')
            campo_titulo.send_keys(f"{titulo_prefixo} - {nome}")

            # TODO É REEMBOLSO?
            campo_titulo = navegador.find_element(By.ID, 'issue_subject')
            campo_titulo.send_keys(f"{titulo_prefixo} - {nome}")

            # TODO REMESSA FUTURA?
            campo_titulo = navegador.find_element(By.ID, 'issue_subject')
            campo_titulo.send_keys(f"{titulo_prefixo} - {nome}")

            # TODO SOLICITAÇÃO RPA?
            campo_titulo = navegador.find_element(By.ID, 'issue_subject')
            campo_titulo.send_keys(f"{titulo_prefixo} - {nome}")

            # VALOR BRUTO VALOR BRUTO
            campo_valor_total = navegador.find_element(By.ID, 'issue_custom_field_values_424')
            campo_valor_total.send_keys(f"{valor_bruto}")

            # TODO POSSUI NOTA FISCAL?
            campo_valor_liquido = navegador.find_element(By.ID, 'issue_custom_field_values_2772')
            campo_valor_liquido.send_keys(f"{valor_liquido}")

            # VALOR LÍQUIDO
            campo_valor_liquido = navegador.find_element(By.ID, 'issue_custom_field_values_2772')
            campo_valor_liquido.send_keys(f"{valor_liquido}")

            # TODO DATA VENCIMENTO NF
            campo_valor_liquido = navegador.find_element(By.ID, 'issue_custom_field_values_2772')
            campo_valor_liquido.send_keys(f"{valor_liquido}")

            # TODO DATA ENTREGA ORÇAMENTÁRIA OBS DATE +5)
            campo_valor_liquido = navegador.find_element(By.ID, 'issue_custom_field_values_2772')
            campo_valor_liquido.send_keys(f"{valor_liquido}")

            # TODO HAVERÁ RATEIO)
            campo_valor_liquido = navegador.find_element(By.ID, 'issue_custom_field_values_2772')
            campo_valor_liquido.send_keys(f"{valor_liquido}")


            print(f"--> Formuário de {nome} preenchido!")
            
            # ATENÇÃO: Botão de salvar comentado para segurança
            # navegador.find_element(By.NAME, 'commit').click() 
            
            time.sleep(2)

        except Exception as e:
            print(f"Erro ao preencher linha {index}: {e}")

    print("Processo finalizado.")

if __name__ == "__main__":
    iniciar_automacao()