import os
import re
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class ExtratorQAcademico:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        print("[SISTEMA] Abrindo navegador Chrome...")
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()

    def extrair_com_observacao(self):
        try:
            print("\n[INFO] Lendo dados da tela atual...")
            
            # Identifica a avaliação para nomear o arquivo
            try:
                nome_eval = self.driver.find_element(By.XPATH, "//td[contains(text(), 'Avaliação:')]/following-sibling::td").text.strip()
            except:
                nome_eval = "Avaliacao_QAcademico"

            wait = WebDriverWait(self.driver, 10)
            tabela = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "conteudoTexto")))
            
            linhas = tabela.find_elements(By.TAG_NAME, "tr")[1:] 
            lista_dados = []
            
            for linha in linhas:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if len(colunas) >= 7:
                    matricula = colunas[1].text.strip()
                    nome_aluno = colunas[2].text.strip()
                    
                    input_nota = colunas[5].find_element(By.TAG_NAME, "input")
                    id_nota = input_nota.get_attribute("name")
                    valor_nota = input_nota.get_attribute("value")
                    
                    input_obs = colunas[6].find_element(By.TAG_NAME, "input")
                    id_obs = input_obs.get_attribute("name")
                    valor_obs = input_obs.get_attribute("value")
                    
                    lista_dados.append({
                        "Matrícula": matricula,
                        "Aluno": nome_aluno,
                        "Nota": valor_nota,
                        "Observação": valor_obs,
                        "ID_Nota_Interno": id_nota,
                        "ID_Obs_Interno": id_obs
                    })

            df = pd.DataFrame(lista_dados)
            nome_arquivo = f"Extração_{re.sub(r'[^A-Za-z0-9]+', '_', nome_eval)}.xlsx"
            df.to_excel(nome_arquivo, index=False)
            
            print(f"[OK] Extração concluída! {len(df)} alunos processados.")
            return nome_arquivo
        except Exception as e:
            print(f"[ERRO NA EXTRAÇÃO] Não foi possível ler a tabela: {e}")
            return None

    def importar_notas_do_excel(self, caminho_arquivo):
        try:
            print(f"[INFO] Importando notas de: {caminho_arquivo}")
            df = pd.read_excel(caminho_arquivo)
            sucessos = 0
            
            for _, row in df.iterrows():
                try:
                    # Regra: Separador de vírgula e máximo 10
                    if pd.isna(row['Nota']):
                        nota_formatada = ""
                    else:
                        valor = float(str(row['Nota']).replace(',', '.'))
                        if valor > 10:
                            print(f"  [!] Nota {valor} de {row['Aluno']} ignorada (maior que 10).")
                            continue
                        nota_formatada = str(valor).replace('.', ',')
                    
                    # Preenche no site
                    self.driver.find_element(By.NAME, row['ID_Nota_Interno']).clear()
                    self.driver.find_element(By.NAME, row['ID_Nota_Interno']).send_keys(nota_formatada)

                    if not pd.isna(row['Observação']):
                        self.driver.find_element(By.NAME, row['ID_Obs_Interno']).clear()
                        self.driver.find_element(By.NAME, row['ID_Obs_Interno']).send_keys(str(row['Observação']))
                    
                    sucessos += 1
                except Exception:
                    continue # Se falhar um aluno, tenta o próximo

            print(f"[OK] {sucessos} notas inseridas com sucesso.")
            return True
        except Exception as e:
            print(f"[ERRO NA IMPORTAÇÃO] Falha ao ler arquivo ou preencher site: {e}")
            return False

# ============================================================
# LOOP DE FUNCIONAMENTO
# ============================================================
if __name__ == "__main__":
    bot = ExtratorQAcademico()
    
    try:
        # Abre o site uma única vez
        bot.driver.get("https://academico.ifes.edu.br/qacademico/index.asp?t=3061")
        
        while True:
            print("\n" + "="*60)
            print("NOVO CICLO DE LANÇAMENTO")
            print("="*60)
            print("1. No navegador, vá até a tela de 'Lançar Notas' da turma desejada.")
            input("2. Quando a lista de alunos estiver visível, aperte ENTER aqui...")

            try:
                # --- PASSO 1: EXTRAIR ---
                arquivo = bot.extrair_com_observacao()

                if arquivo:
                    print(f"\n[EDITAR] Abra o arquivo '{arquivo}'")
                    print("Preencha as notas, SALVE e FECHE o Excel.")
                    input("Após fechar o Excel, aperte ENTER para enviar as notas ao site...")

                    # --- PASSO 2: IMPORTAR ---
                    if os.path.exists(arquivo):
                        bot.importar_notas_do_excel(arquivo)
                        print("\n[AVISO] Notas inseridas! Lembre-se de clicar em 'SALVAR' no site.")
                
            except Exception as e:
                print(f"\n[ALERTA] Algo deu errado neste ciclo: {e}")
                print("O programa continuará aberto para você tentar novamente.")

            # --- PASSO 3: PERGUNTAR SE CONTINUA ---
            resposta = input("\n[?] Deseja processar outra turma/lista? (s/n): ").strip().lower()
            if resposta != 's':
                break

        print("\n[FINALIZADO] Automação concluída com sucesso.")

    except KeyboardInterrupt:
        print("\n[SISTEMA] Execução interrompida pelo usuário.")
    except Exception as e:
        print(f"\n[ERRO CRÍTICO] Ocorreu uma falha grave: {e}")
    finally:
        # Mantém o navegador aberto se o usuário quiser conferir
        finalizar = input("\nDeseja fechar o navegador agora? (s/n): ").strip().lower()
        if finalizar == 's':
            bot.driver.quit()