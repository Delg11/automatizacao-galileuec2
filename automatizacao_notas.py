"""
Sistema de Automação de Preenchimento de Notas - Galileu EC2
===========================================================

Este programa automatiza o processo de preenchimento de notas no sistema Galileu EC2.
Ele extrai a planilha de notas atual, permite edição offline e depois preenche
automaticamente todas as notas no sistema.

Versão: 2.1 - Compatível com Windows (Sem Warnings)
Autor: Gabriel Delgado
Data: 2025
"""

import os
import re
import sys
import time
import warnings
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, 
    InvalidElementStateException, 
    ElementNotInteractableException,
    TimeoutException
)
from openpyxl import Workbook
from pathlib import Path

# Suprimir warnings desnecessários
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning)
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'  # Suprimir logs do TensorFlow


class AutomacaoNotasGalileu:
    """
    Classe principal para automação do sistema de notas Galileu EC2
    """
    
    def __init__(self):
        """Inicializa o sistema de automação"""
        self.driver = None
        self.df_usuario = None
        self.df_interno = None
        self.nome_arquivo_excel = None
        self.configuracao_curso = {
            'FUND2': '3533',
            'MEDIO': '3532'
        }
        
    def inicializar_navegador(self):
        """
        Inicializa o navegador Chrome com configurações para reduzir warnings
        Returns:
            bool: True se inicializado com sucesso, False caso contrário
        """
        try:
            print("[INFO] Iniciando navegador Chrome...")
            
            # Configurações do Chrome para reduzir logs/warnings
            chrome_options = Options()
            chrome_options.add_experimental_option('useAutomationExtension', False)
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_argument('--disable-extensions')
            chrome_options.add_argument('--disable-plugins')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-logging')
            chrome_options.add_argument('--disable-log-messages')
            chrome_options.add_argument('--silent')
            chrome_options.add_argument('--log-level=3')
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Suprimir logs do ChromeDriver
            chrome_options.add_argument('--disable-web-security')
            chrome_options.add_argument('--allow-running-insecure-content')
            
            # Inicializar o driver com as opções configuradas
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()
            
            print("[OK] Navegador iniciado com sucesso!")
            return True
        except Exception as e:
            print(f"[ERRO] Erro ao inicializar navegador: {e}")
            return False
    
    def fazer_login(self, usuario=None, senha=None):
        """
        Realiza login no sistema Galileu EC2
        Args:
            usuario (str): Nome de usuário (se None, solicita input)
            senha (str): Senha (se None, solicita input)
        Returns:
            bool: True se login realizado com sucesso
        """
        try:
            print("[INFO] Acessando o sistema Galileu EC2...")
            self.driver.get("https://ec2galileu.com.br/professor")
            
            # Aguarda a página carregar
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "identity"))
            )
            
            # Solicita credenciais se não fornecidas
            if not usuario:
                usuario = input("[INPUT] Digite seu usuário: ")
            if not senha:
                import getpass
                senha = getpass.getpass("[INPUT] Digite sua senha: ")
            
            # Preenche credenciais
            user_input = self.driver.find_element(By.ID, "identity")
            user_input.clear()
            user_input.send_keys(usuario)
            
            password_input = self.driver.find_element(By.ID, "credential")
            password_input.clear()
            password_input.send_keys(senha)
            
            # Clica em entrar
            entrar_button = self.driver.find_element(By.XPATH, "//button[contains(.,' Entrar')]")
            entrar_button.click()
            
            # Verifica se login foi bem-sucedido aguardando redirecionamento
            time.sleep(3)
            if "login" not in self.driver.current_url.lower():
                print("[OK] Login realizado com sucesso!")
                return True
            else:
                print("[ERRO] Falha no login - verifique suas credenciais")
                return False
                
        except Exception as e:
            print(f"[ERRO] Erro durante login: {e}")
            return False
    
    def acessar_registro_notas(self):
        """
        Navega para a página de registro de notas
        Returns:
            bool: True se acessado com sucesso
        """
        try:
            print("[INFO] Acessando registro de notas...")
            self.driver.get("https://ec2galileu.com.br/professor/registro-nota")
            
            # Aguarda página carregar
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "id_curso"))
            )
            print("[OK] Página de registro de notas carregada!")
            return True
            
        except Exception as e:
            print(f"[ERRO] Erro ao acessar registro de notas: {e}")
            return False
    
    def configurar_filtros_interface_amigavel(self):
        """
        Interface amigável para configurar curso, turma e período
        Returns:
            bool: True se configuração realizada com sucesso
        """
        try:
            print("\n" + "="*60)
            print("CONFIGURACAO DE FILTROS")
            print("="*60)
            
            # 1. Selecionar Curso
            print("\nSelecione o curso:")
            print("1. Ensino Fundamental II")
            print("2. Ensino Médio")
            
            while True:
                escolha_curso = input("\nDigite 1 ou 2: ").strip()
                if escolha_curso == "1":
                    curso_id = self.configuracao_curso['FUND2']
                    print("[OK] Selecionado: Ensino Fundamental II")
                    break
                elif escolha_curso == "2":
                    curso_id = self.configuracao_curso['MEDIO']
                    print("[OK] Selecionado: Ensino Médio")
                    break
                else:
                    print("[ERRO] Opção inválida. Digite 1 ou 2.")
            
            # Selecionar curso no sistema
            Select(self.driver.find_element(By.ID, "id_curso")).select_by_value(curso_id)
            time.sleep(3)  # Aguarda carregamento das turmas
            
            # 2. Listar e selecionar turmas disponíveis
            print("\nTurmas disponíveis:")
            select_turma = Select(self.driver.find_element(By.ID, "id_turma"))
            turmas_disponiveis = []
            
            for i, option in enumerate(select_turma.options[1:], 1):  # Pula a primeira opção vazia
                turma_texto = option.text.strip()
                turmas_disponiveis.append((option.get_attribute('value'), turma_texto))
                print(f"{i}. {turma_texto}")
            
            while True:
                try:
                    escolha_turma = int(input(f"\nDigite o número da turma (1-{len(turmas_disponiveis)}): "))
                    if 1 <= escolha_turma <= len(turmas_disponiveis):
                        turma_selecionada = turmas_disponiveis[escolha_turma - 1]
                        select_turma.select_by_value(turma_selecionada[0])
                        print(f"[OK] Selecionada: {turma_selecionada[1]}")
                        break
                    else:
                        print(f"[ERRO] Número inválido. Digite entre 1 e {len(turmas_disponiveis)}.")
                except ValueError:
                    print("[ERRO] Digite apenas números.")
            
            time.sleep(3)  # Aguarda carregamento
            
            # 3. Selecionar período
            print("\nSelecione o período:")
            print("1. 1º Trimestre")
            print("2. 2º Trimestre") 
            print("3. 3º Trimestre")
            
            while True:
                escolha_periodo = input("\nDigite 1, 2 ou 3: ").strip()
                if escolha_periodo in ['1', '2', '3']:
                    Select(self.driver.find_element(By.ID, "nr_periodo")).select_by_value(escolha_periodo)
                    print(f"[OK] Selecionado: {escolha_periodo}º Trimestre")
                    break
                else:
                    print("[ERRO] Opção inválida. Digite 1, 2 ou 3.")
            
            time.sleep(3)  # Aguarda carregamento final
            print("\n[OK] Configuração concluída com sucesso!")
            return True
            
        except Exception as e:
            print(f"[ERRO] Erro durante configuração: {e}")
            return False
    
    def extrair_dados_tabela(self):
        """
        Extrai dados da tabela de alunos e cria planilha Excel
        Returns:
            bool: True se extração realizada com sucesso
        """
        try:
            print("\n" + "="*60)
            print("EXTRAINDO DADOS DA TABELA")
            print("="*60)
            
            # Aguarda tabela carregar
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.ID, "gridAlunos"))
            )
            
            tabela = self.driver.find_element(By.ID, "gridAlunos")
            linhas = tabela.find_elements(By.TAG_NAME, "tr")
            
            dados_tabela = []
            dados_tabela_interna = []
            
            print(f"[INFO] Processando {len(linhas)} linhas da tabela...")
            
            for i, linha in enumerate(linhas):
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if colunas:
                    # Nome do aluno (primeira coluna, primeira linha do texto)
                    nome_aluno = colunas[0].text.strip().split("\n")[0]
                    dados_linha = [nome_aluno]
                    dados_linha_interna = [nome_aluno]
                    
                    # Extrai inputs das demais colunas
                    for coluna in colunas[1:]:
                        inputs = coluna.find_elements(By.CSS_SELECTOR, "input")
                        if inputs:
                            input_element = inputs[0]
                            input_value = input_element.get_attribute("value")
                            input_name = input_element.get_attribute("name")
                            dados_linha.append(input_value)
                            dados_linha_interna.append(input_name)
                    
                    dados_tabela.append(dados_linha)
                    dados_tabela_interna.append(dados_linha_interna)
                    
                    if i % 5 == 0:  # Progresso a cada 5 linhas
                        print(f"  [PROG] Processando linha {i+1}...")
            
            # Nomes das colunas
            nomes_colunas = [
                "Aluno", "VERIFICACAO PARCIAL", "VERIFICACAO GLOBAL", 
                "ATIVIDADE 1", "ATIVIDADE 2", "ATIVIDADE 3", "ATIVIDADE 4", 
                "PONTO OLIMPIADA", "MEDIA MANUAL"
            ]
            
            # Criar DataFrames
            max_cols = max(len(linha) for linha in dados_tabela)
            colunas_usadas = nomes_colunas[:max_cols]
            
            self.df_usuario = pd.DataFrame(dados_tabela, columns=colunas_usadas)
            self.df_interno = pd.DataFrame(dados_tabela_interna, columns=colunas_usadas)
            
            # Gerar nome do arquivo baseado na turma selecionada
            select_turma = self.driver.find_element(By.ID, "id_turma")
            turma_selecionada = select_turma.find_element(By.CSS_SELECTOR, "option:checked").text.strip()
            
            # Limpar nome para usar como arquivo
            nome_base = re.sub(r'\d{2}/\d{2}/\d{4} a \d{2}/\d{2}/\d{4}', '', turma_selecionada)
            nome_base = re.sub(r'[^A-Za-z0-9\s]+', '', nome_base)
            nome_base = re.sub(r'\s+', '_', nome_base.strip())
            
            self.nome_arquivo_excel = f"{nome_base}_Notas_Para_Edicao.xlsx"
            
            # Salvar arquivo Excel
            caminho_completo = os.path.join(os.getcwd(), self.nome_arquivo_excel)
            self.df_usuario.to_excel(caminho_completo, index=False)
            
            print(f"\n[OK] Dados extraídos com sucesso!")
            print(f"[INFO] Arquivo criado: {self.nome_arquivo_excel}")
            print(f"[INFO] Localização: {caminho_completo}")
            print(f"[INFO] Total de alunos: {len(self.df_usuario)}")
            print(f"[INFO] Colunas de notas: {len(self.df_usuario.columns) - 1}")
            
            return True
            
        except Exception as e:
            print(f"[ERRO] Erro ao extrair dados: {e}")
            return False
    
    def aguardar_edicao_planilha(self):
        """
        Pausa o programa para o usuário editar a planilha
        Returns:
            bool: True quando usuário confirmar que editou
        """
        print("\n" + "="*60)
        print("EDICAO DA PLANILHA")
        print("="*60)
        print(f"[INFO] Abra o arquivo: {self.nome_arquivo_excel}")
        print("[INFO] Edite as notas dos alunos conforme necessário")
        print("[INFO] Salve o arquivo após as alterações")
        print("\nINSTRUCOES IMPORTANTES:")
        print("   • Para alunos que faltaram, deixe a célula vazia (não digite nada)")
        print("   • Para atividades que não existiram, escreva 'N/C' na célula")
        print("   • Use vírgula (,) como separador decimal: 7,5 ao invés de 7.5")
        print("   • Não altere os nomes dos alunos")
        print("\nIMPORTANTE: Feche o Excel completamente antes de continuar!")
        
        while True:
            # Limpar o buffer do input para evitar caracteres estranhos
            sys.stdout.flush()
            
            try:
                resposta = input("\n[INPUT] Você terminou de editar e salvou a planilha? (s/n): ").strip().lower()
                if resposta in ['s', 'sim', 'y', 'yes']:
                    print("[OK] Continuando com o preenchimento automático...")
                    return True
                elif resposta in ['n', 'não', 'nao', 'no']:
                    print("[INFO] Aguardando... Termine a edição e tente novamente.")
                else:
                    print("[ERRO] Resposta inválida. Digite 's' para sim ou 'n' para não.")
            except KeyboardInterrupt:
                print("\n[INFO] Operação cancelada pelo usuário.")
                return False
            except Exception as e:
                print(f"[ERRO] Erro no input: {e}")
                continue
    
    def carregar_notas_editadas(self):
        """
        Carrega a planilha editada pelo usuário
        Returns:
            bool: True se carregamento realizado com sucesso
        """
        try:
            print("\n" + "="*60)
            print("CARREGANDO PLANILHA EDITADA")
            print("="*60)
            
            caminho_excel = os.path.join(os.getcwd(), self.nome_arquivo_excel)
            
            if not os.path.exists(caminho_excel):
                print(f"[ERRO] Arquivo não encontrado: {caminho_excel}")
                return False
            
            print(f"[INFO] Carregando: {self.nome_arquivo_excel}")
            
            # Carregar planilha suprimindo warnings do pandas
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                notas = pd.read_excel(caminho_excel, sheet_name=0, header=None)
            
            # Converter floats para string com vírgula
            total_celulas = 0
            celulas_convertidas = 0
            
            for i in range(len(notas)):
                for j in range(1, min(8, len(notas.columns))):
                    total_celulas += 1
                    try:
                        valor = notas.iloc[i, j]
                        if isinstance(valor, float) and not pd.isna(valor):
                            notas.iloc[i, j] = str(valor).replace('.', ',')
                            celulas_convertidas += 1
                    except IndexError:
                        continue
            
            # Processar valores especiais
            notas = notas.replace("nan", np.nan)
            
            # Transformar primeira linha em cabeçalho
            notas.columns = notas.iloc[0]
            self.notas = notas[1:]
            
            print(f"[OK] Planilha carregada com sucesso!")
            print(f"[INFO] Dados processados: {self.notas.shape[0]} alunos, {self.notas.shape[1]} colunas")
            print(f"[INFO] Números convertidos: {celulas_convertidas}/{total_celulas} células")
            
            return True
            
        except Exception as e:
            print(f"[ERRO] Erro ao carregar planilha: {e}")
            return False
    
    def preencher_notas_automaticamente(self):
        """
        Preenche automaticamente as notas no sistema
        Returns:
            bool: True se preenchimento realizado com sucesso
        """
        try:
            print("\n" + "="*60)
            print("PREENCHIMENTO AUTOMATICO DAS NOTAS")
            print("="*60)
            
            num_alunos = len(self.df_interno)
            total_campos = num_alunos * (len(self.df_interno.columns) - 1)
            campos_preenchidos = 0
            campos_com_checkbox = 0
            erros = 0
            
            print(f"[INFO] Processando {num_alunos} alunos...")
            print(f"[INFO] Total de campos estimados: {total_campos}")
            print("\n[INFO] Iniciando preenchimento...")
            
            # Loop principal de preenchimento
            for i in range(num_alunos):
                nome_aluno = self.df_interno.iloc[i, 0]
                print(f"\n[ALUNO] Processando: {nome_aluno}")
                
                # Percorrer colunas de notas
                for j in range(1, len(self.df_interno.columns)):
                    id_campo = self.df_interno.iloc[i, j]
                    
                    # Pular campos de média manual
                    if "media-manual" in str(id_campo).lower():
                        continue
                    
                    try:
                        # Verificar se há nota para este campo
                        if not pd.isna(self.notas.iloc[i, j]):
                            nota = str(self.notas.iloc[i, j]).strip()
                            
                            # Caso especial: N/C (Não Contabilizar)
                            if nota.upper() == "N/C":
                                if self._marcar_checkbox_nc(id_campo):
                                    campos_com_checkbox += 1
                                    print(f"   [NC] N/C marcado")
                            else:
                                # Preencher nota normal
                                if self._preencher_campo_nota(id_campo, nota):
                                    campos_preenchidos += 1
                                    print(f"   [NOTA] Nota: {nota}")
                        else:
                            # Célula vazia = marcar como N/C (Não Contabilizar)
                            if self._marcar_checkbox_nc(id_campo):
                                campos_com_checkbox += 1
                                print(f"   [Vazio] N/C marcada")
                                
                    except Exception as e:
                        erros += 1
                        print(f"   [ERRO] Erro no campo {j}: {e}")
                
                # Mostrar progresso
                progresso = ((i + 1) / num_alunos) * 100
                print(f"   [PROG] Progresso: {progresso:.1f}%")
            
            # Relatório final
            print("\n" + "="*60)
            print("RELATORIO FINAL")
            print("="*60)
            print(f"[OK] Campos preenchidos com notas: {campos_preenchidos}")
            print(f"[OK] Campos marcados como N/C ou falta: {campos_com_checkbox}")
            print(f"[ERRO] Erros encontrados: {erros}")
            print(f"[INFO] Total processado: {campos_preenchidos + campos_com_checkbox + erros}")
            
            if erros == 0:
                print("\n[SUCESSO] Preenchimento concluído com SUCESSO! Todas as notas foram inseridas.")
            else:
                print(f"\n[AVISO] Preenchimento concluído com {erros} erros. Verifique os campos manualmente.")
            
            return True
            
        except Exception as e:
            print(f"[ERRO] Erro durante preenchimento automático: {e}")
            return False
    
    def _preencher_campo_nota(self, id_campo, nota):
        """
        Preenche um campo específico com uma nota
        Args:
            id_campo (str): ID do campo de input
            nota (str): Nota a ser inserida
        Returns:
            bool: True se preenchido com sucesso
        """
        try:
            wait = WebDriverWait(self.driver, 5)
            campo = wait.until(EC.presence_of_element_located((By.ID, id_campo)))
            
            if campo.is_enabled() and campo.is_displayed():
                # Tentar método normal primeiro
                try:
                    campo.clear()
                    campo.send_keys(nota)
                    return True
                except InvalidElementStateException:
                    # Fallback: usar JavaScript
                    self.driver.execute_script(f"arguments[0].value = '{nota}';", campo)
                    return True
            
            return False
            
        except (NoSuchElementException, TimeoutException):
            return False
        except Exception:
            return False
    
    def _marcar_checkbox_nc(self, id_campo):
        """
        Marca checkbox de 'Não Compareceu' para um campo
        Args:
            id_campo (str): ID do campo base
        Returns:
            bool: True se marcado com sucesso
        """
        try:
            id_checkbox = f"chk-nc-{id_campo.lower()}"
            wait = WebDriverWait(self.driver, 5)
            checkbox = wait.until(EC.presence_of_element_located((By.ID, id_checkbox)))
            
            if checkbox.is_enabled() and checkbox.is_displayed():
                if not checkbox.is_selected():
                    checkbox.click()
                return True
            
            return False
            
        except (NoSuchElementException, TimeoutException):
            return False
        except Exception:
            return False
    
    def executar_processo_completo(self):
        """
        Executa o processo completo de automação
        Returns:
            bool: True se todo o processo foi executado com sucesso
        """
        print("\n" + "="*60)
        print("   SISTEMA DE AUTOMACAO DE NOTAS GALILEU EC2")
        print("="*60)
        print("\nVersao 2.1 - Sistema Otimizado (Sem Warnings)")
        print("Este programa ira:")
        print("1. Fazer login no sistema")
        print("2. Configurar filtros (curso, turma, periodo)")
        print("3. Extrair dados atuais em planilha Excel")
        print("4. Aguardar sua edicao da planilha")
        print("5. Preencher automaticamente todas as notas")
        print("\n" + "="*60)
        
        # Etapa 1: Inicializar navegador
        if not self.inicializar_navegador():
            return False
        
        # Etapa 2: Fazer login
        if not self.fazer_login():
            return False
        
        # Etapa 3: Acessar registro de notas
        if not self.acessar_registro_notas():
            return False
        
        # Etapa 4: Configurar filtros
        if not self.configurar_filtros_interface_amigavel():
            return False
        
        # Etapa 5: Extrair dados
        if not self.extrair_dados_tabela():
            return False
        
        # Etapa 6: Aguardar edição
        if not self.aguardar_edicao_planilha():
            return False
        
        # Etapa 7: Carregar dados editados
        if not self.carregar_notas_editadas():
            return False
        
        # Etapa 8: Preencher automaticamente
        if not self.preencher_notas_automaticamente():
            return False
        
        print("\n[SUCESSO] PROCESSO CONCLUIDO COM SUCESSO!")
        print("\nDicas finais:")
        print("   • Revise as notas inseridas antes de finalizar")
        print("   • Salve/submeta as alteracoes no sistema")
        print("   • Mantenha backup da planilha Excel gerada")
        
        return True
    
    def finalizar(self):
        """Finaliza o programa e fecha o navegador"""
        if self.driver:
            try:
                escolha = input("\n[INPUT] Deseja fechar o navegador automaticamente? (s/n): ").strip().lower()
                if escolha in ['s', 'sim', 'y', 'yes']:
                    self.driver.quit()
                    print("[OK] Navegador fechado com sucesso!")
                else:
                    print("[INFO] Navegador mantido aberto para revisao manual.")
            except:
                # Em caso de erro no input, apenas fecha o navegador
                self.driver.quit()
                print("[OK] Navegador fechado automaticamente.")


def main():
    """Função principal do programa"""
    # Suprimir outputs desnecessários do sistema
    sys.stderr = open(os.devnull, 'w') if os.name == 'nt' else sys.stderr
    
    sistema = AutomacaoNotasGalileu()
    
    try:
        sucesso = sistema.executar_processo_completo()
        
        if sucesso:
            print("\n[OK] Todos os processos foram executados com sucesso!")
        else:
            print("\n[ERRO] Houve problemas durante a execução.")
            
    except KeyboardInterrupt:
        print("\n\n[INFO] Programa interrompido pelo usuário.")
    except Exception as e:
        print(f"\n\n[ERRO] Erro inesperado: {e}")
    finally:
        sistema.finalizar()
        
        # Restaurar stderr se foi redirecionado
        if os.name == 'nt' and hasattr(sys.stderr, 'close'):
            sys.stderr.close()
            sys.stderr = sys.__stderr__


if __name__ == "__main__":
    main()