import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from datetime import date

# =================================================================
#                         CONFIGURAÇÕES
# =================================================================

DRIVER_PATH = r"C:\Users\AV\Desktop\Automação SAR\chromedriver.exe"
EXCEL_PATH = r"C:\Users\AV\Desktop\Automação SAR\Vendas_Do_Dia.xlsx"
NOME_GRUPO = "Fora de Rota"

# Diretório para salvar a sessão do WhatsApp (evita fazer login toda vez)
USER_DATA_DIR = r"C:\Users\AV\Desktop\Automação SAR\ChromeProfile"

# Regex para capturar números com 5 ou mais dígitos
REGEX_CODIGOS = r"\b(\d{5,})\b"

# Configurações de rolagem INFINITA até pegar tudo
MAX_TENTATIVAS_SEM_NOVOS = 8  # Para após 8 rolagens sem códigos novos
TEMPO_ENTRE_ROLAGENS = 3
SCROLL_DISTANCE = 3  # Quantos Page Ups por iteração

# =================================================================
#                         FUNÇÕES
# =================================================================

def extrair_codigos_com_contato(driver):
    """Extrai códigos E os nomes dos contatos que enviaram"""
    codigos_com_contato = {}  # {codigo: contato}
    
    try:
        # Pega o texto completo da conversa
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        texto_completo = main.text
        
        # Divide em linhas para processar
        linhas = texto_completo.split('\n')
        
        contato_atual = "Desconhecido"
        
        for i, linha in enumerate(linhas):
            linha = linha.strip()
            
            # Detecta linha de contato (formato: "123 - Nome" ou "Nome")
            # Exemplos: "803 - Átila", "151 - - Álvaro", "201 - - Daniele"
            if re.match(r'^\d{3,4}\s*-\s*-?\s*.+', linha):
                # Extrai o nome após os números e traços
                match = re.search(r'^\d{3,4}\s*-\s*-?\s*(.+)', linha)
                if match:
                    contato_atual = match.group(1).strip()
            
            # Procura códigos na linha atual
            codigos_na_linha = re.findall(REGEX_CODIGOS, linha)
            
            for codigo in codigos_na_linha:
                if 5 <= len(codigo) <= 10:
                    # Se o código já existe, mantém o primeiro contato encontrado
                    if codigo not in codigos_com_contato:
                        codigos_com_contato[codigo] = contato_atual
        
    except Exception as e:
        print(f"        ⚠ Erro na extração com contato: {str(e)[:50]}")
    
    return codigos_com_contato

def salvar_debug(driver):
    """Salva conteúdo completo para debug"""
    try:
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        with open('debug_extracao.txt', 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("DEBUG - CONTEÚDO COMPLETO DA CONVERSA\n")
            f.write("=" * 80 + "\n\n")
            f.write(main.text)
        return True
    except:
        return False

# =================================================================
#                         FUNÇÃO PRINCIPAL
# =================================================================

def iniciar_automacao():
    service = Service(DRIVER_PATH)
    
    # Configura o Chrome para manter a sessão (não precisa fazer login toda vez)
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    
    driver = webdriver.Chrome(service=service, options=options)
    
    try:
        print("=" * 80)
        print("🤖 AUTOMAÇÃO SAR - EXTRAÇÃO COMPLETA DE CÓDIGOS DO WHATSAPP")
        print("=" * 80)
        
        # PASSO 1: ABRE WHATSAPP WEB
        driver.get("https://web.whatsapp.com/")
        print("\n[1/7] 📱 Conectando ao WhatsApp Web...")

        wait = WebDriverWait(driver, 90)
        
        # Verifica se já está logado ou se precisa escanear QR Code
        try:
            # Tenta encontrar a barra de pesquisa (indica que já está logado)
            search_box = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            ))
            print("        ✅ Já está logado! (sessão salva)")
        except:
            print("        ⚠️ Primeira vez ou sessão expirada")
            print("        📱 Escaneie o QR Code no celular...")
            search_box = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            ))
            print("        ✅ Login realizado! (sessão será salva para próximas vezes)")

        # PASSO 2: BUSCA O GRUPO
        print(f"\n[2/7] 🔍 Buscando grupo '{NOME_GRUPO}'...")
        search_box.click()
        time.sleep(1)
        search_box.clear()
        search_box.send_keys(NOME_GRUPO)
        time.sleep(3)
        
        grupo = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//span[@title='{NOME_GRUPO}']")
        ))
        grupo.click()
        time.sleep(5)
        print("        ✅ Grupo aberto!")

        # PASSO 3: FOCA NA ÁREA DE MENSAGENS E VAI PARA O FINAL
        print("\n[3/7] 📜 Focando na área de mensagens do grupo...")
        
        # CRÍTICO: Clica na área de mensagens para garantir foco correto
        try:
            # Tenta clicar na área de mensagens (não na lista de conversas)
            area_mensagens = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[@id="main"]//div[@role="application"]')
            ))
            area_mensagens.click()
            time.sleep(2)
            print("        ✅ Foco na área de mensagens!")
        except:
            # Fallback: clica no meio da tela direita
            print("        ⚠️ Usando método alternativo de foco...")
            driver.execute_script("window.scrollTo(500, 500);")
            time.sleep(1)
        
        # Agora rola DENTRO da área de mensagens
        print("        Posicionando no final da conversa...")
        
        # Método 1: Tenta rolar usando JavaScript no container correto
        try:
            container = driver.find_element(By.XPATH, 
                '//div[@id="main"]//div[contains(@class, "copyable-area")]')
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", container)
            time.sleep(2)
        except:
            pass
        
        # Método 2: Usa Keys.END na área de mensagens
        for i in range(5):
            try:
                area_msgs = driver.find_element(By.XPATH, '//div[@id="main"]')
                area_msgs.send_keys(Keys.END)
            except:
                pass
            time.sleep(0.5)
        
        time.sleep(3)
        print("        ✅ Posicionado nas mensagens mais recentes!")
        
        # PASSO 4: EXTRAÇÃO INICIAL
        print("\n[4/7] 🔢 Extraindo códigos da tela inicial...")
        todos_codigos_com_contato = {}  # {codigo: contato}
        
        codigos_iniciais = extrair_codigos_com_contato(driver)
        todos_codigos_com_contato.update(codigos_iniciais)
        print(f"        → {len(codigos_iniciais)} códigos encontrados na tela inicial")
        
        # PASSO 5: SCROLL APENAS NAS MENSAGENS DE HOJE
        print(f"\n[5/7] 🔄 Iniciando rolagem apenas nas mensagens de HOJE...")
        print(f"        (Para quando encontrar mensagens de dias anteriores)\n")
        
        tentativas_sem_novos = 0
        contador_rolagem = 0
        encontrou_data_antiga = False
        
        # Pega referência da área de mensagens para rolar nela especificamente
        try:
            area_mensagens = driver.find_element(By.XPATH, '//div[@id="main"]')
        except:
            area_mensagens = driver.find_element(By.TAG_NAME, 'body')
        
        while tentativas_sem_novos < MAX_TENTATIVAS_SEM_NOVOS and not encontrou_data_antiga:
            contador_rolagem += 1
            
            # Rola para CIMA DENTRO da área de mensagens
            for _ in range(SCROLL_DISTANCE):
                try:
                    area_mensagens.send_keys(Keys.PAGE_UP)
                except:
                    # Fallback: usa JavaScript
                    driver.execute_script("""
                        var main = document.getElementById('main');
                        if (main) {
                            var scrollable = main.querySelector('[data-tab="8"]') || main;
                            scrollable.scrollTop -= 1000;
                        }
                    """)
            
            time.sleep(TEMPO_ENTRE_ROLAGENS)
            
            # VERIFICA SE SAIU DA ÁREA DE "HOJE"
            try:
                texto_visivel = driver.find_element(By.XPATH, '//div[@id="main"]').text
                
                # Procura por indicadores de data antiga (dias da semana em português)
                dias_anteriores = ['ontem', 'Ontem', 'sexta', 'Sexta', 'quinta', 'Quinta', 
                                 'quarta', 'Quarta', 'terça', 'Terça', 'segunda', 'Segunda',
                                 'sábado', 'Sábado', 'domingo', 'Domingo']
                
                # Procura por datas no formato dd/mm/yyyy ou dd/mm
                tem_data = re.search(r'\b\d{1,2}/\d{1,2}(/\d{2,4})?\b', texto_visivel)
                
                # Se encontrou dia da semana anterior OU data, para
                if any(dia in texto_visivel for dia in dias_anteriores) or tem_data:
                    print(f"\n        ⚠️  Detectadas mensagens de dias anteriores!")
                    print(f"        ⏹️  Parando rolagem para evitar capturar códigos antigos...")
                    encontrou_data_antiga = True
                    break
                    
            except:
                pass
            
            # Extrai códigos da posição atual
            codigos_atuais = extrair_todos_codigos_visiveis(driver)
            tamanho_antes = len(todos_codigos)
            todos_codigos.update(codigos_atuais)
            novos = len(todos_codigos) - tamanho_antes
            
            if novos > 0:
                print(f"        ✅ Rolagem {contador_rolagem:3d}: +{novos:2d} novos códigos | Total: {len(todos_codigos):3d}")
                tentativas_sem_novos = 0  # Reseta contador
            else:
                tentativas_sem_novos += 1
                print(f"        ⚪ Rolagem {contador_rolagem:3d}: Sem novos códigos ({tentativas_sem_novos}/{MAX_TENTATIVAS_SEM_NOVOS}) | Total: {len(todos_codigos):3d}")
            
            # Segurança: para em 100 rolagens de qualquer forma
            if contador_rolagem >= 100:
                print("\n        ⚠️  Limite de 100 rolagens atingido, finalizando...")
                break
        
        print(f"\n        ✅ Rolagem finalizada após {contador_rolagem} iterações!")
        print(f"        📊 Total de códigos únicos capturados: {len(todos_codigos_com_contato)}")
        
        if encontrou_data_antiga:
            print(f"        ℹ️  Rolagem parou ao detectar mensagens de dias anteriores (filtro ativo)")
        
        # PASSO 6: SALVA DEBUG
        print("\n[6/7] 💾 Salvando arquivo de debug...")
        if salvar_debug(driver):
            print("        ✅ Debug salvo: debug_extracao.txt")
        
        # PASSO 7: SALVA PLANILHA COM CONTATOS
        print("\n[7/7] 📊 Processando e salvando planilha...")
        
        if todos_codigos_com_contato:
            # Remove códigos inválidos conhecidos
            codigos_invalidos = ['99999', '00000', '11111', '22222', '33333', 
                               '44444', '55555', '66666', '77777', '88888', '12345']
            
            # Filtra e cria lista de tuplas (codigo, contato)
            dados_filtrados = [(cod, contato) for cod, contato in todos_codigos_com_contato.items() 
                              if cod not in codigos_invalidos]
            
            # Ordena por código numericamente
            dados_ordenados = sorted(dados_filtrados, key=lambda x: int(x[0]))
            
            # Cria DataFrame com 3 colunas
            df = pd.DataFrame(dados_ordenados, columns=['Código', 'Contato'])
            df['Data'] = date.today().strftime('%d/%m/%Y')
            
            # Reordena colunas: Código, Data, Contato
            df = df[['Código', 'Data', 'Contato']]
            
            # SOBRESCREVE completamente a planilha
            df.to_excel(EXCEL_PATH, index=False)
            
            print(f"        ✅ Planilha atualizada: {len(dados_ordenados)} códigos salvos")
            print(f"        📁 Arquivo: {EXCEL_PATH}")
            
            # RESULTADO FINAL
            print("\n" + "=" * 80)
            print("✅ AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
            print("=" * 80)
            print(f"\n📊 CÓDIGOS CAPTURADOS DE HOJE (Total: {len(dados_ordenados)}):\n")
            
            # Mostra códigos com contatos
            print("    " + "-" * 76)
            print(f"    {'#':>3}  {'Código':>8}  {'Contato':<40}")
            print("    " + "-" * 76)
            
            for i, (codigo, contato) in enumerate(dados_ordenados, 1):
                # Limita o nome do contato a 40 caracteres
                contato_curto = contato[:37] + "..." if len(contato) > 40 else contato
                print(f"    {i:3d}. {codigo:>8s}  {contato_curto:<40}")
            
            print("    " + "-" * 76)
            
            print(f"\n📅 Data da extração: {date.today().strftime('%d/%m/%Y - %A')}")
            print(f"📁 Planilha salva: Vendas_Do_Dia.xlsx")
            print(f"🔢 Total de rolagens: {contador_rolagem}")
            
            # Resultado final simplificado
            print(f"\n💡 Resultado:")
            print(f"   ✅ {len(dados_ordenados)} códigos capturados com sucesso!")
            print(f"   📋 Todos os dados foram salvos na planilha")
            
            print("=" * 80)
            
        else:
            print("\n" + "=" * 80)
            print("❌ NENHUM CÓDIGO ENCONTRADO!")
            print("=" * 80)
            print("\n🔍 Possíveis causas:")
            print("   • O grupo não tem mensagens com códigos de 5+ dígitos")
            print("   • A estrutura do WhatsApp mudou")
            print("   • Problema de conexão ou carregamento")
            print("\n💡 Verifique o arquivo 'debug_extracao.txt' para diagnóstico")

    except Exception as e:
        print("\n" + "=" * 80)
        print("❌ ERRO FATAL NA AUTOMAÇÃO")
        print("=" * 80)
        print(f"\nErro: {str(e)}\n")
        import traceback
        print("Detalhes técnicos:")
        print(traceback.format_exc())
        print("\n💡 Tente executar novamente ou verifique:")
        print("   • ChromeDriver está atualizado")
        print("   • WhatsApp Web está funcionando normalmente")
        print("   • Caminho dos arquivos está correto")

    finally:
        print("\n⏳ Fechando navegador em 10 segundos...")
        print("   (Tempo para você visualizar os resultados)")
        time.sleep(10)
        driver.quit()
        print("✅ Navegador fechado com sucesso.\n")

# =================================================================
#                         EXECUÇÃO
# =================================================================

if __name__ == "__main__":
    iniciar_automacao()