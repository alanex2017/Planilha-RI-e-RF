import gdown
import pandas as pd
import os
import time
import sys

# === CAPTURA DE PARÂMETROS DO VBA ===
if len(sys.argv) < 3:
    print("Erro: são esperados dois argumentos: caminho_planilha e pasta_destino.")
    input("Pressione ENTER para sair.")
    sys.exit()

caminho_planilha = sys.argv[1]
pasta_destino = sys.argv[2]

# Criar pasta de destino, se não existir
try:
    os.makedirs(pasta_destino, exist_ok=True)
except Exception as e:
    print(f"Erro ao criar a pasta de destino: {e}")
    input("Pressione ENTER para sair...1")
    sys.exit()

# === FUNÇÃO PARA OBTER ID DO LINK GOOGLE DRIVE ===
def extrair_id_google_drive(link):
    if pd.isna(link):
        return None
    link = str(link).strip()
    if "drive.google.com/open?id=" in link:
        return link.split("id=")[-1]
    elif "drive.google.com/file/d/" in link:
        return link.split("/d/")[1].split("/")[0]
    else:
        return None

# === LEITURA DA PLANILHA ===
try:
    df = pd.read_excel(caminho_planilha, sheet_name='RI-FOTOS')
except PermissionError:
    print("Erro: o arquivo está aberto no Excel. Feche-o e tente novamente.")
    input("Pressione ENTER para sair...2")
    sys.exit()
except Exception as e:
    print(f"Erro ao abrir a planilha: {e}")
    input("Pressione ENTER para sair...3")
    sys.exit()

# === LOOP PARA BAIXAR AS IMAGENS ===
contador = 0
total_linhas = len(df)

# Abre arquivo de log
log_path = os.path.join(pasta_destino, "log_download.txt")
try:
    log = open(log_path, "w", encoding="utf-8")
except Exception as e:
    print(f"Erro ao criar o log de download: {e}")
    input("Pressione ENTER para sair...4")
    sys.exit()

for idx, row in df.iterrows():
    link = row.get("LF")
    nf = row.get("NF")
    file_id = extrair_id_google_drive(link)

    if not file_id:
        msg = f"[IGNORADO] Linha {idx+2}: link inválido ou ausente -> {link}"
        print(msg)
        log.write(msg + "\n")
        continue

    url = f"https://drive.google.com/uc?id={file_id}"
    nome_arquivo = os.path.join(pasta_destino, f"{nf}.jpg")

    # Verifica se o arquivo já existe
    if os.path.exists(nome_arquivo):
        msg = f"[IGNORADO] Linha {idx+2}: {nf}.jpg já existe. Pulando download."
        print(msg)
        log.write(msg + "\n")
        continue

    # Tenta até 3 vezes
    for tentativa in range(3):
        try:
            time.sleep(1)  # Pausa antes de cada tentativa
            gdown.download(url, nome_arquivo, quiet=False)
            msg = f"[OK] {nf} salvo em {nome_arquivo}"
            print(msg)
            log.write(msg + "\n")
            contador += 1
            break  # sucesso, sai do loop de tentativas

        except Exception as e:
            msg = f"[ERRO] Linha {idx+2} - tentativa {tentativa+1} - {url}: {e}"
            print(msg)
            log.write(msg + "\n")

            # Se for erro de conexão anulada localmente, tenta novamente após pausa
            if "10053" in str(e) and tentativa < 2:
                print("⚠ Conexão anulada. Aguardando 60 segundos antes de tentar novamente...")
                time.sleep(60)
            else:
                break  # outro erro ou última tentativa

    # Pausa de 30 minutos a cada 30 downloads válidos
    if contador > 0 and contador % 30 == 0 and contador < total_linhas:
        pausa_segundos = 1800
        print(f"\n⚡ Foram baixados {contador} arquivos. Pausando 30 minutos para evitar bloqueio...\n")
        while pausa_segundos > 0:
            minutos = pausa_segundos // 60
            segundos = pausa_segundos % 60
            sys.stdout.write(f"\r⏳ Retomando em: {minutos:02d}:{segundos:02d}")
            sys.stdout.flush()
            time.sleep(1)
            pausa_segundos -= 1
        print("\n✅ Retomando downloads...\n")

log.close()

