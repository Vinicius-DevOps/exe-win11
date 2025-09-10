# setup_info_installer.py
import subprocess
import shutil
import tempfile
import os
import re
import requests

# ====== CONFIG: ID do arquivo no Google Drive (Due Studio) ======
DUE_STUDIO_DRIVE_FILE_ID = "1NVMlI1mzSxcFXsbxi-wcDLR69ARxzS_Y"

# ===============================================================
# Utilidades
# ===============================================================
def run(cmd, check=False):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, shell=False, encoding='utf-8', errors='ignore')
        if check and p.returncode != 0:
            raise subprocess.CalledProcessError(p.returncode, cmd, p.stdout, p.stderr)
        return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()
    except FileNotFoundError as e:
        return 127, "", str(e)

def run_ps(ps_script):
    return run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script])

def has_admin_rights():
    ps = "([bool]([Security.Principal.WindowsPrincipal]" \
          " [Security.Principal.WindowsIdentity]::GetCurrent())." \
          "IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)"
    _, out, _ = run_ps(ps)
    return out.strip().lower() == "true"

# ===============================================================
# Coleta de informações
# ===============================================================
def get_windows_key():
    ps_cmd = r"(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrada)"

def get_bios_serial():
    ps_cmd = r"(Get-WmiObject Win32_BIOS | Select-Object -ExpandProperty SerialNumber)"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrada)"

def get_device_name():
    ps_cmd = r"(Get-ComputerInfo -Property 'CsName').CsName"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrado)"

# ===============================================================
# Winget helpers
# ===============================================================
def winget_available(): return shutil.which("winget") is not None

def winget_install_or_upgrade(ids_or_queries):
    if not winget_available():
        return False, "Winget não encontrado. Instale a Microsoft Store 'App Installer' e tente novamente."
    
    for identifier in ids_or_queries:
        identifier = identifier.strip()
        if not identifier: continue
        
        # O comando 'install' do winget também atualiza o pacote se ele já estiver instalado.
        # Usamos --accept-package-agreements e --accept-source-agreements para automatizar.
        cmd = [
            "winget", "install", "--id", identifier, 
            "--silent", "--accept-package-agreements", "--accept-source-agreements"
        ]
        
        code, stdout, stderr = run(cmd)
        
        if code == 0:
            # Sucesso pode significar instalado ou já atualizado.
            return True, f"Instalado/Atualizado com sucesso: {identifier}"
        else:
            # Tenta o próximo identificador em caso de falha
            print(f"Tentativa com '{identifier}' falhou (código: {code})...")
            print(f"  stdout: {stdout}")
            print(f"  stderr: {stderr}")

    return False, f"Falhou instalar/atualizar para todas as tentativas: {', '.join(ids_or_queries)}"

# ===============================================================
# Google Drive download (com requests)
# ===============================================================
def gdrive_download(file_id, dest_path, timeout=120):
    """
    Faz download de um arquivo público do Google Drive de forma robusta usando a biblioteca requests.
    """
    URL = "https://docs.google.com/uc?export=download"
    
    try:
        with requests.Session() as session:
            # Primeira requisição para obter o cookie de confirmação
            response = session.get(URL, params={'id': file_id}, stream=True, timeout=timeout)
            
            token = None
            for key, value in response.cookies.items():
                if key.startswith('download_warning'):
                    token = value
                    break
            
            # Se um token foi encontrado, significa que o Google mostrou um aviso.
            # Precisamos fazer uma segunda requisição com o token de confirmação.
            if token:
                params = {'id': file_id, 'confirm': token}
                response = session.get(URL, params=params, stream=True, timeout=timeout)

            # Agora, faz o download do conteúdo em streaming
            with open(dest_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk: # Filtra chunks de keep-alive
                        f.write(chunk)
            
            # Verifica se o arquivo baixado não é uma página de erro HTML
            with open(dest_path, 'r', encoding='utf-8', errors='ignore') as f:
                content_start = f.read(100).lower()
                if content_start.strip().startswith('<!doctype html'):
                    raise IOError("O arquivo baixado parece ser uma página HTML, não o instalador.")

        return True, None
    except Exception as e:
        # Retorna a mensagem de erro para ser exibida
        return False, str(e)


# ===============================================================
# Due Studio (Drive -> instalar silencioso se possível)
# ===============================================================
SILENT_SWITCHES = [
    ["/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],  # Inno Setup
    ["/S"],  # NSIS
]

def install_duestudio_from_drive():
    print("→ Baixando Due Studio do Google Drive...")
    tmpdir = tempfile.mkdtemp(prefix="duestudio_")
    filename = os.path.join(tmpdir, "Instalador_DueStudio.exe")
    
    ok, msg = gdrive_download(DUE_STUDIO_DRIVE_FILE_ID, filename)
    if not ok:
        # A msg de erro agora vem diretamente da função de download
        return False, f"Falha ao baixar do Google Drive: {msg}"

    # Tenta instalação silenciosa
    for switches in SILENT_SWITCHES:
        code, _, _ = run([filename] + switches)
        if code == 0:
            shutil.rmtree(tmpdir) # Limpa o diretório temporário
            return True, "Instalado (silencioso) a partir do Google Drive."

    # Fallback para instalação interativa
    code, _, _ = run([filename])
    if code == 0:
        shutil.rmtree(tmpdir) # Limpa o diretório temporário
        return True, "Instalado (interativo) a partir do Google Drive."
    
    # Não apaga o tmpdir em caso de falha para permitir depuração
    return False, f"Falha ao executar o instalador (código: {code}). Arquivo em: {filename}"

# ===============================================================
# Fluxo principal
# ===============================================================
PROGRAMAS_WINGET = [
    # Tenta primeiro o ID completo, depois o nome mais comum
    ["Inkscape.Inkscape", "Inkscape"],
    ["Ultimaker.Cura", "Cura"],
]

def coletar_informacoes():
    print("-------------------------------------------")
    print("Coleta de informações")
    print("-------------------------------------------\
")
    print(f"Nome do dispositivo: {get_device_name()}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Chave do Windows: {get_windows_key()}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Número de série do BIOS/Windows: {get_bios_serial()}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")

def instalar_programas():
    print("\n-------------------------------------------")
    print("Instalação/Atualização de Programas")
    print("-------------------------------------------")
    if not has_admin_rights():
        print("⚠️ Recomenda-se executar este programa COMO ADMINISTRADOR.\n")

    # Winget
    for variantes in PROGRAMAS_WINGET:
        ok, msg = winget_install_or_upgrade(variantes)
        print(("✅ " if ok else "❌ ") + msg)

    # Due Studio via Drive
    ok, msg = install_duestudio_from_drive()
    print(("✅ " if ok else "❌ ") + f"Due Studio: {msg}")

def main():
    coletar_informacoes()
    instalar_programas()
    print()
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
