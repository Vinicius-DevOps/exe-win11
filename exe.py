# setup_info_installer.py
import subprocess
import shutil
import os

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
    ps = ('[bool]([Security.Principal.WindowsPrincipal]'
          ' [Security.Principal.WindowsIdentity]::GetCurrent()).'
          "IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)")

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

def main():
    coletar_informacoes()
    instalar_programas()
    print()
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()