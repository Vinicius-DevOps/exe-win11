# setup_info_installer.py
import subprocess
import shutil
import os

# ===============================================================
# Utilidades
# ===============================================================
def run(cmd, check=False):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, shell=False)
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

def winget_search(query):
    code, out, _ = run(["winget", "search", "--name", query])
    if code != 0: return None
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 2 and "." in parts[-1] and not line.lower().startswith("name"):
            return parts[-1]
    return None

def winget_list_id_startswith(prefix):
    code, out, _ = run(["winget", "list"])
    if code != 0: return None
    for line in out.splitlines():
        parts = line.split()
        if parts:
            candidate = parts[-1]
            if candidate.lower().startswith(prefix.lower()): return candidate
    return None

def winget_install_or_upgrade(ids_or_queries):
    if not winget_available():
        return False, "Winget não encontrado. Instale a Microsoft Store 'App Installer' e tente novamente."
    tentativas = []
    for token in ids_or_queries:
        token = token.strip()
        if not token: continue
        pkg_id = token if "." in token else (winget_list_id_startswith(token.replace(" ","")) or winget_search(token))
        if not pkg_id:
            tentativas.append(f"{token} (nenhum ID encontrado)"); continue
        up_code,_,_ = run(["winget","upgrade","--id",pkg_id,"--silent","--accept-package-agreements","--accept-source-agreements"])
        if up_code == 0: return True, f"Atualizado (ou já estava na última versão): {pkg_id}"
        in_code,_,_ = run(["winget","install","--id",pkg_id,"--silent","--accept-package-agreements","--accept-source-agreements"])
        if in_code == 0: return True, f"Instalado: {pkg_id}"
        tentativas.append(f"{pkg_id} (upgrade:{up_code}; install:{in_code})")
    return False, "Falhou instalar/atualizar. Tentativas: " + "; ".join(tentativas)

# ===============================================================
# Fluxo principal
# ===============================================================
PROGRAMAS_WINGET = [
    ["Inkscape.Inkscape", "Inkscape"],
    ["Ultimaker.Cura", "UltiMaker Cura", "Cura"],
]

def coletar_informacoes():
    print("-------------------------------------------")
    print("Coleta de informações")
    print("-------------------------------------------\n")
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