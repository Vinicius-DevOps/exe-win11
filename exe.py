# setup_info_installer.py
import subprocess
import sys
import shutil
import tempfile
import os
import re
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

# -----------------------------
# Utilidades gerais
# -----------------------------
def run(cmd, check=False):
    try:
        p = subprocess.run(cmd, capture_output=True, text=True, shell=False)
        if check and p.returncode != 0:
            raise subprocess.CalledProcessError(p.returncode, cmd, p.stdout, p.stderr)
        return p.returncode, (p.stdout or "").strip(), (p.stderr or "").strip()
    except FileNotFoundError as e:
        return 127, "", str(e)

def run_ps(ps_script):
    return run([
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script
    ])

def has_admin_rights():
    ps = "[bool]([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)"
    code, out, _ = run_ps(ps)
    return out.strip().lower() == "true"

# -----------------------------
# Coleta de informações
# -----------------------------
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

# -----------------------------
# Winget helpers
# -----------------------------
def winget_available():
    return shutil.which("winget") is not None

def winget_search(query):
    code, out, _ = run(["winget", "search", "--name", query])
    if code != 0:
        return None
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 2 and "." in parts[-1] and not line.lower().startswith("name"):
            return parts[-1]
    return None

def winget_list_id_startswith(prefix):
    code, out, _ = run(["winget", "list"])
    if code != 0:
        return None
    for line in out.splitlines():
        parts = line.split()
        if parts:
            candidate = parts[-1]
            if candidate.lower().startswith(prefix.lower()):
                return candidate
    return None

def winget_install_or_upgrade(ids_or_queries):
    if not winget_available():
        return False, "Winget não encontrado. Instale a Microsoft Store 'App Installer' e tente novamente."
    tentativas = []
    for token in ids_or_queries:
        token = token.strip()
        if not token:
            continue
        if "." in token:
            pkg_id = token
        else:
            installed_guess = winget_list_id_startswith(token.replace(" ", ""))
            pkg_id = installed_guess or winget_search(token)
        if not pkg_id:
            tentativas.append(f"{token} (nenhum ID encontrado)")
            continue

        up_code, _, _ = run(["winget", "upgrade", "--id", pkg_id, "--silent",
                             "--accept-package-agreements", "--accept-source-agreements"])
        if up_code == 0:
            return True, f"Atualizado (ou já estava na última versão): {pkg_id}"

        in_code, _, _ = run(["winget", "install", "--id", pkg_id, "--silent",
                             "--accept-package-agreements", "--accept-source-agreements"])
        if in_code == 0:
            return True, f"Instalado: {pkg_id}"

        tentativas.append(f"{pkg_id} (upgrade:{up_code}; install:{in_code})")

    return False, "Falhou instalar/atualizar. Tentativas: " + "; ".join(tentativas)

# -----------------------------
# Due Studio (download direto do fabricante)
# -----------------------------
DUE_STUDIO_SOURCES = [
    # Páginas oficiais onde o botão "DUE STUDIO - 64 BITS" aparece com link direto
    # (mantemos mais de uma para redundância)
    "https://duelaser.zendesk.com/hc/pt-br/articles/4421456877581-V-Due-Studio-download",
    "https://duemax.duelaser.com/contents/a2c06a93-b1df-41af-8561-0225e9c04a4b",
    "https://duemax.duelaser.com/contents/1443694f-c096-4028-9ca3-cc13a0380f36",
]

# padrões comuns de instaladores (Inno Setup/NSIS/MSI)
SILENT_SWITCHES = [
    ["/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"],
    ["/S"],  # NSIS
]

def _http_get(url, timeout=30):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=timeout) as resp:
        return resp.read().decode("utf-8", errors="ignore")

def _find_duestudio_download_url(html):
    # Procura por links .exe/.msi que contenham due e studio no nome
    m = re.findall(r'href=["\'](https?://[^"\']+(?:due|duestudio)[^"\']+\.(?:exe|msi))["\']', html, flags=re.I)
    if m:
        # Heurística: escolha o mais longo (geralmente o mais específico/recente)
        m.sort(key=len, reverse=True)
        return m[0]
    # fallback: qualquer .exe/.msi na página
    m = re.findall(r'href=["\'](https?://[^"\']+\.(?:exe|msi))["\']', html, flags=re.I)
    return m[0] if m else None

def download_latest_duestudio():
    for src in DUE_STUDIO_SOURCES:
        try:
            html = _http_get(src)
            url = _find_duestudio_download_url(html)
            if url:
                return url
        except (URLError, HTTPError, TimeoutError):
            continue
    return None

def install_duestudio():
    print("→ Procurando o instalador mais recente do Due Studio no site da Due Laser...")
    url = download_latest_duestudio()
    if not url:
        return False, "Não foi possível localizar o link de download do Due Studio nas páginas de suporte da Due Laser."

    print(f"  Encontrado: {url}")
    # Baixa para pasta temporária
    tmpdir = tempfile.mkdtemp(prefix="duestudio_")
    filename = os.path.join(tmpdir, os.path.basename(url.split("?")[0]))
    try:
        print("→ Baixando instalador...")
        req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urlopen(req, timeout=120) as resp, open(filename, "wb") as f:
            f.write(resp.read())
    except Exception as e:
        return False, f"Falha ao baixar o instalador: {e}"

    # Tenta instalar silenciosamente
    for switches in SILENT_SWITCHES:
        code, _, _ = run([filename] + switches)
        if code == 0:
            return True, f"Instalado (silencioso): {os.path.basename(filename)}"

    # Se os switches não funcionarem, tenta instalação normal
    code, _, _ = run([filename])
    if code == 0:
        return True, f"Instalado (interativo): {os.path.basename(filename)}"
    return False, f"Falha ao executar o instalador ({code}). Tente executar manualmente: {filename}"

# -----------------------------
# Lógica principal
# -----------------------------
PROGRAMAS = [
    ["Inkscape.Inkscape", "Inkscape"],
    ["Ultimaker.Cura", "UltiMaker Cura", "Cura"],
]

def instalar_programas():
    print()
    print("-------------------------------------------")
    print("Instalação/Atualização de Programas")
    print("-------------------------------------------")

    if not has_admin_rights():
        print("⚠️ Recomenda-se executar este programa COMO ADMINISTRADOR.")
        print()

    # Winget apps
    for variantes in PROGRAMAS:
        ok, msg = winget_install_or_upgrade(variantes)
        print(("✅ " if ok else "❌ ") + msg)

    # Due Studio (download direto)
    ok, msg = install_duestudio()
    print(("✅ " if ok else "❌ ") + f"Due Studio: {msg}")

def coletar_informacoes():
    key = get_windows_key()
    ns = get_bios_serial()
    name = get_device_name()

    print("-------------------------------------------")
    print("Coleta de informações")
    print("-------------------------------------------")
    print()
    print(f"Nome do dispositivo: {name}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Chave do Windows: {key}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Número de série do BIOS/Windows: {ns}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")

def main():
    coletar_informacoes()
    instalar_programas()
    print()
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
