# setup_info_installer.py
import subprocess
import sys
import shutil
from typing import Optional, Tuple

# -----------------------------
# Utilidades gerais
# -----------------------------
def run(cmd: list[str], check: bool = False) -> Tuple[int, str, str]:
    """Executa um comando e retorna (returncode, stdout, stderr)."""
    try:
        p = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            shell=False
        )
        if check and p.returncode != 0:
            raise subprocess.CalledProcessError(p.returncode, cmd, p.stdout, p.stderr)
        return p.returncode, p.stdout.strip(), p.stderr.strip()
    except FileNotFoundError as e:
        return 127, "", str(e)

def run_ps(ps_script: str) -> Tuple[int, str, str]:
    """Executa um comando PowerShell sem perfil e com ExecutionPolicy bypass."""
    return run([
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-Command", ps_script
    ])

def has_admin_rights() -> bool:
    ps = "[bool]([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)"
    code, out, _ = run_ps(ps)
    return out.strip().lower() == "true"

# -----------------------------
# Coleta de informações
# -----------------------------
def get_windows_key() -> str:
    # Tenta OEM/retail quando disponível
    ps_cmd = r"(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrada)"

def get_bios_serial() -> str:
    ps_cmd = r"(Get-WmiObject Win32_BIOS | Select-Object -ExpandProperty SerialNumber)"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrada)"

def get_device_name() -> str:
    # CsName = nome do computador
    ps_cmd = r"(Get-ComputerInfo -Property 'CsName').CsName"
    _, out, _ = run_ps(ps_cmd)
    return out.strip() or "(não encontrado)"

# -----------------------------
# Winget helpers
# -----------------------------
def winget_available() -> bool:
    return shutil.which("winget") is not None

def winget_search(query: str) -> Optional[str]:
    """
    Retorna um possível 'Id' a partir do 'winget search'.
    Critérios simples: pega a primeira linha que contenha 'Id'.
    """
    code, out, err = run(["winget", "search", "--name", query])
    if code != 0:
        return None

    # A saída do winget é tabular. Tentamos achar uma linha com "Id"
    # Ex.: "Inkscape Inkscape  ...  Inkscape.Inkscape"
    lines = [l for l in out.splitlines() if l.strip()]
    # Ignora o cabeçalho até atingir linhas de resultados (heurística)
    for line in lines:
        # Normalmente o ID está na última coluna.
        parts = line.split()
        if len(parts) >= 2 and "." in parts[-1] and not line.lower().startswith("name"):
            candidate = parts[-1]
            # Filtra resultados "non-package" óbvios (heurística bem leve)
            if len(candidate) >= 5 and candidate.count(".") >= 1:
                return candidate
    return None

def winget_list_id_startswith(prefix: str) -> Optional[str]:
    """
    Checa se algo instalado tem ID começando com `prefix`.
    Retorna o ID completo se achar.
    """
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

def winget_install_or_upgrade(ids_or_queries: list[str]) -> Tuple[bool, str]:
    """
    Tenta instalar/atualizar usando:
      1) ID exato, se fornecido
      2) Busca por nome (winget search) e instala o primeiro ID compatível
    Retorna (sucesso, mensagem).
    """
    if not winget_available():
        return False, "Winget não encontrado. Instale a Microsoft Store App Installer (winget) e tente novamente."

    # Tenta na ordem:
    # - se item parece um ID (tem ponto), tenta direto
    # - senão, busca ID por 'winget search'
    tried = []

    for token in ids_or_queries:
        token = token.strip()
        if not token:
            continue

        # Descobre ID
        if "." in token:
            pkg_id = token
        else:
            # primeiro: vê se já existe algo instalado com esse prefixo (heurística)
            installed_guess = winget_list_id_startswith(token.replace(" ", ""))
            if installed_guess:
                pkg_id = installed_guess
            else:
                found = winget_search(token)
                pkg_id = found if found else None

        if not pkg_id:
            tried.append(f"{token} (nenhum ID encontrado)")
            continue

        # Se já está instalado, tenta upgrade primeiro
        # upgrade --silent
        up_code, up_out, up_err = run(["winget", "upgrade", "--id", pkg_id, "--silent",
                                       "--accept-package-agreements", "--accept-source-agreements"])
        if up_code == 0:
            return True, f"Atualizado (ou já estava na última versão): {pkg_id}"

        # Se upgrade não rolou (pode não haver atualização), tenta instalar (idempotente)
        in_code, in_out, in_err = run(["winget", "install", "--id", pkg_id, "--silent",
                                       "--accept-package-agreements", "--accept-source-agreements"])
        if in_code == 0:
            return True, f"Instalado: {pkg_id}"

        tried.append(f"{pkg_id} (upgrade:{up_code}; install:{in_code})")

    return False, "Falhou instalar/atualizar. Tentativas: " + "; ".join(tried)

# -----------------------------
# Lógica principal
# -----------------------------
PROGRAMAS = [
    # IDs preferenciais conhecidos (quando possível) seguidos de termos de busca alternativos
    # Inkscape
    ["Inkscape.Inkscape", "Inkscape"],
    # UltiMaker Cura
    ["Ultimaker.Cura", "UltiMaker Cura", "Cura"],
    # DueStudio (não há um ID universal conhecido; fazemos busca por nome)
    ["DueStudio", "Due Studio"]
]

def instalar_programas():
    print()
    print("-------------------------------------------")
    print("Instalação/Atualização de Programas (Winget)")
    print("-------------------------------------------")

    if not has_admin_rights():
        print("⚠️ Recomenda-se executar este programa COMO ADMINISTRADOR para instalar/atualizar aplicativos.")
        print()

    for variantes in PROGRAMAS:
        nome_exibicao = variantes[0]
        ok, msg = winget_install_or_upgrade(variantes)
        status = "✅" if ok else "❌"
        print(f"{status} {msg}")

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
