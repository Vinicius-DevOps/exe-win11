# win_key.py
import subprocess

def main():
    # PowerShell: pega a chave OEM/retail quando disponível
    ps_cmd = r"(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"

    # Pegar o numero  de serie do Windows via PowerShell
    ns_cmd = r"Get-WmiObject Win32_BIOS | Select-Object SerialNumber"

    # chama powershell sem perfil e ignorando ExecutionPolicy (só na sessão)
    result_win = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd],
        capture_output=True, text=True
    )

    # chama powershell sem perfil e ignorando ExecutionPolicy (só na sessão)
    result_key = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ns_cmd],
        capture_output=True, text=True
    )

    key = result_win.stdout.strip() or "(não encontrada)"
    ns = result_key.stdout.strip() or "(não encontrada)"
    print(f"Chave do Windows: {key}")
    print(f"Número de série do Windows: {ns}")
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
