# win_key.py
import subprocess

def main():
    # PowerShell: pega a chave OEM/retail quando disponível
    ps_cmd = r"(Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"

    # Pegar o numero  de serie do Windows via PowerShell
    ns_cmd = r"Get-WmiObject Win32_BIOS | Select-Object SerialNumber"

    # Coletando o nome do dispositivo
    dv_name_cmd = r'Get-ComputerInfo -Property "CsName"'

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

    # chama powershell sem perfil e ignorando ExecutionPolicy (só na sessão)
    result_dv = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", dv_name_cmd],
        capture_output=True, text=True
    )

    key = result_win.stdout.strip() or "(não encontrada)"
    ns = result_key.stdout.strip() or "(não encontrada)"
    name = result_dv.stdout.strip() or "(não encontrado)"

    print("-------------------------------------------")
    print("Coleta de informações")
    print("-------------------------------------------")
    # Quebra de linha para melhor visualização
    print()
    print(f"Nome do dispositivo: {name}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Chave do Windows: {key}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
    print(f"Número de série do Windows: {ns}")
    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")

    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
