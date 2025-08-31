# win_key.py
import subprocess

def main():
    # PowerShell: pega a chave OEM/retail quando disponível
    ps_cmd = r"powershell (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey"

    # chama powershell sem perfil e ignorando ExecutionPolicy (só na sessão)
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd],
        capture_output=True, text=True
    )

    key = result.stdout.strip() or "(não encontrada)"
    print(f"Chave do Windows: {key}")
    input("Pressione Enter para sair...")

if __name__ == "__main__":
    main()
