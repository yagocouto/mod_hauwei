from pathlib import Path
import pandas as pd
import re


def ler_arquivo(caminho):
    for enc in ["utf-8", "latin-1", "cp1252", "utf-16"]:
        try:
            with open(caminho, "r", encoding=enc) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(
        "Não foi possível abrir o arquivo com codificações conhecidas."
    )


def extrair_device_name(linhas):
    for linha in linhas:
        if linha.startswith("sysname"):
            return linha.split("sysname")[1].strip()
    return ""


def extrair_interface_brief(linhas):
    interfaces = {}
    capturar = False
    for linha in linhas:
        if linha.strip().startswith("Interface") and "outErrors" in linha:
            capturar = True
            continue
        if capturar:
            if not linha.strip() or linha.startswith("<") or linha.startswith("==="):
                break
            partes = linha.split()
            if len(partes) < 7:
                continue

            if not (partes[-2].isdigit() and partes[-1].isdigit()):
                continue

            iface = partes[0]
            status = partes[1]
            in_errors = int(partes[-2])
            out_errors = int(partes[-1])
            observacao = ""
            if in_errors >= 5 or out_errors >= 5:
                observacao = f"inErrors: {in_errors}, outErrors: {out_errors}"
            interfaces[iface] = {
                "Current Port": iface,
                "Status": status,
                "Observação": observacao,
            }
    return interfaces


def display_current_configuration(linhas, interfaces):
    iface_atual = None
    for linha in linhas:
        if linha.strip().lower().startswith("interface "):
            nome = linha.split()[1].strip()
            for iface_nome in interfaces:
                if iface_nome.lower() == nome.lower():
                    iface_atual = iface_nome
                    break
            else:
                iface_atual = None

        if iface_atual:
            if "port hybrid tagged vlan" in linha:
                interfaces[iface_atual]["Tagged"] = linha.split("tagged vlan")[
                    1
                ].strip()
            if "port hybrid untagged vlan" in linha:
                interfaces[iface_atual]["Untagged"] = linha.split("untagged vlan")[
                    1
                ].strip()
            if "voice-vlan" in linha and "enable" in linha:
                interfaces[iface_atual]["Voice-vlan"] = linha.split()[1]

    return interfaces


def extrair_detalhes_display_interface(linhas, interfaces):
    iface_atual = None
    for linha in linhas:
        # Detecta interface antes de "current state"
        match_iface = re.match(r"^(\S+)\s+current state", linha)
        if match_iface:
            nome = match_iface.group(1)
            if nome in interfaces:
                iface_atual = nome
        elif iface_atual:
            if "Description" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Description"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Description"] = ""
            if "Link-type" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Link-type"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Description"] = ""
            if "PVID" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["PVID"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["PVID"] = ""
            if "Link-type" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Link-type"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Link-type"] = ""
            if "Speed" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Speed"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Speed"] = ""
            if "Duplex" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Duplex"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Duplex"] = ""
    return interfaces


def extrair_lldp(linhas, interfaces):
    iface_local = None  # Interface local atual
    iface_atual = None  # Interface que vamos gravar no dict
    port_id = None
    mgmt_addr = None

    for linha in linhas:
        linha = linha.strip()

        # Captura a interface local
        match_iface = re.match(r"^(\S+) has \d+ neighbor\(s\):", linha)
        if match_iface:
            # se já estávamos em uma interface, salva antes de resetar
            if iface_local and iface_local in interfaces:
                if mgmt_addr:  # caso com Management Address
                    interfaces[iface_local]["Neighbor Dest. Port"] = port_id
                    interfaces[iface_local]["Neighbor IP Address"] = mgmt_addr
                else:  # caso sem Management Address
                    interfaces[iface_local]["LLDP Device ID"] = port_id

            # inicia novo bloco
            iface_local = match_iface.group(1)
            port_id = None
            mgmt_addr = None
            continue

        # Captura Port ID
        match_port = re.match(r"^Port ID\s*:\s*(.+)", linha)
        if match_port:
            port_id = match_port.group(1).strip()
            continue

        # Captura Management address value
        match_mgmt = re.match(r"^Management address value\s*:\s*(.+)", linha)
        if match_mgmt:
            mgmt_addr = match_mgmt.group(1).strip()
            continue

    # salva o último bloco
    if iface_local and iface_local in interfaces:
        if mgmt_addr:
            interfaces[iface_local]["Neighbor Dest. Port"] = port_id
            interfaces[iface_local]["Neighbor IP Address"] = mgmt_addr
        else:
            interfaces[iface_local]["LLDP Device ID"] = port_id

    return interfaces


def processar_mod_huawei():
    pasta = Path("entrada")
    arquivos_txt = list(pasta.glob("*.txt"))
    caminho_excel = Path("saida/interfaces_huawei.xlsx")

    for arquivo in arquivos_txt:
        print(f"Lendo arquivo: {arquivo.name}")
        linhas = ler_arquivo(arquivo)
        device_name = extrair_device_name(linhas)
        interfaces = extrair_interface_brief(linhas)

        for v in interfaces.values():
            campos = [
                "Description",
                "PVID",
                "Duplex",
                "Speed",
                "Link-type",
                "Tagged",
                "Untagged",
                "Voice-vlan",
                "LLDP Device ID",
                "Neighbor Dest. Port",
                "Neighbor IP Address",
            ]
            for c in campos:
                v.setdefault(c, "")

        interfaces = display_current_configuration(linhas, interfaces)
        interfaces = extrair_detalhes_display_interface(linhas, interfaces)
        interfaces = extrair_lldp(linhas, interfaces)

        dados = []
        for iface in interfaces.values():
            iface["Device Name"] = device_name
            dados.append(iface)

        colunas = [
            "Device Name",
            "Current Port",
            "Description",
            "LLDP Device ID",
            "Neighbor Dest. Port",
            "Neighbor IP Address",
            "Status",
            "Link-type",
            "PVID",
            "Tagged",
            "Untagged",
            "Voice-vlan",
            "Duplex",
            "Speed",
            "Observação",
        ]
        df = pd.DataFrame(dados, columns=colunas)

        if caminho_excel.exists():
            with pd.ExcelWriter(
                caminho_excel, engine="openpyxl", mode="a", if_sheet_exists="replace"
            ) as writer:
                df.to_excel(writer, sheet_name=device_name, index=False)
        else:
            with pd.ExcelWriter(caminho_excel, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=device_name, index=False)


if __name__ == "__main__":
    processar_mod_huawei()
