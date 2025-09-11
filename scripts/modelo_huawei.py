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
            if "port link-type" in linha:
                if "trunk" in linha:
                    interfaces[iface_atual]["Tagged"] = extrair_link_type_trunk(linhas)

                interfaces[iface_atual]["Link-type"] = linha.split("port link-type")[
                    1
                ].strip()

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
            if "PVID" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["PVID"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["PVID"] = ""

            if "Link-type" in linha and ":" in linha:
                trunk = linha.split(":", 1)[1].split(",")[0].strip()
                if "trunk" in trunk:
                    interfaces[iface_atual]["Tagged"] = extrair_link_type_trunk(linhas)
                try:
                    interfaces[iface_atual]["Link-type"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Link-type"] = ""

            if "Speed :" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Speed"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Speed"] = ""
            if "Duplex:" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Duplex"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Duplex"] = ""
    return interfaces


def extrair_link_type_trunk(linhas):
    for linha in linhas:
        if "port trunk allow-pass vlan" in linha:
            return " ".join(linha.split()[4:])


def extrair_lldp(linhas, interfaces):
    iface_atual = None  # Interface local atual

    for linha in linhas:
        linha = linha.strip()

        # Captura a interface local
        match_iface = re.match(r"^(\S+) has \d+ neighbor\(s\):", linha)
        if match_iface:
            nome = match_iface.group(1)
            if nome in interfaces:
                iface_atual = nome
        elif iface_atual:

            if "System name" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["LLDP Device ID"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["LLDP Device ID"] = ""
            if "Port ID" in linha and ":" in linha:
                try:
                    interfaces[iface_atual]["Neighbor Dest. Port"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Neighbor Dest. Port"] = ""
            if (
                "Management address value" in linha
                and ":" in linha
                or "Management address" in linha
                and ":" in linha
            ):
                try:
                    interfaces[iface_atual]["Neighbor IP Address"] = (
                        linha.split(":", 1)[1].split(",")[0].strip()
                    )
                except IndexError:
                    interfaces[iface_atual]["Neighbor IP Address"] = ""

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
