from pathlib import Path
import pandas as pd


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


def extrair_detalhes_interfaces(linhas, interfaces):
    iface_atual = None
    for linha in linhas:
        if linha.strip().lower().startswith("interface "):
            nome = linha.split()[1].strip()
            # normaliza para minúsculas ao comparar
            for iface_nome in interfaces:
                if iface_nome.lower() == nome.lower():
                    iface_atual = iface_nome
                    break
            else:
                iface_atual = None

        if iface_atual:
            # captura description ignorando espaços e maiúsculas
            if "description" in linha.lower():
                partes = linha.split("description", 1)
                if len(partes) > 1:
                    descricao = partes[1]
                    interfaces[iface_atual]["Description"] = descricao

            if "port link-type" in linha:
                interfaces[iface_atual]["Link-type"] = linha.split()[-1]
            if "port hybrid pvid vlan" in linha:
                interfaces[iface_atual]["PVID"] = linha.split()[-1]
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
        if linha.startswith("GigabitEthernet"):
            nome = linha.split()[0]
            if nome in interfaces:
                iface_atual = nome
        elif iface_atual:
            if "Speed :" in linha:
                interfaces[iface_atual]["Speed"] = (
                    linha.split(":")[1].split(",")[0].strip()
                )
            if "Duplex:" in linha:
                interfaces[iface_atual]["Duplex"] = (
                    linha.split(":")[1].split(",")[0].strip()
                )
    return interfaces


import re


def extrair_lldp(linhas, interfaces):
    """
    Extrai informações de LLDP de um conjunto de linhas de saída
    e organiza os dados por interface.
    """

    iface_atual = None
    for linha in linhas:
        linha = linha.strip()

        # Identifica início de interface
        if linha.startswith("GigabitEthernet"):
            nome = linha.split()[0]
            if nome in interfaces:
                iface_atual = nome
            continue

        if iface_atual:
            # Captura Port ID (ignora 'Port ID type')
            if re.search(r"^Port ID\s*:", linha) and "Port ID type" not in linha:
                interfaces[iface_atual]["Neighbor Dest. Port"] = linha.split(":", 1)[
                    1
                ].strip()

            # Captura LLDP Device ID
            elif re.search(r"^Chassis ID\s*:", linha):
                interfaces[iface_atual]["LLDP Device ID"] = linha.split(":", 1)[
                    1
                ].strip()

            # Captura IP do vizinho
            elif re.search(r"^Management address value\s*:", linha):
                interfaces[iface_atual]["Neighbor IP Address"] = linha.split(":", 1)[
                    1
                ].strip()

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

        # Inicializa campos apenas se não existirem (não sobrescreve Description)
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

        interfaces = extrair_detalhes_interfaces(linhas, interfaces)
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
