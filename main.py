import os, ssl, requests, urllib3
import certifi
import yfinance as yf
from datetime import datetime, timedelta
import pandas as pd


# Função para imprimir o nome do aplicativo
def nome_do_aplicativo():
    print(
        """
▒█▀▄▀█ ░█▀▀█ ▒█▀▀█ ▒█░▄▀ ▒█▀▀▀ ▀▀█▀▀ ▀█▀ ▒█▄░▒█ ▒█▀▀█ 　 ▀▀█▀▀ ▒█░░░ ▒█▀▀▄ 
▒█▒█▒█ ▒█▄▄█ ▒█▄▄▀ ▒█▀▄░ ▒█▀▀▀ ░▒█░░ ▒█░ ▒█▒█▒█ ▒█░▄▄ 　 ░▒█░░ ▒█░░░ ▒█░▒█ 
▒█░░▒█ ▒█░▒█ ▒█░▒█ ▒█░▒█ ▒█▄▄▄ ░▒█░░ ▄█▄ ▒█░░▀█ ▒█▄▄█ 　 ░▒█░░ ▒█▄▄█ ▒█▄▄▀
\n"""
    )


# Função para finalizar o aplicativo
def finalizar_app():
    os.system("cls" if os.name == "nt" else "clear")
    print("Finalizando app")


# Função para exibir as opções do menu
def exibir_opcoes():
    print("1. Calcular PK")
    print("2. Ver cotação USDxBRL")
    print("3. Sair")


# Função para lidar com opções inválidas
def opcao_invalida():
    print("Opção inválida!\n")
    input("ENTER para voltar")
    main()


# Função para escolher uma opção do menu
def escolher_opcao():
    try:
        opcao_escolhida = int(input("Seleciona a opção:"))
        print("Você escolheu a opção:", opcao_escolhida, "\n")

        if opcao_escolhida == 1:
            calcular_pk()
        elif opcao_escolhida == 2:
            cotacao_usd_brl()
        elif opcao_escolhida == 3:
            print("Saindo do aplicativo, volte sempre!\n")
            finalizar_app()
        else:
            finalizar_app()
    except ValueError as e:
        print("Erro: ", e)
        opcao_invalida()


# Desabilitar warnings relacionados à verificação SSL (use apenas em situações controladas)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Substitua pela sua chave de API Alpha Vantage
API_KEY = "FMCXG1U8M58I9KC7"


def cotacao_usd_brl():
    try:
        # Endpoint da API Alpha Vantage para obter a taxa de câmbio USD/BRL
        url = f"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=USD&to_currency=BRL&apikey={API_KEY}"

        # Fazer a requisição para a API, ignorando a verificação SSL
        response = requests.get(url, verify=False)
        data = response.json()

        # Verificar se a resposta contém o erro padrão ou está vazia
        if response.status_code != 200:
            print(f"Erro na requisição da API: {response.status_code}")
            input("ENTER para voltar")
            main()
            return

        if not data or "Error Message" in data:
            print(
                "Erro ao obter os dados da API Alpha Vantage. Verifique sua chave de API ou o limite de chamadas."
            )
            input("ENTER para voltar")
            main()
            return

        # Verificar se o resultado contém a taxa de câmbio
        if "Realtime Currency Exchange Rate" not in data:
            print("Erro: Dados de taxa de câmbio não encontrados.")
            input("ENTER para voltar")
            main()
            return

        # Extrair os dados da resposta
        exchange_rate = data["Realtime Currency Exchange Rate"]["5. Exchange Rate"]
        last_refreshed = data["Realtime Currency Exchange Rate"]["6. Last Refreshed"]

        # Exibir a cotação e a última data de atualização
        print(f"Taxa de câmbio USD/BRL: {exchange_rate}")
        print(f"Última atualização: {last_refreshed}")

        input("ENTER para voltar")
        main()
    except Exception as e:
        print(f"Ocorreu um erro ao obter a cotação: {e}")
        opcao_invalida()


# Função para calcular PK e verificar se o PN já existe antes
def calcular_pk():
    try:
        # Função auxiliar para tratar a entrada do usuário, substituindo ',' por '.'
        def tratar_entrada(input_str):
            return float(input_str.replace(",", "."))

        # Caminho do arquivo Excel e nome da aba
        caminho_excel = r"C:\Users\brctf\Desktop\PK.xlsx"
        nome_aba = "PK"

        # Ler o arquivo Excel e garantir que a coluna 'PN' seja tratada como string
        try:
            df = pd.read_excel(caminho_excel, sheet_name=nome_aba, dtype={"PN": str})
        except FileNotFoundError:
            print("Erro: O arquivo PK.xlsx não foi encontrado.")
            return
        except ValueError:
            print(f"Erro: A aba {nome_aba} não foi encontrada no arquivo.")
            return

        # Limpar os PNs no DataFrame, removendo espaços e convertendo para maiúsculas
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()

        # Solicitar o PN do usuário e limpar a entrada
        PN = input("Digite o PN (Part Number): ").strip().upper()

        # Verificar se o PN já existe na base de dados
        if PN in df["PN"].values:
            print(f"O PN {PN} já existe na base de dados.")
            prosseguir = (
                input("Deseja prosseguir com o cálculo mesmo assim? (s/n): ")
                .strip()
                .lower()
            )
            if prosseguir != "s":
                print("Operação cancelada.")
                return main()
        else:
            print(f"O PN {PN} não foi encontrado. Será adicionado após o cálculo.")

        # Solicitar as demais variáveis
        ACIGN = input("Digite o número do chamado (ACIGN): ")
        TP = tratar_entrada(input("Digite o TP em USD: "))
        LCF = tratar_entrada(input("Digite o LCF entre 0 - 100: "))
        CNP = tratar_entrada(input("Digite o CNP: "))
        CGP = tratar_entrada(input("Digite o % CGP entre 0 - 100: ")) / 100

        # Cálculo de PK
        PK = float((CNP - (CNP * CGP)) - (TP * LCF / 100))
        print("O PK deste item é: {:.2f}\n".format(PK))

        # Chama a função para salvar o resultado em Excel
        salvar_pk_database(PN, PK, ACIGN)

        input("ENTER para voltar")
        main()
    except ValueError:
        print("Erro: Entrada inválida. Certifique-se de inserir números corretos.")
        opcao_invalida()
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        opcao_invalida()


# Função para salvar ou atualizar PK no arquivo Excel
def salvar_pk_database(PN, PK, ACIGN):
    try:
        # Caminho do arquivo Excel
        caminho_excel = r"C:\Users\brctf\Desktop\PK.xlsx"
        nome_aba = "PK"

        # Data do dia
        data_atual = datetime.now().strftime("%d/%m/%Y")

        # Limpar o PN fornecido pelo usuário: remover espaços em branco e garantir que seja string
        PN = str(PN).strip().upper()

        # Ler o arquivo Excel e garantir que a coluna 'PN' seja tratada como string
        try:
            df = pd.read_excel(caminho_excel, sheet_name=nome_aba, dtype={"PN": str})
        except FileNotFoundError:
            print("Erro: O arquivo PK.xlsx não foi encontrado.")
            return
        except ValueError:
            print(f"Erro: A aba {nome_aba} não foi encontrada no arquivo.")
            return

        # Limpar os PNs no DataFrame, removendo espaços e convertendo para maiúsculas
        df["PN"] = df["PN"].astype(str).str.strip().str.upper()

        # Verificar se o PN já existe
        if PN in df["PN"].values:
            # Encontrar a linha onde o PN está
            idx = df[df["PN"] == PN].index[0]
            pk_existente = df.at[idx, "PK"]

            # Verificar se o PK existente é diferente do novo PK
            if float(pk_existente) != float(
                PK
            ):  # Convertendo ambos para float para garantir comparação correta
                # Perguntar ao usuário qual PK deseja manter
                print(
                    f"O PN {PN} já existe com o PK {pk_existente}. O novo PK calculado é {PK}."
                )
                escolha = (
                    input(
                        "Qual PK deseja manter? (Digite 'antigo' para manter o existente ou 'novo' para atualizar): "
                    )
                    .strip()
                    .lower()
                )

                if escolha == "antigo":
                    print(f"Mantendo o PK existente: {pk_existente}")
                elif escolha == "novo":
                    df.at[idx, "PK"] = PK  # Atualizar o PK
                    print(f"PK atualizado para: {PK}")
                else:
                    print("Opção inválida. Nenhuma alteração foi feita.")
            else:
                print(f"O PK do PN {PN} já está atualizado.")

            # Atualizar a data independentemente da escolha
            df.at[idx, "Data"] = data_atual
            print(f"Data atualizada para {data_atual}.")

        else:
            # Criar um DataFrame para a nova linha
            nova_linha = pd.DataFrame(
                {"Data": [data_atual], "PN": [PN], "PK": [PK], "ACIGN": [ACIGN]}
            )

            # Verificar se o DataFrame original está vazio
            if df.empty:
                df = nova_linha  # Se o DataFrame estiver vazio, usar a nova linha
            else:
                # Usar pd.concat para adicionar a nova linha ao DataFrame existente
                df = pd.concat([df, nova_linha], ignore_index=True)

            print(f"Novo PN {PN} adicionado com PK {PK}.")

        # Salvar em um arquivo temporário para evitar problemas de permissão
        caminho_temp = r"C:\Users\brctf\Desktop\PK_temp.xlsx"
        df.to_excel(caminho_temp, sheet_name=nome_aba, index=False)

        # Substituir o arquivo original pelo temporário
        os.replace(caminho_temp, caminho_excel)

        print(f"Dados salvos com sucesso em {caminho_excel}")

    except PermissionError:
        print(
            f"Erro: Permissão negada ao tentar salvar em {caminho_excel}. Verifique se o arquivo está aberto e tente novamente."
        )
    except Exception as e:
        print(f"Ocorreu um erro ao salvar os dados: {e}")
        input("ENTER para voltar")


# Função principal
def main():
    os.system("cls" if os.name == "nt" else "clear")
    nome_do_aplicativo()
    exibir_opcoes()
    escolher_opcao()


# Declaração para iniciar o aplicativo pela função main
if __name__ == "__main__":
    main()
