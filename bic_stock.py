import pandas as pd
import datetime

def get_sheet(file_path: str, sheet_name: str):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def main():
    df_mov = get_sheet("data/estoque-bic.xlsx", "movimentacao")
    df_bal = get_sheet("data/estoque-bic.xlsx", "saldos")

    df_mov = df_mov.sort_values(by=["item-code", "dt-mov", "tip-mov"])
    balances = df_bal.set_index("Item")["Saldo Inicial"].to_dict() # Cria um dicionário com os saldos iniciais, onde a chave é o código do item

    # Para cada item único na movimentação
    for item in df_mov["item-code"].unique():
        df_item = df_mov[df_mov["item-code"] == item]
        bl_previous = balances.get(item, 0)
        bl_actual = bl_previous
        

        print(f"\nITEM: {item}")
        print(f"\nSaldo Anterior: {bl_previous}\n")
        print("Data\t\tTipo Movimentação\tQuantidade\tSaldo Atualizado")

        for _, row in df_item.iterrows():
            data = row["dt-mov"].strftime("%d/%m/%Y")
            tipo = row["tip-mov"].strip() # Strip remove os espaços em branco do ínicio e fim de uma string
            qtd = row["qtd"]
            if tipo == "C":
                tipo_txt = "Compra" 
            else:
                tipo_txt = "Venda"

            if tipo == "C":
                bl_actual += qtd
            elif tipo == "V":
                bl_actual -= qtd

            print(f"{data:20} {tipo_txt:21} {qtd:4} {bl_actual:18}")

        print(f"\nSaldo Final: {bl_actual}")

if __name__ == "__main__":
    main()
