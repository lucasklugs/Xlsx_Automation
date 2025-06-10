import pandas as pd


def get_sheet(file_path:str, sheet_name):
    df = pd.read_excel(file_path, sheet_name= sheet_name)
    return df

def main():
    df_mov = get_sheet("data/estoque-bic.xlsx", "movimentacao")
    df_bal = get_sheet("data/estoque-bic.xlsx", "saldos")
    df_mov = df_mov.sort_values(by=["item-code", "dt-mov", "tip-mov"])
    balances = df_bal.set_index("Item")["Saldo Inicial"].to_dict()
    for index,row in df_mov.iterrows():
        print(row["item-code"])

    print(df_mov.head())
    print(balances)

if __name__ == "__main__":
    main()