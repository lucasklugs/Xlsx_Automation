import pandas as pd

def get_sheet(file_path: str, sheet_name: str):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def main():
    df_acess = get_sheet("data/access_point.xlsx", sheet_name=0)

    df_acess = df_acess.sort_values(by=["user", "date", "hour"])

    print(f"{'ID':<12} {'Data':<18} {'Usuário':<13} {'Hora':<20} {'Movimentação':<26} {'Ponto de acesso':<23} {'Tipo de acesso'}")

    for id in df_acess["id"].unique():
        df_ac = df_acess[df_acess["id"] == id]

        for _, row in df_ac.iterrows():
            id_acess = row['id']
            date = row['date'].strftime("%d/%m/%Y") if pd.notna(row['date']) else "Faltante"
            user = row['user'].strip() if pd.notna(row['user']) else "Faltante"
            hour = row['hour'] if pd.notna(row['hour']) else "Faltante"
            mov = row['mov'] if pd.notna(row['mov']) else "Faltante"
            acesspoint = row['accesspoint'] if pd.notna(row['accesspoint']) else "1.0"
            finger = row['finger'] if pd.notna(row['finger']) else "1.0"

            print(f"{str(id_acess):<12} {date:<18} {user:<13} {str(hour):<20} {mov:<26} {acesspoint:<23} {finger}")

if __name__ == "__main__":
    main()
