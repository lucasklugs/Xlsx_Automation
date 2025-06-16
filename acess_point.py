import pandas as pd
import locale

locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')

def get_sheet(file_path: str, sheet_name: str):
    return pd.read_excel(file_path, sheet_name=sheet_name)

def preprocess_records(df):
    
    df['accesspoint'] = df['accesspoint'].fillna(1)

    
    df = df[df['hour'].notna()].copy()

    df['date_only'] = df['date'].dt.date

    # Ordena os dados por usuário, data e hora
    df = df.sort_values(by=['user', 'date', 'hour'], ascending=True)

    # Regra: Se dois registros tiverem o mesmo horário na mesma data, mantém o mais recente (último)
    df = df.drop_duplicates(subset=['user', 'date_only', 'hour'], keep='last')

    df['mov'] = df['mov'].fillna('auto')

    return df


def generate_access_report(records):
    access = {}  # Dicionário final para armazenar os dados por usuário
    user_day_records = {}  # Armazena registros temporários por (usuário, data)

    # Função auxiliar para adicionar um par de entrada/saída no dicionário final
    def append_record(user, date, entry, exit_):
        if user not in access:
            access[user] = []
        access[user].append({
            'date': date,
            'entry': entry,
            'exit': exit_
        })


    for record in records:
        user = record['user']
        date = record['date'].date()
        hour = record['hour']
        mov = record['mov']

        key = (user, date)

        if key not in user_day_records:
            user_day_records[key] = []

        user_day_records[key].append((hour, mov))

    for (user, date), actions in user_day_records.items():
        actions.sort()
        i = 0
        while i < len(actions):
            hour, mov = actions[i]

            if mov == 'exit':
                append_record(user, date, "Faltante", hour)
                i += 1
            elif mov == 'entry':
                if i + 1 < len(actions) and actions[i + 1][1] == 'exit':
                    append_record(user, date, hour, actions[i + 1][0])
                    i += 2
                else:
                    append_record(user, date, hour, "Faltante")
                    i += 1
            elif mov == 'auto':
                # Se é o único do dia → entrada
                if len(actions) == 1:
                    append_record(user, date, hour, "Faltante")
                    i += 1
                else:
                    # Se é par, assume como entrada
                    if i % 2 == 0:
                        if i + 1 < len(actions):
                            append_record(user, date, hour, actions[i + 1][0])
                            i += 2
                        else:
                            append_record(user, date, hour, "Faltante")
                            i += 1
                    else:
                        append_record(user, date, "Faltante", hour)
                        i += 1
            else:
                i += 1

    return access

def main():
    df = get_sheet("data/access_point.xlsx", sheet_name=0)
    df = preprocess_records(df)
    df = df.sort_values(by=["user", "date", "hour"])

    records = df.to_dict(orient="records")
    access_report = generate_access_report(records)

    print(f"\n{'Usuário':<15} {'Data':<12} {'Entrada':<10} {'Saída':<10}")
    for user, entries in access_report.items():
        for f in entries:
            data = f['date'].strftime("%d/%b")
            entry = f['entry']
            exit_ = f['exit']
            entry = entry.strftime("%H:%M:%S") if hasattr(entry, 'strftime') else str(entry)
            exit_ = exit_.strftime("%H:%M:%S") if hasattr(exit_, 'strftime') else str(exit_)
            print(f"{user:<15} {data:<12} {entry:<10} {exit_:<10}")

if __name__ == "__main__":
    main()
