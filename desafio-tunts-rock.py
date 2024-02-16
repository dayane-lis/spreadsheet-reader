import gspread
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from math import ceil

def authenticate_google_sheets():
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file('/assets/json/credential-file.json', scope=scope)
    client = gspread.authorize(creds)
    return client

def calculate_final_grade(m, naf):
    return ceil((m + naf) / 2)

def process_student(spreadsheet, sheet_name):
    sheet = spreadsheet.worksheet(sheet_name)

    header = sheet.row_values(2)  
    matricula_col = header.index("Matricula") + 1
    aluno_col = header.index("Nome") + 1
    faltas_col = header.index("Faltas") + 1
    p1_col = header.index("P1") + 1
    p2_col = header.index("P2") + 1
    p3_col = header.index("P3") + 1
    situacao_col = header.index("Situacao") + 1
    naf_col = header.index("Nota para Aprovação Final") + 1

    for i in range(3, sheet.row_count + 1):
        matricula = sheet.cell(i, matricula_col).value
        aluno_name = sheet.cell(i, aluno_col).value
        faltas = float(sheet.cell(i, faltas_col).value)
        p1 = float(sheet.cell(i, p1_col).value)
        p2 = float(sheet.cell(i, p2_col).value)
        p3 = float(sheet.cell(i, p3_col).value)

        total_classes = 60  
        absent_classes = faltas        

        average = (p1 + p2 + p3) / 3
        print(f"Calculating for {aluno_name} - Average: {average:.2f}")

        if absent_classes > 0.25 * total_classes:
            status = "Reprovado por Falta"
        elif average < 5:
            status = "Reprovado por Nota"
        elif 5 <= average < 7:
            status = "Exame Final"
            naf = calculate_final_grade(average, 0)
            print(f"Calculating NAF for {aluno_name} - NAF: {naf}")
            sheet.update_cell(i, naf_col, naf)
        else:
            status = "Aprovado"

        sheet.update_cell(i, situacao_col, status)
        print(f"Result for {aluno_name} - Status: {status}")

if __name__ == "__main__":
    
    client = authenticate_google_sheets()

    spreadsheet_link = 'https://docs.google.com/spreadsheets/d/1Q54X0fTkyycifu93OYaPiESaSTICNwg0VYB0r8r5V3U/edit#gid=0'

    spreadsheet_key = spreadsheet_link.split('/')[5]

    sheet_name = 'Engenharia de Software'

    process_student(client.open_by_key(spreadsheet_key), sheet_name)
