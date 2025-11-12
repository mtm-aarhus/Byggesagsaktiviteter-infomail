from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import os
from datetime import datetime
import pandas as pd
import pyodbc
import mimetypes
from email.message import EmailMessage
import smtplib
def make_excel(data, filename, tablename ):
    ExcelFileName = f"{datetime.today().strftime('%Y%m%d')} {filename}.xlsx"
    excel_file_path = os.path.join(os.getcwd(), ExcelFileName)

    # Sørg for at 'Sagsdato' er i datetime-format
    if "Sagsdato" in data.columns:
        data["Sagsdato"] = pd.to_datetime(data["Sagsdato"], errors='coerce')

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter', datetime_format='dd-mm-yyyy') as writer:
        data.to_excel(writer, sheet_name='Ark1', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Ark1']

        # Lav Excel-tabel over hele området
        (max_row, max_col) = data.shape
        col_range = chr(65 + max_col - 1)
        table_range = f"A1:{col_range}{max_row + 1}"

        worksheet.add_table(table_range, {
            'name': f'{tablename}',
            'columns': [{'header': col} for col in data.columns]
        })

        # Excel-datoformat
        date_format = workbook.add_format({'num_format': 'dd-mm-yyyy'})

        # Kolonnebredde og formatering
        for i, column in enumerate(data.columns):
            column_length = max(data[column].astype(str).map(len).max(), len(column))
            fmt = date_format if column == "Sagsdato" else None
            worksheet.set_column(i, i, column_length + 2, fmt)
    return excel_file_path
def get_data_from_sql(sql_file_path, orchestrator_connection):
    with open(sql_file_path, "r", encoding="utf-8") as file:
        query = file.read()

    # Database connection setup
    sql_server = orchestrator_connection.get_constant("SqlServer").value
    conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};' + f'SERVER={sql_server};DATABASE=LOIS;Trusted_Connection=yes'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Step 2: Execute the actual SELECT query
    cursor.execute(query)

    # Step 3: Fetch results
    rows = cursor.fetchall()
    orchestrator_connection.log_info(f'rows {rows}')

    # Step 4: Load the data into a Pandas DataFrame
    data = pd.read_sql(query, conn)

    # Step 5: Close database connection
    cursor.close()
    conn.close()
    return data

def process(orchestrator_connection: OrchestratorConnection| None = None) -> None:

    orchestrator_connection.log_info('Starting process Byggersager uden aktivitet -bot')

    # Read the SQL query from file
    sql_file_path_1 = "SQL LOIS-tabeller 2 - sager med åbne aktiviteter uden startdato.sql"
    data1 = get_data_from_sql(sql_file_path_1, orchestrator_connection)

    excel_file_name_1 = 'Sager uden aktivitet eller kun med tidsbegrænsede aktiviteter'
    table_name_1 = 'Byggesager_uden_aktiviteter'
    excel_file_path_1 = make_excel(data1, excel_file_name_1, table_name_1)

    # Read the SQL query from file
    sql_file_path_2 = "SQL LOIS-tabeller 2 - sager kun med tidsregistrering.sql"
    data2 = get_data_from_sql(sql_file_path_2, orchestrator_connection)

    excel_file_name_2 = 'Sager kun med tidsregistreringsaktivitet'
    table_name_2 = 'Byggesager_kun_med_tidsregistreringsaktivitet'
    excel_file_path_2 = make_excel(data2, excel_file_name_2, table_name_2)


    # SMTP Configuration
    SMTP_SERVER = "smtp.adm.aarhuskommune.dk"
    SMTP_PORT = 25
    SCREENSHOT_SENDER = "byggesager@aarhus.dk"
    subject = f"Byggesager uden aktiviteter"

    # Email body (HTML)
    body = f"""
    <html>
        <body>
            <p>Hej Anne 😊,</p>
            <br>
            <p>Her er excelarket med byggesager uden aktiviteter eller med tidsbegrænsede aktiviteter,</p>
            <p>samt excelarket med byggesager, der kun har tidsregistreringsaktiviteter. </p>
            <br>
            <p>Med venlig hilsen</p>
            <br>
            <p>Teknik og Miljø</p>
            <p>Digitalisering</p>
            <p>Aarhus Kommune</p>
        </body>
    </html>
    """

    # Hent mailadresse fra Orchestrator
    UdviklerMail = orchestrator_connection.get_constant('balas').value
    ModtagerMail = orchestrator_connection.get_constant('AnneKVMail').value

    # Opret besked
    msg = EmailMessage()
    msg['To'] = UdviklerMail
    msg['From'] = SCREENSHOT_SENDER
    msg['Subject'] = subject
    msg.set_content("Please enable HTML to view this message.")
    msg.add_alternative(body, subtype='html')
    msg['Reply-To'] = UdviklerMail
    msg['Bcc'] = UdviklerMail

    # Vedhæft Excel-filer
    try:
        with open(excel_file_path_1, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(excel_file_path_1)
            mime_type, _ = mimetypes.guess_type(excel_file_path_1)
            maintype, subtype = mime_type.split('/') if mime_type else ('application', 'octet-stream')

            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)
    except Exception as e:
        orchestrator_connection.log_info(f"Fejl under vedhæftning af fil: {e}")
    try:
        with open(excel_file_path_1, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(excel_file_path_2)
            mime_type, _ = mimetypes.guess_type(excel_file_path_2)
            maintype, subtype = mime_type.split('/') if mime_type else ('application', 'octet-stream')

            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)
    except Exception as e:
        orchestrator_connection.log_info(f"Fejl under vedhæftning af fil: {e}")

    # Send mail
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.send_message(msg)
    except Exception as e:
        orchestrator_connection.log_info(f"Failed to send byggesagsemail: {e}")


    if os.path.exists(excel_file_path_1):
        os.remove(excel_file_path_1)
    if os.path.exists(excel_file_path_2):
        os.remove(excel_file_path_2)