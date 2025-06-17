import streamlit as st
import csv
import pandas as pd
import io
from datetime import datetime, timedelta

# Utility: format date to MM/DD/YYYY 00:00

def formatta_data(data_str: str) -> str:
    for sep in ['-', '/']:
        try:
            g, m, a = map(int, data_str.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime('%m/%d/%Y 00:00')
        except ValueError:
            continue
    return data_str

# AD modification header fields
header_modifica = [
    'sAMAccountName', 'Creation', 'OU', 'Name', 'DisplayName', 'cn', 'GivenName', 'Surname',
    'employeeNumber', 'employeeID', 'department', 'Description', 'passwordNeverExpired',
    'ExpireDate', 'userprincipalname', 'mail', 'mobile', 'RimozioneGruppo', 'InserimentoGruppo',
    'disable', 'moveToOU', 'telephoneNumber', 'company'
]

st.title('Gestione Modifiche AD')

# File naming and row count
file_name = st.text_input('Nome file di output', 'modifiche_utenti.csv')
num_righe = st.number_input('Quante righe inserire?', min_value=1, max_value=20, value=1)

modifiche = []
for i in range(num_righe):
    with st.expander(f'Riga {i+1}'):
        m = dict.fromkeys(header_modifica, '')
        m['sAMAccountName'] = st.text_input(f'[Riga {i+1}] sAMAccountName *', key=f'sam_{i}')
        campi = st.multiselect(
            f'[Riga {i+1}] Campi da modificare',
            [f for f in header_modifica if f != 'sAMAccountName'],
            key=f'fields_{i}'
        )
        for campo in campi:
            v = st.text_input(f'[Riga {i+1}] {campo}', key=f'{campo}_{i}')
            # Format special fields
            if campo == 'ExpireDate':
                v = formatta_data(v)
            if campo == 'mobile' and v and not v.startswith('+'):
                v = f'+39 {v.strip()}'
            m[campo] = v
        modifiche.append(m)

# Generate CSV
if st.button('Genera CSV Modifiche'):
    buf = io.StringIO()
    # Use no quoting and escape spaces
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar='\\')
    # Write header
    writer.writerow(header_modifica)
    # Write each row, escaping spaces in each field
    for m in modifiche:
        row = [str(m.get(field, '')) for field in header_modifica]
        # Escape spaces by prefixing with backslash
        escaped = [val.replace(' ', '\\ ') for val in row]
        writer.writerow(escaped)
    buf.seek(0)

    # Show preview
    df = pd.DataFrame(modifiche, columns=header_modifica)
    st.dataframe(df)

    # Instruction
    file_path = r"\\srv_dati.consip.tesoro.it\AreaCondivisa\DEPSI\IC\AD_Modifiche"
    st.markdown(f"""
Ciao.

Si richiede modifica come da file **{file_name}**
archiviato al percorso `{file_path}`

Grazie
""")

    # Download
    st.download_button(
        label='📥 Scarica CSV Modifiche',
        data=buf.getvalue(),
        file_name=file_name,
        mime='text/csv'
    )
    st.success('✅ CSV modifiche generato con successo.')
