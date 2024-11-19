import pandas as pd

excel_file = './setup/transformer_data.xlsx'

def create_or_load_file():
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        columns = ['Province', 'Area Code', 'Date', 'Transformer Serial No', 'KV', 'KVA',
                   'Reason for Movement', 'Present Condition', 
                   'Movement From Cost Code', 'Movement From SIN Location',
                   'Movement To Cost Code', 'Movement To SIN Location',
                   'TV/WB No', 'Remark']
        df = pd.DataFrame(columns=columns)
        df.to_excel(excel_file, index=False)
