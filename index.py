import pandas as pd
import os
import sys
from currency_converter import CurrencyConverter
import xlsxwriter  # Importe a biblioteca xlsxwriter

clid = sys.argv[1].upper()
file_path = f"CLID{clid}.xlsx"

# Ler arquivo Excel
workbook = pd.ExcelFile(file_path)
sheet_name = workbook.sheet_names[0]  # Trabalhando com a primeira aba/planilha

if sheet_name:
    sheet = workbook.parse(sheet_name)
    data = sheet.to_dict(orient='records')

    # Remove a linha do cabeçalho e mantém apenas as linhas dos dados
    row_data = data[1:]

    # Agrupar por contrato 'Contrato'
    grouped = sheet.groupby('Contrato')
    contract_groups = {key: group.to_dict(orient='records') for key, group in grouped}
    
    # Define os campos como String
    Plant = 'Plant'
    Quantity = 'Quantity'
    Bundle = 'Bundle'
    Description = 'Description'

    def map_international(e):
        c = CurrencyConverter()
        
        #plant_value = '{:04d}'.format(int(e['Plant'])) if pd.notnull(e['Plant']) and e['Plant'].isdigit() else ''
        
        unit_price = "{:.2f}".format(e['Unit Price']).replace(',', '.') if pd.notnull(e['Unit Price']) else ''
        
        if pd.notnull(e['Plant']):
            if e['Plant'].isdigit():
                plant_value = '{:04d}'.format(int(e['Plant']))
            else:
                plant_value = str(e['Plant'])  # Mantenha o valor original se não for um número
        else:
            plant_value = ''        

        plu = "{:.2f}".format(e['Preço Líquido Unitário']).replace(',', '.') if pd.notnull(e['Preço Líquido Unitário']) else ''
        
        return {
        'Item Number': str(e['Item Number']),
        'Short Name': e['Short Name'],
        'Material Number': str(e['Material Number']),
        Plant: str(plant_value),  # Use o valor tratado para 'Plant'
        'Unit Price': unit_price,  # Converter para o formato String e substituir vírgulas por ponto
        'Quantity': "{:.2f}".format(e['Quantity']).replace(',', '.'),
        'Unit Of Measure': e['Unit Of Measure'],
        'Preço Líquido Unitário': plu,  # Convert to string and replace commas with periods
        'Unit Price Currency': e['Unit Price Currency'],
        'Resultado Serviço': e['Resultado Serviço'],
        'Ônus de IRRF e ISS': e['Ônus de IRRF e ISS'],
        'Material Group': str(e['Material Group']),
        'Material Type': e['Material Type'],
        'Discount Amount': e['Discount Amount'],
        'Supplier Discount(%)': e['Supplier Discount(%)'],
        Bundle: e['Bundle'],
        Description: e['Description'],
        'Extended Description': e['Extended Description'],
        'Supplier Part Number': e['Supplier Part Number'],
        'Classification Domain': e['Classification Domain'],
        'Classification Code': e['Classification Code'],
        'Minimum Quantity': e['Minimum Quantity'],
        'Maximum Quantity': e['Maximum Quantity'],
        'Minimum Amount': e['Minimum Amount'],
        'Maximum Amount': e['Maximum Amount'],
        'Manufacturer Name': e['Manufacturer Name'],
        'Manufacturer Part Number': e['Manufacturer Part Number'],
        'Limit Type': e['Limit Type'],
        'Number': e['Number'],
        'Item Status': e['Item Status'],
        'External System Line Number': str(e['External System Line Number']).zfill(10),  # Formatar campo [External System Line Number] com 10 dígitos
      # ... (Outras colunas caso precisar adicionar)
        }

    def map_national(e):
        c = CurrencyConverter()
        
        plant_value = '{:04d}'.format(int(e['Plant'])) if pd.notnull(e['Plant']) and e['Plant'].isdigit() else ''
        unit_price = "{:.2f}".format(e['Unit Price']).replace(',', '.') if pd.notnull(e['Unit Price']) else ''

        if pd.notnull(e['Plant']):
            if e['Plant'].isdigit():
                plant_value = '{:04d}'.format(int(e['Plant']))
            else:
                plant_value = str(e['Plant'])  # Mantenha o valor original se não for um número
        else:
            plant_value = ''
    
        return {
        'Item Number': str(e['Item Number']),
        'Short Name': e['Short Name'],
        'Material Number': str(e['Material Number']),
        Plant: str(plant_value),  # Use o valor tratado para 'Plant'
        'Unit Price': unit_price,  # Converter para o formato String e substituir vírgulas por ponto
        'Quantity': "{:.2f}".format(e['Quantity']).replace(',', '.'), # Converter para o formato String e substituir vírgulas por ponto
        'Unit Of Measure': e['Unit Of Measure'],
        'Material Group': str(e['Material Group']),
        'Material Type': e['Material Type'],
        Bundle: e['Bundle'],
        Description: e['Description'],
        'Extended Description': e['Extended Description'],
        'Supplier Part Number': e['Supplier Part Number'],
        'Discount Amount': e['Discount Amount'],
        'Supplier Discount(%)': e['Supplier Discount(%)'],
        'Unit Price Currency': e['Unit Price Currency'],
        'Classification Domain': e['Classification Domain'],
        'Classification Code': e['Classification Code'],
        'Minimum Quantity': e['Minimum Quantity'],
        'Maximum Quantity': e['Maximum Quantity'],
        'Minimum Amount': e['Minimum Amount'],
        'Maximum Amount': e['Maximum Amount'],
        'Manufacturer Name': e['Manufacturer Name'],
        'Manufacturer Part Number': e['Manufacturer Part Number'],
        'Limit Type': e['Limit Type'],
        'Number': e['Number'],
        'Item Status': e['Item Status'],
        'External System Line Number': str(e['External System Line Number']).zfill(10),  # Formatar campo [External System Line Number] com 10 dígitos
        # ... (Outras colunas caso precisar adicionar)
        }

    contract_keys = list(contract_groups.keys())

    output_dir = os.path.join(os.getcwd(), 'output', clid)
    os.makedirs(output_dir, exist_ok=True)

    for key in contract_keys:
        group_list = list(map(lambda x: map_national(x) if sys.argv[1] == 'national' or sys.argv[1] == 'NATIONAL' else map_international(x), contract_groups[key]))
        group_list.sort(key=lambda x: int(x['Item Number'])) # Coloca na ordem crescente os dados do Item Nmuber
        new_workbook = pd.ExcelWriter(os.path.join(output_dir, f"CLID{key}.xlsx"), engine='xlsxwriter')
        
        pd.DataFrame(group_list).to_excel(new_workbook, sheet_name='Contract Item Information', index=False)
        pd.DataFrame([{
            'Item Number': '',
            'Attribute Name': '',
            'Attribute Value': '',
            'Display Text': '',
            'Type': '',
            'Description': ''
        }]).to_excel(new_workbook, sheet_name='Item Attributes', index=False)

        new_workbook.close()
        print(f"Arquivo excel CLID{key}.xlsx {sys.argv[1]} criado com sucesso!")

else:
    print('Nenhuma aba encontrada na planilha, favor verificar!')
