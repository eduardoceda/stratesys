from flask import Flask, render_template, request, redirect, url_for
from flask import jsonify
from flask import render_template
from flask_wtf.csrf import CSRFProtect
import pandas as pd
import os
from currency_converter import CurrencyConverter
import xlsxwriter

app = Flask(__name__)

clid = 'default'  # Valor padrão para evitar erro ao chamar sys.argv[1]
file_path = f"CLID{clid}.xlsx"
contract_groups = {}  # Para evitar erros ao acessar essa variável

Plant = 'Plant'
Quantity = 'Quantity'
Bundle = 'Bundle'
Description = 'Description'

def map_international(e):
    c = CurrencyConverter()
    unit_price = "{:.2f}".format(e['Unit Price']).replace(',', '.') if pd.notnull(e['Unit Price']) else ''
    plu = "{:.2f}".format(e['Preço Líquido Unitário']).replace(',', '.') if pd.notnull(e['Preço Líquido Unitário']) else ''
        
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
        'Plant': str(plant_value),
        'Unit Price': unit_price,
        'Quantity': "{:.2f}".format(e['Quantity']).replace(',', '.'),
        'Unit Of Measure': e['Unit Of Measure'],
        'Preço Líquido Unitário': plu,
        'Unit Price Currency': e['Unit Price Currency'],
        'Resultado Serviço': e['Resultado Serviço'],
        'Ônus de IRRF e ISS': e['Ônus de IRRF e ISS'],
        'Material Group': str(e['Material Group']),
        'Material Type': e['Material Type'],
        'Discount Amount': e['Discount Amount'],
        'Supplier Discount(%)': e['Supplier Discount(%)'],
        'Bundle': e['Bundle'],
        'Description': e['Description'],
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
        'External System Line Number': str(e['External System Line Number']).zfill(10),
    }

def map_national(e):
    c = CurrencyConverter()
    unit_price = "{:.2f}".format(e['Unit Price']).replace(',', '.') if pd.notnull(e['Unit Price']) else ''  
        
    if pd.notnull(e['Plant']):
        plant_value = e['Plant']
        if isinstance(plant_value, (int, float)) or (isinstance(plant_value, str) and plant_value.isdigit()):
            plant_value = '{:04d}'.format(int(plant_value))
        else:
            plant_value = str(e['Plant'])  # Mantenha o valor original se não for um número ou não puder ser convertido para número
    else:
        plant_value = '' 
    
    return {
        'Item Number': str(e['Item Number']),
        'Short Name': e['Short Name'],
        'Material Number': str(e['Material Number']),
        'Plant': str(plant_value),
        'Unit Price': unit_price,
        'Quantity': "{:.2f}".format(e['Quantity']).replace(',', '.'),
        'Unit Of Measure': e['Unit Of Measure'],
        'Material Group': str(e['Material Group']),
        'Material Type': e['Material Type'],
        'Bundle': e['Bundle'],
        'Description': e['Description'],
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
        'External System Line Number': str(e['External System Line Number']).zfill(10),
    }

# Criação do diretório 'uploads'
uploads_dir = os.path.join(app.instance_path, 'uploads')
os.makedirs(uploads_dir, exist_ok=True)    

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/arquivos_gerados', methods=['GET'])
def listar_arquivos_gerados():
    output_dir = os.path.join(os.getcwd(), 'output', 'national')  # Diretório de saída
    if os.path.exists(output_dir):
        lista_arquivos = os.listdir(output_dir)
        return jsonify({'arquivos': lista_arquivos})
    else:
        return jsonify({'arquivos': []})

@app.route('/success')
def success():
    return render_template('arquivos_gerados.html')

@app.route('/process', methods=['POST'])
def process():
    option = request.form['option']
    #global clid
    #clid = 'default'  # Define o valor padrão para evitar erros ao chamar sys.argv[1]
    uploaded_file = request.files['fileUpload']  # Obtém o arquivo enviado pelo formulário
    
    
    # Verifica se um arquivo foi enviado
    if uploaded_file.filename == '':
        return 'Nenhum arquivo selecionado.'
    
    if uploaded_file.filename != '':
        # Obtém o caminho absoluto para o diretório 'uploads'
        uploads_dir = os.path.join(app.root_path, 'uploads')
        os.makedirs(uploads_dir, exist_ok=True)
        
        # Verifica se o tipo de arquivo corresponde à opção escolhida
        file_option = 'national' if 'NATIONAL' in uploaded_file.filename.upper() else 'international'
        
        if option != file_option:
            error_message = f"O arquivo selecionado não corresponde à opção {option.upper()}"
            return render_template('index.html', error=error_message)
    
        # Salva o arquivo carregado no servidor
        file_path = os.path.join(uploads_dir, uploaded_file.filename)
        uploaded_file.save(file_path)
    
        # Verifica se o arquivo Excel foi corretamente carregado
        if os.path.exists(file_path):
           workbook = pd.ExcelFile(file_path)
        else:
            print(f"Caminho do arquivo: {file_path}")  # Adicione uma declaração de impressão para depuração
            return f"Arquivo para a opção {option.upper()} não encontrado!"
            
            # Processamento do arquivo com base na opção escolhida
            sheet_name = workbook.sheet_names[0]
    
        if os.path.exists(file_path):
           workbook = pd.ExcelFile(file_path)
        else:
            print(f"Caminho do arquivo: {file_path}")  # Adicione uma declaração de impressão para depuração
            return f"Arquivo para a opção {option.upper()} não encontrado!"

        file_path = f"CLID{clid}.xlsx"

        sheet_name = workbook.sheet_names[0]
        
        if sheet_name:
            sheet = workbook.parse(sheet_name)
            data = sheet.to_dict(orient='records')
            row_data = data[1:]
            grouped = sheet.groupby('Contrato')
            global contract_groups
            contract_groups = {key: group.to_dict(orient='records') for key, group in grouped}
            contract_keys = list(contract_groups.keys())
            
            #output_dir = os.path.join(os.getcwd(), 'output', clid)
            output_dir = os.path.join(os.getcwd(), 'output', option.upper())
            os.makedirs(output_dir, exist_ok=True)

            for key in contract_keys:
                group_list = list(map(lambda x: map_national(x) if option == 'national' else map_international(x), contract_groups[key]))
                group_list.sort(key=lambda x: int(x['Item Number']))

                new_workbook = pd.ExcelWriter(os.path.join(output_dir, f"CLID{key}.xlsx"), engine='xlsxwriter')
                
                pd.DataFrame(group_list).to_excel(new_workbook, sheet_name='Contract Item Information', index=False)
                pd.DataFrame([
                    {'Item Number': '', 'Attribute Name': '', 'Attribute Value': '', 'Display Text': '', 'Type': '', 'Description': ''}
                ]).to_excel(new_workbook, sheet_name='Item Attributes', index=False)

                new_workbook.close()
                print(f"Arquivo excel CLID{key}.xlsx {option.upper()} criado com sucesso!")
                
                    # Exemplo de redirecionamento para uma página de sucesso após processamento
            return redirect(url_for('success'))
        else:
            return f"Arquivo carregado não encontrado para a opção {option.upper()}!"

            #return f"Arquivos excel para CLID{clid} criados com a opção: {option.upper()}"
        #else:
            #return 'Nenhuma aba encontrada na planilha, favor verificar!'

if __name__ == '__main__':
    app.run(debug=True)
