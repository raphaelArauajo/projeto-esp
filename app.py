import os
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Garante que a pasta exista
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        fopa_file = request.files.get('fopa')
        matriz_file = request.files.get('matriz')

        if not fopa_file or not matriz_file:
            return "Por favor, envie os dois arquivos.", 400

        fopa_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(fopa_file.filename))
        matriz_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(matriz_file.filename))

        fopa_file.save(fopa_path)
        matriz_file.save(matriz_path)

        # Processamento dos dados
        fopa = pd.read_excel(fopa_path)
        especialidade_original = pd.read_excel(matriz_path)

        fopa_crm = fopa[fopa['DS_SIGLA_CONSELHO_REGIONAL'] == 'CRM']
        fopa_ativo = fopa_crm[fopa_crm['DS_STATUS'] == 'Ativo']
        fopa_desligados = fopa_crm[fopa_crm['DS_STATUS'] == 'Saiu da empresa']

        fopa_esp = pd.merge(fopa_ativo, especialidade_original, left_on='CD_DRT', right_on='DRT', how='left')
        fopa_drt = fopa_esp.drop(columns=[
            'Descrição da Área', 'Data da Contratação', 'Nome do gestor',
            'Endereço de Email', 'Descrição do Código do Cargo', 'Descrição de Unidade',
            'Status da ocupação', 'ID do gestor', 'Nome Completo'
        ])

        fopa_drt = fopa_drt.rename(columns={
            'DS_NOME': 'Nome Completo',
            'CD_DRT': 'DRT',
            'DS_AREA_RH': 'Descrição de Unidade',
            'DS_UNIDADE_ORGANIZACIONAL': 'Descrição da Área',
            'DT_ENTRADA': 'Data da Contratação',
            'Gestor': 'Nome do gestor',
            'DS_EMAIL': 'Endereço de Email',
            'DS_CARGO': 'Descrição do Código do Cargo',
            'DS_STATUS': 'Status da ocupação',
            'FL_AFASTAMENTO': 'afastado',
            'FL_DUPLO_CONTRATO': 'Duplo Contrato',
            'CD_DRT_GESTOR': 'ID do gestor',
            'QT_IDADE': 'Idade',
            'Matriz de Especialidades': 'Especialidades'
        })

        nova_ordem = ['DRT', 'CD_CPF','Nome Completo',"Idade", 'ID do Colaborador', 'Especialidades', 'Descrição do Código do Cargo','Descrição de Unidade',
                    'ID de unidade', 'afastado', 'Status da ocupação', 'Descrição da Área',
                    'Data da Contratação', 'Endereço de Email', 'ID do gestor', 'DS_NOME_GESTOR',
                    'Duplo Contrato','Observação']

        especialidade_atualizada = fopa_drt[nova_ordem]
        especialidade_atualizada = especialidade_atualizada.rename(columns={'CD_CPF': 'CPF'})

        email_gestor = pd.merge(fopa, especialidade_atualizada, left_on='CD_DRT', right_on='ID do gestor', how='inner')
        email_gestor = email_gestor.rename(columns={
            'CD_CPF_y': 'CPF',
            'DS_NOME_GESTOR_y': 'Nome do gestor',
            'DS_EMAIL': 'E-mail Gestor'
        })

        nova_ordem2 = ['DRT','CPF', 'Nome Completo','Idade', 'ID do Colaborador', 'Especialidades', 'Descrição do Código do Cargo','Descrição de Unidade',
                    'ID de unidade', 'afastado', 'Status da ocupação',  'Descrição da Área',
                    'Data da Contratação', 'Endereço de Email', 'ID do gestor', 'Nome do gestor',
                    'E-mail Gestor', 'Duplo Contrato','Observação']

        especialidade_final = email_gestor[nova_ordem2]
        especialidade_final = especialidade_final.rename(columns={'Especialidades': 'Matriz de Especialidades'})

        consultaEsp = pd.read_csv("Consulta_ESP.csv", sep=";")

        mapearEspecialidade = pd.merge(
            especialidade_final,
            consultaEsp,
            how='left',
            left_on=['Descrição do Código do Cargo', 'Descrição de Unidade', 'Descrição da Área'],
            right_on=['Descrição do Código do Cargo', 'Descrição de Unidade', 'Descrição da Área']
        )

        mapearEspecialidade = mapearEspecialidade.rename(columns={
            'Matriz de Especialidades_y': 'Matriz de Treinamento Sugerida'
        })
        mapearEspecialidade['Matriz de Treinamento Sugerida'] = mapearEspecialidade['Matriz de Treinamento Sugerida'].fillna('Nenhum padrão encontrado')

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'mapearEspecialidade.xlsx')
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            mapearEspecialidade.to_excel(writer, sheet_name='Especialidade Atualizada', index=False)
            fopa_desligados.to_excel(writer, sheet_name='Desligados', index=False)

        return send_file(output_path, as_attachment=True)

    return render_template('index.html')
    
if __name__ == '__main__':
    app.run(debug=True)
