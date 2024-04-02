import pandas as pd
import os
import glob

# caminho que vai ler os arquivos

folder_path = 'src\\data\\raw'

# Lista todos os arquivos de excel

excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print("Não foi encontrado arquivo compatível")
else:
    dfs = [] # Onde vai ficar alocado a tabela final

    for excel_file in excel_files:
        try:

            # Lê o arquivo de excel
            df_temp = pd.read_excel(excel_file)

            # Pega o nome do arquivo
            file_name = os.path.basename(excel_file)

            df_temp['filename'] = file_name

            # Criação de uma nova coluna chamada location
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'BR'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'FR'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'IT'

            # Criando uma nova coluna chamada campaign
            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            # Guarda os dados tratados dentro de um dataframe comum
            dfs.append(df_temp)


        except Exception as e:
            print(f"Erro ao ler o arquivo {excel_file} : {e}")

    if dfs:

        # Concatena todas as tabelas salvas no dfs em uma única tabela
        result = pd.concat(dfs, ignore_index=True)

        # Caminho de saída
        output_files = os.path.join('src', 'data', 'ready', 'clean.xlsx')

        # Configurando do motor de escrita
        writer = pd.ExcelWriter(output_files, engine='xlsxwriter')

        # Leva os dados do resultado a serem escritos no motor de excel configurado
        result.to_excel(writer, index=False)

        # Salva o arquivo de excel
        writer._save()
    else:
        print("Nenhum dado para ser salvo")




