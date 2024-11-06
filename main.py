import os
import re
import json
import pandas as pd
import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from tkinter import *
from tkinter import filedialog

from tkinter import messagebox

#from tkinter.ttk import Progressbar
from tkinter import ttk
#import threading
import random

#import undetected_chromedriver as uc

import glob

import matplotlib.pyplot as plt

from dash.dependencies import Output, Input
import tkinter as tk
from matplotlib import pyplot as plt
from dash import Dash
from dash import dcc
from dash import html
from threading import Thread
import webbrowser


from fake_useragent import UserAgent

# Versão chromedriver estável: https://googlechromelabs.github.io/chrome-for-testing/#stable

list_of_files = glob.glob('classificacoes*.xlsx') # * significa qualquer coisa
latest_file = max(list_of_files, key=os.path.getctime)
qualis_data = pd.read_excel(latest_file)[['ISSN', 'Título', 'Estrato']]

#qualis_data = pd.read_excel(r'classificacoes_publicadas_biotecnologia_2022_1672761158359.xlsx')[['ISSN', 'Título', 'Estrato']]

os.environ["webdriver.chrome.driver"] = "chromedriver.exe"

#instructions = True

def alternar(Instructions):
    # Verifica o estado atual e altera para o oposto
    if instructions:
        instructions = False
    else:
        instructions = True
    return instructions
        
def extrair_dados(data, output_dir):
  file_path = os.path.join(output_dir, 'out_json_professores.json')
  with open(file_path, "r") as f:
    data = json.load(f)
  print(output_dir)
  

  info_list = []
  
  for professor_nome, professor_dados in data.items():

      issn_pattern = r"issn=(\w+)"
      matches = re.findall(issn_pattern, professor_dados)


      soup = BeautifulSoup(professor_dados, 'html.parser')
      nome = soup.find('h2', class_='nome').text.strip()
      bolsista_info = None
      try:
        bolsista_info = soup.find_all('h2', class_='nome')[1].text.strip()


        h2_elements = soup.find_all('h2', class_='nome')
        if len(h2_elements) > 1:
            bolsista_info = h2_elements[1].text.strip()
        else:
            bolsista_info = '---'
      except:
          pass 
        
      informacoes_autor = soup.find('ul', class_='informacoes-autor').find_all('li')
      endereco_cv = informacoes_autor[0].text.strip()
      endereco_cv = endereco_cv.split()[-1]
      id_lattes = informacoes_autor[1].text.strip().split(': ')[1]
      ultima_atualizacao = informacoes_autor[2].text.strip()
      ultima_atualizacao = ultima_atualizacao.split()[-1]
      total_artigos = len(soup.find_all('div', class_='artigo-completo'))
      total_orientacoes = str(professor_dados).count(nome + " - Coordenador")


      info_dict = {'Docente':nome, 'Orientações':total_orientacoes, 'Bolsista Info':bolsista_info, 'Endereço CV':endereco_cv, 'ID Lattes':id_lattes, 'Última Atualização':ultima_atualizacao, 'Total de Artigos':total_artigos}
      info_list.append(info_dict)
      dfgeral = pd.DataFrame(info_list)
      dfgeral.to_excel(os.path.join(output_dir, 'DocentesPPG.xlsx'))
      print(dfgeral)
    

      soup = BeautifulSoup(professor_dados, 'html.parser')
      articles = soup.find_all('div', class_='artigo-completo')
      articles_info_list = []
      for i, article in enumerate(articles):
          article_dict = {}

          try:
            issn = matches[i]
            issn = issn[:4] + '-' + issn[4:]
            article_dict['ISSN'] = issn
          except:
              article_dict['ISSN'] = None

          qualis_info = find_issn_qualis(article_dict['ISSN'], qualis_data)
          article_dict['Qualis'] = qualis_info


          article_dict['Índice do Artigo'] = i+1


          #informacoes_gerais_artigo = ' '.join(article.text.split()[5:-3])
          #article_dict['Título do artigo'] = ' '.join(article.text.split()[5:-3])
          
          citation_div = article.find('div', class_='citado')
          
          revista_element = article.find('img', {'class': 'ajaxJCR'})


          if citation_div and 'nomePeriodico=' in citation_div.get('cvuri', ''):
              periodico = citation_div.get('cvuri').split('nomePeriodico=')[1].split('&')[0]

              article_dict['Periódico'] = periodico
              
              
          elif revista_element and revista_element.get('original-title'):
              periodico = revista_element.get('original-title').split('<br')[0]
              article_dict['Periódico'] = periodico
              article
          
          else:
            try:
              layout_cell_pad = article.find_all('div', class_='layout-cell-pad-5')
              nome_periodico = 'N/A'
              
              if len(layout_cell_pad) > 1:
                  periódico_data = layout_cell_pad[1].text.strip()
                  nome_periodico_search = re.search(r'\.\s(.*?), v\.', periódico_data)

                  if nome_periodico_search:
                      nome_periodico_split = nome_periodico_search.group(1).strip().split('. ')
                      nome_ultimo_elemento = nome_periodico_split[-1]

                      if nome_ultimo_elemento.endswith(')') and '(' not in nome_ultimo_elemento:
                          nome_periodico_split2 = nome_periodico_search.group(1).strip().split(' (')
                          nome_ultimo_elemento2 = ' ('.join(nome_periodico_split2[:-1])
                          nome_ultimo_elemento2 = nome_ultimo_elemento2.split('. ')
                          nome_ultimo_elemento_final = nome_ultimo_elemento2[-1]
                          nome_ultimo_elemento = nome_ultimo_elemento_final + " (" + nome_periodico_split2[-1]
                      nome_periodico = nome_ultimo_elemento
                      article_dict['Periódico'] = nome_periodico
            except:
              article_dict['Periódico'] = "---"

          jcr_element = article.find('span', {'data-tipo-ordenacao': 'jcr'})
          if jcr_element:
              article_dict['JCR'] = jcr_element.text 
          else:
              article_dict['JCR'] = "---"

          article_soup = BeautifulSoup(str(article), 'html.parser')
          
            


          
          article_year = article_soup.find('span', attrs = {'data-tipo-ordenacao':'ano'}).text # Para extrair o ano
          article_dict['Ano'] = article_year  # Adicionar o ano do artigo ao dicionário
          

          # find publication type and format it
          #publication_type_span = article_soup.find('a', attrs = {'name': True})
          #if publication_type_span:
          #    publication_type = publication_type_span['name']
          #    formatted_publication_type = re.sub(r"(?<!^)(?=[A-Z])", " ", publication_type)  # adds a space before each capital letter that's not at the start
          #    article_dict['Tipo de Publicação'] = formatted_publication_type


          articles_info_list.append(article_dict)

      df = pd.DataFrame(articles_info_list).set_index('Índice do Artigo')
      df.to_excel(os.path.join(output_dir, professor_nome+'.xlsx'))
      #df.to_csv(os.path.join(output_dir, professor_nome+'.csv'))
      

    #dfgeral.to_excel(f'Informações_Gerais.xlsx')

def find_issn_qualis(issn: str, qualis_data: pd.DataFrame):
    linha = qualis_data[qualis_data['ISSN'] == issn]
    return linha['Estrato'].values[0] if not linha.empty else '---'

def wait_and_find(driver, by_what, identifier):
    return WebDriverWait(driver, 10).until(EC.presence_of_element_located((by_what, identifier)))

def get_html(driver, proff):
    url = "http://buscatextual.cnpq.br/buscatextual/busca.do"
    
    #driver = uc.Chrome()
    #driver = webdriver.Chrome()
    driver.get(url)

    search_bar = wait_and_find(driver, By.ID, 'textoBusca')
    search_bar.send_keys(proff)
    time.sleep(1)
    search_bar.send_keys(Keys.ENTER)
    
    link = wait_and_find(driver, By.XPATH, f'//a[starts-with(@href, "javascript:abreDetalhe")]')
    time.sleep(1)
    link.click()

    wait_and_find(driver, By.ID, "idbtnabrircurriculo").click()
    time.sleep(2)
    
    driver.switch_to.window(driver.window_handles[1])
    html = driver.page_source
    
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return html

def get_htmls(professors):
    #chrome_options = Options()

    #ua = UserAgent()
    #user_agent = ua.random

    #chrome_options.add_argument("--incognito")

    #chrome_options.add_argument(f'--user-agent={user_agent}')

    #chrome_options.add_argument("--start-maximized")
    #driver = ""
    def get_random_proxy(filename):
        with open(filename, 'r') as f:
            proxies = f.read().splitlines()
            return random.choice(proxies)
    #driver = uc.Chrome()
    def change_useragent(instructions):

        chrome_options = Options()

        ua = UserAgent()
        user_agent = ua.random

        if instructions == 1:
            chrome_options.add_argument("--incognito")
            chrome_options.add_argument(f'--user-agent={user_agent}')
            #chrome_options.add_argument("--headless=new") Se atentar aos proxies (estão com erro)
            #proxies = get_random_proxy('proxies.txt')
            #chrome_options.add_argument(f'--proxy-server=http://{proxies}')
            print()

            driver = webdriver.Chrome(options=chrome_options)
        else:
            #chrome_options.add_argument("--headless=new")
            driver = webdriver.Chrome()
        
        return driver


    html_dict = {}
    errored_professors = []
    final_error = []

    try:
        
        driver = change_useragent(1)
        for index, proff in enumerate(professors, 1):

            print(f"Procurando: {proff}")
            try:
                html_dict[proff] = get_html(driver, proff)
            except Exception as e:
                print(f'Erro com o professor: {proff}. Detalhes: {str(e)}')
                errored_professors.append(proff)
                driver = change_useragent(0)
                driver.delete_all_cookies() 

            if index % 3 == 0:
                print("Deletando Coockies")
                driver.delete_all_cookies() 

        while errored_professors:  # Se ainda há professores com erro.
            proff = errored_professors.pop()
            driver = webdriver.Chrome()
            try:
                print(f"Tentando novamente: {proff}")
                #driver = change_useragent(False)
                html_dict[proff] = get_html(driver, proff)
            except Exception as e:
                print(f'Erro repetido com o professor: {proff}. Detalhes: {str(e)}')
                final_error.append(proff)

        # professores que não puderam ser encontrados
        if final_error:
            error_message = "\n".join(final_error)
            messagebox.showinfo("Não foi encontrado", f"Os seguintes professores não foram encontrados:\n\n{error_message}")
            for proff in final_error:
                print(proff)
    finally:
        #driver.service.process.send_signal(15)
        pass
        #driver.quit()

    return html_dict

def start(input_file, output_dir):
    # input file
    professors = pd.read_excel(input_file)['Docente'].tolist()
    html_dict = get_htmls(professors)

    # HTML p/ JSON
    with open(os.path.join(output_dir, 'out_json_professores.json'), 'w') as f:
        json.dump(html_dict, f)

    messagebox.showinfo("Iniciando a criação das planilhas...", f"As planilhas serão criadas e salvas no seguinte diretório: {output_dir}")
    # Extract the data from the HTML
    extrair_dados(html_dict, output_dir)
    messagebox.showinfo("Concluído!", f"Salvo em: {output_dir}")
    resposta = messagebox.askyesno("Gerar Dashboard?", "Deseja gerar um dashboard agora?")
    if resposta:
        gerar_dashboard(output_dir)


def gerar_dashboard(path):



    def load_dataframes_from_directory(path):
        all_dfs = []

        for file in os.listdir(path):
            if file.endswith(".xlsx"):
                file_path = os.path.join(path, file)
                df = pd.read_excel(file_path, engine='openpyxl')
                professor_name = os.path.splitext(file)[0]
                df['Docente'] = professor_name
                all_dfs.append(df)

        if not all_dfs:
            print('Nenhum arquivo .xlsx encontrado na pasta selecionada.')
            return
        return pd.concat(all_dfs, ignore_index=True)

    def run_dash(path, url):
        app = Dash(__name__)
        total_df = load_dataframes_from_directory(path)

        # pegar os anos para o slider
        min_year = total_df['Ano'].min()
        max_year = total_df['Ano'].max()

        app.layout = html.Div([
            # coluna para os controles
            html.Div([
                html.P(f"Intervalo de Anos ({int(min_year)} - {int(max_year)}):"),
                dcc.RangeSlider(
                    id='year-slider',
                    min=min_year,
                    max=max_year,
                    step=1,
                    value=[min_year, max_year],
                    marks={i: '{}'.format(i) if i % 5 == 0 or i == min_year else '' for i in range(int(min_year), int(max_year) + 1)},
                    tooltip={"placement": "bottom", "always_visible": True},
                    
                ),
                html.P("Professores:"),
                dcc.Dropdown(
                    id='professor-dropdown',
                    options=[{'label': i, 'value': i} for i in total_df['Docente'].unique()],
                    value=[i for i in total_df['Docente'].unique()],
                    multi=True
                ),
                html.Button('Recarregar Dados', id='reload-button', n_clicks=0)
            ], style={'position': 'fixed', 'width': '35%', 'height': '100%', 'display': 'inline-block', 'vertical-align': 'top'}),#style={'position': 'fixed', 'top': 0, 'left': 0, 'width': '50%', 'height': '100%'})

            # Espaço para os gráficos
            html.Div([   
                dcc.Graph(id="bar-chart-docentes"),
                dcc.Graph(id="line-chart-publicacoes"),
                dcc.Graph(id="bar-chart-periodicos"),
                dcc.Graph(id="bar-chart-qualis"),
                dcc.Graph(id="line-chart-professors"),
            ], style={'position': 'relative', 'marginLeft': '35%', 'overflow': 'auto', 'maxHeight': '100vh', 'width': '65%'})
        ])

        # Callbacks
        # Callback para recarregar dados
        @app.callback(
            [Output('bar-chart-docentes', 'figure'),
            Output('line-chart-publicacoes', 'figure'),
            Output('bar-chart-periodicos', 'figure'),
            Output('bar-chart-qualis', 'figure'),
            Output('line-chart-professors', 'figure')],
            [Input('year-slider', 'value'),
            Input('professor-dropdown', 'value'),
            Input('reload-button', 'n_clicks')])
        
        def reload_data(year_range, selected_professors, n):
            total_df = load_dataframes_from_directory(path)
            filtered_df = total_df[(total_df['Ano'] >= year_range[0]) & (total_df['Ano'] <= year_range[1]) & total_df['Docente'].isin(selected_professors)]
            return update_graph(filtered_df)
        
        def update_data(year_range, selected_professors):
            global total_df
            filtered_df = total_df[(total_df['Ano'] >= year_range[0]) & (total_df['Ano'] <= year_range[1]) & total_df['Docente'].isin(selected_professors)]
            return update_graph(filtered_df)
        


        def update_graph(data):
            professor_data = []
            for professor in data['Docente'].unique():
                professor_row = data[data['Docente'] == professor]['Ano'].value_counts().sort_index()
                professor_data.append({'x': professor_row.index, 'y': professor_row.values, 'type': 'line', 'name': professor})

            publications_per_year = data['Ano'].value_counts().sort_index()



            figures = []

            figures.append(
                {
                    'data': [
                        {
                            'x': data['Docente'].value_counts().index,
                            'y': data['Docente'].value_counts().values,
                            'type': 'bar',
                            'name': 'Número de Artigos por Docente',
                            'marker': {'color': '#008bff'}
                        }
                    ],
                    'layout': {'title': 'Número total de artigos por docente'}
                }
            )


            figures.append(
                {
                    'data': [
                        {
                            'x': data['Periódico'].value_counts().index[:10],
                            'y': data['Periódico'].value_counts().values[:10],
                            'type': 'bar',
                            'name': 'Número de Publicações',
                            'marker': {'color': '#008bff'}
                        }
                    ],
                    'layout': {'title': 'Top 10 periódicos por número de publicações'}
                }
            )

            figures.append(
                {
                    'data': [
                        {
                            'x': data['Qualis'].value_counts().index,
                            'y': data['Qualis'].value_counts().values,
                            'type': 'bar',
                            'name': 'Número de Artigos',
                            'marker': {'color': '#008bff'}
                        }
                    ],
                    'layout': {'title': 'Distribuição de artigos por Qualis'}
                }
            )

            figures.append(
                {
                    'data': [
                        {
                            'x': publications_per_year.index,
                            'y': publications_per_year.values,
                            'type': 'line',
                            'name': 'Número total de Publicações por Ano',
                        }
                    ],
                    'layout': {'title': 'Publicações por ano'}
                }
            )
            
            figures.append(
                {
                    'data': professor_data,
                    'layout': {'title': 'Evolução de publicações ao longo dos anos por docente'}
                }
            )

            return figures

        time.sleep(2)
        webbrowser.open(url)
        
        app.run_server(debug=True, use_reloader=False)

    def select_directory_and_run_dash():
        root = tk.Tk()
        root.withdraw()

        df = load_dataframes_from_directory(path)

        if df is None:
            return
        dash_thread = Thread(target=run_dash, args=(path, "http://127.0.0.1:8050/"))
        dash_thread.start()

    select_directory_and_run_dash()




def main():
    root = Tk()
    root.title("Busca Lattes")
    root.geometry("640x640")  # tamanho geral da janela
    root.configure(bg='#f0f0f0')
    root.resizable(False, False)  # Desabilita a maximização da janela
    style = ttk.Style()
    style.configure('TButton', font=('Arial', 16, 'bold'), borderwidth='1', background='#67a8a6')  # Letras bold e maiores com background azul acinzentado
    style.configure('TLabel', font=('Arial', 12), background='#f0f0f0')

    progress_bar = ttk.Progressbar(root, mode='indeterminate', length=500)

    
    def select_input_file():
        input_file = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*" )))
        input_file_entry.delete(0, END)
        input_file_entry.insert(0, input_file)

    def select_output_dir():
        output_dir = filedialog.askdirectory()
        output_dir_entry.delete(0, END)
        output_dir_entry.insert(0, output_dir)


    def process():
        # Função para processar as informações
        start(input_file_entry.get(), output_dir_entry.get())

    def check_thread(thread):
        if thread.is_alive():
            root.after(500, check_thread, thread)
        else:
            progress_bar.stop()
            start_button['state'] = NORMAL

    # Define a imagem de fundo
    bg_image = PhotoImage(file="background.png") # "background.png" pelo caminho para a imagem de fundo
    bg_label = Label(root, image=bg_image)
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    input_frame = Frame(root, bg='#f0f0f0')
    input_frame.pack(pady=60) 
    input_file_label = ttk.Label(input_frame, text="Selecione o arquivo com a coluna |Docente|")
    input_file_label.grid(row=1, column=1, padx=10, sticky='w')
    input_file_entry = ttk.Entry(input_frame, width=40)
    input_file_entry.grid(row=2, column=1, padx=10, pady=5)
    input_file_button = ttk.Button(input_frame, text="Selecione o arquivo", command=select_input_file)
    input_file_button.grid(row=3, column=1, padx=10, pady=5)

    output_frame = Frame(root, bg='#f0f0f0')
    output_frame.pack(pady=20)
    output_dir_label = ttk.Label(output_frame, text="Selecione o diretório de saída")
    output_dir_label.grid(row=1, column=1, padx=10, sticky='w')
    output_dir_entry = ttk.Entry(output_frame, width=40)
    output_dir_entry.grid(row=2, column=1, padx=10, pady=5)
    output_dir_button = ttk.Button(output_frame, text="Selecione o diretório", command=select_output_dir)
    output_dir_button.grid(row=3, column=1, padx=10, pady=5)

    start_button = ttk.Button(root, text="Iniciar busca", command=process)
    start_button.pack(padx=100, pady=40)

    # Barra de Progresso
    progress_bar.pack(pady=20)

    root.mainloop()


  
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
