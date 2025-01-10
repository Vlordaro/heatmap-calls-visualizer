import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import calendar
import tkinter as tk
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.colors as mcolors
from matplotlib.widgets import Cursor
from tkinter import filedialog
import matplotlib.backends.backend_pdf
import io
import xlsxwriter

# Carregar credenciais do arquivo
credentials = Credentials.from_service_account_file('credentials.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
gc = gspread.authorize(credentials)
sh = gc.open_by_key('your-spreadsheet-id').sheet1

# Função para ler os dados da planilha
def read_data_from_worksheet():
   data = sh.get_all_records()
   df = pd.DataFrame(data)
   df['Created'] = pd.to_datetime(df['Created'], dayfirst=True)
   return df

# Função para retornar os anos disponíveis com base nos dados da planilha
def get_available_years():
   df = read_data_from_worksheet()
   valid_years = df['Created'].dt.year.dropna().unique().astype(int)
   return sorted(valid_years)

# Função para plotar o heatmap
def plot_monthly_heatmap(year, month):
   plt.close()  # Fechar a figura anterior
   if hasattr(app, 'canvas'):
       app.canvas.get_tk_widget().destroy()
   if hasattr(app, 'export_button'):  # Verificar se o botão existe
       app.export_button.destroy()  # Remover o botão antigo
   df = read_data_from_worksheet()
   if df is None:
       print("Não foi possível ler os dados da planilha.")
       return
   df_selected_month = df[(df['Created'].dt.year == year) & (df['Created'].dt.month == month)].copy()
   if df_selected_month.empty:
       print("Não há dados disponíveis para este mês.")
       return

   # Criar uma sequência de todas as datas e horas do mês
   all_days_month = pd.date_range(start=f'{year}-{month:02d}-01', end=f'{year}-{month:02d}-{calendar.monthrange(year, month)[1]} 23:00:00', freq='h')

   # Criar um DataFrame com todas as datas e horas do mês
   all_days_hours_df = pd.DataFrame({'datetime': all_days_month})
   all_days_hours_df['day'] = all_days_hours_df['datetime'].dt.day
   all_days_hours_df['hour'] = all_days_hours_df['datetime'].dt.hour

   df_selected_month['day'] = df_selected_month['Created'].dt.day
   df_selected_month['hour'] = df_selected_month['Created'].dt.hour

   # Contar o número de ocorrências para cada combinação de dia e hora
   hour_counts = df_selected_month.groupby(['day', 'hour']).size().reset_index(name='count')

   # Combinar todas as datas e horas do mês com as contagens de atendimento
   heatmap_data = all_days_hours_df.merge(hour_counts, how='left', on=['day', 'hour']).fillna(0)
   heatmap_data = heatmap_data.pivot(index='hour', columns='day', values='count').fillna(0)
   heatmap_data = heatmap_data.iloc[::-1]

   colors = ['#009929', 'yellow', 'orange', '#db0e0e']
   cmap = mcolors.LinearSegmentedColormap.from_list("custom_cmap", colors)

   fig, ax = plt.subplots(figsize=(10, 6))
   sns.heatmap(heatmap_data, cmap=cmap, linecolor='white', linewidths=1, ax=ax)
   ax.set_title(f'Distribuição de chamados ao longo do mês - {calendar.month_name[month]} de {year}')
   ax.set_xlabel('Dia do mês')
   ax.set_ylabel('Hora do dia')

   # Calcular o número total de chamados e o tempo médio de atendimento
   total_chamados = len(df_selected_month)
   df_selected_month['Closed'] = pd.to_datetime(df_selected_month['Closed'], dayfirst=True)
   df_selected_month['Tempo de Atendimento'] = (df_selected_month['Closed'] - df_selected_month['Created']).dt.total_seconds() / 3600
   tempo_medio_atendimento = df_selected_month['Tempo de Atendimento'].mean()

  # Criar um quadrado com as informações
   textstr = f'Chamados: {total_chamados } \nTempo Médio: {tempo_medio_atendimento:.2f}h/atendimento'
   props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
   
   # Ajustar a posição do quadrado
   ax.text(1.15, 0.95, textstr, transform=ax.transAxes, fontsize=10,
           verticalalignment='top', bbox=props)

   # Criar o cursor interativo
   cursor = Cursor(ax, useblit=True, color='white', linewidth=1)

   # Adicionar evento de clique no heatmap
   def on_click(event):
       if event.inaxes == ax:
           row = int(event.ydata)
           col = int(event.xdata)
           hour = 23 - row  # Invertemos a ordem das horas
           day = col + 1
           
           # Filtrar os dados da planilha pelo dia e hora selecionados
           filtered_df = df_selected_month[(df_selected_month['hour'] == hour) & (df_selected_month['day'] == day)]
           
           # Exibir as informações na janela
           if not filtered_df.empty:
               info_window = tk.Toplevel(app)
               info_window.title(f"Informações do dia {day}, {hour:02d}:00")
               for i, row in filtered_df.iterrows():
                   customer = row['Customer']
                   category = row['Category']
                   tk.Label(info_window, text=f"BKN: {customer} - Categoria: {category}").pack()
           else:
               tk.messagebox.showinfo("Informações", "Nenhum chamado registrado para este dia e hora.")

   fig.canvas.mpl_connect('button_press_event', on_click)

   canvas = FigureCanvasTkAgg(fig, master=app)
   canvas.draw()
   canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
   app.canvas = canvas
   app.after(60000, refresh_heatmap)

   # Botão para exportar o gráfico
   app.export_button = tk.Button(app, text="Exportar Gráfico", command=lambda: export_graph(fig))  # Criar o botão
   app.export_button.pack(side=tk.BOTTOM)

# Função para exportar o gráfico
def export_graph(fig):
   def export_as_png():
       file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
       if file_path:
           fig.savefig(file_path)
           export_options.destroy()  # Fechar a janela de opções

   def export_as_pdf():
       file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
       if file_path:
           pdf = matplotlib.backends.backend_pdf.PdfPages(file_path)
           pdf.savefig(fig)
           pdf.close()
           export_options.destroy()  # Fechar a janela de opções

   def export_as_excel():
       file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
       if file_path:
           # Criar um buffer em memória para o Excel
           buffer = io.BytesIO()
           workbook = xlsxwriter.Workbook(buffer)
           worksheet = workbook.add_worksheet()

           # Obter os dados do heatmap
           heatmap_data = fig.axes[0].get_array()

           # Escrever os dados no Excel
           for row_index, row in enumerate(heatmap_data):
               for col_index, value in enumerate(row):
                   worksheet.write(row_index, col_index, value)

           workbook.close()

           # Salvar o arquivo Excel
           with open(file_path, "wb") as f:
               f.write(buffer.getvalue())
           export_options.destroy()  # Fechar a janela de opções

   # Criar a janela de opções de exportação
   export_options = tk.Toplevel(app)
   export_options.title("Exportar Gráfico")

   # Criar os botões para cada formato
   png_button = tk.Button(export_options, text="PNG", command=export_as_png)
   png_button.pack(side=tk.TOP, padx=10, pady=10)

   pdf_button = tk.Button(export_options, text="PDF", command=export_as_pdf)
   pdf_button.pack(side=tk.TOP, padx=10, pady=10)

   excel_button = tk.Button(export_options, text="Excel", command=export_as_excel)
   excel_button.pack(side=tk.TOP, padx=10, pady=10)

# Função para atualizar o heatmap
def refresh_heatmap():
   plot_monthly_heatmap(ano_var.get(), int(mes_var.get()))

# Função para lidar com a seleção de ano ou mês
def handle_selection(event):
   plot_monthly_heatmap(ano_var.get(), int(mes_var.get()))

# Função para alternar a visibilidade dos widgets de seleção de período
def toggle_period_selection_visibility(visible):
   if visible:
       ano_label.pack(side=tk.TOP)
       ano_dropdown.pack(side=tk.TOP)
       mes_label.pack(side=tk.TOP)
       mes_dropdown.pack(side=tk.TOP)
   else:
       ano_label.pack_forget()
       ano_dropdown.pack_forget()
       mes_label.pack_forget()
       mes_dropdown.pack_forget()

# Função para colocar a aplicação em tela cheia
def toggle_fullscreen(event=None):
   fullscreen = not app.attributes("-fullscreen")
   app.attributes("-fullscreen", fullscreen)
   toggle_period_selection_visibility(not fullscreen)

# Criar a aplicação
app = tk.Tk()
app.title("Heatmap Visualizer")

# Criar widgets para seleção de ano e mês
ano_label = ttk.Label(app, text="Ano:")
ano_label.pack(side=tk.TOP)
ano_var = tk.IntVar()
ano_dropdown = ttk.Combobox(app, textvariable=ano_var, values=get_available_years())
ano_dropdown.pack(side=tk.TOP)
ano_dropdown.bind("<<ComboboxSelected>>", handle_selection)

mes_label = ttk.Label(app, text="Mês:")
mes_label.pack(side=tk.TOP)
mes_var = tk.StringVar()  # Trocando para StringVar aqui
mes_dropdown = ttk.Combobox(app, textvariable=mes_var, values=[str(i) for i in range(1, 13)])  # Convertendo para string aqui
mes_dropdown.current(0)
mes_dropdown.pack(side=tk.TOP)
mes_dropdown.bind("<<ComboboxSelected>>", handle_selection)

# Atualizar o heatmap inicialmente
handle_selection(None)

# Binding da tecla F11 para alternar tela cheia
app.bind("<F11>", toggle_fullscreen)

# Alterar a chamada para toggle_period_selection_visibility para garantir que os widgets estejam visíveis inicialmente
toggle_period_selection_visibility(True)

app.mainloop()
