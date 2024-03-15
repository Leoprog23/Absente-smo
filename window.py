# Importando as bibliotecas necessárias

from tkinter import *
import pyodbc
from tkinter import messagebox
from tkcalendar import DateEntry
import pandas as pd
from tkinter import ttk
from tkinter import filedialog, simpledialog

# Criando a conexão com o banco de dados utilizado (SQLITE):
caminho_db = r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Banco de dados\Absenteísmo.db'
conexão = pyodbc.connect("Driver={SQLite3 ODBC Driver};"
                        "Server=localhost;"
                        f"Database={caminho_db}")
cursor = conexão.cursor()
# Criando a variável que irá receber as informações da lista de Funcionários:
Tabela = pd.read_excel(r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Lista funcionários.xlsx')




########################################################### Botão que adiciona as informações no banco de dados: ###########################################################

def BotãoConfirmarInformações():
    # Verificando se o campo de matrícula está preenchido (Tornando- o um campo obrigatório para funcionamento do sistema)
    if not campo_matrícula.get('1.0', END).strip():
        messagebox.showerror(title='Erro', message='Por favor, preencha o campo de matrícula.')
        return
    Área = campo_área.get()
    Matrícula = campo_matrícula.get('1.0', END)
    Nome = campo_nome.get('1.0', END)
    Turno = campo_turno.get('1.0', END)
    Condutor = campo_condutor.get('1.0', END)
    Supervisor = campo_supervisor.get('1.0', END)
    Motivo = campo_motivo.get()
    Data = campo_data.get()
    Observação = campo_observação.get('1.0', END)
    comando = f"""INSERT INTO Absenteísmo(Área, Matrícula, Nome, Turno, Condutor, Supervisor, Motivo, Data, Observação)
        VALUES
            ('{Área}', '{Matrícula}', '{Nome}', '{Turno}', '{Condutor}', '{Supervisor}', '{Motivo}', '{Data}', '{Observação}')"""
    cursor.execute(comando)
    cursor.commit()
    # Executando a função de preencher a tabela para inserir as informações dinamicamente:
    preencher_tabela()
    # Exibindo um mensagem para o usuário de que a inserção foi um sucesso:
    messagebox.showinfo(title='Alerta de inserção!', message=f'As informações para a matrícula {Matrícula} do(a) colaborador(a) {Nome} foram inseridas com sucesso!\nClique em "Ok" para prosseguir!')


########################################################### Botão que limpa as informações dos campos de inserção para cadastro dos dados: ###########################################################

def BotãoLimparInformações():
    campo_matrícula.delete('1.0', END)
    campo_nome.delete('1.0', END)
    campo_turno.delete('1.0', END)
    campo_condutor.delete('1.0', END)
    campo_supervisor.delete('1.0', END)
    campo_observação.delete('1.0', END)
    campo_motivo.set("")
    campo_área.set("")
    messagebox.showinfo(title='Alerta de modificação!', message='Informações redefinidas, você já pode inserir novos dados!')


########################################################### Botão que extrai o banco de dados para um Excel: ###########################################################
def BotãoExportarExcel():
    # Obtendo os dados da tabela:
    dados_tabela = []
    for linha in campo_tabela.get_children():
        dados_linha = campo_tabela.item(linha, 'values')
        dados_tabela.append(dados_linha)
    # Criar um DataFrame pandas com os dados:
    df = pd.DataFrame(dados_tabela, columns=['Área', 'Matrícula', 'Nome', 'Turno', 'Condutor', 'Supervisor', 'Motivo', 'Data', 'Observação'])
    # Funcionalidade que interage com o usuário, requerendo um local para salvar o arquivo que será criado:
    nome_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
    if nome_arquivo:
        df.to_excel(nome_arquivo, index=False)
        messagebox.showinfo(title='Alerta de exportação', message=f"Os dados foram exportados com sucesso no caminho {nome_arquivo}")


########################################################### Botão que atualiza as informações diretamente no banco de dados: ###########################################################

def BotãoAtualizarBD():
    # Mapear as opções de coluna para o nome da coluna correspondente no banco de dados
    opções_coluna = {
        "Área": "Área",
        "Matrícula": "Matrícula",
        "Nome": "Nome",
        "Turno": "Turno",
        "Condutor": "Condutor",
        "Supervisor": "Supervisor",
        "Motivo": "Motivo",
        "Data": "Data",
        "Observação": "Observação"
    }
    # Obter a linha selecionada na tabela
    seleção = campo_tabela.selection()
    if seleção:
        # Obter o ID da linha selecionada
        id_linha = campo_tabela.item(seleção, 'values')[0]
        # Abra uma janela para o usuário inserir a nova informação
        nova_coluna = simpledialog.askstring("Modificar Informação", "Qual coluna você deseja modificar?")
        if nova_coluna is not None:
            # Verificar se a opção de coluna inserida pelo usuário está no dicionário de opções
            if nova_coluna in opções_coluna:
                # Obter o nome da coluna correspondente à opção inserida
                nome_coluna = opções_coluna[nova_coluna]
                # Abra uma nova janela para o usuário inserir a nova informação
                nova_informação = simpledialog.askstring("Modificar Informação", f"Insira a nova informação para a coluna {nova_coluna}:")
                if nova_informação is not None:
                    # Atualizar a informação no banco de dados
                    comando_sql = f"UPDATE Absenteísmo SET {nome_coluna} = ? WHERE Id = ?"
                    cursor.execute(comando_sql, (nova_informação, id_linha))
                    conexão.commit()
                    messagebox.showinfo("Alerta de modificação", "Informação atualizada com sucesso!")
                    # Atualizar a tabela após a modificação
                    preencher_tabela()
            else:
                messagebox.showerror("Erro", "Coluna inválida.")
    else:
        messagebox.showerror("Erro", "Selecione uma linha para modificar.")


############################################################ Criando a janela: ############################################################

window = Tk()
# Modificando o título da janela:
window.title('Controle de Absenteísmo')
window.geometry("1737x785")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 785,
    width = 1737,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")

canvas.place(x = 0, y = 0)

############################################################ Criando e definindo o background: ############################################################

background_img = PhotoImage(file = r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Imagens\background.png')
background = canvas.create_image(
    844.0, 388.0,
    image=background_img)

############################################################ Criando e definindo o botão "Exportar para Excel": ############################################################

imagem_exportar_excel = PhotoImage(file = r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Imagens\img0.png')
exportar_excel = Button(
    image = imagem_exportar_excel,
    borderwidth = 0,
    highlightthickness = 0,
    command = BotãoExportarExcel,
    relief = "flat")

exportar_excel.place(
    x = 1621,
    y = 223,
    width = 53,
    height = 60)


############################################################ Criando e definindo as informações do botão de "Confirmar Informações": ############################################################

style = ttk.Style()
style.configure("RoundedButton.TButton", borderwidth=5, relief="flat", border="10")

# Criando o botão:
botão_confirmar_informações = ttk.Button(
    text="Confirmar Dados",
    compound="left",
    command=BotãoConfirmarInformações,
    style="RoundedButton.TButton",
    cursor='hand2'
)
# Colocando o botão na janela:
botão_confirmar_informações.place(
    x=140,
    y=730,
    width=110,
    height=30
)

############################################################ Criando e definindo as informações do botão de "Atualizar o Banco de Dados" ############################################################

imagem_atualizar_BD = PhotoImage(file = r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Imagens\img3.png')
botão_atualizar_BD = Button(
    image = imagem_atualizar_BD,
    borderwidth = 0,
    highlightthickness = 0,
    command = BotãoAtualizarBD,
    relief = "flat")

botão_atualizar_BD.place(
    x = 1665,
    y = 225,
    width = 59,
    height = 54)

############################################################ Criando e definindo as informações do botão de "Limpar Informações" ############################################################
style = ttk.Style()
style.configure("RoundedButton.TButton", borderwidth=5, relief="flat", border="10")
# Criando o botão:
botão_limpar_informações = ttk.Button(
    text="Limpar Dados",
    compound="left",
    command=BotãoLimparInformações,
    style="RoundedButton.TButton",
    cursor='hand2'
)

# Inserindo o botão na janela:
botão_limpar_informações.place(
    x=15,
    y=730,
    width=110,
    height=30
)

############################################################ Criando e definindo o campo de filtragem "Área": ############################################################

opções_área = ["Todas as Áreas", "Fábrica", "Centro Logístico"]
# Variáveis para armazenar as opções selecionadas nos dropdowns
filtro_área = StringVar(window)
# Criando a estilização do menu dropdown:
dropdown_filtro_área = ttk.Combobox(window,
        textvariable=filtro_área,
        values=opções_área)

dropdown_filtro_área.place(
        x=300,
        y=243,
        width=130,
        height=30
        )

############################################################ Criando e definindo os dados do campo "Filtrar_matrícula": ############################################################

# Criar o dropdown de matrículas
dropdown__filtro_matricula = ttk.Combobox(
    window,
    values=[],
)
dropdown__filtro_matricula.place(
    x=455,
    y=243,
    width=130,
    height=30,
)

# Função para carregar todas as matrículas únicas do banco de dados SQLite
def carregar_matriculas():
    cursor.execute("SELECT DISTINCT Matrícula FROM Absenteísmo")
    matriculas = cursor.fetchall()
    matriculas = [str(matricula[0]) for matricula in matriculas]
    return matriculas
# Obter todas as matrículas do banco de dados
matriculas = carregar_matriculas()
# Inserir uma opção adicional para selecionar todas as matrículas
matriculas.insert(0, "Todas as Matrículas")
# Definir as opções do dropdown de matrícula
dropdown__filtro_matricula['values'] = matriculas

############################################################ Criando e definindo os dados do campo "dropdown_nome":   ############################################################

dropdown_filtro_nome = ttk.Combobox(
    window,
    values=[],
    )

dropdown_filtro_nome.place(
    x = 610, y = 243,
    width = 170,
    height = 30)

# Função para carregar todos os nomes únicos do banco de dados SQLite
def carregar_nome():
    cursor.execute("SELECT DISTINCT Nome FROM Absenteísmo")
    nomes = cursor.fetchall()
    nomes = [str(nome[0]) for nome in nomes]
    return nomes

# Obter todas as matrículas do banco de dados
nomes = carregar_nome()
# Inserir uma opção adicional para selecionar todas as matrículas
nomes.insert(0, "Todos os Nomes")
# Definir as opções do dropdown de matrícula
dropdown_filtro_nome['values'] = nomes


############################################################ Criando e definindo os dados do campo "dropdown_turno": ############################################################

opções_turno = ["Todos os Turnos", "1º TURNO", "2º TURNO"]
# Variáveis para armazenar as opções selecionadas nos dropdowns
filtro_turno = StringVar(window)
# Criando a estilização do menu dropdown:
dropdown_filtro_turno = ttk.Combobox(window,
        textvariable=filtro_turno,
        values=opções_turno)

dropdown_filtro_turno.place(
        x=805,
        y=243,
        width=130,
        height=30
        )

# Função para carregar todos os turnos do banco de dados SQLite:
def carregar_turno():
    cursor.execute("SELECT DISTINCT Turno FROM Absenteísmo")
    turnos = cursor.fetchall()
    turnos = [str(turno[0]) for turno in turnos]
    return turnos
# Obter todos os turnos do banco de dados
turnos = carregar_turno()
# Inserir uma opção adicional para selecionar todos os turnos
turnos.insert(0, "Todos os Turnos")
# Definir as opções do dropdown de turnos
dropdown_filtro_turno['values'] = turnos

############################################################ Criando e definindo os dados do campo "dropdown_condutor" ############################################################

dropdown_filtro_condutor = ttk.Combobox(
    window,
    values=[],
    )

dropdown_filtro_condutor.place(
    x = 960, y = 243,
    width = 133,
    height = 30)

# Função para carregar todos os nomes únicos do banco de dados SQLite
def carregar_condutor():
    cursor.execute("SELECT DISTINCT Condutor FROM Absenteísmo")
    condutores = cursor.fetchall()
    condutores = [str(condutor[0]) for condutor in condutores]
    return condutores

# Obter todas as matrículas do banco de dados
condutores = carregar_condutor()
# Inserir uma opção adicional para selecionar todas as matrículas
condutores.insert(0, "Todos os condutores")
# Definir as opções do dropdown de matrícula
dropdown_filtro_condutor['values'] = condutores


############################################################ Criando e definindo os dados do campo "dropdown_supervisor" ############################################################

dropdown_filtro_supervisor = ttk.Combobox(
    window,
    values=[],
    )

dropdown_filtro_supervisor.place(
    x = 1115, y = 243,
    width = 133,
    height = 30)

# Função para carregar todos os nomes únicos do banco de dados SQLite
def carregar_supervisor():
    cursor.execute("SELECT DISTINCT Supervisor FROM Absenteísmo")
    supervisores = cursor.fetchall()
    supervisores = [str(supervisor[0]) for supervisor in supervisores]
    return supervisores

# Obter todas as matrículas do banco de dados
supervisores = carregar_supervisor()
# Inserir uma opção adicional para selecionar todas as matrículas
supervisores.insert(0, "Todos os supervisores")
# Definir as opções do dropdown de matrícula
dropdown_filtro_supervisor['values'] = supervisores


############################################################ Criando e definindo os dados do campo "dropdown_motivo" ############################################################

opções_motivo = ["COVID", "Dengue", 'Outros']
# Variáveis para armazenar as opções selecionadas nos dropdowns
filtro_motivo = StringVar(window)
# Criando a estilização do menu dropdown:
dropdown_filtro_motivo = ttk.Combobox(window,
        textvariable=filtro_motivo,
        values=opções_motivo)

dropdown_filtro_motivo.place(
        x=1270,
        y=243,
        width=130,
        height=30
        )

# Função para carregar todos os turnos do banco de dados SQLite:
def carregar_motivo():
    cursor.execute("SELECT DISTINCT Motivo FROM Absenteísmo")
    motivos = cursor.fetchall()
    motivos = [str(motivo[0]) for motivo in motivos]
    return motivos
# Obter todos os turnos do banco de dados
motivos = carregar_motivo()
# Inserir uma opção adicional para selecionar todos os turnos
motivos.insert(0, "Todos os Motivos")
# Definir as opções do dropdown de turnos
dropdown_filtro_motivo['values'] = motivos


############################################################ Criando e definindo os dados do campo "filtro_data" ############################################################

filtro_data = DateEntry(
    window,
    foreground='white',
    bordercolor='black',
    borderwidth=5, 
    highlightthickness=1, 
    year=2024,
    locale='pt_br',
)

filtro_data.delete(0, 'end')  # Limpa o valor inicial do DateEntry

filtro_data.place(
    x=1425,
    y=243,
    width=130,
    height=30
)

############################################################ Criando e definindo as funções que irão filtrar a tabela" ############################################################

# Estrutura de dados para armazenar o estado dos filtros
filtros = {
    "área": None,
    "matrícula": None,
    "nome": None,
    "turno": None,
    "condutor": None,
    "supervisor": None,
    "motivo": None,
    "data": None
}

# Função para atualizar o estado dos filtros
def atualizar_filtros():
    filtros["área"] = dropdown_filtro_área.get()
    filtros["matrícula"] = dropdown__filtro_matricula.get()
    filtros["nome"] = dropdown_filtro_nome.get()
    filtros["turno"] = dropdown_filtro_turno.get()
    filtros["condutor"] = dropdown_filtro_condutor.get()
    filtros["supervisor"] = dropdown_filtro_supervisor.get()
    filtros["motivo"] = dropdown_filtro_motivo.get()
    filtros["data"] = filtro_data.get()

# Função que filtra a tabela:
def filtrar_tabela():
    # Limpar a tabela antes de preencher com os dados filtrados
    for linha in campo_tabela.get_children():
        campo_tabela.delete(linha)
    # Construir a consulta SQL com base nos filtros selecionados
    comando = "SELECT * FROM Absenteísmo WHERE 1=1"
    # Verificar se algum filtro está vazio para renderizar todas as informações da tabela
    if all(value == "" or value == "Todas as Áreas" for value in filtros.values()):
        comando = "SELECT * FROM Absenteísmo"
    else:
        # Adicionar condições para lidar com a interdependência entre os filtros
        if filtros["área"] and filtros["área"] != "Todas as Áreas":
            comando += f" AND Área = '{filtros['área']}'"

        if filtros["matrícula"] and filtros["matrícula"] != "Todas as Matrículas":
            comando += f" AND Matrícula = '{filtros['matrícula']}'"

        if filtros["nome"] and filtros["nome"] != "Todos os Nomes":
            comando += f" AND Nome = '{filtros['nome']}'"

        if filtros["turno"] and filtros["turno"] != "Todos os Turnos":
            comando += f" AND Turno = '{filtros['turno']}'"

        if filtros["condutor"] and filtros["condutor"] != "Todos os condutores":
            comando += f" AND Condutor = '{filtros['condutor']}'"

        if filtros["supervisor"] and filtros["supervisor"] != "Todos os supervisores":
            comando += f" AND Supervisor = '{filtros['supervisor']}'"

        if filtros["motivo"] and filtros["motivo"] != "Todos os Motivos":
            comando += f" AND Motivo = '{filtros['motivo']}'"

        if filtros["data"] and filtros["data"]:
            comando += f" AND Data = '{filtros['data']}'"

    # Executar a consulta SQL
    cursor.execute(comando)
    dados_filtrados = cursor.fetchall()
    # Adicionar os dados filtrados ao Treeview
    for linha in dados_filtrados:
        linha_formatada = [str(item).replace('\n', '') for item in linha[1:]]
        campo_tabela.insert('', 'end', values=linha_formatada)

# Vincular a função de atualização de filtros ao evento de mudança de seleção de cada dropdown
dropdown_filtro_área.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown__filtro_matricula.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown_filtro_nome.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown_filtro_turno.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown_filtro_condutor.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown_filtro_supervisor.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
dropdown_filtro_motivo.bind("<<ComboboxSelected>>", lambda event: [atualizar_filtros(), filtrar_tabela()])
filtro_data.bind("<FocusOut>", lambda event: [atualizar_filtros(), filtrar_tabela()])


# Função para obter todas as áreas únicas do banco de dados
def obter_areas():
    cursor.execute("SELECT DISTINCT Área FROM Absenteísmo")
    areas = cursor.fetchall()
    return ["Todas as Áreas"] + [area[0] for area in areas]

# Botão que reseta o aplicativo para o estado inicial:
def BotãoLimparInformaçõesFiltros():
    dropdown_filtro_área['values'] = obter_areas()
    dropdown_filtro_área.set("") 
    dropdown__filtro_matricula.set("")
    dropdown_filtro_nome.set("")
    dropdown_filtro_turno.set("")
    dropdown_filtro_condutor.set("")
    dropdown_filtro_supervisor.set("")
    dropdown_filtro_motivo.set("")
    filtro_data.delete(0, 'end')
    # Consultar todas as informações da tabela
    cursor.execute("SELECT * FROM Absenteísmo")
    todas_informações = cursor.fetchall()
    # Limpar a tabela antes de preencher
    for linha in campo_tabela.get_children():
        campo_tabela.delete(linha)
    # Adicionar todas as informações ao Treeview
    for linha in todas_informações:
        linha_formatada = [str(item).replace('\n', '') for item in linha[1:]]
        campo_tabela.insert('', 'end', values=linha_formatada)


############################################################ Criando e definindo as informações do botão "Limpar Informações Filtros": ############################################################

imagem_limpar_filtros = PhotoImage(file = r'C:\Users\F89074d\Desktop\Python\Projeto Absenteísmo Tkinter\Imagens\img1.png')
botão_limpar_filtros = Button(
    image = imagem_limpar_filtros,
    borderwidth = 0,
    highlightthickness = 0,
    command = BotãoLimparInformaçõesFiltros,
    relief = "flat")

botão_limpar_filtros.place(
    x = 1571,
    y = 223,
    width = 55,
    height = 60)


def obter_matriculas_por_area(area):
    cursor.execute("SELECT DISTINCT Matrícula FROM Absenteísmo WHERE Área = ?", (area,))
    matriculas = cursor.fetchall()
    matriculas = [str(matricula[0]) for matricula in matriculas]
    return matriculas

# Função para obter os nomes correspondentes à área especificada
def obter_nomes_por_area(area):
    cursor.execute("SELECT DISTINCT Nome FROM Absenteísmo WHERE Área = ?", (area,))
    nomes = cursor.fetchall()
    nomes = [str(nome[0]) for nome in nomes]
    return nomes

# Função para obter os nomes correspondentes à área especificada
def obter_turno_por_area(area):
    cursor.execute("SELECT DISTINCT Turno FROM Absenteísmo WHERE Área = ?", (area,))
    turnos = cursor.fetchall()
    turnos = [str(turno[0]) for turno in turnos]
    return turnos

# Função para obter os nomes correspondentes à área especificada
def obter_condutor_por_area(area):
    cursor.execute("SELECT DISTINCT Condutor FROM Absenteísmo WHERE Área = ?", (area,))
    condutores = cursor.fetchall()
    condutores = [str(condutor[0]) for condutor in condutores]
    return condutores

# Função para obter os nomes correspondentes à área especificada
def obter_supervisor_por_area(area):
    cursor.execute("SELECT DISTINCT Supervisor FROM Absenteísmo WHERE Área = ?", (area,))
    supervisores = cursor.fetchall()
    supervisores = [str(supervisor[0]) for supervisor in supervisores]
    return supervisores

# Função para obter os nomes correspondentes à área especificada
def obter_motivo_por_area(area):
    cursor.execute("SELECT DISTINCT Motivo FROM Absenteísmo WHERE Área = ?", (area,))
    motivos = cursor.fetchall()
    motivos = [str(motivo[0]) for motivo in motivos]
    return motivos

# Função para atualizar as opções dos dropdowns com base na seleção do filtro de área
def atualizar_opcoes_dropdown():
    area_selecionada = filtro_área.get()
    if area_selecionada == "Fábrica":
        # Atualizar as opções do dropdown de matrícula para refletir apenas as matrículas correspondentes à fábrica
        matriculas_fabrica = obter_matriculas_por_area("Fábrica")
        dropdown__filtro_matricula['values'] = ["Todas as Matrículas"] + matriculas_fabrica
        # Atualizar as opções do dropdown de nome para refletir apenas os nomes correspondentes ao Centro logístico
        nomes_fabrica = obter_nomes_por_area("Fábrica")
        dropdown_filtro_nome['values'] = ["Todos os Nomes"] + nomes_fabrica

        turnos_fábrica = obter_turno_por_area('Fábrica')
        dropdown_filtro_turno['values'] = ['Todos os Turnos'] + turnos_fábrica

        condutores_fábrica = obter_condutor_por_area('Fábrica')
        dropdown_filtro_condutor['values'] = ['Todos os condutores'] + condutores_fábrica

        supervisores_fábrica = obter_supervisor_por_area('Fábrica')
        dropdown_filtro_supervisor['values'] = ['Todos os supervisores'] + supervisores_fábrica

        motivos_fábrica = obter_motivo_por_area('Fábrica')
        dropdown_filtro_motivo['values'] = ['Todos os Motivos'] + motivos_fábrica

    if area_selecionada == "Centro Logístico":
        matriculas_centro_logístico = obter_matriculas_por_area("Centro Logístico") 
        dropdown__filtro_matricula['values'] = ["Todas as Matrículas"] + matriculas_centro_logístico

        nomes_centro_logístico = obter_nomes_por_area("Centro Logístico") 
        dropdown_filtro_nome['values'] = ["Todos os Nomes"] + nomes_centro_logístico

        turnos_centro_logístico = obter_turno_por_area('Centro Logístico')
        dropdown_filtro_turno['values'] = ['Todos os Turnos'] + turnos_centro_logístico

        condutores_centro_logístico = obter_condutor_por_area('Centro Logístico')
        dropdown_filtro_condutor['values'] = ['Todos os condutores'] + condutores_centro_logístico

        supervisores_centro_logístico = obter_supervisor_por_area('Centro Logístico')
        dropdown_filtro_supervisor['values'] = ['Todos os supervisores'] + supervisores_centro_logístico

        motivos_centro_logístico = obter_motivo_por_area('Centro Logístico')
        dropdown_filtro_motivo['values'] = ['Todos os Motivos'] + motivos_centro_logístico

# Vincular a função de atualização de opções de dropdown ao evento de mudança de seleção do filtro de área
filtro_área.trace_add("write", lambda *args: atualizar_opcoes_dropdown())

############################################################ Definindo as informações do campo "Matrícula": ############################################################

campo_matrícula = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_matrícula.place(
    x = 15, y = 112,
    width = 240,
    height = 35)

############################################################ Definindo as informações do campo "Nome": ############################################################

campo_nome = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_nome.place(
    x = 15, y = 191,
    width = 240,
    height = 35)

# Criando a função que vai atualizar os campos dinamicamente e também bloquear inserções erradas no campo "Matrícula":
def atualizar_campos(event):
    matricula_inserida = campo_matrícula.get('1.0', END).strip()  # Obter a matrícula inserida pelo usuário
    area_selecionada = campo_área.get()  # Obter a área selecionada
    # Verificar se a matrícula inserida está dentro da área selecionada
    matriculas_validas = Tabela.loc[Tabela['Área'] == area_selecionada, 'Matrícula'].tolist()
    if str(matricula_inserida) not in map(str, matriculas_validas):
        messagebox.showerror(title='Alerta de dados!', message=f'A matrícula {matricula_inserida} não está cadastrada na área {area_selecionada}. Verifique a informação e tente novamente.')
        # Limpar o campo de matrícula
        campo_matrícula.delete('1.0', END)
    matricula_inserida = campo_matrícula.get('1.0', END).strip()
    nome_correspondente = Tabela.loc[Tabela['Matrícula'] == int(matricula_inserida), 'Nome'].values
    turno_correspondente = Tabela.loc[Tabela['Matrícula'] == int(matricula_inserida), 'Turno'].values
    condutor_correspondente = Tabela.loc[Tabela['Matrícula'] == int(matricula_inserida), 'Condutor'].values
    supervisor_correspondente = Tabela.loc[Tabela['Matrícula'] == int(matricula_inserida), 'Supervisor'].values
    
    if nome_correspondente:
        campo_nome.delete('1.0', END)
        campo_nome.insert(END, nome_correspondente[0])

    if turno_correspondente:
        campo_turno.delete('1.0', END)
        campo_turno.insert(END, turno_correspondente[0])

    if condutor_correspondente:
        campo_condutor.delete('1.0', END)
        campo_condutor.insert(END, condutor_correspondente[0])

    if supervisor_correspondente:
        campo_supervisor.delete('1.0', END)
        campo_supervisor.insert(END, supervisor_correspondente[0])

# Associe a função de atualização ao evento de perda de foco no campo de matrícula
campo_matrícula.bind("<FocusOut>", atualizar_campos)

############################################################ Definindo as informações do campo "Turno": ############################################################

campo_turno = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_turno.place(
    x = 15, y = 270,
    width = 240,
    height = 35)

############################################################ Definindo as informações do campo "Condutor": ############################################################

campo_condutor = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_condutor.place(
    x = 15, y = 349,
    width = 240,
    height = 35)

############################################################ Definindo as informações do campo "Supervisor": ############################################################

campo_supervisor = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_supervisor.place(
    x = 15, y = 428,
    width = 240,
    height = 35)

############################################################ Definindo as informações dos dropdowns "Área" e "Motivo": ############################################################

# Lista de opções para o dropdown "Área":
opções_área = ["Fábrica", "Centro Logístico"]
# Variáveis para armazenar as opções selecionadas nos dropdowns
campo_área = StringVar(window)
# Dropdown para Área
dropdown_área = ttk.Combobox(
    window,
    textvariable=campo_área,
    values=opções_área)

dropdown_área.place(x=15, y=33, width=240, height=35)
# Estilo para o dropdown:
style = ttk.Style()
style.theme_use('clam')
style.configure('TMenubutton', background='white', foreground='black')

# Lista de opções para o dropdown "Motivo":
opções_motivo = ["Covid", "Dengue", "Outros"]
# Variáveis para armazenar as opções selecionadas nos dropdowns:
campo_motivo = StringVar(window)
# Dropdown para o campo "Motivo"
dropdown_motivo = ttk.Combobox(
    window,
    textvariable=campo_motivo,
    values=opções_motivo)

dropdown_motivo.place(x=15, y=511, width=240, height=35)
# Estilo para o dropdown:
style = ttk.Style()
style.theme_use('clam')
style.configure('TMenubutton', background='white', foreground='black')

############################################################ Definindo as informações do campo "Data": ############################################################

campo_data = DateEntry(
    window,
    width=12,
    foreground='white',
    bordercolor='black',  # Define a cor da borda
    borderwidth=5,  # Define a largura da borda
    highlightthickness=1,  # Define a espessura do destaque da borda
    year=2024,
    locale='pt_br',
)
campo_data.place(x=15, y=590, width=240, height=35)

campo_data.delete(0, 'end')

############################################################ Definindo as informações do campo "Observação": ############################################################

campo_observação = Text(
    bd=3,
    background="#ffffff",
    highlightthickness=0,
    relief="solid",
    borderwidth=1)

campo_observação.place(
    x = 15, y = 668,
    width = 240,
    height = 50)

############################################################ Definindo as informações da Tabela exibida": ############################################################

campo_tabela = ttk.Treeview(
    window,
    columns=('Área', 'Matrícula', 'Nome', 'Turno', 'Condutor', 'Supervisor', 'Motivo', 'Data', 'Observação'),
    show='headings',  # Isso irá ocultar uma coluna vazia à esquerda
)

# Criando um estilo para o Treeview:
style = ttk.Style()
style.configure("Custom.Treeview.Heading")
# Aplicando o estilo:
campo_tabela.heading('Área', text='Área')
campo_tabela.heading('Matrícula', text='Matrícula')
campo_tabela.heading('Nome', text='Nome')
campo_tabela.heading('Turno', text='Turno')
campo_tabela.heading('Condutor', text='Condutor')
campo_tabela.heading('Supervisor', text='Supervisor')
campo_tabela.heading('Motivo', text='Motivo')
campo_tabela.heading('Data', text='Data',)
campo_tabela.heading('Observação', text='Observação')
# Definindo a largura das colunas:
campo_tabela.column('Área', width=100, anchor='center')
campo_tabela.column('Matrícula', width=100, anchor='center')
campo_tabela.column('Nome', width=150,anchor='center')
campo_tabela.column('Turno', width=100,anchor='center')
campo_tabela.column('Condutor', width=100,anchor='center')
campo_tabela.column('Supervisor', width=100,anchor='center')
campo_tabela.column('Motivo', width=100,anchor='center')
campo_tabela.column('Data', width=100,anchor='center')
campo_tabela.column('Observação', width=100,anchor='center')


# Preenchendo o Treeview com os dados do banco de dados:
def preencher_tabela():
    # Limpe a tabela antes de preencher
    for linha in campo_tabela.get_children():
        campo_tabela.delete(linha)
    # Consultando o banco de dados e obtenha os dados:
    cursor.execute("SELECT * FROM Absenteísmo")
    dados = cursor.fetchall()
    # Adicionando os dados ao Treeview:
    for linha in dados:
        # Concatenar os elementos da tupla e remover caracteres de quebra de linha (\n)
        linha_formatada = [str(item).replace('\n', '') for item in linha[1:]]
        campo_tabela.insert('', 'end', values=linha_formatada)
# Chamando a função para preencher a tabela:
preencher_tabela()
# Posicione o Treeview na janela
campo_tabela.place(x=285, y=295, width=1435, height=479)

window.resizable(False, False)
window.mainloop()