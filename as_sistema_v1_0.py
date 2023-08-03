from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Treeview
from tkcalendar import DateEntry
from mysql.connector import connect
from contextlib import contextmanager
from mysql.connector import ProgrammingError, DataError
from requests import get
import json.decoder
import pandas as pd
import os
import openpyxl

parametros = dict(
    host='localhost',
    passwd='132,mysql',
    port=3306,
    user='root',
    database='estacionamento'
)


@contextmanager
def nova_conexao():
    conexao = connect(**parametros)
    try:
        yield conexao
    finally:
        if conexao and conexao.is_connected():
            conexao.close()


# criar tabela
mensalistas = '''CREATE TABLE IF NOT EXISTS mensalistas (
                            id INT AUTO_INCREMENT NOT NULL,
                            nome VARCHAR(50) NOT NULL,
                            cpf VARCHAR(14) NOT NULL,
                            valor VARCHAR(10) NOT NULL,
                            data_vencimento VARCHAR(15) NOT NULL,
                            endereco VARCHAR(150),
                            primary key(id)
)'''

with nova_conexao() as conexao:
    cursor = conexao.cursor()
    cursor.execute(mensalistas)


def buscar():  # cep
    try:
        url = get(f'http://viacep.com.br/ws/{entry_cep.get()}/json/')
        resultado = url.json()
        texto_logradouro['text'] = resultado['logradouro']
        texto_bairro['text'] = resultado['bairro']
        texto_cidade['text'] = resultado['localidade']
        texto_uf['text'] = resultado['uf']
        texto_complemento['text'] = resultado['complemento'] if len(resultado['complemento']) > 0 else 'Sem complemento'
    except json.decoder.JSONDecodeError:
        messagebox.showerror('Erro!', 'CEP não encontrado!')
        entry_cep.delete(first=0, last=len(entry_cep.get()))

    except KeyError:
        messagebox.showerror('Erro!', 'CEP não encontrado!')
        entry_cep.delete(first=0, last=len(entry_cep.get()))


def adicionar():
    sql = 'INSERT INTO mensalistas (nome, cpf, valor, data_vencimento, endereco) VALUES (%s, %s, %s, %s, %s)'
    args = (f'{entry_nome.get()}'.title(),
            
            f'{entry_cpf.get()}' if "." in entry_cpf.get() and "-" in entry_cpf.get() else 
            '{}.{}.{}-{}'.format(entry_cpf.get()[0:3], entry_cpf.get()[3:6], entry_cpf.get()[6:9],
            entry_cpf.get()[9:11]),

            f'{entry_valor.get()}',

            f'{entry_vencimento.get()}',

            f'''{texto_logradouro["text"] +

                 ', ' + texto_complemento["text"] +

                 (' N° ' + entry_numero.get() if len(entry_numero.get()) > 0 else ' S/N') +

                 ', ' + texto_bairro["text"] +

                 ' ' + texto_cidade["text"] +

                 '-' + texto_uf["text"] if len(entry_cep.get()) > 0 else ""}'''
            )

    with nova_conexao() as conexao:
        if len(entry_nome.get()) == 0 or len(entry_cpf.get()) == 0 or len(entry_valor.get()) == 0:
            messagebox.showerror('Erro!', 'Todos os campos da aba "Dados" são obrigatórios!')
        else:
            try:
                cursor = conexao.cursor()
                cursor.execute(sql, args)
                conexao.commit()
                atualizar_tabela()
            except ProgrammingError as e:
                messagebox.showerror('Erro:', f'{e.msg}')
            else:
                messagebox.showinfo('Sucesso!', 'Mensalista adicionado!')
                entry_cep.delete(first=0, last=len(entry_cep.get()))
                entry_nome.delete(first=0, last=len(entry_nome.get()))
                entry_cpf.delete(first=0, last=len(entry_cpf.get()))
                entry_valor.delete(first=0, last=len(entry_valor.get()))
                entry_vencimento.delete(first=0, last=len(entry_vencimento.get()))
                entry_numero.delete(first=0, last=len(entry_numero.get()))
                texto_logradouro["text"] = ''
                texto_complemento["text"] = ''
                texto_bairro["text"] = ''
                texto_cidade["text"] = ''
                texto_uf["text"] = ''

def limpar_endereco():
    entry_cep.delete(first=0, last=len(entry_cep.get()))
    entry_numero.delete(first=0, last=len(entry_numero.get()))
    texto_logradouro["text"] = ''
    texto_complemento["text"] = ''
    texto_bairro["text"] = ''
    texto_cidade["text"] = ''
    texto_uf["text"] = ''


def limpar_dados():
    adicionar['state'] = NORMAL
    adicionar['bg'] = '#289c2e'
    salvar['state'] = DISABLED
    salvar['bg'] = '#9e9e9e'
    entry_nome.delete(first=0, last=len(entry_nome.get()))
    entry_cpf.delete(first=0, last=len(entry_cpf.get()))
    entry_valor.delete(first=0, last=len(entry_valor.get()))
    entry_vencimento.delete(first=0, last=len(entry_vencimento.get()))


def atualizar_tabela():
    for item in tab.get_children():
        tab.delete(item)
    sql = 'SELECT * FROM mensalistas'

    with nova_conexao() as conexao:
        try:
            cursor = conexao.cursor()
            cursor.execute(sql)
            mensalistas = cursor.fetchall()
        except ProgrammingError as e:
            messagebox.showerror('Erro', e.msg)
        for mensalista in mensalistas:
            tab.insert('', END, values=mensalista)

def pesquisar():
    if len(entry_pesquisar.get()) == 0:
        messagebox.showerror('Erro!', 'Digite um nome para pesquisar!')

    for item in tab.get_children():
        tab.delete(item)
    sql = "SELECT * FROM mensalistas WHERE nome like %s"
    args = (f'%{entry_pesquisar.get()}%',)

    with nova_conexao() as conexao:
        cursor = conexao.cursor()
        cursor.execute(sql, args)
        for x in cursor:
            tab.insert('', END, values=x)


def excluir():
    variavel = messagebox.askyesno('Atencão!', 'Tem certeza que deseja excluir o mensalista selecionado?')
    if variavel:
        try:
            item_selecionado = tab.focus()
            detalhe = tab.item(item_selecionado)
            resultado = detalhe['values'][1]
            sql = "DELETE FROM mensalistas WHERE nome = %s"
            args = (f'{resultado}',)
            with nova_conexao() as conexao:
                cursor = conexao.cursor()
                cursor.execute(sql, args)
                conexao.commit()
            messagebox.showinfo('', 'Mensalista excluido!')
            atualizar_tabela()
        except IndexError:
            messagebox.showerror('Erro!', 'Selecione um mensalista na tabela para excluir!')


def editar_mensalista():
    adicionar['state'] = DISABLED
    adicionar['bg'] = '#9e9e9e'
    salvar['state'] = NORMAL
    salvar['bg'] = '#289c2e'
    try:
        if len(entry_nome.get()) <= 0: 
            item_selecionado = tab.focus()
            detalhe = tab.item(item_selecionado)
            resultado = detalhe['values']
            str(entry_nome.insert(0, resultado[1]))
            str(entry_cpf.insert(0, resultado[2]))
            str(entry_valor.insert(0, resultado[3]))
            entry_vencimento.delete(first=0, last=len(entry_vencimento.get()))
            str(entry_vencimento.insert(0, resultado[4]))
        else:
            entry_nome.delete(first=0, last=len(entry_nome.get()))
            entry_cpf.delete(first=0, last=len(entry_cpf.get()))
            entry_valor.delete(first=0, last=len(entry_valor.get()))
            entry_vencimento.delete(first=0, last=len(entry_vencimento.get()))

            item_selecionado = tab.focus()
            detalhe = tab.item(item_selecionado)
            resultado = detalhe['values']
            entry_nome.insert(0, resultado[1])
            entry_cpf.insert(0, resultado[2])
            entry_valor.insert(0, resultado[3])
            entry_vencimento.insert(0, resultado[4])
    except IndexError:
        messagebox.showerror('', 'Selecione um mensalista para editar!')
        adicionar['state'] = NORMAL
        adicionar['bg'] = '#289c2e'
        salvar['state'] = DISABLED
        salvar['bg'] = '#9e9e9e'


def to_excel():
    for item in tab.get_children():
        tab.delete(item)
    sql = 'SELECT * FROM mensalistas'

    with nova_conexao() as conexao:
        try:
            cursor = conexao.cursor()
            cursor.execute(sql)
            mensalistas = cursor.fetchall()
            dados = pd.DataFrame(data=mensalistas)
            dados.to_excel('tabela_estacionamento.xlsx', columns=[1, 2, 3, 4, 5], header=['Nome', 'CPF', 'Valor', 'Início', 'Endereço'])
            wb = openpyxl.load_workbook('tabela_estacionamento.xlsx')
            sheet = wb.active
            sheet.column_dimensions['A'].width = 5
            sheet.column_dimensions['B'].width = 25
            sheet.column_dimensions['C'].width = 16
            sheet.column_dimensions['D'].width = 12
            sheet.column_dimensions['E'].width = 20
            sheet.column_dimensions['F'].width = 62
            wb.save('tabela_estacionamento.xlsx')
            os.startfile('tabela_estacionamento.xlsx')
        except ProgrammingError as e:
            messagebox.showerror('Erro', e.msg)
        except PermissionError:
            messagebox.showinfo('', 'A tabela ja foi aberta!')


def salvar_alteracoes():
    adicionar['state'] = NORMAL
    adicionar['bg'] = '#289c2e'
    item_selecionado = tab.focus()
    detalhe = tab.item(item_selecionado)
    resultado = detalhe['values']
    id_selecionado = resultado[0]
    sql = "UPDATE mensalistas SET nome = %s, cpf = %s, valor = %s, data_vencimento = %s, endereco = %s WHERE id = %s"
    args = (f'{entry_nome.get()}'.title(),
            
            f'{entry_cpf.get()}' if "." in entry_cpf.get() and "-" in entry_cpf.get() else 
            '{}.{}.{}-{}'.format(entry_cpf.get()[0:3], entry_cpf.get()[3:6], entry_cpf.get()[6:9],
            entry_cpf.get()[9:11]),

            f'{entry_valor.get()}',

            f'{entry_vencimento.get()}',

            f'''{texto_logradouro["text"] +

                 ', ' + texto_complemento["text"] +

                 (' N° ' + entry_numero.get() if len(entry_numero.get()) > 0 else ' S/N') +

                 ', ' + texto_bairro["text"] +

                 ' ' + texto_cidade["text"] +

                 '-' + texto_uf["text"] if len(entry_cep.get()) > 0 else ""}''', id_selecionado
            )

    with nova_conexao() as conexao:
        if len(entry_nome.get()) == 0 or len(entry_cpf.get()) == 0 or len(entry_valor.get()) == 0:
            messagebox.showerror('Erro!', 'Todos os campos da aba "Dados" são obrigatórios!')
        else:
            try:
                cursor = conexao.cursor()
                cursor.execute(sql, args)
                conexao.commit()
                atualizar_tabela()
            except KeyError as e:
                messagebox.showerror('Erro:', f'{e.msg}')
            except DataError:
                messagebox.showerror('', 'Os dados informados estão errados! Verifique os dados e tente novamente.')

            else:
                messagebox.showinfo('Sucesso!', f'Mensalista atualizado!\nId: {resultado[0]}\nNome: {entry_nome.get()}\nCPF: {entry_cpf.get()}\nValor: {entry_valor.get()}\nData inicial: {entry_vencimento.get()}')
                entry_cep.delete(first=0, last=len(entry_cep.get()))
                entry_nome.delete(first=0, last=len(entry_nome.get()))
                entry_cpf.delete(first=0, last=len(entry_cpf.get()))
                entry_valor.delete(first=0, last=len(entry_valor.get()))
                entry_vencimento.delete(first=0, last=len(entry_vencimento.get()))
                entry_numero.delete(first=0, last=len(entry_numero.get()))
                texto_logradouro["text"] = ''
                texto_complemento["text"] = ''
                texto_bairro["text"] = ''
                texto_cidade["text"] = ''
                texto_uf["text"] = ''



janela = Tk()
janela.title('AS Sistemas v1.0')
janela.geometry('834x610+400+50')
janela.config(bg='#454545')
janela.iconphoto(False, PhotoImage(file='logo.png'))
janela.resizable(width=False, height=False)

# labelframe 'dados do mensalista'
cadastrar_mensalista = LabelFrame(janela, width=834, height=1000, text='Cadastrar mensalista', font='Calibre 11',
                                  border=1, fg='white')
cadastrar_mensalista.pack()
cadastrar_mensalista.config(bg='#454545')  # cor de fundo

# --------dados-------------------------
dados = LabelFrame(cadastrar_mensalista, width=300, height=160, text='Dados', font='Calibre 11')
dados.place(x=5, y=10)

label_nome = Label(dados, text='Nome', font='Calibre 10')
label_nome.place(x=20, y=10)
entry_nome = Entry(dados, width=30, border=1, borderwidth=2, font='Calinre 10')
entry_nome.place(x=70, y=10)

label_cpf = Label(dados, text='CPF', font='Calibre 10')
label_cpf.place(x=20, y=40)
entry_cpf = Entry(dados, width=15, border=1, borderwidth=2, font='Calinre 10')
entry_cpf.place(x=70, y=40)

label_valor = Label(dados, text='Valor R$', font='Calibre 10')
label_valor.place(x=180, y=40)
entry_valor = Entry(dados, width=6, border=1, borderwidth=2, font='Calinre 10')
entry_valor.place(x=237, y=40)

label_vencimento = Label(dados, text='Data inicial', font='Calibre 10')
label_vencimento.place(x=20, y=70)
entry_vencimento = DateEntry(dados, width=10, border=1, borderwidth=2, font='Calinre 10', justify=CENTER, locale='pt_br')
entry_vencimento.place(x=100, y=70)

limpar_dados = Button(dados, width=6, text='Limpar', font='Impact 9', bg='#d10404', fg='white',
                      command=limpar_dados)
limpar_dados.place(x=62, y=110)

adicionar = Button(dados, width=8, text='Adicionar', font='Impact 9', bg='#289c2e', fg='white', command=adicionar)
adicionar.place(x=2, y=110)

# ------------------ endereço -----------------------------------------------------------
endereco = LabelFrame(cadastrar_mensalista, width=518, height=160, text='Endereço', font='Calibre 11')
endereco.place(x=310, y=10)

label_cep = Label(endereco, text='CEP', font='Calibre 10')
label_cep.place(x=20, y=10)
entry_cep = Entry(endereco, width=13, border=1, borderwidth=2, font='Calinre 10')
entry_cep.place(x=70, y=10)

buscar = Button(endereco, width=6, border=1, text='Buscar', font='Impact 9', bg='#289c2e', fg='white', command=buscar)
buscar.place(x=170, y=8)

limpar_endereco = Button(endereco, width=6, border=1, text='Limpar', font='Impact 9', bg='#d10404', fg='white',
                         command=limpar_endereco)
limpar_endereco.place(x=220, y=8)

label_numero = Label(endereco, text='Número', font='Calibre 10')
label_numero.place(x=360, y=10)
entry_numero = Entry(endereco, width=10, border=1, borderwidth=2, font='Calinre 10', justify=CENTER)
entry_numero.place(x=422, y=10)

label_logradouro = Label(endereco, text='Endereço/Logradouro', font='Calibre 10')
label_logradouro.place(x=10, y=35)
texto_logradouro = Label(endereco, width=24, height=1, text='', font='Calibre', bg='white', justify='left')
texto_logradouro.place(x=10, y=55)

label_bairro = Label(endereco, text='Bairro', font='Calibre 10')
label_bairro.place(x=10, y=80)
texto_bairro = Label(endereco, width=24, height=1, text='', font='Calibre', bg='white', justify='left')
texto_bairro.place(x=10, y=100)

# ---------------------------------

label_complemento = Label(endereco, text='Complemento', font='Calibre 10')
label_complemento.place(x=275, y=35)
texto_complemento = Label(endereco, width=24, height=1, text='', font='Calibre', bg='white', justify='left')
texto_complemento.place(x=275, y=55)

label_cidade = Label(endereco, text='Cidade', font='Calibre 10')
label_cidade.place(x=275, y=80)
texto_cidade = Label(endereco, width=16, height=1, text='', font='Calibre', bg='white', justify='left')
texto_cidade.place(x=275, y=100)

label_uf = Label(endereco, text='UF', font='Calibre 10')
label_uf.place(x=436, y=80)
texto_uf = Label(endereco, width=6, height=1, text='', font='Calibre', bg='white')
texto_uf.place(x=436, y=100)

# ---------------------------------

# tabela e colunas ---------
colunas = ('id', 'Nome', 'CPF', 'Valor', 'Data inicial', 'Endereço')
tab = Treeview(janela, selectmode='browse', columns=colunas, show='headings', height=17)
tab.place(x=5, y=240)

# coluna id
tab.column('id', width=30, minwidth=50, stretch=NO)
tab.heading(0, text='id')

# coluna nome
tab.column('Nome', width=175, minwidth=50, stretch=NO)
tab.heading(1, text='Nome')

# coluna cpf
tab.column('CPF', width=100, minwidth=50, stretch=NO)
tab.heading(2, text='CPF')

# coluna valor
tab.column('Valor', width=47, minwidth=50, stretch=NO)
tab.heading(3, text='Valor')

# coluna data inicial
tab.column('Data inicial', width=90, minwidth=50, stretch=NO)
tab.heading(4, text='Data inicial')

tab.column('Endereço', width=380, minwidth=50, stretch=NO)
tab.heading(5, text='Endereço')

# ---------------labelframe utilidades-----------------
info = LabelFrame(janela, width=824, height=42)
info.place(x=5, y=194)

label_pesquisar_nome = Label(info, text='Nome', font='Calibre 10', fg='black')
label_pesquisar_nome.place(x=530, y=6)
entry_pesquisar = Entry(info, width=24, border=6, borderwidth=5, font='Calinre 10')
entry_pesquisar.place(x=572, y=5)

botao_pesquisar = Button(info, width=8, height=0, text='Pesquisar', font='Impact 9', bg='#289c2e', fg='white',
                         command=pesquisar)
botao_pesquisar.place(x=757, y=4)

# botão de atualizar tabela
atualizar_tab = Button(info, width=13, height=1, text='Atualizar Tabela', font='Impact 9', bg='#4353fa', fg='white',
                       command=atualizar_tabela)
atualizar_tab.place(x=0, y=4)

# botão de excluir
excluir = Button(info, text='Excluir', width=6, height=1, font='Impact 9', bg='#d10404', fg='white',
                 command=excluir)
excluir.place(x=90, y=4)

# transformar em excel
abrir_excel = Button(info, width=11, height=1, text='Abrir no excel', font='Impact 9', bg='#289c2e', fg='white',
                     command=to_excel)
abrir_excel.place(x=138, y=4)

# botão editar mensalista
editar = Button(info, width=6, height=1, text='Editar', font='Impact 9', bg='#b59f12', fg='white',
                     command=editar_mensalista)
editar.place(x=230, y=4)

# botão salvar alterações
salvar = Button(info, width=14, height=1, text='Salvar alterações', font='Impact 9', bg='#9e9e9e', fg='white', 
command=salvar_alteracoes, state=DISABLED)
salvar.place(x=278, y=4)

janela.mainloop()
