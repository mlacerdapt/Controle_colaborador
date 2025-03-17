import tkinter as tk
from tkinter import messagebox
import serial
import time
from openpyxl import Workbook
from datetime import datetime

# Configurações dos leitores de cartão serial
PORTA_SERIAL_ENTRADA = '/dev/ttyUSB0'  # Substitua pela porta serial correta do leitor de entrada
PORTA_SERIAL_SAIDA = '/dev/ttyUSB1'    # Substitua pela porta serial correta do leitor de saída
BAUD_RATE = 9600

# Nome do arquivo Excel
NOME_ARQUIVO_EXCEL = 'registros_ponto.xlsx'

def ler_cartao(porta_serial):
    try:
        ser = serial.Serial(porta_serial, BAUD_RATE)
        time.sleep(2)
        codigo_cartao = ser.readline().decode('utf-8').strip()
        ser.close()
        return codigo_cartao
    except serial.SerialException as e:
        messagebox.showerror("Erro", f"Erro ao ler o cartão na porta {porta_serial}: {e}")
        return None

def registrar_ponto_excel(codigo_cartao, tipo):
    try:
        workbook = openpyxl.load_workbook(NOME_ARQUIVO_EXCEL)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(['Código do Cartão', 'Data e Hora', 'Tipo'])

    sheet.append([codigo_cartao, datetime.now(), tipo])
    workbook.save(NOME_ARQUIVO_EXCEL)
    messagebox.showinfo("Sucesso", f"Registro de {tipo} para o cartão {codigo_cartao} salvo em Excel.")

def registrar_ponto():
    def realizar_registro():
        porta_serial = PORTA_SERIAL_ENTRADA if entrada_var.get() == 1 else PORTA_SERIAL_SAIDA
        codigo_cartao = ler_cartao(porta_serial)
        if codigo_cartao:
            tipo = "Entrada" if entrada_var.get() == 1 else "Saída"
            registrar_ponto_excel(codigo_cartao, tipo)
        janela_registro.destroy()

    janela_registro = tk.Toplevel(janela_principal)
    janela_registro.title(["Registrar Ponto","Registrar Ponto","Registrar Ponto"])

    entrada_var = tk.IntVar()
    entrada_radio = tk.Radiobutton(janela_registro, text="Entrada", variable=entrada_var, value=1)
    saida_radio = tk.Radiobutton(janela_registro, text="Saída", variable=entrada_var, value=0)
    registrar_button = tk.Button(janela_registro, text="Registrar", command=realizar_registro)

    entrada_radio.pack()
    saida_radio.pack()
    registrar_button.pack()

janela_principal = tk.Tk()
janela_principal.title("Controle de Ponto")

registrar_ponto_button = tk.Button(janela_principal, text="Registrar Ponto", command=registrar_ponto)
registrar_ponto_button.pack()

janela_principal.mainloop()