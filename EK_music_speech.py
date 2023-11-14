########################################################################
#       Pequeño script para calcular los parametros acústicos          #
# EK music y EK speech definidos por von L. Dietsch & W. Kraak en 1986 #
#                   Autor: Leandro Damian Gómez :)                     #
#                               Enjoy!                                 #
########################################################################

import scipy as sp
from tkinter import filedialog as fd
from tkinter import Tk
import numpy as np
import soundfile as sf
import xlsxwriter as xls

# def ir_path():
#     filetypes = (
#         ('WAV files', '*.wav'),
#     )

#     window = Tk()
#     window.wm_attributes('-topmost', 1)
#     window.withdraw()

#     path = fd.askopenfilenames(
#         parent=window,
#         title='Open a file',
#         initialdir='/',
#         filetypes=filetypes)

#     return path

################################################################
#
# Esta función sirve para detectar el comienzo de la RIR y así
# (si se guardó el "total delay" al guardar la RIR en EASE)
# obtener el tiempo de arriba en el punto de escucha, utilizado
# para calcular el EK music y EK speech
#
# def ir_delay(ir, fs):
#     ir_start = np.argwhere(np.abs(ir) >= 0.05)[0][0]
#     print(ir_start)
#     delay = ir_start * 1000 / fs
#     return delay
#
################################################################


# Funcion para filtrar la señal por octavas o tercios de octavas
def filtrado(senial, fs, orden):
    senial_filtrada = []
    # Se cean los array con las frecuencias centrales de los bancos de filtro
    bandas_octavas = [31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000]
    # Se obtienen los límites de las bandas segun la norma ISO-8832
    G = 10 ** (3 / 10)

    bandas = bandas_octavas
    lim_inf = G ** (-1 / 2)
    lim_sup = G ** (1 / 2)

    for i, banda in enumerate(bandas):
        # Se aplican los límites a cada banda
        sup = lim_sup * banda / (0.5 * fs)

        if sup >= 1:
            sup = 0.999999
        inf = lim_inf * banda / (0.5 * fs)

        # Se crea filtro Buttterworth
        filtro = sp.signal.butter(orden, [inf, sup], 'bp', output='sos')
        filtrado = sp.signal.sosfilt(filtro, senial)  # Se realiza el filtrado

        # Se toma la porción util de la señal
        senial_filtrada.append(filtrado[:int(len(filtrado) * 0.95)])

    return senial_filtrada


# Función que calcula el EK music
def EK_music(bandas, tau, fs):
    delta_tauE = int(14 * fs / 1000)  # Delta Tau_e para ek music
    ts_tau = []
    ts_tau_deltaE = []
    for i in range(len(bandas)):
        banda = bandas[i][:tau]
        ts_num = 0
        ts_den = 0
        for t in range(len(banda)):
            ts_num = ts_num + (t * abs(banda[t]))
            ts_den = ts_den + (abs(banda[t]))

        ts_tau.append(ts_num/ts_den)

    for j in range(len(bandas)):
        banda = bandas[j][:tau+delta_tauE]
        ts_num = 0
        ts_den = 0
        for t in range(len(banda)):
            ts_num = ts_num + (t * abs(banda[t]))
            ts_den = ts_den + (abs(banda[t]))

        ts_tau_deltaE.append(ts_num/ts_den)

    ek_music = (np.array(ts_tau_deltaE) - np.array(ts_tau)) / delta_tauE

    return ek_music.tolist()


# Funcion que calcula el EK speech
def EK_speech(bandas, tau, fs):
    delta_tauE = int(9 * fs / 1000)  # Delta tau_e para speech
    n = 2/3
    ts_tau = []
    ts_tau_deltaE = []
    for i in range(len(bandas)):
        banda = bandas[i][:tau]
        ts_num = 0
        ts_den = 0
        for t in range(len(banda)):
            ts_num = ts_num + (t * (abs(banda[t]))**n)
            ts_den = ts_den + (abs(banda[t]))**n

        ts_tau.append(ts_num/ts_den)

    for j in range(len(bandas)):
        banda = bandas[j][:tau+delta_tauE]
        ts_num = 0
        ts_den = 0
        for t in range(len(banda)):
            ts_num = ts_num + (t * (abs(banda[t]))**n)
            ts_den = ts_den + (abs(banda[t]))**n

        ts_tau_deltaE.append(ts_num/ts_den)

    ek_music = (np.array(ts_tau_deltaE) - np.array(ts_tau)) / delta_tauE

    return ek_music


# Valores de total delay ingresados a mano en caso de no guardar el "total delay" en EASE
ir_delay = np.array([30.78, 52.76, 30.59, 55.01, 73.36, 44.53, 50.90, 68.35, 79.88,
                     67.76, 37.53, 45.30, 58.68, 47.43, 93.34, 66.53, 56.94, 83.68, 78.35, 74.86])

ek_music = []
ek_speech = []

# Cambiar rango depende la cantidad de posiciones de escucha analizados
for i in range(20):
    data, fs = sf.read(
        'C:/Users/leand/Desktop/PARCIAL IMA/Auralisation/' + str(i+1) + '.wav')  # Cambiar el path por donde se encuentre la respuesta al impulso

    delay_samples = int(ir_delay[i]*fs/1000)

    ir = []
    for i in range(len(data)):
        ir.append(data[i][0])

    bands = filtrado(ir, fs, 6)

    ek_music.append(EK_music(bands, delay_samples, fs))
    ek_speech.append(EK_speech(bands, delay_samples, fs))

# Guardado de los datos en .XLSX
with xls.Workbook('C:/Users/leand/Desktop/PARCIAL IMA/Auralisation/EK_values.xlsx') as file:
    worksheet = file.add_worksheet()
    worksheet.write_row(0, 0, ['EK Music'])
    worksheet.write_row(1, 0,
                        [31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000])

    for i, data in enumerate(ek_music):
        worksheet.write_row(i+2, 0, data)

    i = i + 6
    worksheet.write_row(i, 0, ['EK Speech'])
    worksheet.write_row(
        i+1, 0, [31.5, 63, 125, 250, 500, 1000, 2000, 4000, 8000, 16000])

    for j, data2 in enumerate(ek_speech):
        worksheet.write_row(i+j+2, 0, data2)
