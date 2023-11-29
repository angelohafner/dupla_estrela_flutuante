import streamlit as st
import numpy as np
from engineering_notation import EngNumber
from funcoes_desbalanco_neutro import *
from diagramas_fasoriais import *
import os
import openpyxl
from openpyxl.styles import Border, Side
import gc

# %%
def text_to_int(num):
    try:
        num = int(num)
        saida = num
    except ValueError:
        st.error("Por favor, insira um número válido")
        saida = "Por favor, insira um número válido"
    return saida

# %%
st.markdown("# Calculadora de Tensões em Circuitos Dupla Estrela não Aterrado")
file_path = "correntes.xlsx"
if os.path.exists(file_path):
    os.remove(file_path)

file_path = "tensoes.xlsx"
if os.path.exists(file_path):
    os.remove(file_path)

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("### Internos")
    nr_lin_int = st.text_input("Quantidade Série", value="5")
    nr_lin_int = text_to_int(nr_lin_int)
    nr_col_int = st.text_input("Quantidade Paralelo", value="4")
    nr_col_int = text_to_int(nr_col_int)
    st.markdown("### Externos")
    nr_lin_ext = st.text_input("Quantidade Série", value="3")
    nr_lin_ext = text_to_int(nr_lin_ext)
    nr_col_ext = st.text_input("Quantidade Paralelo", value="2")
    nr_col_ext = text_to_int(nr_col_ext)

with col2:
    st.markdown("### Sistema")
    potencia_nominal_trifásica = st.text_input("Potência Reativa Trifásica [MVAr]", value="10")
    potencia_nominal_trifásica = 1e6 * text_to_int(potencia_nominal_trifásica)
    frequencia_fudamental_Hz = st.text_input("Frequência Fundamental [Hz]", value="60")
    frequencia_fudamental_Hz = text_to_int(frequencia_fudamental_Hz)
    tensao_nominal_fase_fase = st.number_input("Tensao de Linha [kV]", value=69.0, step=0.5)
    tensao_nominal_fase_fase = 1e3*tensao_nominal_fase_fase

a = 1 * np.exp(1j * 2 * np.pi / 3)
omega = 2 * np.pi * frequencia_fudamental_Hz
Vff = tensao_nominal_fase_fase * np.exp(1j * np.pi / 6)
tensao_fase_neutro = tensao_nominal_fase_fase / np.sqrt(3)
corrente_fase_neutro = (potencia_nominal_trifásica / 3) / tensao_fase_neutro
reatancia = tensao_fase_neutro / corrente_fase_neutro
corrente_nominal_lata = corrente_fase_neutro
cap_total = 1 / (omega * reatancia)
cap_externa = (cap_total/2*nr_col_ext) * nr_lin_ext
cap_interna = (cap_externa/nr_col_int) * nr_lin_int

with col3:
    st.markdown(f"### Pré-processamento")
    st.markdown(f"$$I_{{ext}} = {EngNumber(corrente_fase_neutro)} \\, \\rm A $$")
    st.markdown(f"$$X_{{ext}} = {EngNumber(reatancia)} \\, \\Omega$$")
    st.markdown(f"$$C_{{ext}} = {EngNumber(cap_externa)/1e-6} \\rm \\mu F$$")
    st.markdown(f"$$C_{{int}} = {EngNumber(cap_interna)/1e-6} \\rm \\mu F$$")

des_int = 0.0

# %% gerar matriz no excel
wb, matriz = matriz_fases_ramos(nr_lin_ext=nr_lin_ext, nr_col_ext=nr_col_ext, nr_lin_int=nr_lin_int,
                                nr_col_int=nr_col_int, cap_interna=cap_interna, des_int=des_int)
wb.save('matriz_total.xlsx')

st.markdown("## Planilha com os valores das capacitâncias")
st.markdown("Download da planilha para preenchimento das capacitâncias.")
with open("matriz_total.xlsx", "rb") as file:
    st.download_button(label="Download Planilha Modelo de Capacitâncias", data=file, file_name="matriz_total.xlsx",
                       mime="application/vnd.ms-excel")

st.markdown("Depois faça o upload dela preenchida.")
uploaded_file = st.file_uploader("Faça upload da sua planilha Excel", type=["xlsx"])
if uploaded_file is not None:
    # Salve o arquivo carregado temporariamente
    with open("temp.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    # Renomeie o arquivo temporário
    if os.path.exists("matriz_total.xlsx"):
        os.remove("matriz_total.xlsx")
    os.rename("temp.xlsx", "matriz_total.xlsx")

# %% ler matriz do excel, depois de editada manualmente
data = load_excel_to_numpy('matriz_total.xlsx')
matriz_A = data['A']
matriz_B = data['B']
matriz_C = data['C']

# %%
eq_paral_ramos_A, \
    eq_serie_externos_A1, eq_paral_externos_A1, \
    eq_serie_internos_A1, eq_paral_internos_A1, \
    eq_serie_externos_A2, eq_paral_externos_A2, \
    eq_serie_internos_A2, eq_paral_internos_A2, \
    matriz_A1, matriz_A2, super_matriz_A1, super_matriz_A2 = \
    fase(matriz_A, nr_col_ext, nr_col_int, nr_lin_ext, nr_lin_int, 1, 2)

eq_paral_ramos_B, \
    eq_serie_externos_B1, eq_paral_externos_B1, eq_serie_internos_B1, eq_paral_internos_B1, \
    eq_serie_externos_B2, eq_paral_externos_B2, eq_serie_internos_B2, eq_paral_internos_B2, \
    matriz_B1, matriz_B2, super_matriz_B1, super_matriz_B2 = \
    fase(matriz_B, nr_col_ext, nr_col_int, nr_lin_ext, nr_lin_int, 1, 2)

eq_paral_ramos_C, \
    eq_serie_externos_C1, eq_paral_externos_C1, eq_serie_internos_C1, eq_paral_internos_C1, \
    eq_serie_externos_C2, eq_paral_externos_C2, eq_serie_internos_C2, eq_paral_internos_C2, \
    matriz_C1, matriz_C2, super_matriz_C1, super_matriz_C2 = \
    fase(matriz_C, nr_col_ext, nr_col_int, nr_lin_ext, nr_lin_int, 1, 2)

# %% Salvando no excel para acomapanhar
# Paralelos Internos
datasets = [eq_paral_internos_A1, eq_paral_internos_A2,
            eq_paral_internos_B1, eq_paral_internos_B2,
            eq_paral_internos_C1, eq_paral_internos_C2]

sheet_names = ["eq_paral_internos_A1", "eq_paral_internos_A2",
               "eq_paral_internos_B1", "eq_paral_internos_B2",
               "eq_paral_internos_C1", "eq_paral_internos_C2"]

save_multiple_datasets_to_excel(datasets, "matriz_total.xlsx", sheet_names,
                                nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int)


# Séries Internos
append_df_to_excel(matriz=eq_serie_internos_A1, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_A1")
append_df_to_excel(matriz=eq_serie_internos_A2, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_A2")
append_df_to_excel(matriz=eq_serie_internos_B1, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_B1")
append_df_to_excel(matriz=eq_serie_internos_B2, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_B2")
append_df_to_excel(matriz=eq_serie_internos_C1, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_C1")
append_df_to_excel(matriz=eq_serie_internos_C2, filename="matriz_total.xlsx", sheet_name="eq_serie_internos_C2")

# Paralelos Externos
append_df_to_excel(matriz=eq_paral_externos_A1, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_A1")
append_df_to_excel(matriz=eq_paral_externos_A2, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_A2")
append_df_to_excel(matriz=eq_paral_externos_B1, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_B1")
append_df_to_excel(matriz=eq_paral_externos_B2, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_B2")
append_df_to_excel(matriz=eq_paral_externos_C1, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_C1")
append_df_to_excel(matriz=eq_paral_externos_C2, filename="matriz_total.xlsx", sheet_name="eq_paral_externos_C2")

# Séries Externos
append_df_to_excel(matriz=[eq_serie_externos_A1], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_A1")
append_df_to_excel(matriz=[eq_serie_externos_A2], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_A2")
append_df_to_excel(matriz=[eq_serie_externos_B1], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_B1")
append_df_to_excel(matriz=[eq_serie_externos_B2], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_B2")
append_df_to_excel(matriz=[eq_serie_externos_C1], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_C1")
append_df_to_excel(matriz=[eq_serie_externos_C2], filename="matriz_total.xlsx", sheet_name="eq_serie_externos_C2")

# Parelelos dos Ramos
data = [["eq_paral_ramos_A", "eq_paral_ramos_A", "eq_paral_ramos_A"],
        [eq_paral_ramos_A, eq_paral_ramos_A, eq_paral_ramos_A]]
append_df_to_excel(matriz=data, filename="matriz_total.xlsx", sheet_name="eq_paral_ramos")


# %% Impedâncias equivalentes das fases
Za = np.complex128(1 / (1j * omega * eq_paral_ramos_A))
Zb = np.complex128(1 / (1j * omega * eq_paral_ramos_B))
Zc = np.complex128(1 / (1j * omega * eq_paral_ramos_C))

#
matriz_impedancia_sistema = np.array([[Za, 0, 0], [0, Zb, 0], [0, 0, Zc]])
matriz_correntes_fase, matriz_tensoes_Vabco, tensao_deslocamento_netro = \
    calcular_correntes_tensoes(Za, Zb, Zc, Vff, a, matriz_impedancia_sistema)

V_ao = matriz_tensoes_Vabco[0]
V_bo = matriz_tensoes_Vabco[1]
V_co = matriz_tensoes_Vabco[2]
V_on = (V_ao + V_bo + V_co) / 3

I_ao = matriz_correntes_fase[0]
I_bo = matriz_correntes_fase[1]
I_co = matriz_correntes_fase[2]
I_on = I_ao + I_bo + I_co
#
save_fasorial_image(V_ao, V_bo, V_co, V_on,
                    string_Vabco=['V_ao', 'V_bo', 'V_co', 'V_on'],
                    cores=['red', 'brown', 'yellow', 'blue'],
                    img_filename='Fasorial_Tensoes.png',
                    excel_filename='tensoes.xlsx',
                    sheet_name='Vabco')

save_fasorial_image(I_ao, I_bo, I_co, I_on,
                    string_Vabco=['I_ao', 'I_bo', 'I_co', 'I_o'],
                    cores=['red', 'brown', 'yellow', 'blue'],
                    img_filename='Fasorial_Correntes.png',
                    excel_filename='correntes.xlsx',
                    sheet_name='Iabco')

#
# %% Começa o caminho inverso

# %% Fase A, Ramo 1
I_a1_ext, V_a1_ser_ext, I_a1_par_ext, I_a1_par_int, \
    V_a1_ser_int, V_a1_par_int, I_a1_par_int_excel, V_a1_par_int_excel = \
    compute_values(V_ao, omega, eq_serie_externos_A1,
                   eq_paral_externos_A1,
                   eq_serie_internos_A1, eq_paral_internos_A1,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_A1)

# %% Fase A, Ramo 2
I_a2_ext, V_a2_ser_ext, I_a2_par_ext, I_a2_par_int, \
    V_a2_ser_int, V_a2_par_int, I_a2_par_int_excel, V_a2_par_int_excel = \
    compute_values(V_ao, omega,
                   eq_serie_externos_A2, eq_paral_externos_A2,
                   eq_serie_internos_A2, eq_paral_internos_A2,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_A2)

# %% Fase B, Ramo 1
I_b1_ext, V_b1_ser_ext, I_b1_par_ext, I_b1_par_int, \
    V_b1_ser_int, V_b1_par_int, I_b1_par_int_excel, V_b1_par_int_excel = \
    compute_values(V_bo, omega,
                   eq_serie_externos_B1, eq_paral_externos_B1,
                   eq_serie_internos_B1, eq_paral_internos_B1,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_B1)

# %% Fase B, Ramo 2
I_b2_ext, V_b2_ser_ext, I_b2_par_ext, I_b2_par_int, \
    V_b2_ser_int, V_b2_par_int, I_b2_par_int_excel, V_b2_par_int_excel = \
    compute_values(V_bo, omega,
                   eq_serie_externos_B2, eq_paral_externos_B2,
                   eq_serie_internos_B2, eq_paral_internos_B2,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_B2)

# %% Fase C, Ramo 1
I_c1_ext, V_c1_ser_ext, I_c1_par_ext, I_c1_par_int, \
    V_c1_ser_int, V_c1_par_int, I_c1_par_int_excel, V_c1_par_int_excel = \
    compute_values(V_co, omega,
                   eq_serie_externos_C1, eq_paral_externos_C1,
                   eq_serie_internos_C1, eq_paral_internos_C1,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_C1)

# %% Fase C, Ramo 2
I_c2_ext, V_c2_ser_ext, I_c2_par_ext, I_c2_par_int, \
    V_c2_ser_int, V_c2_par_int, I_c2_par_int_excel, V_c2_par_int_excel = \
    compute_values(V_co, omega,
                   eq_serie_externos_C2, eq_paral_externos_C2,
                   eq_serie_internos_C2, eq_paral_internos_C2,
                   nr_lin_ext, nr_col_ext, nr_lin_int, nr_col_int,
                   super_matriz_C2)

# %% Correntes nas latas equivalentes em série de cada ramo
I_o1_ext = I_a1_ext + I_b1_ext + I_c1_ext
I_o2_ext = I_a2_ext + I_b2_ext + I_c2_ext

save_fasorial_image(I_a1_ext, I_b1_ext, I_c1_ext, I_o1_ext,
                    string_Vabco=['I_a1_ext', 'I_b1_ext', 'I_c1_ext', 'I_o1_ext'],
                    cores=['red', 'brown', 'yellow', 'blue'],
                    img_filename='Fasorial_Correntes1.png',
                    excel_filename='correntes.xlsx',
                    sheet_name='Iabco1')

save_fasorial_image(I_a2_ext, I_b2_ext, I_c2_ext, I_o2_ext,
                    string_Vabco=['I_a2_ext', 'I_b2_ext', 'I_c2_ext', 'I_o2_ext'],
                    cores=['red', 'brown', 'yellow', 'blue'],
                    img_filename='Fasorial_Correntes2.png',
                    excel_filename='correntes.xlsx',
                    sheet_name='Iabco2')


# %% Exportando para o Excel
### Corrente Internos A ###
salvar_dataframes_com_exporta(I_a1_par_int_excel, I_a2_par_int_excel, arquivo='correntes.xlsx',
                              planilha_c1='internos-a1',
                              planilha_c2='internos-a2', planilha_c='internos-a')
### Corrente Internos B ###
salvar_dataframes_com_exporta(I_b1_par_int_excel, I_b2_par_int_excel, arquivo='correntes.xlsx',
                              planilha_c1='internos-b1',
                              planilha_c2='internos-b2', planilha_c='internos-b')
### Corrente Internos C ###
salvar_dataframes_com_exporta(I_c1_par_int_excel, I_c2_par_int_excel, arquivo='correntes.xlsx',
                              planilha_c1='internos-c1',
                              planilha_c2='internos-c2', planilha_c='internos-c')
### Corrente Externos A ###
salvar_dataframes_com_exporta(I_a1_par_ext, I_a2_par_ext, arquivo='correntes.xlsx', planilha_c1='externos-a1',
                              planilha_c2='externos-a2', planilha_c='externos-a')
### Corrente Externos B ###
salvar_dataframes_com_exporta(I_b1_par_ext, I_b2_par_ext, arquivo='correntes.xlsx', planilha_c1='externos-b1',
                              planilha_c2='externos-b2', planilha_c='externos-b')
### Corrente Externos C ###
salvar_dataframes_com_exporta(I_c1_par_ext, I_c2_par_ext, arquivo='correntes.xlsx', planilha_c1='externos-c1',
                              planilha_c2='externos-c2', planilha_c='externos-c')

### Tensões Internos A ###
salvar_dataframes_com_exporta(V_a1_par_int_excel, V_a2_par_int_excel, arquivo='tensoes.xlsx',
                              planilha_c1='internos-a1',
                              planilha_c2='internos-a2', planilha_c='internos-a')
### Tensões Internos B ###
salvar_dataframes_com_exporta(V_b1_par_int_excel, V_b2_par_int_excel, arquivo='tensoes.xlsx',
                              planilha_c1='internos-b1', planilha_c2='internos-b2', planilha_c='internos-b')
### Tensões Internos C ###
salvar_dataframes_com_exporta(V_c1_par_int_excel, V_c2_par_int_excel, arquivo='tensoes.xlsx',
                              planilha_c1='internos-c1', planilha_c2='internos-c2', planilha_c='internos-c')

### Identificando células com sobretensao
highlight_cells_above_threshold("tensoes.xlsx", ["internos-a_abs", "internos-b_abs", "internos-c_abs"],
                                1.1 * tensao_fase_neutro/(nr_lin_int*nr_lin_ext))








lista_temp_Tensoes = [V_a1_par_int_excel, V_a2_par_int_excel,
                      V_b1_par_int_excel, V_b2_par_int_excel,
                      V_c1_par_int_excel, V_c2_par_int_excel]

lista_temp = ["fase a, ramo 1", "fase a, ramo 2",
              "fase b, ramo 1", "fase b, ramo 2",
              "fase c, ramo 1", "fase c, ramo 2"]

st.markdown("## Resultados")
pos = 0
for elemento in lista_temp_Tensoes:
    indice_max = np.argmax(np.abs(elemento))
    linha, coluna = np.unravel_index(indice_max, elemento.shape)
    st.markdown(f"Valor máximo: {EngNumber(np.abs(elemento[linha, coluna]))} V,  na posição: ({linha}, {coluna}), na {lista_temp[pos]}.")
    pos = pos + 1

st.markdown("#### Para mais detalhes baixe as planilhas")
with open("tensoes.xlsx", "rb") as file:
    st.download_button(label="Download Tensões", data=file, file_name="tensoes.xlsx", mime="application/vnd.ms-excel")
with open("correntes.xlsx", "rb") as file:
    st.download_button(label="Download Correntes", data=file, file_name="correntes.xlsx", mime="application/vnd.ms-excel")