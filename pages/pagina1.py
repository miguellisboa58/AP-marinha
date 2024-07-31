import pandas as pd
import streamlit as st

xls = pd.ExcelFile('big planilha.xlsx')
planilha = st.selectbox('Selecione a planilha desejada', ['BP_Pagamento','Condomínio Papem', 'Taxa_de_Condomínio', 'Despesas', 'ReceitasxDespesas', 'Previsão orçamentaria', 'Taxa complementar'
                                                          'Empréstimo', 'TLP_2023', 'TLP_2024', 'TLP_2017', 'TLP_2019', 'TLP_2020', 'TLP_2021', 'TLP_2022', 'TLP_2023 (1)', 'TLP_2024 (1)',
                                                            'Planilha_JUN_para_SGM', 'Desocupados Junho 2024', 'Taxa Extra', 'JAN_24', 'FEV_24', 'MAR_24', 'ABR_2024', 'MAI_2024', 'JUN_2024',
                                                              'JUL_2024', 'AGO_2024', 'SET_2024', 'OUT_2024', 'NOV_2024', 'DEZ_2024', 'Restituições', 'Diversos', 'Descontos VNAVI', 'GERAL', 'BB',
                                                                'CE', 'COND. SQS 202 BL G	', 'COND. SQS 202 BL C	', 'COND. SQS 202 BL A	',  'COND. SQS 202 BL E	', 'COND. SQS 414 BL Q	', 'COND. SQS 414 BL R	',
                                                                  'COND. SQS 414 BL S	', 'COND. SQS 414 BL T	', 'COND. SHCGN 708 BL K	', 'COND. SHCGN 708 BL A	', 'COND. SHCGN 710 BL D	', 'COND. SHCGN 710 BL K	',
                                                                    'COND. SHCGN 711 BL G	', 'COND. SHCGN 712 BL D	', 'COND. SHCGN 712 BL F	', 'COND. SHCGN 712 BL I	', 'COND. SHCGN 714 BL G	', 'COND. SHCGN 714 BL K	', 
                                                                      'COND. SHCGN 714 BL D	', 'COND. SHCGN 715 BL F	', 'COND. SHCGN 715 BL H', 'COND. SHCGN 715 BL I	', 'COND. SQN 407 BL D	', 'COND. SQS 111 BL J	', 'COND. SQS 113 BL I	',
                                                                        'COND. SQS 113 BL G	', 'COND. SHCES 1601 BL B	', 'COND. SHCES 1601 BL A	', 'COND. SHCES 1505 BL E	', 'COND. SHCES 1501 BL D	', 'A.CLARAS BL A	', 'A. CLARAS BL D	',
                                                                          'A. CLARAS BL J	', 'A. CLARAS BL B	', 'A. CLARAS BL E	', 'A. CLARAS BL F-G', 'A. CLARAS BL I	', 'A. CLARAS BL C', 'A. CLARAS BL H	', 'Demonstrativo Janeiro',
                                                                            'Demonstrativo Fevereiro', 'Demonstrativo Março', 'Demonstrativo Abril', 'Demonstrativo Maio', 'Demonstrativo Junho'])
df = pd.read_excel('big planilha.xlsx', sheet_name=planilha)
df = st.data_editor(df, num_rows='dynamic')



