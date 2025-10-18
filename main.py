import pandas as pd
import openpyxl
import numpy as np
import re
from thefuzz import fuzz
from itertools import combinations


class ACHA_MATRICULA_CONSIGFACIL:
    def __init__(self):

        self.credbase_bruto = r"C:\Users\Guilherme\PycharmProjects\AchaMatriculas\matriculas gov pi\CREDBASE TRABALHADO GOV PI 08.2025.xlsx"

        self.averbacao_bruto = r"C:\Users\Guilherme\PycharmProjects\AchaMatriculas\matriculas gov pi\AVERBADOS GOV PI 08.2025.xlsx"

        self.caminho = r'C:\Users\Guilherme\PycharmProjects\AchaMatriculas\matriculas gov pi'

        df_credbase = pd.read_excel(self.credbase_bruto)

        df_averbacao = pd.read_excel(self.averbacao_bruto)

        self.acha_matricula(df_credbase, df_averbacao)

    # Aqui vai a função de procurar as matrículas corretas por similaridmat
    def acha_matricula(self, credbase: pd.DataFrame, averbacao: pd.DataFrame):
        # --- PREPARAÇÃO INICIAL ---
        credbase = credbase.rename(columns={'Codigo Credbase': 'Codigo_Credbase'})
        colunas_necessarias = ['Codigo_Credbase', 'Banco(s) quitado(s)', 'Esteira',
                               'Esteira(dias)', 'Tipo', 'Operacao', 'Situacao', 'Inicio', 'Cliente',
                               'Data Averbacao', 'CPF', 'Convenio', 'Banco', 'Parcela', 'Prazo',
                               'Tabela', 'Matricula']
        crebase_reduzido = credbase[colunas_necessarias].copy()

        print("Iniciando o processo de extração das matrículas...")

        # --- DEFINIÇÃO DAS FUNÇÕES AUXILIARES ---

        def limpar_matricula(texto: str) -> str:
            if not isinstance(texto, str):
                texto = str(texto)
            return re.sub(r'[^0-9a-zA-Z]', '', texto)

        def encontrar_matriculas_na_linha(row, mapa_cpf_matriculas):
            cpf = str(row['CPF']).strip()
            texto_matriculas_sujo = str(row['Matricula'])

            if cpf not in mapa_cpf_matriculas:
                return ['CPF não encontrado na Averbação']

            matriculas_validos_para_cpf = mapa_cpf_matriculas[cpf]
            partes_sujas = [p for p in re.split(r'[/,;\s-]+', texto_matriculas_sujo) if p]
            if not partes_sujas:
                return []

            encontrados_nesta_linha = []
            matriculas_disponiveis = list(matriculas_validos_para_cpf)
            LIMIAR_DE_SIMILARIDADE = 70

            for parte in partes_sujas:
                parte_limpa = limpar_matricula(parte)
                if not parte_limpa:
                    continue

                melhor_match_para_parte = None
                maior_score = 0
                for matricula_valida in matriculas_disponiveis:
                    matricula_valida_limpo = limpar_matricula(matricula_valida)
                    score = fuzz.ratio(parte_limpa, matricula_valida_limpo)
                    if score > LIMIAR_DE_SIMILARIDADE and score > maior_score:
                        maior_score = score
                        melhor_match_para_parte = matricula_valida

                if melhor_match_para_parte:
                    encontrados_nesta_linha.append(melhor_match_para_parte)
                    matriculas_disponiveis.remove(melhor_match_para_parte)

            return encontrados_nesta_linha

        def soma_por_cpf(credbase_para_somar, averbacao_para_somar):
            print("\n--- Iniciando Método 1: Lógica de Soma por CPF ---")
            credbase_tratado = credbase_para_somar.copy()
            averbacao_geral = averbacao_para_somar.copy()

            averbacao_geral['Matrícula'] = averbacao_geral['Matrícula'].astype(str)

            # Preparação dos dados para a soma
            credbase_tratado['CPF'] = credbase_tratado['CPF'].astype(str).str.strip()
            averbacao_geral['CPF'] = averbacao_geral['CPF'].astype(str).str.strip()
            credbase_tratado['Parcela'] = pd.to_numeric(credbase_tratado['Parcela'], errors='coerce')

            '''averbacao_geral['Valor da reserva'] = (averbacao_geral['Valor da reserva'].astype(str)
                                                   .str.replace(".", "", regex=False)
                                                   .str.replace(",", ".", regex=False))'''
            averbacao_geral['Valor da reserva'] = pd.to_numeric(averbacao_geral['Valor da reserva'], errors='coerce')

            b_lookup = averbacao_geral.dropna(subset=['CPF', 'Valor da reserva', 'Matrícula']) \
                .groupby('CPF') \
                .apply(lambda x: dict(zip(round(x['Valor da reserva'], 2), x['Matrícula']))) \
                .to_dict()

            # print(f'b_lookup {b_lookup}')

            cpfs_unicos = credbase_tratado['CPF'].unique()
            tolerancias = [0, 20, 40, 60]

            for cpf in cpfs_unicos:
                if cpf not in b_lookup:
                    indices_para_marcar = credbase_tratado[credbase_tratado['CPF'] == cpf].index
                    credbase_tratado.loc[indices_para_marcar, 'MATRICULA_ENCONTRADA_1'] = 'CPF não encontrado na averbação (Soma)'
                    continue

                itens_a_combinar = list(credbase_tratado[credbase_tratado['CPF'] == cpf][['Parcela']].itertuples())
                while itens_a_combinar:
                    match_encontrado_nesta_iteracao = False
                    for tamanho_comb in range(1, len(itens_a_combinar) + 1):
                        for comb in combinations(itens_a_combinar, tamanho_comb):
                            soma_parcelas = round(sum(item[1] for item in comb), 2)
                            for tol in tolerancias:
                                valor_alvo = round(soma_parcelas + tol, 2)
                                if valor_alvo in b_lookup.get(cpf, {}):
                                    mat_disponivel = b_lookup[cpf].pop(valor_alvo)
                                    indices_para_atualizar = [item.Index for item in comb]

                                    # Atualiza a coluna 'Matricula' original
                                    credbase_tratado.loc[indices_para_atualizar, 'MATRICULA_ENCONTRADA_1'] = mat_disponivel
                                    # Adiciona colunas de auditoria
                                    credbase_tratado.loc[indices_para_atualizar, 'Soma_Calculada'] = soma_parcelas
                                    credbase_tratado.loc[indices_para_atualizar, 'Parcela_Encontrada'] = valor_alvo
                                    credbase_tratado.loc[indices_para_atualizar, 'Metodo_Encontrado'] = 'SOMAS'

                                    itens_a_combinar = [item for item in itens_a_combinar if
                                                        item.Index not in indices_para_atualizar]
                                    match_encontrado_nesta_iteracao = True
                                    break
                            if match_encontrado_nesta_iteracao: break
                        if match_encontrado_nesta_iteracao: break
                    if not match_encontrado_nesta_iteracao: break

            print("--- Processo de soma concluído! ---")
            # print(f"\n\n{credbase_tratado.loc[credbase_tratado['MATRICULA_ENCONTRADA_1'].notna()]}\n\n")
            return credbase_tratado  # <<-- IMPORTANTE: Retorna o DataFrame modificado

        def achar_por_contse(credbase_contse, averbacao_contse):
            print("\n--- Iniciando Método 3: Lógica de CONT.SE 1 ---")
            # 1 Copiar os DataFrames
            df_averbacao = averbacao_contse.copy()
            df_credbase = credbase_contse.copy()

            # 2 Contar quantidade de registros por CPF
            df_averbacao['CONTSE LOCAL'] = df_averbacao.groupby('CPF')['CPF'].transform('count')

            # 3 Apenas CPFs com 1 contrato (sem duplicidade)
            df_averbacao_unico = df_averbacao[df_averbacao['CONTSE LOCAL'] == 1]

            # 4 Criar dicionário CPF -> Matrícula
            mapa_matricula = df_averbacao_unico.set_index('CPF')['Matrícula'].to_dict()

            # 5 Preencher no credbase apenas para esses CPFs
            df_credbase['MATRICULA_ENCONTRADA_1'] = df_credbase['CPF'].map(mapa_matricula)
            df_credbase.loc[df_credbase['MATRICULA_ENCONTRADA_1'] != '', 'Metodo_Encontrado'] = 'CONT.SE'

            return df_credbase

        def achar_por_saldo_restante(credbase_restante, averbacao_original, credbase_com_resultados_parciais):
            """
            Atribui matrículas com base no maior saldo restante disponível para um CPF,
            consumindo o saldo dinamicamente.

            Args:
                credbase_restante (pd.DataFrame): As linhas da credbase que ainda não têm matrícula.
                averbacao_original (pd.DataFrame): O DataFrame de averbação original.
                credbase_com_resultados_parciais (pd.DataFrame): A credbase completa com os resultados dos passos anteriores.

            Returns:
                pd.DataFrame: O DataFrame credbase_restante com as matrículas encontradas por este método.
            """
            print("\n--- Iniciando Método 4: Análise de Saldo Restante ---")
            df_restante = credbase_restante.copy()

            # --- Passo 1: Calcular o "gasto" de cada matrícula com base nos resultados já encontrados ---
            # Usamos o DataFrame completo com os resultados parciais para ter a visão total.
            gasto_por_matricula = (credbase_com_resultados_parciais
                                   .dropna(subset=['MATRICULA_ENCONTRADA_1', 'Parcela'])
                                   .groupby('MATRICULA_ENCONTRADA_1')['Parcela']
                                   .sum())

            print(f'tipo da coluna matrícula {credbase_com_resultados_parciais['MATRICULA_ENCONTRADA_1'].dtype}')

            print(f'\nGASTO POR MATRICULA: \n{gasto_por_matricula['96253']}\n')

            # --- Passo 2: Calcular o saldo restante em cada matrícula da averbação ---
            df_averbacao_saldos = averbacao_original.copy()
            print(f'tipo da coluna matrícula {df_averbacao_saldos['Matrícula'].dtype}')
            df_averbacao_saldos['Matrícula'] = df_averbacao_saldos['Matrícula'].astype(str)
            # Limpa a coluna de valor para garantir que é numérica
            '''df_averbacao_saldos['Valor da reserva'] = (df_averbacao_saldos['Valor da reserva'].astype(str)
                                                       .str.replace(".", "", regex=False)
                                                       .str.replace(",", ".", regex=False))'''
            df_averbacao_saldos['Valor da reserva'] = pd.to_numeric(df_averbacao_saldos['Valor da reserva'],
                                                                    errors='coerce')

            # Mapeia o gasto para a tabela de averbação. fillna(0) para as que não tiveram gasto.
            df_averbacao_saldos['gasto_calculado'] = df_averbacao_saldos['Matrícula'].map(gasto_por_matricula).fillna(0)
            df_averbacao_saldos['saldo_restante'] = df_averbacao_saldos['Valor da reserva'] - df_averbacao_saldos[
                'gasto_calculado']

            print(f'\nGASTO POR MATRICULA: \n{df_averbacao_saldos.loc[df_averbacao_saldos['Matrícula'] == '96253']}\n')

            # --- Passo 3: Criar um mapa de busca otimizado: CPF -> [(Matrícula, Saldo), ...] ---
            # ATENÇÃO: Convertemos para lista de listas para que seja mutável (listas são, tuplas não são)
            mapa_saldos_por_cpf = (df_averbacao_saldos
                                   .groupby('CPF')[['Matrícula', 'saldo_restante']]
                                   .apply(
                lambda x: [list(item) for item in x.to_records(index=False)])  # Converte para lista de listas
                                   .to_dict())

            # --- Passo 4: Iterar sobre as linhas restantes e encontrar a melhor matrícula (LÓGICA ALTERADA) ---
            # É importante ordenar por Parcela (da maior para a menor) para alocar as parcelas grandes primeiro
            # Isso ajuda a otimizar o uso do saldo
            df_restante_ordenado = df_restante.sort_values(by=['CPF', 'Parcela'], ascending=[True, False])

            for index, row in df_restante_ordenado.iterrows():
                cpf = str(row['CPF']).strip()
                parcela_a_cobrir = row['Parcela']

                # Pega a lista de matrículas e saldos disponíveis para este CPF
                saldos_disponiveis = mapa_saldos_por_cpf.get(cpf)

                if saldos_disponiveis:

                    # 1. Tenta encontrar matrículas que TÊM saldo suficiente
                    #    item[0] = Matrícula, item[1] = saldo_restante
                    opcoes_com_saldo = [item for item in saldos_disponiveis if item[1] >= parcela_a_cobrir]

                    if opcoes_com_saldo:
                        # Se encontrou, pega a que tem o maior saldo *entre elas*
                        melhor_opcao = max(opcoes_com_saldo, key=lambda item: item[1])
                    else:
                        # 2. Fallback: Se NENHUMA tem saldo, pega a que tem o maior saldo (mesmo que estoure)
                        melhor_opcao = max(saldos_disponiveis, key=lambda item: item[1])

                    melhor_matricula = melhor_opcao[0]
                    saldo_atual_da_matricula = melhor_opcao[1]

                    # Atribui a matrícula encontrada e adiciona uma coluna de auditoria
                    df_restante.loc[index, 'MATRICULA_ENCONTRADA_1'] = melhor_matricula
                    df_restante.loc[index, 'Metodo_Encontrado'] = 'SALDO_RESTANTE'
                    df_restante.loc[index, 'Saldo_da_Matricula_no_Momento'] = saldo_atual_da_matricula

                    # 3. ATUALIZA O SALDO EM MEMÓRIA (Recálculo)
                    #    Subtrai a parcela do saldo da matrícula escolhida
                    #    para que a próxima iteração deste CPF veja o valor atualizado.
                    melhor_opcao[1] -= parcela_a_cobrir  # Modifica o saldo na lista

            print("--- Método 4 concluído! ---")
            return df_restante

        # =========================================================================
        # --- FLUXO PRINCIPAL DA EXECUÇÃO (NOVA ORDEM: 1.Soma, 2.Saldo, 3.ContSe, 4.Fuzzy) ---
        # =========================================================================

        # --- PASSO 1: EXECUTAR MÉTODO DE SOMA POR CPF ---
        print("\n--- Iniciando Método 1: Lógica de Soma por CPF ---")
        df_resultado_passo_1 = soma_por_cpf(crebase_reduzido, averbacao)

        # Identifica os resolvidos pela coluna de auditoria 'Soma_Calculada'
        mask_resolvidos_p1 = df_resultado_passo_1['Soma_Calculada'].notna()
        df_resolvidos_passo_1 = df_resultado_passo_1[mask_resolvidos_p1]
        df_para_passo_2 = df_resultado_passo_1[~mask_resolvidos_p1]  # Restantes

        df_resolvidos_passo_1.to_excel(fr'{self.caminho}\df_resolvidos_passo_1_Soma.xlsx', index=False)
        print(f"--- Método 1 (Soma) concluído: {len(df_resolvidos_passo_1)} registros encontrados. ---")
        print(f"--- {len(df_para_passo_2)} registros restantes para o Método 2 (Saldo). ---")

        # --- PASSO 2: EXECUTAR MÉTODO DE SALDO RESTANTE (Novo Passo 2) ---
        df_resultado_passo_2 = pd.DataFrame()
        if not df_para_passo_2.empty:
            print(f"\n--- Iniciando Método 2: Análise de Saldo Restante ({len(df_para_passo_2)} registros) ---")

            # CRUCIAL: Neste ponto, os únicos resultados parciais são os do Passo 1
            df_com_resultados_parciais = df_resolvidos_passo_1

            # Chama a função de saldo com os dados restantes do passo 1
            df_resultado_passo_2 = achar_por_saldo_restante(df_para_passo_2, averbacao, df_com_resultados_parciais)

            # Identifica os resolvidos pela coluna de auditoria 'Metodo_Encontrado'
            mask_resolvidos_p2 = df_resultado_passo_2['Metodo_Encontrado'] == 'SALDO_RESTANTE'

            df_resolvidos_passo_2 = df_resultado_passo_2[mask_resolvidos_p2]
            df_para_passo_3 = df_resultado_passo_2[~mask_resolvidos_p2]  # Restantes

            df_resolvidos_passo_2.to_excel(fr'{self.caminho}\df_resolvidos_passo_2_Saldo.xlsx', index=False)
            print(f"--- Método 2 (Saldo) concluído: {len(df_resolvidos_passo_2)} registros encontrados. ---")
            print(f"--- {len(df_para_passo_3)} registros restantes para o Método 3 (Cont.se). ---")
        else:
            # Se não havia nada para processar, os próximos passos recebem DataFrames vazios
            df_resolvidos_passo_2 = pd.DataFrame()
            df_para_passo_3 = pd.DataFrame()

        # --- PASSO 3: MÉTODO CONT.SE (Novo Passo 3) ---
        df_resultado_passo_3 = pd.DataFrame()
        if not df_para_passo_3.empty:
            print(f"\n--- Iniciando Método 3: Busca por Cont.se ({len(df_para_passo_3)} registros) ---")

            # Chama a função 'achar_por_contse' com os dados restantes do passo 2
            df_resultado_passo_3 = achar_por_contse(df_para_passo_3, averbacao)

            # Assumindo que 'achar_por_contse' também preenche a coluna 'MATRICULA_ENCONTRADA_1'
            mask_resolvidos_p3 = df_resultado_passo_3['MATRICULA_ENCONTRADA_1'].notna()

            df_resolvidos_passo_3 = df_resultado_passo_3[mask_resolvidos_p3]
            df_para_passo_4 = df_resultado_passo_3[~mask_resolvidos_p3]  # Restantes

            df_resolvidos_passo_3.to_excel(fr'{self.caminho}\df_resolvidos_passo_3_ContSe.xlsx', index=False)
            print(f"--- Método 3 (Cont.se) concluído: {len(df_resolvidos_passo_3)} registros encontrados. ---")
            print(f"--- {len(df_para_passo_4)} registros restantes para o Método 4 (Fuzzy). ---")
        else:
            df_resolvidos_passo_3 = pd.DataFrame()
            df_para_passo_4 = pd.DataFrame()

        # --- PASSO 4: EXECUTAR MÉTODO DE SIMILARIDADE (FUZZY MATCH) (Novo Passo 4) ---
        df_resultado_passo_4 = pd.DataFrame()
        if not df_para_passo_4.empty:
            print(f"\n--- Iniciando Método 4: Busca por Similaridade ({len(df_para_passo_4)} registros) ---")

            # Prepara os dados necessários para o método de similaridade
            averbacao['Matrícula'] = averbacao['Matrícula'].astype(str).str.strip()
            mapa_cpf_matriculas = averbacao.groupby('CPF')['Matrícula'].apply(list).to_dict()

            # Aplica a função de busca por similaridade apenas no DataFrame dos restantes
            lista_de_matriculas_encontrados = df_para_passo_4.apply(
                lambda row: encontrar_matriculas_na_linha(row, mapa_cpf_matriculas), axis=1
            )

            # Cria um DataFrame temporário com os resultados
            df_resultados_novos = pd.DataFrame(lista_de_matriculas_encontrados.tolist(),
                                               index=df_para_passo_4.index)

            # Itera sobre as colunas do DataFrame de resultados (0, 1, 2...)
            for i, col_index in enumerate(df_resultados_novos.columns):
                nova_coluna_nome = f'MATRICULA_ENCONTRADA_{i + 1}'
                df_para_passo_4[nova_coluna_nome] = df_resultados_novos[col_index]

            # O resultado final desta etapa é o próprio DataFrame modificado (contendo resolvidos e não resolvidos)
            df_resultado_passo_4 = df_para_passo_4

            df_resultado_passo_4.to_excel(fr'{self.caminho}\df_resultado_passo_4_Fuzzy.xlsx', index=False)
            print("--- Método 4 (Fuzzy) concluído! ---")

        # --- PASSO 5: CONSOLIDAR TODOS OS RESULTADOS ---
        print("\nConsolidando resultados de todos os métodos...")
        df_final_combinado = pd.concat([
            df_resolvidos_passo_1,
            df_resolvidos_passo_2,
            df_resolvidos_passo_3,
            df_resultado_passo_4  # Contém tanto os resolvidos quanto os não resolvidos do último passo
        ], ignore_index=True)

        # --- PASSO 6: LIMPEZA FINAL E SALVAMENTO ---
        try:
            # Tenta preencher NaN em 'MATRICULA_ENCONTRADA_1' antes de converter
            df_final_combinado['MATRICULA_ENCONTRADA_1'] = df_final_combinado['MATRICULA_ENCONTRADA_1'].fillna('N/A')
            mask = df_final_combinado['MATRICULA_ENCONTRADA_1'] != 'CPF não encontrado na Averbação'
            df_final_combinado.loc[mask, 'MATRICULA_ENCONTRADA_1'] = df_final_combinado.loc[
                mask, 'MATRICULA_ENCONTRADA_1'].astype(str)
        except Exception as e:
            print(f"Aviso: Não foi possível processar a coluna de matrícula. Erro: {e}")

        print("Processo geral concluído com sucesso!")
        caminho_arquivo_saida = fr'{self.caminho}\Matriculas_Tratadas_FINAL_NovaOrdem_2.xlsx'
        df_final_combinado.to_excel(caminho_arquivo_saida, index=False)
        print(f"Arquivo final salvo em: {caminho_arquivo_saida}")
        # return df_resultado

        # Chama a função principal com os dataframes preparados

    '''df_codigos_tratados = acha_matricula(credbase_reduzido, averbacao)
    return df_codigos_tratados'''

if __name__ == "__main__":
    # Esta linha cria o objeto e inicia todo o processo
    # chamando o método __init__
    processador = ACHA_MATRICULA_CONSIGFACIL()
    print("Script principal finalizado.")
