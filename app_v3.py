import itertools
import math
import pulp
import pandas as pd

def otimizar_corte_simplex(df, comprimento_vara=6000):
    resultados_materiais = []

    # Agrupar o DataFrame por 'mp' (matéria-prima)
    grupos = df.groupby('mp')

    for mp, grupo in grupos:
        print(f"\nProcessando matéria-prima: {mp}")

        # Obter as peças e suas quantidades e tamanhos
        codigos_pecas = grupo['codigo'].tolist()
        qtd_maxima = grupo['qtd_maxima'].tolist()
        tamanhos = grupo['tamanhos'].tolist()
        conjuntos = grupo['conjuntos'].tolist()

        while any(qtd > 0 for qtd in qtd_maxima):
            # Criação do problema de otimização (Minimização da folga)
            prob = pulp.LpProblem("Minimizar Perda de Material", pulp.LpMinimize)

            # Definindo as variáveis de decisão com limites baseados na quantidade máxima atual
            variaveis = []
            for i in range(len(codigos_pecas)):
                nome_var = f"x_{codigos_pecas[i]}"
                variavel = pulp.LpVariable(nome_var, lowBound=0, upBound=qtd_maxima[i], cat='Integer')
                variaveis.append(variavel)

            # Variável de folga (perda de material)
            perda = pulp.LpVariable('Perda', lowBound=0, cat='Continuous')

            # Definindo a função objetivo (minimizar a perda)
            prob += perda, "Minimizar Perda"

            # Restrição: soma dos comprimentos das peças <= comprimento da vara
            prob += pulp.lpSum([variaveis[i] * tamanhos[i] for i in range(len(codigos_pecas))]) + perda == comprimento_vara, "Restrição de Comprimento da Vara"

            # Resolvendo o problema
            prob.solve()

            # Verificar se uma solução viável foi encontrada
            if prob.status != pulp.LpStatusOptimal:
                print(f"Não foi possível encontrar uma solução ótima para a MP {mp}. Encerrando este grupo.")
                break

            # Armazenar os resultados para esta vara
            descricao_combinacao = []
            total_cortado = 0
            for i, variable in enumerate(variaveis):
                if variable.varValue > 0:
                    qtd_cortada = int(variable.varValue)
                    descricao_combinacao.append(f"{qtd_cortada} peças {codigos_pecas[i]}")
                    total_cortado += qtd_cortada * tamanhos[i]
                    qtd_maxima[i] -= qtd_cortada  # Decrementa a quantidade processada

            perda_atual = comprimento_vara - total_cortado
            combinacoes_usadas = {
                'mp': mp,
                'descricao_combinacao': ", ".join(descricao_combinacao),
                'varas_usadas': 1,
                'perda_total': perda_atual
            }

            print(f"MP {mp}: {', '.join(descricao_combinacao)} = cortando 1x de {comprimento_vara}mm, perda {perda_atual}mm")
            resultados_materiais.append(combinacoes_usadas)

    return resultados_materiais

# Teste da função com um DataFrame de exemplo
df = pd.DataFrame({
    'mp': ['A', 'A', 'A', 'B', 'B'],
    'codigo': ['P1', 'P2', 'P3', 'P4', 'P5'],
    'qtd_maxima': [20, 15, 10, 12, 16],
    'tamanhos': [500, 700, 1200, 1500, 1000],
    'conjuntos': ['C1', 'C1', 'C2', 'C3', 'C3']
})

resultados = otimizar_corte_simplex(df)

# Criação do DataFrame com os resultados
resultados_df = pd.DataFrame(resultados)

# Agrupando por 'mp' e 'descricao_combinacao'
agrupados = resultados_df.groupby(['mp', 'descricao_combinacao']).agg({
    'varas_usadas': 'sum',       # Soma o número de varas usadas para cada combinação
    'perda_total': 'sum'         # Soma a perda total para cada combinação
}).reset_index()