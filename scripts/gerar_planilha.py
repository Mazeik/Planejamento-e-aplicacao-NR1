#!/usr/bin/env python3
"""
Gerador da planilha de gestão da NR-1 (GRO/PGR + riscos psicossociais)
para a empresa FICTÍCIA "Indústria Exemplo S.A.".

Todos os dados são inventados. Nenhuma informação de empresa real é usada.

Uso:
    python gerar_planilha.py            # gera a versão DEMO e o MODELO VAZIO
    python gerar_planilha.py --prints   # gera a versão DEMO só para tirar prints

Observação técnica: o openpyxl grava fórmulas com nomes em INGLÊS e vírgula
(IF, COUNTIFS, INDEX). O Excel em português mostra SE, CONT.SES, ÍNDICE com
ponto e vírgula automaticamente. Por isso nunca gravamos "=SE(...)" no arquivo.
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ----------------------------------------------------------------------------
# Estilo (poucas cores: azul-escuro, amarelo, cinza e laranja para alertas)
# ----------------------------------------------------------------------------
FONTE = "Calibri"
AZUL = "1F3864"
F_HDR = PatternFill("solid", fgColor=AZUL)
F_AMARELO = PatternFill("solid", fgColor="FFF2A8")   # preencher
F_CINZA = PatternFill("solid", fgColor="E7E6E6")     # automático / fixo
F_FAIXA = PatternFill("solid", fgColor="DDEBF7")     # faixa explicativa
F_ALERTA = PatternFill("solid", fgColor="F8CBAD")    # laranja de alerta
FONT = Font(name=FONTE, size=11)
FONT_B = Font(name=FONTE, size=11, bold=True)
FONT_HDR = Font(name=FONTE, size=11, bold=True, color="FFFFFF")
FONT_TIT = Font(name=FONTE, size=16, bold=True, color=AZUL)
FONT_CINZA = Font(name=FONTE, size=9, color="7F7F7F")
FINO = Side(style="thin", color="BFBFBF")
BORDA = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
QUEBRA = Alignment(wrap_text=True, vertical="top")
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA = "dd/mm/yyyy"

# ----------------------------------------------------------------------------
# Dados de base (fictícios)
# ----------------------------------------------------------------------------
EMPRESA = "Indústria Exemplo S.A. (empresa fictícia)"

# (GHE, unidade, lotação, setor, turnos, nº pessoas, líder, observação)
SETORES = [
    ("GHE-01", "Planta Principal", "Escritório", "Administrativo", "1º turno", 12, "Gerente administrativo", ""),
    ("GHE-02", "Planta Principal", "Escritório", "Recursos Humanos", "1º turno", 4, "Coordenação de RH", ""),
    ("GHE-03", "Planta Principal", "Laboratório e sala de PCP", "Qualidade e PCP", "1º turno", 9, "Coordenação da qualidade", ""),
    ("GHE-04", "Planta Principal", "Galpão 1", "Produção - Linha A", "1º e 2º turnos", 38, "Supervisão de produção A", ""),
    ("GHE-05", "Planta Principal", "Galpão 2", "Produção - Linha B", "1º, 2º e 3º turnos", 34, "Supervisão de produção B",
     "Há turno noturno; tarefa de inspeção repetitiva (hipótese a investigar)"),
    ("GHE-06", "Centro de Distribuição", "Almoxarifado e docas", "Almoxarifado e Expedição", "1º e 2º turnos", 16, "Supervisão de logística", ""),
    ("GHE-07", "Planta Principal", "Oficina", "Manutenção", "1º turno + sobreaviso", 10, "Supervisão de manutenção", ""),
    ("GHE-08", "Planta Principal", "Portaria", "Portaria e Vigilância", "12x36 (inclui noturno)", 6, "Coordenação de serviços gerais",
     "Vigia noturno trabalha sozinho (hipótese a investigar)"),
]
NOMES_SETORES = [s[3] for s in SETORES]

# (fator, como reconhecer) - lista exemplificativa baseada no Guia do MTE
FATORES = [
    ("Assédio de qualquer natureza no trabalho", "Relatos de humilhação, ameaça, perseguição ou condutas abusivas."),
    ("Má gestão de mudanças organizacionais", "Mudanças de sistema, chefia ou rotina sem aviso ou preparo."),
    ("Baixa clareza de papel ou função", "Pessoas sem saber o que se espera delas ou recebendo ordens conflitantes."),
    ("Baixas recompensas e reconhecimento", "Sensação de esforço não reconhecido, sem retorno ou valorização."),
    ("Falta de suporte ou apoio no trabalho", "Pouca ajuda da liderança ou dos colegas, novatos sem acompanhamento."),
    ("Baixo controle no trabalho / falta de autonomia", "Pouca margem para decidir como, quando e em que ordem trabalhar."),
    ("Baixa justiça organizacional", "Percepção de regras aplicadas de forma desigual ou decisões sem explicação."),
    ("Eventos violentos ou traumáticos", "Acidentes graves, assaltos, ameaças ou agressões no ambiente de trabalho."),
    ("Baixa demanda no trabalho (subcarga)", "Tarefas monótonas, repetitivas ou abaixo da capacidade da pessoa."),
    ("Excesso de demandas no trabalho (sobrecarga)", "Prazos curtos, picos de volume, muitas tarefas ao mesmo tempo."),
    ("Maus relacionamentos no local de trabalho", "Conflitos frequentes, isolamento ou clima de hostilidade."),
    ("Trabalho em condições de difícil comunicação", "Ruído, distância ou barreiras que impedem pedir ajuda ou avisar."),
    ("Trabalho remoto e isolado", "Pessoa trabalhando sozinha ou sem contato regular com a equipe."),
]

# Matriz 5x5: regra didática deste exemplo (provisória, registrada em Premissas)
def classe(p, s):
    n = p * s
    if n <= 2:
        return "Trivial"
    if n <= 4:
        return "Tolerável"
    if n <= 9:
        return "Moderado"
    if n <= 14:
        return "Substancial"
    return "Intolerável"

CLASSES = ["Trivial", "Tolerável", "Moderado", "Substancial", "Intolerável"]

DIAGNOSTICO = [
    ("Gestão de riscos", "Incluir os fatores psicossociais no gerenciamento de riscos da empresa", "NR-1 (GRO/PGR)", "Inventário e plano de ação com os fatores psicossociais tratados"),
    ("Gestão de riscos", "Considerar as condições de trabalho da NR-17 no gerenciamento de riscos", "NR-1 e NR-17", "Inventário e plano que citam os itens da NR-17"),
    ("Ergonomia", "Fazer a Avaliação Ergonômica Preliminar (AEP), incluindo os fatores psicossociais", "NR-17", "Documento da AEP datado e assinado"),
    ("Levantamento de perigos", "Levantar os perigos de cada setor", "NR-1 (GRO/PGR)", "Levantamento preliminar por setor"),
    ("Levantamento de perigos", "Identificar os perigos psicossociais, suas fontes e os grupos expostos", "NR-1 (GRO/PGR)", "Aba Levantamento preenchida, com data e forma de identificação"),
    ("Avaliação de riscos", "Avaliar e classificar o nível de cada risco", "NR-1 (GRO/PGR)", "Probabilidade, severidade e classificação registradas"),
    ("Avaliação de riscos", "Ter um documento com os critérios de severidade, probabilidade e classificação", "NR-1 (GRO/PGR)", "Documento de critérios aprovado pela empresa"),
    ("Inventário de riscos", "Manter o inventário de riscos com o conteúdo mínimo, datado e assinado", "NR-1 (GRO/PGR)", "Inventário assinado pelo responsável"),
    ("Inventário de riscos", "Manter o inventário atualizado e guardar o histórico por 20 anos", "NR-1 (GRO/PGR)", "Versões anteriores arquivadas"),
    ("Plano de ação", "Ter plano de ação com cronograma, responsáveis e forma de acompanhamento", "NR-1 (GRO/PGR)", "Aba Plano de ação com responsável e prazo em todas as linhas"),
    ("Acompanhamento", "Registrar o acompanhamento das medidas e a verificação de eficácia", "NR-1 (GRO/PGR)", "Coluna 'Funcionou?' preenchida e evidências guardadas"),
    ("Participação", "Consultar os trabalhadores sobre a percepção de riscos", "NR-1 (GRO/PGR)", "Atas de escuta e de reuniões da CIPA"),
    ("Participação", "Comunicar aos trabalhadores os riscos e as medidas", "NR-1 (GRO/PGR)", "Registros de comunicação e listas de presença"),
    ("Revisão", "Revisar a avaliação a cada 2 anos ou nas situações previstas na norma", "NR-1 (GRO/PGR)", "Registro da revisão e dos motivos"),
    ("Análise de ocorrências", "Analisar acidentes e doenças relacionadas ao trabalho", "NR-1 (GRO/PGR)", "Relatórios de análise das ocorrências"),
    ("Saúde ocupacional", "Integrar o gerenciamento de riscos ao PCMSO (médico informado dos riscos)", "NR-7", "Comunicação formal ao médico responsável"),
    ("Responsabilidades", "Designar responsável ou equipe pelo tema", "Perguntas e Respostas do MTE", "Documento de designação assinado"),
    ("Documentação", "Manter os documentos disponíveis para a Inspeção do Trabalho", "NR-1 (GRO/PGR)", "Pasta organizada e atualizada"),
    ("CIPA", "Ter CIPA constituída ou representante nomeado", "NR-5", "Ata de posse e calendário de reuniões"),
    ("Assédio e violência", "Ter regras de conduta sobre assédio e violência", "Lei 14.457/2022 e NR-5", "Documento de regras divulgado aos empregados"),
    ("Assédio e violência", "Ter canal de denúncias com anonimato e procedimento de apuração", "Lei 14.457/2022", "Canal em funcionamento e procedimento escrito"),
    ("Capacitação", "Fazer capacitação anual sobre assédio, violência, igualdade e diversidade", "Lei 14.457/2022", "Lista de presença e conteúdo aplicado"),
    ("Capacitação", "Treinar as lideranças sobre fatores psicossociais", "Evidência de medida de prevenção", "Registro do treinamento com lista de presença"),
    ("Integração", "Citar os riscos nas ordens de serviço e na integração de novos empregados", "NR-1 (GRO/PGR)", "Ordem de serviço assinada e registro de integração"),
]

TREINAMENTOS = [  # nome, para quem, por que, tipo, validade (meses), interno?
    ("Capacitação anual sobre assédio, violência, igualdade e diversidade", "Todos os níveis", "Exigência da Lei 14.457/2022", "Obrigatório por norma", 12, "Sim"),
    ("Treinamento da CIPA com o tema assédio", "Membros da CIPA", "Exigência da Lei 14.457/2022 e NR-5", "Obrigatório por norma", 12, "Sim"),
    ("Formação de lideranças em fatores psicossociais", "Supervisores e coordenadores", "Evidência de medida de prevenção", "Evidência de medida de prevenção", 12, "Sim"),
    ("Orientação sobre os riscos do setor e as medidas (integração e reciclagem)", "Todos, por setor", "Comunicação dos riscos e medidas", "Evidência de medida de prevenção", 12, "Sim"),
    ("Capacitação de quem recebe e apura relatos", "Comissão de apuração", "Canal de denúncias funcionando", "Evidência de medida de prevenção", 12, "Sim"),
    ("Capacitação técnica de quem conduz o levantamento", "RH e Segurança do Trabalho", "Qualidade do levantamento psicossocial", "Recomendado", 24, "Não"),
]

INDICADORES = [  # nome, tipo, melhor quando, origem
    ("% de conformidade", "%", "Maior", "Automático"),
    ("% de ações no prazo", "%", "Maior", "Automático"),
    ("% de fatores avaliados", "%", "Maior", "Automático"),
    ("% de pessoas treinadas", "%", "Maior", "Automático"),
    ("Nº de escutas no mês", "nº", "Maior", "Automático"),
    ("Nº de relatos no canal", "nº", "Menor", "Manual (a levantar)"),
    ("Afastamentos acima de 15 dias", "nº", "Menor", "Manual (a levantar)"),
    ("Rotatividade", "%", "Menor", "Manual (a levantar)"),
    ("Absenteísmo", "%", "Menor", "Manual (a levantar)"),
]

# Linhas das tabelas
D1, D2 = 4, 27                       # Diagnóstico (24 linhas)
S1, S2 = 4, 11                       # Setores (8 linhas)
L1 = 4
L2 = L1 + len(SETORES) * len(FATORES) - 1   # Levantamento (104 linhas)
P1, P2 = 4, 203                      # Plano de ação
T_CAT1, T_CAT2 = 4, 9                # Treinamentos: catálogo
T1, T2 = 15, 64                      # Treinamentos: registro
E1, E2 = 4, 203                      # Escuta
A1, A2 = 4, 300                      # Acompanhamento

N_DIAG = "'Diagnóstico NR-1'"
N_PLANO = "'Plano de ação'"
N_LEV = "Levantamento"


def rng(sheet, col, r1, r2):
    """Monta uma referência absoluta, ex.: 'Plano de ação'!$Q$4:$Q$203."""
    return f"{sheet}!${col}${r1}:${col}${r2}"


# ----------------------------------------------------------------------------
# Funções auxiliares de formatação
# ----------------------------------------------------------------------------
def faixa(ws, texto, ate_col):
    """Linha 1: O que a aba responde + o que o usuário precisa fazer."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ate_col)
    c = ws.cell(row=1, column=1, value=texto)
    c.font = Font(name=FONTE, size=11, bold=True, color=AZUL)
    c.fill = F_FAIXA
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 48


def cabecalho(ws, row, titulos, larguras=None, altura=34):
    for i, t in enumerate(titulos, start=1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = FONT_HDR
        c.fill = F_HDR
        c.alignment = CENTRO
        c.border = BORDA
    ws.row_dimensions[row].height = altura
    if larguras:
        for i, w in enumerate(larguras, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def pintar(ws, r1, r2, cols_amarelas, cols_cinzas, ate_col, fmt_datas=(), altura=None):
    """Pinta as linhas: amarelo = preencher, cinza = automático/fixo."""
    for r in range(r1, r2 + 1):
        for c in range(1, ate_col + 1):
            cel = ws.cell(row=r, column=c)
            cel.font = FONT
            cel.border = BORDA
            cel.alignment = QUEBRA
            if c in cols_amarelas:
                cel.fill = F_AMARELO
            elif c in cols_cinzas:
                cel.fill = F_CINZA
            if c in fmt_datas:
                cel.number_format = DATA
        if altura:
            ws.row_dimensions[r].height = altura


def lista(ws, ref, intervalo, aviso="Escolha um item da lista."):
    dv = DataValidation(type="list", formula1=ref, allow_blank=True, showErrorMessage=True,
                        errorTitle="Valor inválido", error=aviso)
    ws.add_data_validation(dv)
    dv.add(intervalo)


def validar_data(ws, intervalo):
    dv = DataValidation(type="date", operator="between", formula1="36526", formula2="73050",
                        allow_blank=True, showErrorMessage=True, errorTitle="Data inválida",
                        error="Digite uma data no formato dd/mm/aaaa (ex.: 15/03/2026).")
    ws.add_data_validation(dv)
    dv.add(intervalo)


def validar_inteiro(ws, intervalo, minimo, maximo, msg):
    dv = DataValidation(type="whole", operator="between", formula1=str(minimo), formula2=str(maximo),
                        allow_blank=True, showErrorMessage=True, errorTitle="Valor inválido", error=msg)
    ws.add_data_validation(dv)
    dv.add(intervalo)


def realce_alerta(ws, intervalo, col_alerta, r1):
    """Linha fica laranja quando a coluna Alerta tem texto."""
    ws.conditional_formatting.add(
        intervalo,
        FormulaRule(formula=[f'${col_alerta}{r1}<>""'], fill=F_ALERTA, stopIfTrue=False))


def config_aba(ws, zoom=100, congelar=None, paisagem=True):
    ws.sheet_view.zoomScale = zoom
    ws.sheet_view.showGridLines = False
    if congelar:
        ws.freeze_panes = congelar
    ws.page_setup.orientation = "landscape" if paisagem else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ----------------------------------------------------------------------------
# Construção do arquivo
# ----------------------------------------------------------------------------
def construir(demo: bool, prints: bool = False):
    wb = Workbook()
    wb.remove(wb.active)
    nomes = ["Como usar", "Painel", "Diagnóstico NR-1", "Setores", "Levantamento", "Plano de ação",
             "Treinamentos", "Escuta", "Acompanhamento", "Alertas", "Listas", "Premissas", "Log"]
    ws = {n: wb.create_sheet(n) for n in nomes}

    aba_listas(ws["Listas"])
    aba_diagnostico(ws["Diagnóstico NR-1"], demo)
    aba_setores(ws["Setores"])
    aba_levantamento(ws["Levantamento"], demo)
    aba_plano(ws["Plano de ação"], demo)
    aba_treinamentos(ws["Treinamentos"], demo)
    aba_escuta(ws["Escuta"], demo)
    aba_acompanhamento(ws["Acompanhamento"], demo)
    aba_alertas(ws["Alertas"])
    aba_painel(ws["Painel"], demo)
    aba_como_usar(ws["Como usar"], demo)
    aba_premissas(ws["Premissas"])
    aba_log(ws["Log"], demo)

    if prints:  # limita a área de impressão para cada aba caber em uma página
        areas = {"Diagnóstico NR-1": "A1:K27", "Setores": "A1:J11", "Levantamento": "A1:N30",
                 "Plano de ação": "A1:S22", "Treinamentos": "A1:L22", "Escuta": "A1:H14",
                 "Acompanhamento": "A1:K22", "Alertas": "A1:D18", "Listas": "A1:Y16",
                 "Premissas": "A1:D16", "Log": "A1:D14", "Painel": "A1:N47", "Como usar": "A1:H28"}
        for n, a in areas.items():
            ws[n].print_area = a
            ws[n].page_setup.fitToHeight = 1
    return wb


# --- Listas -----------------------------------------------------------------
def aba_listas(ws):
    config_aba(ws, 100)
    cols = {
        "A": ("Situação (Diagnóstico)", ["Atendido", "Parcial", "Não atendido", "A levantar", "Não se aplica"], 16),
        "B": ("O fator existe?", ["Sim", "Não", "A investigar"], 14),
        "C": ("Como foi identificado", ["Observação do trabalho", "Escuta coletiva", "Entrevista com líder", "Indicador",
                                         "Canal de relatos", "Documento do PGR", "Outro"], 22),
        "D": ("Origem da ação", ["Diagnóstico", "Levantamento", "Escuta", "Auditoria", "Outro"], 14),
        "E": ("Tipo de medida", ["Eliminar a causa", "Mudança na organização do trabalho", "Medida coletiva",
                                  "Treinamento/orientação", "Apoio individual"], 30),
        "F": ("Situação informada", ["Não iniciada", "Em andamento", "Concluída", "Cancelada"], 14),
        "G": ("Funcionou?", ["Sim", "Parcial", "Não", "Ainda não avaliado"], 16),
        "H": ("Tipo de escuta", ["Reunião da CIPA", "Escuta coletiva por setor", "Diálogo de segurança",
                                  "Devolutiva de resultados", "Relato pelo canal"], 24),
        "I": ("Setor (Painel)", ["Todos"] + NOMES_SETORES, 24),
        "J": ("Setor (ações e escutas)", ["Toda a empresa"] + NOMES_SETORES, 24),
        "K": ("Indicador", [i[0] for i in INDICADORES], 28),
        "L": ("Tipo", [i[1] for i in INDICADORES], 6),
        "M": ("Melhor quando", [i[2] for i in INDICADORES], 12),
        "N": ("Origem do valor", [i[3] for i in INDICADORES], 20),
        "O": ("Sim / Não", ["Sim", "Não"], 10),
    }
    for col, (titulo, itens, larg) in cols.items():
        ws.column_dimensions[col].width = larg
        c = ws[f"{col}1"]
        c.value = titulo
        c.font, c.fill, c.alignment, c.border = FONT_HDR, F_HDR, CENTRO, BORDA
        for i, v in enumerate(itens, start=2):
            x = ws[f"{col}{i}"]
            x.value = v
            x.font, x.border, x.fill = FONT, BORDA, F_CINZA
    ws.row_dimensions[1].height = 34
    # fatores psicossociais e "como reconhecer"
    ws.column_dimensions["Q"].width = 40
    ws.column_dimensions["R"].width = 60
    for col, t in (("Q", "Fator psicossocial"), ("R", "Como reconhecer (linguagem de RH)")):
        c = ws[f"{col}1"]
        c.value, c.font, c.fill, c.alignment, c.border = t, FONT_HDR, F_HDR, CENTRO, BORDA
    for i, (f, como) in enumerate(FATORES, start=2):
        for col, v in (("Q", f), ("R", como)):
            x = ws[f"{col}{i}"]
            x.value, x.font, x.border, x.fill, x.alignment = v, FONT, BORDA, F_CINZA, QUEBRA
    # matriz 5x5 (célula a célula): linhas = probabilidade, colunas = severidade
    ws["T1"].value = "Prob. \\ Sev."
    for j in range(1, 6):
        for (r, c, v) in ((1, 20 + j, j), (1 + j, 20, j)):
            x = ws.cell(row=r, column=c, value=v)
            x.font, x.fill, x.alignment, x.border = FONT_HDR, F_HDR, CENTRO, BORDA
    ws["T1"].font, ws["T1"].fill, ws["T1"].alignment, ws["T1"].border = FONT_HDR, F_HDR, CENTRO, BORDA
    for p in range(1, 6):
        for s in range(1, 6):
            x = ws.cell(row=1 + p, column=20 + s, value=classe(p, s))
            x.font, x.fill, x.alignment, x.border = FONT, F_CINZA, CENTRO, BORDA
    for col in "TUVWXY":
        ws.column_dimensions[col].width = 13
    ws["T9"].value = "Matriz 5x5 (provisória): linha = probabilidade (1 a 5); coluna = severidade (1 a 5)."
    ws["T10"].value = "Regra deste exemplo: produto 1-2 Trivial; 3-4 Tolerável; 5-9 Moderado; 10-14 Substancial; 15-25 Intolerável."
    ws["T11"].value = "Não altere esta aba sem conversar com o responsável pela planilha."
    for a in ("T9", "T10", "T11"):
        ws[a].font = FONT_CINZA
    ws.freeze_panes = "A2"


# --- Diagnóstico NR-1 -------------------------------------------------------
def aba_diagnostico(ws, demo):
    config_aba(ws, 90, "C4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: o que falta para a empresa estar em conformidade com a NR-1?    "
              "O QUE VOCÊ FAZ AQUI: nas colunas amarelas, escolha a situação de cada exigência, diga onde está o "
              "documento, quem é o responsável e o prazo.", 11)
    cabecalho(ws, 3, ["Nº", "Tema", "Exigência (em linguagem simples)", "Base legal", "O que o auditor pede como evidência",
                      "Situação", "Onde está o documento", "Responsável", "Prazo", "Observação", "Alerta"],
              [5, 18, 40, 20, 34, 14, 22, 18, 12, 24, 26])
    pintar(ws, D1, D2, {6, 7, 8, 9, 10}, {1, 2, 3, 4, 5, 11}, 11, fmt_datas={9}, altura=48)
    situacoes = (["Parcial", "Parcial", "Não atendido", "Atendido", "Parcial", "Parcial", "Não atendido", "Parcial",
                  "Atendido", "Parcial", "Não atendido", "Parcial", "Não atendido", "A levantar", "Atendido", "Parcial",
                  "Atendido", "Atendido", "Atendido", "Parcial", "Não atendido", "Parcial", "Não atendido",
                  "A levantar"] if demo else ["A levantar"] * 24)
    # dados de exemplo (só na versão demo): local, responsável, prazo, observação
    ex = {1: ("Pasta SST/Inventário v1", "Coordenação de RH", date(2026, 12, 15), "Falta tratar subcarga e sobrecarga"),
          2: ("Pasta SST/Inventário v1", "Segurança do Trabalho", date(2026, 12, 15), ""),
          3: ("", "Segurança do Trabalho", date(2026, 11, 30), "Aguardando proposta da consultoria"),
          4: ("Pasta SST/Levantamento", "", None, ""),
          5: ("Planilha NR-1 / Levantamento", "Coordenação de RH", date(2026, 12, 20), "4 de 8 setores avaliados"),
          6: ("Planilha NR-1 / Levantamento", "Coordenação de RH", date(2026, 12, 20), ""),
          7: ("", "Segurança do Trabalho", date(2026, 12, 10), "Usar a matriz da aba Listas até validar"),
          8: ("Pasta SST/Inventário v1", "Segurança do Trabalho", date(2027, 1, 15), "Falta assinatura"),
          9: ("Pasta SST/Histórico", "", None, ""),
          10: ("Planilha NR-1 / Plano de ação", "Coordenação de RH", date(2026, 12, 5), ""),
          11: ("", "Coordenação de RH", date(2027, 1, 31), ""),
          12: ("Pasta Escutas", "Coordenação de RH", date(2026, 12, 20), "5 escutas feitas"),
          13: ("", "", date(2027, 1, 31), ""),
          15: ("Pasta SST/Ocorrências", "", None, ""),
          16: ("Pasta SST/PCMSO", "Segurança do Trabalho", date(2026, 12, 5), ""),
          17: ("Pasta Gestão NR-1/Portaria 01", "", None, ""),
          18: ("", "", None, "Pasta ainda sem organização"),
          19: ("Pasta CIPA/Ata de posse", "", None, ""),
          20: ("Intranet", "Coordenação de RH", date(2026, 12, 20), "Versão inicial publicada"),
          21: ("", "Coordenação de RH", date(2026, 10, 30), ""),
          22: ("Pasta Treinamentos 2026/lista 03", "Coordenação de RH", date(2026, 12, 15), "76% da equipe presente na primeira turma"),
          23: ("", "Coordenação de RH", None, "")}
    for i, (tema, exig, base, evid) in enumerate(DIAGNOSTICO):
        r = D1 + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, tema)
        ws.cell(r, 3, exig)
        ws.cell(r, 4, base)
        ws.cell(r, 5, evid)
        ws.cell(r, 6, situacoes[i])
        if demo and (i + 1) in ex:
            loc, resp, prazo, obs = ex[i + 1]
            ws.cell(r, 7, loc or None)
            ws.cell(r, 8, resp or None)
            ws.cell(r, 9, prazo)
            ws.cell(r, 10, obs or None)
        ws.cell(r, 11, (f'=IF(F{r}="","Informe a situação",'
                        f'IF(AND(F{r}="Atendido",G{r}=""),"Atendido sem local do documento",'
                        f'IF(AND(OR(F{r}="Não atendido",F{r}="Parcial"),H{r}=""),"Sem responsável",'
                        f'IF(AND(OR(F{r}="Não atendido",F{r}="Parcial"),I{r}=""),"Sem prazo",'
                        f'IF(AND(OR(F{r}="Não atendido",F{r}="Parcial"),I{r}<TODAY()),"Prazo vencido","")))))'))
    lista(ws, "=Listas!$A$2:$A$6", f"F{D1}:F{D2}")
    validar_data(ws, f"I{D1}:I{D2}")
    realce_alerta(ws, f"A{D1}:K{D2}", "K", D1)


# --- Setores ----------------------------------------------------------------
def aba_setores(ws):
    config_aba(ws, 100, "E4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais são os setores e há quanto tempo cada um foi ouvido?    "
              "O QUE VOCÊ FAZ AQUI: confira turnos, número de pessoas e líder de cada setor (colunas amarelas). "
              "As colunas cinza se preenchem sozinhas a partir da aba Escuta.", 10)
    cabecalho(ws, 3, ["GHE", "Unidade", "Lotação", "Setor", "Turnos", "Nº de pessoas", "Líder do setor",
                      "Observação do inventário", "Última escuta", "Alerta"],
              [9, 22, 24, 26, 20, 11, 28, 40, 13, 30])
    pintar(ws, S1, S2, {5, 6, 7}, {1, 2, 3, 4, 8, 9, 10}, 10, fmt_datas={9}, altura=36)
    for i, s in enumerate(SETORES):
        r = S1 + i
        for c, v in enumerate(s, start=1):
            ws.cell(r, c, v if v != "" else None)
        esc_c = rng("Escuta", "C", E1, E2)
        esc_a = rng("Escuta", "A", E1, E2)
        ws.cell(r, 9, f'=IF(D{r}="","",IF(COUNTIF({esc_c},D{r})=0,"Nunca",SUMPRODUCT(MAX(({esc_c}=D{r})*{esc_a}))))')
        ws.cell(r, 10, (f'=IF(D{r}="","",IF(F{r}="","Falta nº de pessoas",IF(I{r}="Nunca","Nunca houve escuta neste setor",'
                        f'IF(TODAY()-I{r}>90,"Mais de 90 dias sem escuta",""))))'))
        ws.cell(r, 9).alignment = Alignment(horizontal="center", vertical="top")
    validar_inteiro(ws, f"F{S1}:F{S2}", 0, 99999, "Digite um número inteiro de pessoas.")
    realce_alerta(ws, f"A{S1}:J{S2}", "J", S1)


# --- Levantamento psicossocial ---------------------------------------------
def aba_levantamento(ws, demo):
    config_aba(ws, 85, "C4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais fatores psicossociais existem em cada setor e qual o nível de risco?    "
              "O QUE VOCÊ FAZ AQUI: para cada linha, responda se o fator existe. Se existir, descreva o que foi observado "
              "(sem nomes de pessoas) e dê notas de 1 a 5 para probabilidade e severidade.", 14)
    cabecalho(ws, 3, ["Setor", "Fator psicossocial", "O fator existe neste setor?", "Como foi identificado", "Data",
                      "O que foi observado (sem nomes)", "Medidas que já existem", "Prob. (1-5)", "Sev. (1-5)",
                      "Nível (P x S)", "Classificação", "Precisa de ação?", "Nº da ação", "Alerta"],
              [20, 32, 12, 19, 11, 30, 22, 6, 6, 7, 12, 9, 7, 24], altura=48)
    pintar(ws, L1, L2, {3, 4, 5, 6, 7, 8, 9, 13}, {1, 2, 10, 11, 12, 14}, 14, fmt_datas={5}, altura=34)
    # (setor, fator) -> (existe, como, data, observado, medidas, P, S, nº ação)
    d = {
        (4, 9): ("Sim", "Observação do trabalho", date(2026, 11, 5), "Tarefa repetitiva em parte do turno (hipótese)", "Rodízio informal", 3, 2, None),
        (4, 10): ("Sim", "Escuta coletiva", date(2026, 11, 12), "Picos de pedido no fim do mês", "Hora extra pontual", 3, 3, 13),
        (4, 5): ("Sim", "Escuta coletiva", date(2026, 11, 12), "Operadores novos relatam pouco acompanhamento", "Colega de apoio", 2, 3, 7),
        (4, 3): ("Não", "Escuta coletiva", date(2026, 11, 12), "Sem queixas", "Descrição de funções atualizada", None, None, None),
        (4, 1): ("Não", "Escuta coletiva", date(2026, 11, 12), "Sem relatos", None, None, None, None),
        (5, 9): ("Sim", "Observação do trabalho", date(2026, 10, 10), "Monotonia na tarefa de inspeção (hipótese)", None, 4, 3, 5),
        (5, 10): ("Sim", "Observação do trabalho", date(2026, 10, 10), "Possível sobrecarga no turno noturno em dias de pico (hipótese)", None, 3, 4, 5),
        (5, 8): ("Não", "Observação do trabalho", date(2026, 10, 10), "Sem ocorrências", None, None, None, None),
        (6, 10): ("Sim", "Entrevista com líder", date(2026, 11, 3), "Picos de expedição no fim do mês", "Reforço temporário", 4, 2, 12),
        (6, 2): ("Sim", "Observação do trabalho", date(2026, 11, 3), "Troca do sistema de controle de estoque sem treinamento prévio", None, 3, 3, None),
        (6, 13): ("Não", "Observação do trabalho", date(2026, 11, 3), "Equipe sempre em dupla", None, None, None, None),
        (8, 13): ("Sim", "Observação do trabalho", date(2026, 10, 20), "Vigia noturno trabalha sozinho", "Rádio de comunicação", 2, 4, 14),
        (1, 10): ("Sim", "Escuta coletiva", date(2026, 10, 15), "Fechamento mensal concentra tarefas", None, 2, 2, None),
        (1, 1): ("Não", "Escuta coletiva", date(2026, 10, 15), "Sem relatos", None, None, None, None),
        (7, 10): ("Sim", "Entrevista com líder", date(2026, 11, 5), "Chamados emergenciais fora do horário", "Escala informal", 3, 3, 15),
        (3, 6): ("Sim", "Canal de relatos", date(2026, 11, 8), "Pouca autonomia para priorizar laudos", None, None, None, None),
    }
    r = L1
    for si, setor in enumerate(NOMES_SETORES, start=1):
        for fi, (fator, _) in enumerate(FATORES, start=1):
            ws.cell(r, 1, setor)
            ws.cell(r, 2, fator)
            ws.cell(r, 3, "A investigar")
            if demo and (si, fi) in d:
                ex, como, dt, obs, med, p, s, na = d[(si, fi)]
                ws.cell(r, 3, ex)
                ws.cell(r, 4, como)
                ws.cell(r, 5, dt)
                ws.cell(r, 6, obs)
                ws.cell(r, 7, med)
                ws.cell(r, 8, p)
                ws.cell(r, 9, s)
                ws.cell(r, 13, na)
            ws.cell(r, 10, f'=IF(OR(H{r}="",I{r}=""),"",H{r}*I{r})')
            ws.cell(r, 11, f'=IF(OR(H{r}="",I{r}=""),"",INDEX(Listas!$U$2:$Y$6,H{r},I{r}))')
            ws.cell(r, 12, f'=IF(K{r}="","",IF(OR(K{r}="Moderado",K{r}="Substancial",K{r}="Intolerável"),"Sim","Não"))')
            ws.cell(r, 14, (f'=IF(AND(C{r}="Sim",OR(H{r}="",I{r}="")),"Informe probabilidade e severidade",'
                            f'IF(AND(L{r}="Sim",M{r}=""),"Risco sem ação vinculada",'
                            f'IF(AND(C{r}<>"A investigar",C{r}<>"",E{r}=""),"Falta a data",'
                            f'IF(AND(E{r}<>"",E{r}>TODAY()),"Data futura em fato ocorrido",""))))'))
            for c in (3, 8, 9, 10, 12, 13):
                ws.cell(r, c).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            r += 1
    lista(ws, "=Listas!$B$2:$B$4", f"C{L1}:C{L2}")
    lista(ws, "=Listas!$C$2:$C$8", f"D{L1}:D{L2}")
    validar_data(ws, f"E{L1}:E{L2}")
    validar_inteiro(ws, f"H{L1}:I{L2}", 1, 5, "Digite um número de 1 a 5.")
    realce_alerta(ws, f"A{L1}:N{L2}", "N", L1)
    ws.conditional_formatting.add(f"K{L1}:K{L2}", FormulaRule(
        formula=[f'OR(K{L1}="Substancial",K{L1}="Intolerável")'], font=Font(name=FONTE, bold=True, color="C00000")))


# --- Plano de ação 5W2H -----------------------------------------------------
def aba_plano(ws, demo):
    config_aba(ws, 75, "E4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais ações estão atrasadas e quem é o responsável por cada uma?    "
              "O QUE VOCÊ FAZ AQUI: uma ação por linha, sempre com responsável e prazo. Quando terminar, preencha a data "
              "de conclusão e diga se funcionou. A situação calculada aparece sozinha.", 19)
    cabecalho(ws, 3, ["Nº", "Origem", "Setor", "O quê (ação)", "Por quê", "Onde", "Quem (responsável)", "Início", "Prazo",
                      "Como", "Quanto custa", "Tipo de medida", "Situação informada", "Data de conclusão", "Funcionou?",
                      "Evidência / onde está guardada", "Situação calculada", "Dias de atraso", "Alerta"],
              [4, 12, 13, 24, 18, 11, 14, 11, 11, 16, 10, 16, 12, 11, 10, 14, 12, 7, 22], altura=48)
    pintar(ws, P1, P2, {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}, {1, 17, 18, 19}, 19,
           fmt_datas={8, 9, 14}, altura=52)
    base = [  # Nº, origem, setor, o quê, por quê, onde, quem, início, prazo, como, quanto, tipo, situação, conclusão, funcionou, evidência
        (1, "Diagnóstico", "Toda a empresa", "Designar formalmente o responsável pelo tema NR-1", "É a primeira pergunta do auditor",
         "Diretoria", "Diretoria", date(2026, 9, 1), date(2026, 9, 30), "Portaria assinada", "Sem custo",
         "Mudança na organização do trabalho", "Concluída", date(2026, 9, 28), "Sim", "Pasta Gestão NR-1/Portaria 01"),
        (2, "Diagnóstico", "Toda a empresa", "Pedir à consultoria a AEP com fatores psicossociais e a revisão do inventário",
         "Fecha lacunas do inventário", "Consultoria", "RH e Segurança do Trabalho", date(2026, 10, 1), date(2026, 11, 30),
         "Solicitação formal por e-mail", "A levantar", "Mudança na organização do trabalho", "Em andamento", None, None, None),
        (3, "Levantamento", "Produção - Linha B", "Concluir a escuta coletiva e o preenchimento do levantamento",
         "Evidência de participação e base do inventário", "Galpão 2", "Coordenação de RH", date(2026, 10, 5), date(2026, 10, 31),
         "Roteiro curto, grupos por turno, registro sem nomes", "Horas internas", "Medida coletiva", "Em andamento", None, None, None),
        (4, "Levantamento", "Produção - Linha A", "Fazer a escuta coletiva do setor", "Evidência de participação",
         "Galpão 1", "Coordenação de RH", date(2026, 11, 3), date(2026, 11, 15), "Roteiro curto, registro sem nomes",
         "Horas internas", "Medida coletiva", "Concluída", date(2026, 11, 12), "Ainda não avaliado", "Pasta Escutas/ata 01"),
        (5, "Levantamento", "Produção - Linha B", "Revisar pausas e rodízio de tarefas no turno noturno (hipótese a investigar)",
         "Reduzir monotonia e sobrecarga em dias de pico", "Galpão 2", "Supervisão de produção B", date(2026, 11, 20), date(2026, 12, 15),
         "Teste de rodízio por 30 dias com a equipe", "Sem custo", "Mudança na organização do trabalho", "Não iniciada", None, None, None),
        (6, "Diagnóstico", "Toda a empresa", "Realizar a capacitação anual sobre assédio, violência, igualdade e diversidade",
         "Exigência legal", "Auditório", "Coordenação de RH", date(2026, 9, 1), date(2026, 12, 15), "Turmas por turno, lista de presença",
         "Horas internas", "Treinamento/orientação", "Em andamento", None, None, "Pasta Treinamentos 2026"),
        (7, "Diagnóstico", "Toda a empresa", "Formar as lideranças em fatores psicossociais e acolhimento de novatos",
         "Evidência de medida de prevenção", "Sala de reunião", "Coordenação de RH", date(2026, 12, 1), date(2027, 1, 31),
         "Oficina de 4 horas", "A levantar", "Treinamento/orientação", "Não iniciada", None, None, None),
        (8, "Diagnóstico", "Toda a empresa", "Implantar canal de denúncias com anonimato e procedimento de apuração",
         "Exigência legal", "Intranet", "Coordenação de RH", date(2026, 10, 15), date(2026, 11, 30), "Formulário anônimo e fluxo escrito",
         "A levantar", "Medida coletiva", "Em andamento", None, None, None),
        (9, "Diagnóstico", "Toda a empresa", "Publicar as regras de conduta sobre assédio e violência", "Exigência legal",
         "Intranet e murais", None, date(2026, 11, 1), date(2026, 12, 20), "Texto aprovado pela diretoria", "Sem custo",
         "Medida coletiva", "Não iniciada", None, None, None),
        (10, "Auditoria", "Toda a empresa", "Comunicar riscos e medidas na integração de novos empregados", "Comunicação dos riscos",
         "Sala de integração", "Segurança do Trabalho", date(2026, 7, 1), date(2026, 8, 31), "Slide único por setor", "Sem custo",
         "Treinamento/orientação", "Concluída", date(2026, 8, 25), "Ainda não avaliado", "Pasta Integração"),
        (11, "Diagnóstico", "Toda a empresa", "Reunir com o médico do trabalho para integrar o inventário ao PCMSO", "Integração com o PCMSO",
         "Ambulatório", "Segurança do Trabalho", date(2026, 11, 20), None, "Reunião com ata", "Sem custo",
         "Mudança na organização do trabalho", "Não iniciada", None, None, None),
        (12, "Levantamento", "Almoxarifado e Expedição", "Ajustar o planejamento de pessoal nos picos de expedição",
         "Reduzir sobrecarga no fim do mês", "Centro de Distribuição", "Supervisão de logística", date(2026, 12, 1), date(2027, 2, 28),
         "Calendário de picos e reforço temporário", "A levantar", "Mudança na organização do trabalho", "Não iniciada", None, None, None),
        (13, "Escuta", "Produção - Linha A", "Rever o planejamento dos picos de pedido com o PCP", "Reduzir sobrecarga",
         "Galpão 1", "Coordenação da qualidade", date(2026, 12, 1), date(2027, 1, 31), "Reunião mensal PCP e produção", "Sem custo",
         "Mudança na organização do trabalho", "Não iniciada", None, None, None),
        (14, "Levantamento", "Portaria e Vigilância", "Criar rotina de contato periódico e botão de pedido de ajuda para o vigia noturno",
         "Reduzir isolamento do trabalho noturno", "Portaria", "Coordenação de serviços gerais", date(2026, 11, 25), date(2027, 1, 15),
         "Ronda de contato a cada 2 horas", "A levantar", "Medida coletiva", "Não iniciada", None, None, None),
        (15, "Levantamento", "Manutenção", "Formalizar a escala de sobreaviso e compensação", "Reduzir chamados fora do horário",
         "Oficina", "Supervisão de manutenção", date(2026, 12, 1), date(2027, 1, 31), "Escala escrita e aprovada", "Sem custo",
         "Mudança na organização do trabalho", "Não iniciada", None, None, None),
    ]
    if not demo:   # modelo vazio: só as 5 ações iniciais, sem responsável nem prazo
        base = [
            (1, "Diagnóstico", "Toda a empresa", "Designar formalmente o responsável pelo tema NR-1", "É a primeira pergunta do auditor",
             None, None, None, None, "Documento assinado", "Sem custo", "Mudança na organização do trabalho", "Não iniciada", None, None, None),
            (2, "Diagnóstico", "Toda a empresa", "Pedir à consultoria a AEP com fatores psicossociais e a revisão do inventário",
             "Fecha lacunas do inventário", None, None, None, None, "Solicitação formal", "A levantar", "Mudança na organização do trabalho",
             "Não iniciada", None, None, None),
            (3, "Levantamento", "Toda a empresa", "Fazer a escuta coletiva por setor e preencher o levantamento",
             "Evidência de participação e base do inventário", None, None, None, None, "Roteiro curto, registro sem nomes", "Horas internas",
             "Medida coletiva", "Não iniciada", None, None, None),
            (4, "Diagnóstico", "Toda a empresa", "Confirmar a situação da CIPA e do canal de denúncias e regularizar", "Exigência legal",
             None, None, None, None, "Verificar atas e canal", "Sem custo", "Medida coletiva", "Não iniciada", None, None, None),
            (5, "Diagnóstico", "Toda a empresa", "Capacitação anual sobre assédio e violência e formação de lideranças", "Exigência legal e evidência",
             None, None, None, None, "Turmas com lista de presença", "A levantar", "Treinamento/orientação", "Não iniciada", None, None, None),
        ]
    for i, row in enumerate(base):
        r = P1 + i
        n, org, setor, oque, porque, onde, quem, ini, prazo, como, quanto, tipo, sit, conc, func, evid = row
        for c, v in zip(range(1, 17), (n, org, setor, oque, porque, onde, quem, ini, prazo, como, quanto, tipo, sit, conc, func, evid)):
            ws.cell(r, c, v)
    for r in range(P1, P2 + 1):
        ws.cell(r, 17, (f'=IF(D{r}="","",IF(M{r}="Cancelada","Cancelada",IF(M{r}="Concluída",'
                        f'IF(O{r}="Sim","Concluída e eficaz","Concluída"),IF(I{r}="","Sem prazo",'
                        f'IF(I{r}<TODAY(),"Atrasada","No prazo")))))'))
        ws.cell(r, 18, f'=IF(Q{r}="Atrasada",TODAY()-I{r},"")')
        ws.cell(r, 19, (f'=IF(D{r}="","",IF(C{r}="","Falta o setor",IF(G{r}="","Ação sem responsável",'
                        f'IF(I{r}="","Ação sem prazo",IF(AND(H{r}<>"",I{r}<H{r}),"Prazo anterior ao início",'
                        f'IF(AND(N{r}<>"",H{r}<>"",N{r}<H{r}),"Conclusão anterior ao início",'
                        f'IF(AND(N{r}<>"",N{r}>TODAY()),"Data de conclusão futura",'
                        f'IF(Q{r}="Atrasada","Prazo vencido",'
                        f'IF(AND(M{r}="Concluída",N{r}<>"",TODAY()-N{r}>60,OR(O{r}="",O{r}="Ainda não avaliado")),'
                        f'"Concluída há mais de 60 dias sem avaliar eficácia","")))))))))'))
        ws.cell(r, 18).number_format = "0"
        ws.cell(r, 1).alignment = Alignment(horizontal="center", vertical="top")
    lista(ws, "=Listas!$D$2:$D$6", f"B{P1}:B{P2}")
    lista(ws, "=Listas!$J$2:$J$10", f"C{P1}:C{P2}")
    lista(ws, "=Listas!$E$2:$E$6", f"L{P1}:L{P2}")
    lista(ws, "=Listas!$F$2:$F$5", f"M{P1}:M{P2}")
    lista(ws, "=Listas!$G$2:$G$5", f"O{P1}:O{P2}")
    validar_data(ws, f"H{P1}:I{P2}")
    validar_data(ws, f"N{P1}:N{P2}")
    realce_alerta(ws, f"A{P1}:S{P2}", "S", P1)


# --- Treinamentos -----------------------------------------------------------
def aba_treinamentos(ws, demo):
    config_aba(ws, 90, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais treinamentos estão vencidos ou vão vencer e quanto da equipe foi treinada?    "
              "O QUE VOCÊ FAZ AQUI: no REGISTRO (abaixo), escolha o treinamento, informe a data, as presenças e onde está a "
              "lista de presença. O vencimento é calculado sozinho.", 12)
    cabecalho(ws, 3, ["Treinamento (catálogo)", "Para quem", "Por que é exigido ou recomendado", "Tipo", "Validade (meses)",
                      "Pode ser interno?"], [44, 22, 30, 24, 14, 14, 12, 18, 12, 16, 11, 28], altura=34)
    pintar(ws, T_CAT1, T_CAT2, set(), {1, 2, 3, 4, 5, 6}, 6)
    for i, t in enumerate(TREINAMENTOS):
        for c, v in enumerate(t, start=1):
            ws.cell(T_CAT1 + i, c, v)
    ws.cell(12, 1, "REGISTRO DOS TREINAMENTOS REALIZADOS (preencha as células amarelas)").font = Font(name=FONTE, size=12, bold=True, color=AZUL)
    cabecalho(ws, 14, ["Treinamento", "Data", "Instrutor / empresa", "Carga horária (h)", "Setores", "Nº de convocados",
                       "Nº de presentes", "Onde está a lista de presença", "Vencimento", "Situação", "% de presença", "Alerta"],
              None, altura=48)
    pintar(ws, T1, T2, {1, 2, 3, 4, 5, 6, 7, 8}, {9, 10, 11, 12}, 12, fmt_datas={2, 9}, altura=32)
    if demo:
        reg = [
            (TREINAMENTOS[0][0], date(2026, 9, 10), "Coordenação de RH (interno)", 2, "Todos", 129, 98, "Pasta Treinamentos 2026/lista 03"),
            (TREINAMENTOS[1][0], date(2025, 11, 15), "Consultoria externa", 8, "CIPA", 9, 9, "Pasta CIPA/lista 01"),
            (TREINAMENTOS[3][0], date(2025, 12, 3), "Segurança do Trabalho (interno)", 4, "Produção A e B", 72, 61, "Pasta Treinamentos 2025/lista 07"),
            (TREINAMENTOS[5][0], date(2026, 10, 22), "Consultoria externa", 16, "RH e Segurança do Trabalho", 3, 3, "Pasta Treinamentos 2026/lista 05"),
            (TREINAMENTOS[4][0], date(2026, 11, 5), "Coordenação de RH (interno)", 3, "Comissão de apuração", 4, 4, None),
        ]
        for i, row in enumerate(reg):
            for c, v in enumerate(row, start=1):
                ws.cell(T1 + i, c, v)
    cat_n = f"$A${T_CAT1}:$A${T_CAT2}"
    cat_v = f"$E${T_CAT1}:$E${T_CAT2}"
    for r in range(T1, T2 + 1):
        ws.cell(r, 9, f'=IF(OR(A{r}="",B{r}=""),"",EDATE(B{r},INDEX({cat_v},MATCH(A{r},{cat_n},0))))')
        ws.cell(r, 10, f'=IF(I{r}="","",IF(I{r}<TODAY(),"Vencido",IF(I{r}<=TODAY()+60,"Vence em 60 dias","Em dia")))')
        ws.cell(r, 11, f'=IF(OR(F{r}="",F{r}=0,G{r}=""),"",G{r}/F{r})')
        ws.cell(r, 11).number_format = "0%"
        ws.cell(r, 12, (f'=IF(COUNTA(A{r}:H{r})=0,"",IF(A{r}="","Falta o treinamento",IF(B{r}="","Falta a data",'
                        f'IF(B{r}>TODAY(),"Data futura em treinamento já realizado",IF(J{r}="Vencido","Treinamento vencido",'
                        f'IF(AND(F{r}<>"",G{r}<>"",G{r}>F{r}),"Presentes maior que convocados",'
                        f'IF(H{r}="","Falta o local da lista de presença","")))))))'))
    lista(ws, f"=Treinamentos!$A${T_CAT1}:$A${T_CAT2}", f"A{T1}:A{T2}")
    validar_data(ws, f"B{T1}:B{T2}")
    validar_inteiro(ws, f"F{T1}:G{T2}", 0, 99999, "Digite um número inteiro.")
    realce_alerta(ws, f"A{T1}:L{T2}", "L", T1)
    ws.column_dimensions["A"].width = 46


# --- Escuta -----------------------------------------------------------------
def aba_escuta(ws, demo):
    config_aba(ws, 100, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: os trabalhadores de cada setor foram ouvidos? Quando foi a última vez?    "
              "O QUE VOCÊ FAZ AQUI: registre cada reunião ou escuta. Só contagens e temas, nunca nomes ou relatos "
              "que identifiquem alguém.", 8)
    cabecalho(ws, 3, ["Data", "Tipo", "Setor", "Nº de participantes", "Temas levantados (sem nomes)", "Virou ação? Nº",
                      "Onde está a ata", "Alerta"], [12, 24, 26, 13, 40, 12, 28, 30], altura=34)
    pintar(ws, E1, E2, {1, 2, 3, 4, 5, 6, 7}, {8}, 8, fmt_datas={1}, altura=30)
    if demo:
        reg = [
            (date(2026, 11, 12), "Escuta coletiva por setor", "Produção - Linha A", 14, "Picos de pedido; apoio a novos operadores", 13, "Pasta Escutas/ata 05"),
            (date(2026, 11, 5), "Reunião da CIPA", "Toda a empresa", 9, "Pausas no turno noturno", 5, "Ata CIPA 11/2026"),
            (date(2026, 10, 28), "Diálogo de segurança", "Almoxarifado e Expedição", 8, "Picos de expedição", 12, "Pasta Escutas/ata 04"),
            (date(2026, 10, 15), "Escuta coletiva por setor", "Administrativo", 10, "Fechamento mensal", None, "Pasta Escutas/ata 03"),
            (date(2026, 10, 10), "Escuta coletiva por setor", "Produção - Linha B", 18, "Monotonia na inspeção; escala noturna", 5, None),
        ]
        for i, row in enumerate(reg):
            for c, v in enumerate(row, start=1):
                ws.cell(E1 + i, c, v)
    for r in range(E1, E2 + 1):
        ws.cell(r, 8, (f'=IF(COUNTA(A{r}:G{r})=0,"",IF(A{r}="","Falta a data",IF(A{r}>TODAY(),"Data futura em fato ocorrido",'
                       f'IF(C{r}="","Falta o setor",IF(D{r}="","Falta nº de participantes",'
                       f'IF(G{r}="","Falta onde está a ata",""))))))'))
    lista(ws, "=Listas!$H$2:$H$6", f"B{E1}:B{E2}")
    lista(ws, "=Listas!$J$2:$J$10", f"C{E1}:C{E2}")
    validar_data(ws, f"A{E1}:A{E2}")
    validar_inteiro(ws, f"D{E1}:D{E2}", 0, 99999, "Digite um número inteiro.")
    realce_alerta(ws, f"A{E1}:H{E2}", "H", E1)


# --- Acompanhamento mensal --------------------------------------------------
def aba_acompanhamento(ws, demo):
    config_aba(ws, 100, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: os indicadores estão melhorando em relação ao mês anterior e à meta?    "
              "O QUE VOCÊ FAZ AQUI: no último dia útil do mês, copie o valor do quadro 'Valores de hoje' (à direita) e "
              "cole como VALOR numa nova linha. Os indicadores manuais só entram no Painel com 'Dado validado = Sim'.", 11)
    cabecalho(ws, 3, ["Mês (1º dia)", "Indicador", "Valor", "Meta (estimativa)", "Fonte do dado", "Dado validado?",
                      "Variação x mês anterior", "Atingiu a meta?"], [13, 30, 11, 13, 26, 12, 14, 12], altura=48)
    pintar(ws, A1, A2, {1, 2, 3, 4, 5, 6}, {7, 8}, 8, fmt_datas={1}, altura=None)
    for r in range(A1, A2 + 1):
        ws.cell(r, 1).number_format = DATA
        for c in (3, 4, 7):
            ws.cell(r, c).number_format = "0.0%"
        ws.cell(r, 7, (f'=IF(OR(C{r}="",F{r}="Não"),"",IF(COUNTIFS($B${A1}:$B${A2},B{r},$A${A1}:$A${A2},EDATE(A{r},-1),'
                       f'$C${A1}:$C${A2},"<>")=0,"",C{r}-SUMIFS($C${A1}:$C${A2},$B${A1}:$B${A2},B{r},$A${A1}:$A${A2},EDATE(A{r},-1))))'))
        ws.cell(r, 8, (f'=IF(OR(C{r}="",D{r}="",F{r}="Não"),"",IF(INDEX(Listas!$M$2:$M$10,MATCH(B{r},Listas!$K$2:$K$10,0))="Menor",'
                       f'IF(C{r}<=D{r},"Sim","Não"),IF(C{r}>=D{r},"Sim","Não")))'))
    if demo:
        dados = [
            (date(2026, 5, 1), "% de conformidade", 0.03, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 6, 1), "% de conformidade", 0.06, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 7, 1), "% de conformidade", 0.08, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 8, 1), "% de conformidade", 0.10, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 9, 1), "% de conformidade", 0.21, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 10, 1), "% de conformidade", 0.40, 0.80, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 9, 1), "% de ações no prazo", 0.80, 0.90, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 10, 1), "% de ações no prazo", 0.70, 0.90, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 9, 1), "% de fatores avaliados", 0.10, 1.00, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 10, 1), "% de fatores avaliados", 0.18, 1.00, "Quadro 'Valores de hoje'", "Sim"),
            (date(2026, 10, 1), "Absenteísmo", None, None, "A levantar com o DP", "Não"),
            (date(2026, 10, 1), "Rotatividade", None, None, "A levantar com o DP", "Não"),
        ]
        for i, row in enumerate(dados):
            for c, v in enumerate(row, start=1):
                ws.cell(A1 + i, c, v)
    lista(ws, "=Listas!$K$2:$K$10", f"B{A1}:B{A2}")
    lista(ws, "=Listas!$O$2:$O$3", f"F{A1}:F{A2}")
    validar_data(ws, f"A{A1}:A{A2}")
    # indicadores do tipo "nº" ficam sem formato de porcentagem
    dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=170, formatCode="0"))
    from openpyxl.formatting.rule import Rule
    regra = Rule(type="expression", dxf=dxf, formula=[f'INDEX(Listas!$L$2:$L$10,MATCH($B{A1},Listas!$K$2:$K$10,0))="nº"'])
    ws.conditional_formatting.add(f"C{A1}:D{A2}", regra)
    ws.conditional_formatting.add(f"G{A1}:G{A2}", regra)
    # quadro de valores de hoje
    for c, w in zip("JK", (34, 14)):
        ws.column_dimensions[c].width = w
    ws.column_dimensions["I"].width = 3
    ws["J3"].value, ws["K3"].value = "Valores de hoje (automático)", "Valor"
    for a in ("J3", "K3"):
        ws[a].font, ws[a].fill, ws[a].alignment, ws[a].border = FONT_HDR, F_HDR, CENTRO, BORDA
    dj, dk = rng(N_DIAG, "F", D1, D2), rng(N_PLANO, "Q", P1, P2)
    lv = N_LEV
    tr_a, tr_j, tr_g = rng("Treinamentos", "A", T1, T2), rng("Treinamentos", "J", T1, T2), rng("Treinamentos", "G", T1, T2)
    esc_a = rng("Escuta", "A", E1, E2)
    linhas = [
        ("% de conformidade", f'=IFERROR((COUNTIF({dj},"Atendido")+0.5*COUNTIF({dj},"Parcial"))/(COUNTA({rng(N_DIAG,"A",D1,D2)})-COUNTIF({dj},"Não se aplica")),"")', "0.0%"),
        ("% de ações no prazo", (f'=IFERROR((COUNTIF({dk},"No prazo")+COUNTIF({dk},"Concluída")+COUNTIF({dk},"Concluída e eficaz"))/'
                                 f'(COUNTIF({dk},"No prazo")+COUNTIF({dk},"Atrasada")+COUNTIF({dk},"Concluída")+COUNTIF({dk},"Concluída e eficaz")+COUNTIF({dk},"Sem prazo")),"")'), "0.0%"),
        ("% de fatores avaliados", (f'=IFERROR((COUNTIF({rng(lv,"C",L1,L2)},"Não")+COUNTIFS({rng(lv,"C",L1,L2)},"Sim",{rng(lv,"H",L1,L2)},"<>",{rng(lv,"I",L1,L2)},"<>"))/COUNTA({rng(lv,"A",L1,L2)}),"")'), "0.0%"),
        ("% de pessoas treinadas", (f'=IFERROR(MIN(1,SUMIFS({tr_g},{tr_a},Treinamentos!$A${T_CAT1},{tr_j},"<>Vencido")/SUM({rng("Setores","F",S1,S2)})),"")'), "0.0%"),
        ("Nº de escutas no mês", f'=COUNTIFS({esc_a},">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),{esc_a},"<="&TODAY())', "0"),
    ]
    for i, (nome, f, fmt) in enumerate(linhas):
        r = 4 + i
        ws.cell(r, 10, nome).font = FONT
        ws.cell(r, 11, f).number_format = fmt
        for c in (10, 11):
            ws.cell(r, c).fill, ws.cell(r, c).border = F_CINZA, BORDA
            ws.cell(r, c).font = FONT
    ws["J10"].value = "Não há valor automático para relatos, afastamentos, rotatividade e absenteísmo: dependem de dados validados pelo DP."
    ws["J10"].font = FONT_CINZA
    ws["J10"].alignment = QUEBRA
    ws.merge_cells("J10:K13")


# --- Alertas ----------------------------------------------------------------
def aba_alertas(ws):
    config_aba(ws, 100, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais problemas de preenchimento ou de prazo precisam de atenção agora?    "
              "O QUE VOCÊ FAZ AQUI: leia a coluna 'O que fazer' e corrija na aba indicada. Nada é apagado nem corrigido "
              "sozinho: a linha só fica laranja na aba de origem.", 4)
    cabecalho(ws, 3, ["Regra", "Onde corrigir", "Quantidade", "O que fazer"], [48, 30, 12, 62], altura=30)
    D, P, L, T, E, S = (lambda col: rng(N_DIAG, col, D1, D2)), (lambda col: rng(N_PLANO, col, P1, P2)), \
        (lambda col: rng(N_LEV, col, L1, L2)), (lambda col: rng("Treinamentos", col, T1, T2)), \
        (lambda col: rng("Escuta", col, E1, E2)), (lambda col: rng("Setores", col, S1, S2))
    regras = [
        ("Campo obrigatório vazio", "Setores, Levantamento, Plano, Treinamentos, Escuta",
         f'=COUNTIF({S("J")},"Falta*")+COUNTIF({L("N")},"Falta*")+COUNTIF({P("S")},"Falta*")+COUNTIF({T("L")},"Falta*")+COUNTIF({E("H")},"Falta*")',
         "Complete o campo indicado na coluna Alerta da linha laranja."),
        ("Ação sem responsável", "Plano de ação, Diagnóstico NR-1",
         f'=COUNTIF({P("S")},"Ação sem responsável")+COUNTIF({D("K")},"Sem responsável")',
         "Escolha quem vai tocar a ação. Sem dono, a ação não anda."),
        ("Ação sem prazo", "Plano de ação, Diagnóstico NR-1",
         f'=COUNTIF({P("S")},"Ação sem prazo")+COUNTIF({D("K")},"Sem prazo")', "Combine uma data com o responsável e preencha."),
        ("Prazo vencido", "Plano de ação, Diagnóstico NR-1",
         f'=COUNTIF({P("S")},"Prazo vencido")+COUNTIF({D("K")},"Prazo vencido")', "Conclua a ação ou combine um novo prazo e registre na observação."),
        ("Data de conclusão ou prazo anterior ao início", "Plano de ação",
         f'=COUNTIF({P("S")},"Conclusão anterior ao início")+COUNTIF({P("S")},"Prazo anterior ao início")', "Confira as datas: a data não pode vir antes do início."),
        ("Data futura em algo que já aconteceu", "Plano, Levantamento, Treinamentos, Escuta",
         f'=COUNTIF({P("S")},"Data de conclusão futura")+COUNTIF({L("N")},"Data futura*")+COUNTIF({T("L")},"Data futura*")+COUNTIF({E("H")},"Data futura*")',
         "Corrija a data digitada (talvez o ano esteja errado)."),
        ("Fator 'Sim' sem probabilidade ou severidade", "Levantamento", f'=COUNTIF({L("N")},"Informe*")', "Dê notas de 1 a 5 para probabilidade e severidade."),
        ("Risco Moderado ou acima sem ação vinculada", "Levantamento", f'=COUNTIF({L("N")},"Risco sem ação*")',
         "Crie uma ação no Plano e digite o número dela na coluna 'Nº da ação'."),
        ("Treinamento vencido", "Treinamentos", f'=COUNTIF({T("L")},"Treinamento vencido")', "Programe a reciclagem e registre no REGISTRO."),
        ("Setor sem escuta há mais de 90 dias (ou nunca ouvido)", "Setores, Escuta",
         f'=COUNTIF({S("J")},"Mais de 90*")+COUNTIF({S("J")},"Nunca*")', "Programe uma escuta coletiva com o setor."),
        ("Exigência 'Atendido' sem local do documento", "Diagnóstico NR-1", f'=COUNTIF({D("K")},"Atendido sem*")',
         "Informe onde o documento está guardado (o auditor vai pedir)."),
        ("Ação concluída há mais de 60 dias sem avaliar eficácia", "Plano de ação", f'=COUNTIF({P("S")},"Concluída há mais*")',
         "Verifique se a ação funcionou e preencha a coluna 'Funcionou?'."),
        ("Presentes maior que convocados", "Treinamentos", f'=COUNTIF({T("L")},"Presentes maior*")', "Revise os números de convocados e presentes."),
    ]
    for i, (reg, onde, f, fazer) in enumerate(regras):
        r = 4 + i
        for c, v in enumerate((reg, onde, f, fazer), start=1):
            x = ws.cell(r, c, v)
            x.font, x.border, x.alignment = FONT, BORDA, QUEBRA
        ws.cell(r, 3).fill = F_CINZA
        ws.cell(r, 3).alignment = Alignment(horizontal="center", vertical="top")
        ws.row_dimensions[r].height = 32
    tr = 4 + len(regras)
    ws.cell(tr, 1, "TOTAL DE ALERTAS").font = FONT_B
    ws.cell(tr, 3, f"=SUM(C4:C{tr-1})").font = FONT_B
    ws.cell(tr, 3).alignment = Alignment(horizontal="center")
    ws.cell(tr, 3).fill = F_CINZA
    ws.conditional_formatting.add(f"A4:D{tr-1}", FormulaRule(formula=["$C4>0"], fill=F_ALERTA))


# --- Painel -----------------------------------------------------------------
def aba_painel(ws, demo):
    config_aba(ws, 100)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for c in "BCDEFGHIJKLM":
        ws.column_dimensions[c].width = 13
    ws.merge_cells("B1:M1")
    ws["B1"].value = ("O QUE ESTE PAINEL RESPONDE: onde a empresa está em relação à NR-1 e o que precisa de atenção agora?    "
                      "O QUE VOCÊ FAZ AQUI: só escolha o setor na célula amarela. Todo o resto é calculado.")
    ws["B1"].font, ws["B1"].fill = Font(name=FONTE, size=11, bold=True, color=AZUL), F_FAIXA
    ws["B1"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 40
    ws.merge_cells("B2:M2")
    ws["B2"].value = f"Painel NR-1  |  {EMPRESA}"
    ws["B2"].font = FONT_TIT
    ws.row_dimensions[2].height = 30
    ws["B3"].value = "Setor:"
    ws["B3"].font = FONT_B
    ws.merge_cells("C3:E3")
    ws["C3"].value = "Todos"
    for a in ("C3", "D3", "E3"):
        ws[a].fill, ws[a].border = F_AMARELO, BORDA
    ws["C3"].font = FONT_B
    lista(ws, "=Listas!$I$2:$I$10", "C3")
    ws["F3"].value = "← escolha 'Todos' ou um setor"
    ws["F3"].font = FONT_CINZA

    crit = 'IF($C$3="Todos","*",$C$3)'
    dj = rng(N_DIAG, "F", D1, D2)
    la, lc, lh, li, lk = (rng(N_LEV, x, L1, L2) for x in "ACHIK")
    pc, pd, pg, pq = (rng(N_PLANO, x, P1, P2) for x in "CDGQ")
    ej, ac = rng("Treinamentos", "J", T1, T2), rng("Acompanhamento", "C", A1, A2)
    ea, ec = rng("Escuta", "A", E1, E2), rng("Escuta", "C", E1, E2)

    # dados dos gráficos (bloco no fim da aba) -------------------------------
    DR = 50
    ws.cell(DR - 1, 2, "DADOS DOS GRÁFICOS (automático, não mexer)").font = FONT_B
    ws.cell(DR, 2, "Diagnóstico").font = FONT_CINZA
    ws.cell(DR, 3, "Qtde").font = FONT_CINZA
    for i, s in enumerate(["Atendido", "Parcial", "Não atendido", "A levantar"]):
        ws.cell(DR + 1 + i, 2, s)
        ws.cell(DR + 1 + i, 3, f'=COUNTIF({dj},B{DR+1+i})')
    # setores x classes
    DB = 57
    ws.cell(DB, 2, "Setor").font = FONT_CINZA
    for j, cl in enumerate(CLASSES):
        ws.cell(DB, 3 + j, cl).font = FONT_CINZA
    for i in range(len(SETORES)):
        r = DB + 1 + i
        ws.cell(r, 2, f"=Setores!D{S1+i}")
        for j, cl in enumerate(CLASSES):
            ws.cell(r, 3 + j, f'=COUNTIFS({la},$B{r},{lk},{get_col(3+j)}${DB})')
    # ações por situação
    DC = 68
    ws.cell(DC, 2, "Ações por situação").font = FONT_CINZA
    ws.cell(DC, 3, "Qtde").font = FONT_CINZA
    for i, s in enumerate(["No prazo", "Atrasada", "Concluída", "Concluída e eficaz"]):
        ws.cell(DC + 1 + i, 2, s)
        ws.cell(DC + 1 + i, 3, f'=COUNTIFS({pc},{crit},{pq},B{DC+1+i})')
    # série mensal: conformidade x meta
    DM = 75
    ws.cell(DM, 2, "Mês").font = FONT_CINZA
    ws.cell(DM, 3, "Rótulo").font = FONT_CINZA
    ws.cell(DM, 4, "% de conformidade").font = FONT_CINZA
    ws.cell(DM, 5, "Meta").font = FONT_CINZA
    ab, aa, ad = rng("Acompanhamento", "B", A1, A2), rng("Acompanhamento", "A", A1, A2), rng("Acompanhamento", "D", A1, A2)
    for k in range(6):
        r = DM + 1 + k
        ws.cell(r, 2, f"=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),{k-6})").number_format = DATA
        ws.cell(r, 3, (f'=CHOOSE(MONTH(B{r}),"jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez")'
                       f'&"/"&RIGHT(YEAR(B{r}),2)'))
        ws.cell(r, 4, (f'=IF(COUNTIFS({ab},"% de conformidade",{aa},B{r},{ac},"<>")=0,NA(),'
                       f'SUMIFS({ac},{ab},"% de conformidade",{aa},B{r}))'))
        ws.cell(r, 5, (f'=IF(COUNTIFS({ab},"% de conformidade",{aa},B{r},{ad},"<>")=0,NA(),'
                       f'SUMIFS({ad},{ab},"% de conformidade",{aa},B{r}))'))
        ws.cell(r, 4).number_format = ws.cell(r, 5).number_format = "0%"
    for rr in range(DR, DM + 8):
        for cc in range(2, 9):
            ws.cell(rr, cc).font = FONT_CINZA
    # erros #N/D esperados (mês sem dado) ficam invisíveis
    ws.conditional_formatting.add(f"D{DM+1}:E{DM+6}", FormulaRule(formula=[f"ISNA(D{DM+1})"], font=Font(color="FFFFFF")))

    # cartões ----------------------------------------------------------------
    cartoes = [
        ("% de conformidade do Diagnóstico",
         f'=IFERROR((COUNTIF({dj},"Atendido")+0.5*COUNTIF({dj},"Parcial"))/(COUNTA({rng(N_DIAG,"A",D1,D2)})-COUNTIF({dj},"Não se aplica")),0)', "0%",
         f'=IF(ISNUMBER(D{DM+6}),"Mês anterior: "&ROUND(100*D{DM+6},0)&"%","Sem registro do mês anterior")'
         f'&IF(ISNUMBER(E{DM+6})," | Meta: "&ROUND(100*E{DM+6},0)&"%"," | Meta: a definir")'),
        ("Fatores avaliados / previstos",
         (f'=COUNTIFS({la},{crit},{lc},"Não")+COUNTIFS({la},{crit},{lc},"Sim",{lh},"<>",{li},"<>")&" de "&COUNTIF({la},{crit})'), "@",
         '=IF($C$3="Todos","Todos os setores","Setor: "&$C$3)'),
        ("Riscos Substancial ou Intolerável",
         f'=COUNTIFS({la},{crit},{lk},"Substancial")+COUNTIFS({la},{crit},{lk},"Intolerável")', "0", "Prioridade para o plano de ação"),
        ("Ações atrasadas", f'=COUNTIFS({pc},{crit},{pq},"Atrasada")', "0", "Prazo vencido e ação não concluída"),
        ("Ações sem responsável", f'=COUNTIFS({pc},{crit},{pd},"<>",{pg},"=")', "0", "Toda ação precisa de um dono"),
        ("Treinamentos vencidos ou a vencer em 60 dias", f'=COUNTIF({ej},"Vencido")+COUNTIF({ej},"Vence em 60 dias")', "0", "Reciclagem anual"),
        ("Dias desde a última escuta",
         (f'=IF($C$3="Todos",IF(COUNT({ea})=0,"sem dados",TODAY()-MAX({ea})),'
          f'IF(SUMPRODUCT(MAX(({ec}=$C$3)*{ea}))=0,"sem dados",TODAY()-SUMPRODUCT(MAX(({ec}=$C$3)*{ea}))))'), "0",
         "Sugestão: no máximo 90 dias"),
        ("Total de alertas na planilha", f"=Alertas!C{4+13}", "0", "Veja a aba Alertas"),
    ]
    pos = [(5, 2), (5, 5), (5, 8), (5, 11), (9, 2), (9, 5), (9, 8), (9, 11)]
    for (rtop, ccol), (tit, f, fmt, sub) in zip(pos, cartoes):
        for dr, alt in ((0, 20), (1, 34), (2, 18)):
            ws.merge_cells(start_row=rtop + dr, start_column=ccol, end_row=rtop + dr, end_column=ccol + 2)
            ws.row_dimensions[rtop + dr].height = alt
        for dr in range(3):
            for dc in range(3):
                x = ws.cell(rtop + dr, ccol + dc)
                x.fill = F_CINZA
                x.border = Border(left=FINO if dc == 0 else None, right=FINO if dc == 2 else None,
                                  top=FINO if dr == 0 else None, bottom=FINO if dr == 2 else None)
        a = ws.cell(rtop, ccol, tit)
        a.font, a.alignment = Font(name=FONTE, size=10, bold=True, color="404040"), Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = ws.cell(rtop + 1, ccol, f)
        v.font, v.alignment, v.number_format = Font(name=FONTE, size=22, bold=True, color=AZUL), CENTRO, fmt
        s = ws.cell(rtop + 2, ccol, sub)
        s.font, s.alignment = FONT_CINZA, Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[8].height = 8

    # gráficos ---------------------------------------------------------------
    # a) Rosca: parte de um todo (exigências por situação)
    ch = DoughnutChart()
    ch.title = "Exigências do Diagnóstico por situação"
    ch.add_data(Reference(ws, min_col=3, min_row=DR, max_row=DR + 4), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=DR + 1, max_row=DR + 4))
    cores = ["2E75B6", "9DC3E6", "C00000", "A6A6A6"]
    for i, cor in enumerate(cores):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = cor
        ch.series[0].dPt.append(pt)
    ch.holeSize = 55
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.roundedCorners = False
    ch.width, ch.height = 15, 7.4
    ws.add_chart(ch, "B13")

    # b) Barras horizontais empilhadas: nível de risco por setor
    bc = BarChart()
    bc.type = "bar"
    bc.grouping = "stacked"
    bc.overlap = 100
    bc.title = "Fatores por nível de risco, por setor"
    bc.add_data(Reference(ws, min_col=3, max_col=7, min_row=DB, max_row=DB + len(SETORES)), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=2, min_row=DB + 1, max_row=DB + len(SETORES)))
    for s, cor in zip(bc.series, ["BDD7EE", "9DC3E6", "2E75B6", "1F3864", "C00000"]):
        s.graphicalProperties.solidFill = cor
    bc.legend.position = "b"
    bc.legend.overlay = False
    bc.title.overlay = False
    bc.roundedCorners = False
    bc.x_axis.delete = False
    bc.y_axis.delete = False
    bc.x_axis.scaling.orientation = "maxMin"
    bc.width, bc.height = 15, 7.4
    ws.add_chart(bc, "H13")

    # c) Colunas: ações por situação (destaca o atraso)
    cc = BarChart()
    cc.type = "col"
    cc.title = "Ações por situação"
    cc.add_data(Reference(ws, min_col=3, min_row=DC, max_row=DC + 4), titles_from_data=True)
    cc.set_categories(Reference(ws, min_col=2, min_row=DC + 1, max_row=DC + 4))
    cc.legend = None
    cc.series[0].graphicalProperties.solidFill = "2E75B6"
    for i, cor in enumerate(["2E75B6", "C00000", "9DC3E6", "1F3864"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = cor
        cc.series[0].dPt.append(pt)
    cc.title.overlay = False
    cc.roundedCorners = False
    cc.x_axis.delete = False
    cc.y_axis.delete = False
    cc.width, cc.height = 15, 7.4
    ws.add_chart(cc, "B30")

    # d) Linha: conformidade x meta no tempo
    lc_ = LineChart()
    lc_.title = "% de conformidade x meta (por mês)"
    lc_.add_data(Reference(ws, min_col=4, max_col=5, min_row=DM, max_row=DM + 6), titles_from_data=True)
    lc_.set_categories(Reference(ws, min_col=3, min_row=DM + 1, max_row=DM + 6))
    lc_.series[0].graphicalProperties.line.solidFill = "1F3864"
    lc_.series[0].graphicalProperties.line.width = 28000
    for srs in lc_.series:
        srs.smooth = False
    lc_.series[1].graphicalProperties.line.solidFill = "C00000"
    lc_.series[1].graphicalProperties.line.dashStyle = "dash"
    lc_.y_axis.scaling.min, lc_.y_axis.scaling.max = 0, 1
    lc_.y_axis.number_format = "0%"
    lc_.legend.position = "b"
    lc_.legend.overlay = False
    lc_.title.overlay = False
    lc_.roundedCorners = False
    lc_.x_axis.delete = False
    lc_.y_axis.delete = False
    lc_.width, lc_.height = 15, 7.4
    ws.add_chart(lc_, "H30")

    # frases de leitura (fórmulas que montam o texto) -----------------------
    frases = {
        "B28": (f'=C{DR+1}&" de "&SUM(C{DR+1}:C{DR+4})&" exigências atendidas e "&C{DR+2}&" parciais. Prioridade: as "&C{DR+3}'
                f'&" marcadas como Não atendido"&IF(COUNTIFS({dj},"Não atendido",{rng(N_DIAG,"I",D1,D2)},"<"&TODAY())>0,", sendo "&'
                f'COUNTIFS({dj},"Não atendido",{rng(N_DIAG,"I",D1,D2)},"<"&TODAY())&" com prazo vencido","")&"."'),
        "H28": (f'=COUNTIF({lk},"Substancial")+COUNTIF({lk},"Intolerável")&" fatores Substancial ou Intolerável e "&COUNTIF({lk},"Moderado")'
                f'&" Moderados. Ainda a investigar: "&COUNTIF({lc},"A investigar")&" de "&COUNTA({la})&" combinações setor x fator."'),
        "B45": (f'=IF($C$3="Todos","No plano inteiro: ","No setor "&$C$3&": ")&C{DC+2}&" ação(ões) atrasada(s) de "&SUM(C{DC+1}:C{DC+4})'
                f'&" com situação definida. Concluídas e eficazes: "&C{DC+4}&"."'),
        "H45": (f'=IF(COUNT(D{DM+1}:D{DM+6})=0,"Sem dados mensais ainda. No último dia útil, cole o valor da aba Acompanhamento.",'
                f'"Último registro: "&ROUND(100*LOOKUP(2,1/ISNUMBER(D{DM+1}:D{DM+6}),D{DM+1}:D{DM+6}),0)&"% de conformidade"&'
                f'IF(COUNT(E{DM+1}:E{DM+6})=0,".",", meta de "&ROUND(100*LOOKUP(2,1/ISNUMBER(E{DM+1}:E{DM+6}),E{DM+1}:E{DM+6}),0)&"%."))'),
    }
    for cel, f in frases.items():
        col = ws[cel].column
        ws.merge_cells(start_row=ws[cel].row, start_column=col, end_row=ws[cel].row, end_column=col + 5)
        ws[cel].value = f
        ws[cel].font = Font(name=FONTE, size=10, italic=True, color="404040")
        ws[cel].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[ws[cel].row].height = 34
    ws.merge_cells("B47:M47")
    ws["B47"].value = ("Painel de empresa fictícia criada para fins de portfólio. Os gráficos usam os dados dos blocos abaixo "
                       "(linha 49 em diante).")
    ws["B47"].font = FONT_CINZA


def get_col(n):
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


# --- Como usar --------------------------------------------------------------
def aba_como_usar(ws, demo):
    config_aba(ws, 100, paisagem=False)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 6
    ws.column_dimensions["C"].width = 28
    for c in "DEFGH":
        ws.column_dimensions[c].width = 16
    ws.merge_cells("B2:H2")
    ws["B2"].value = "Planilha de gestão da NR-1 (GRO/PGR e riscos psicossociais)"
    ws["B2"].font = FONT_TIT
    ws.merge_cells("B3:H3")
    ws["B3"].value = (EMPRESA + ". " +
                      ("VERSÃO DEMONSTRATIVA: todos os dados são inventados e as datas foram pensadas para novembro de 2026; "
                       "as situações calculadas mudam com o passar do tempo."
                       if demo else "MODELO EM BRANCO: mesmos setores de exemplo, sem registros. Substitua pelos da sua empresa."))
    ws["B3"].font, ws["B3"].alignment = FONT_B, Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 36

    def bloco(r, titulo):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        c = ws.cell(r, 2, titulo)
        c.font, c.fill = FONT_HDR, F_HDR

    bloco(5, "PARA QUE SERVE")
    ws.merge_cells("B6:H7")
    ws["B6"].value = ("Esta planilha é uma ferramenta de gestão do RH. Ela mostra o que falta para a empresa cumprir a NR-1, "
                      "quais ações estão atrasadas e quem é o responsável por cada uma. O inventário oficial de riscos "
                      "continua sendo o PGR, que deve ser revisado com o resultado deste trabalho.")
    ws["B6"].alignment = QUEBRA
    ws["B6"].font = FONT
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 30
    bloco(9, "PASSO A PASSO")
    passos = [
        "Abra a aba Diagnóstico NR-1 e escolha a situação de cada exigência (Atendido, Parcial, Não atendido...).",
        "Na aba Setores, confira turnos, número de pessoas e líder de cada setor.",
        "Na aba Escuta, registre cada conversa com os trabalhadores (só contagens e temas, nunca nomes).",
        "Na aba Levantamento, para cada setor e fator, responda se o fator existe e dê notas de 1 a 5.",
        "Na aba Plano de ação, crie uma ação para cada risco que precisa de medida, sempre com responsável e prazo.",
        "Olhe o Painel e a aba Alertas. Corrija o que estiver laranja. No último dia útil do mês, registre os números em Acompanhamento.",
    ]
    for i, p in enumerate(passos):
        r = 10 + i
        ws.cell(r, 2, i + 1).alignment = Alignment(horizontal="center", vertical="top")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.cell(r, 3, p).alignment = QUEBRA
        ws.row_dimensions[r].height = 32
        ws.cell(r, 2).font = FONT_B
        ws.cell(r, 3).font = FONT
    bloco(17, "LEGENDA DE CORES")
    leg = [(F_AMARELO, "Amarelo: preencher (digite ou escolha na lista)."),
           (F_CINZA, "Cinza: cálculo automático ou texto fixo. Não mexer."),
           (F_ALERTA, "Laranja: a linha tem um problema; veja a coluna Alerta."),
           (F_HDR, "Azul-escuro: títulos das colunas.")]
    for i, (fill, t) in enumerate(leg):
        r = 18 + i
        ws.cell(r, 2).fill = fill
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        ws.cell(r, 3, t).font = FONT
    bloco(23, "ROTINA SUGERIDA")
    ws.merge_cells("B24:H25")
    ws["B24"].value = ("Toda semana (15 min): revise o Plano de ação. Todo mês (30 min): atualize Escuta e Treinamentos, "
                       "olhe o Painel e registre os indicadores em Acompanhamento.")
    ws["B24"].alignment, ws["B24"].font = QUEBRA, FONT
    ws.row_dimensions[24].height = 30
    ws.row_dimensions[25].height = 20
    bloco(27, "AVISOS")
    ws.merge_cells("B28:H29")
    ws["B28"].value = ("Sem macros e sem senha. Nenhuma aba, coluna ou fórmula está escondida. Não registre CID, diagnóstico ou "
                       "nome de trabalhador nos campos de texto. Dados pessoais: use o mínimo necessário.")
    ws["B28"].alignment, ws["B28"].font = QUEBRA, FONT
    ws.row_dimensions[28].height = 30
    ws.row_dimensions[29].height = 20


# --- Premissas e Log --------------------------------------------------------
def aba_premissas(ws):
    config_aba(ws, 100, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: quais regras e suposições foram adotadas na construção desta planilha?    "
              "O QUE VOCÊ FAZ AQUI: leia antes de usar. Regra provisória deve ser confirmada pela consultoria de segurança.", 4)
    cabecalho(ws, 3, ["Nº", "Premissa adotada", "Data", "Quem decidiu"], [5, 110, 12, 24], altura=28)
    itens = [
        "Empresa fictícia: todos os setores, números e registros são inventados. Nenhum dado de empresa real foi usado.",
        "Matriz de risco 5x5 provisória (aba Listas): produto 1-2 Trivial; 3-4 Tolerável; 5-9 Moderado; 10-14 Substancial; 15-25 Intolerável. Confirmar com a consultoria antes de uso real.",
        "% de conformidade = (Atendido + 0,5 x Parcial) / (total de exigências - Não se aplica).",
        "Indicadores manuais (relatos, afastamentos, rotatividade, absenteísmo) só entram nos gráficos com 'Dado validado = Sim'.",
        "% de pessoas treinadas = presentes na capacitação anual sobre assédio (se não vencida) / total de pessoas informado em Setores. É estimativa simples.",
        "Metas do Acompanhamento são ESTIMATIVAS iniciais, para validar com a diretoria.",
        "Tabelas com linhas pré-formatadas (200 em ações e escutas, 50 em treinamentos) em vez de Tabelas do Excel, para manter compatibilidade e simplicidade.",
        "Funções usadas existem do Excel 2016 em diante. Não há FILTRO, PROCX, ÚNICO nem LET.",
        "Setores e fatores do Levantamento são linhas fixas. Para incluir um setor novo, gere a planilha de novo com o script.",
        "O Painel mostra ações por setor escolhido. Ações marcadas 'Toda a empresa' só aparecem com 'Todos'.",
        "As referências legais citam o nome da norma. Confira o item exato no texto vigente antes de citar em documento oficial.",
        "Os valores do dia (aba Acompanhamento) mudam todo dia; por isso o registro mensal é feito copiando como VALOR, sem macro.",
        "Os campos de texto livre não devem conter nomes de trabalhadores, CID ou diagnóstico.",
        "O valor '#N/D' nos dados dos gráficos (Painel, linhas 76 a 81) é intencional: faz o gráfico de linha pular meses sem registro. Fica invisível.",
    ]
    for i, t in enumerate(itens):
        r = 4 + i
        for c, v in enumerate((i + 1, t, date(2026, 11, 16), "Autor do projeto"), start=1):
            x = ws.cell(r, c, v)
            x.font, x.border, x.alignment = FONT, BORDA, QUEBRA
        ws.cell(r, 3).number_format = DATA
        ws.row_dimensions[r].height = 34


def aba_log(ws, demo):
    config_aba(ws, 100, "A4")
    faixa(ws, "O QUE ESTA ABA RESPONDE: o que foi feito na construção desta planilha?    "
              "O QUE VOCÊ FAZ AQUI: nada. É o registro das transformações feitas pelo autor.", 4)
    cabecalho(ws, 3, ["Nº", "Data", "Etapa", "O que foi feito"], [5, 12, 26, 100], altura=28)
    itens = [
        ("Criação das abas", "Criadas 13 abas na ordem: Como usar, Painel, Diagnóstico NR-1, Setores, Levantamento, Plano de ação, Treinamentos, Escuta, Acompanhamento, Alertas, Listas, Premissas, Log."),
        ("Dados de base", "Inventados 8 setores fictícios e 13 fatores psicossociais (lista exemplificativa). 24 exigências de conformidade escritas em linguagem simples."),
        ("Levantamento", "Pré-geradas 104 linhas (8 setores x 13 fatores), todas iniciando em 'A investigar'."),
        ("Cálculos", "Nível = Probabilidade x Severidade; classificação por busca na matriz 5x5 (ÍNDICE); situação das ações por comparação do prazo com HOJE()."),
        ("Alertas", "Cada aba de entrada ganhou coluna Alerta e realce laranja; a aba Alertas conta os problemas por regra."),
        ("Gráficos", "Rosca (parte de um todo), barras empilhadas (comparar setores), colunas (ações por situação) e linha (conformidade x meta no tempo)."),
        ("Dados de exemplo" if demo else "Modelo em branco",
         "Preenchidos registros fictícios cobrindo cada regra de alerta, para demonstração." if demo else "Nenhum registro de exemplo além das 5 ações iniciais, sem responsável nem prazo."),
        ("Conferência", "Recalculada no LibreOffice; verificados erros de fórmula e contagens por regra de alerta."),
    ]
    for i, (e, t) in enumerate(itens):
        r = 4 + i
        for c, v in enumerate((i + 1, date(2026, 11, 16), e, t), start=1):
            x = ws.cell(r, c, v)
            x.font, x.border, x.alignment = FONT, BORDA, QUEBRA
        ws.cell(r, 2).number_format = DATA
        ws.row_dimensions[r].height = 32


if __name__ == "__main__":
    import os
    saida = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "planilha")
    os.makedirs(saida, exist_ok=True)
    if "--prints" in sys.argv:
        construir(True, prints=True).save(os.path.join(saida, "_prints_temp.xlsx"))
    else:
        construir(True).save(os.path.join(saida, "NR1_Industria_Exemplo_DEMO.xlsx"))
        construir(False).save(os.path.join(saida, "NR1_Modelo_Vazio.xlsx"))
    print("ok")
