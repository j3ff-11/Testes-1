#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador do currículo de Jeferson dos Santos Cardoso.

Especificação visual (padrão 10/10):
- Fonte: Arial em todo o documento
- Nome: 24 pt negrito | Títulos de seção: 14 pt negrito MAIÚSCULAS | Corpo: 11 pt
- Espaçamento entre linhas: 1,15
- Margens: 2,0 cm nos quatro lados
- Alinhamento à esquerda (nunca justificado)
- Cores: fundo #FFFFFF, texto #333333, títulos #1B365D
- Linha horizontal fina (#1B365D) abaixo de cada título de seção
- Sem ícones, sem gráficos, sem colunas, sem tabelas, sem caixas de texto
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEXTO = RGBColor(0x33, 0x33, 0x33)   # cinza-grafite
TITULO = RGBColor(0x1B, 0x36, 0x5D)  # azul-marinho corporativo
FONTE = "Arial"


def formatar_run(run, size, bold=False, italic=False, color=TEXTO):
    run.font.name = FONTE
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    # Garante Arial também para o mapeamento east-asian (compatibilidade Word)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONTE)
    rFonts.set(qn("w:hAnsi"), FONTE)
    rFonts.set(qn("w:cs"), FONTE)


def paragrafo(doc, before=0, after=4, indent=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent is not None:
        pf.left_indent = Cm(indent)
    return p


def texto(doc, partes, before=0, after=4, indent=None, size=11):
    """partes: str ou lista de tuplas (texto, negrito, itálico)."""
    p = paragrafo(doc, before=before, after=after, indent=indent)
    if isinstance(partes, str):
        partes = [(partes, False, False)]
    for t, b, i in partes:
        formatar_run(p.add_run(t), size, bold=b, italic=i)
    return p


def bullet(doc, conteudo, after=2):
    """Marcador simples '•' com recuo deslocado (sem tabelas/estilos de lista)."""
    p = paragrafo(doc, before=0, after=after)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.5)
    pf.first_line_indent = Cm(-0.5)
    if isinstance(conteudo, str):
        conteudo = [(conteudo, False, False)]
    formatar_run(p.add_run("•\t"), 11)
    for t, b, i in conteudo:
        formatar_run(p.add_run(t), 11, bold=b, italic=i)
    return p


def titulo_secao(doc, nome, before=9):
    """Título 14 pt negrito maiúsculas #1B365D com linha fina inferior #1B365D."""
    p = paragrafo(doc, before=before, after=5)
    formatar_run(p.add_run(nome.upper()), 14, bold=True, color=TITULO)
    pPr = p._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")       # linha fina (~0,75 pt)
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), "1B365D")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def emprego(doc, empresa, local, cargo, periodo, bullets, before=6):
    texto(doc, [(empresa, True, False), (f" — {local}", False, False)],
          before=before, after=1)
    texto(doc, [(cargo, True, False), (f"  |  {periodo}", False, True)], after=2)
    for b in bullets:
        bullet(doc, b)


doc = Document()

# Margens 2,0 cm e página em branco puro
for section in doc.sections:
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

# ── Cabeçalho ────────────────────────────────────────────────────────────────
p = paragrafo(doc, before=0, after=2)
formatar_run(p.add_run("JEFERSON DOS SANTOS CARDOSO"), 24, bold=True, color=TITULO)

texto(doc, "Mairinque – SP  |  (11) 95640-3941  |  jeferson20sc@gmail.com", after=1)
texto(doc, "linkedin.com/in/jeferson-santos35  |  CNH: A/B (EAR)", after=2)

# ── Resumo profissional ──────────────────────────────────────────────────────
titulo_secao(doc, "Resumo Profissional", before=10)
texto(doc,
      "Profissional com dupla formação técnica pelo SENAI — Eletromecânica (1.500h) e "
      "Administração (1.200h) — e experiência prática em logística e operações industriais, "
      "em transição estratégica para a área de Tecnologia da Informação. Histórico comprovado "
      "de crescimento acelerado: promovido de Operador de Armazém Júnior a Pleno em menos de "
      "12 meses. Aprovado no curso de Tecnólogo em Análise e Desenvolvimento de Sistemas "
      "(Facens), com domínio de Excel Avançado e Power BI e fundamentos de Python e lógica de "
      "programação. Perfil analítico e disciplinado, que une a visão operacional do chão de "
      "fábrica à capacidade de transformar dados em soluções.",
      after=3)

# ── Experiência profissional ─────────────────────────────────────────────────
titulo_secao(doc, "Experiência Profissional")

emprego(
    doc,
    "Companhia Brasileira de Alumínio (CBA)", "Alumínio/SP",
    "Aprendiz – Administração Industrial", "jan/2025 – jun/2026",
    [
        "Selecionado para o programa Jovem Aprendiz da CBA, uma das maiores produtoras de "
        "alumínio do país, com jornada dupla: prática administrativa na empresa pela manhã "
        "e formação Técnica em Administração no SENAI à tarde, de segunda a sexta.",
        "Apoio às rotinas administrativas do setor: organização e controle de documentos, "
        "atualização de planilhas e elaboração de relatórios em Excel para suporte à gestão.",
        "Conclusão simultânea do contrato de aprendizagem e do diploma técnico (1.200h).",
    ],
    before=2,
)

emprego(
    doc,
    "CEFRI / SuperFrio – Logística e Armazenagem Frigorificada", "Mairinque/SP",
    "Operador de Armazém Júnior → Pleno", "abr/2023 – jan/2025",
    [
        "Promovido de Operador Júnior a Pleno em menos de 12 meses, em reconhecimento ao "
        "desempenho, à confiabilidade e à produtividade na operação.",
        "Operação certificada de empilhadeira (NR-11, com reciclagens periódicas) na "
        "movimentação de cargas em ambiente frigorificado de alto giro.",
        "Atuação em recebimento, conferência, endereçamento, separação (picking) e expedição "
        "de mercadorias, com foco em acuracidade de estoque e cumprimento rigoroso de prazos.",
        "Capacitação contínua em paralelo ao trabalho em turno: PCP (40h), Metrologia (60h), "
        "Excel Avançado (40h), Power BI (32h) e Python (30h).",
    ],
)

emprego(
    doc,
    "Supermercado São Roque", "São Roque/SP",
    "Ajudante Operacional", "ago/2021 – jul/2022",
    [
        "Recebimento, conferência, organização de estoque e reposição de mercadorias, "
        "garantindo o abastecimento contínuo da área de vendas, com atendimento ao cliente "
        "e trabalho em equipe em operação de ritmo intenso.",
    ],
)

emprego(
    doc,
    "AMP GRU Sustentável", "São Roque/SP",
    "Jardineiro", "out/2020 – jan/2021",
    [
        "Primeira experiência com carteira assinada, aos 18 anos, em conservação de áreas "
        "verdes — base da ética de trabalho e do senso de responsabilidade.",
    ],
)

# ── Formação acadêmica ───────────────────────────────────────────────────────
titulo_secao(doc, "Formação Acadêmica")

texto(doc, [("Tecnólogo em Análise e Desenvolvimento de Sistemas (EAD)", True, False),
            (" — Facens, Sorocaba/SP", False, False)], before=2, after=1)
texto(doc, "Aprovado em processo seletivo; início previsto para agosto/2026.", after=6)

texto(doc, [("Técnico em Administração", True, False),
            (" — SENAI “Antônio Ermírio de Moraes” (1.200h)", False, False)], after=1)
texto(doc, "Concluído em junho/2026, com diploma — cursado pelo programa Aprendiz CBA.", after=6)

texto(doc, [("Técnico em Eletromecânica", True, False),
            (" — SENAI Centro de Treinamento Mairinque (1.500h)", False, False)], after=1)
texto(doc, "Concluído em junho/2025, com diploma.", after=4)

# ── Cursos e certificações ───────────────────────────────────────────────────
titulo_secao(doc, "Cursos e Certificações")

texto(doc, [("Tecnologia e Dados", True, False)], before=2, after=3)
bullet(doc, "Excel Avançado – SENAI (40h, 2024) e Excel Básico – SENAI (20h, 2025)")
bullet(doc, "Microsoft Power BI – SENAI (32h, 2024)")
bullet(doc, "Fundamentos do Python 1 (30h, 2024) e Lógica de Programação (14h, 2025) – SENAI")
bullet(doc, "Por Dentro da Segurança Cibernética (4h) e Desvendando a Indústria 4.0 (20h) – "
            "SENAI (2025)")
bullet(doc, "Fundamentos da Inteligência Artificial (8h) e Ética na Inteligência "
            "Artificial (4h) – SENAI (2025)", after=5)

texto(doc, [("Industrial e Logística", True, False)], after=3)
bullet(doc, "Caldeiraria Prática – Construtor de Equipamentos Industriais – AHCX "
            "Treinamentos / MJS Brasil (jan–jul/2026, em fase final de conclusão)")
bullet(doc, "NR-11 – Operação de Empilhadeira – SENAI (32h, 2023); reciclagens NR-11 – "
            "SuperFrio (2024)")
bullet(doc, "Metrologia Aplicada à Mecânica – SENAI (60h, 2024)")
bullet(doc, "Planejamento e Controle da Produção – SENAI (40h, 2024)")
bullet(doc, "Segurança no Trabalho (14h, 2023) e Consumo Consciente de Energia (14h, 2025) "
            "– SENAI", after=5)

texto(doc, [("Gestão", True, False)], after=3)
bullet(doc, "Administração – Nova NETT (96h, 2024)", after=4)

# ── Competências ─────────────────────────────────────────────────────────────
titulo_secao(doc, "Competências")

bullet(doc, [("Análise de dados: ", True, False),
             ("Excel Avançado (dashboards, fórmulas, relatórios gerenciais) e Power BI",
              False, False)])
bullet(doc, [("Programação: ", True, False),
             ("fundamentos de Python e lógica de programação", False, False)])
bullet(doc, [("Logística de armazém: ", True, False),
             ("recebimento, endereçamento, picking, expedição, inventário e operação de "
              "empilhadeira (NR-11)", False, False)])
bullet(doc, [("Indústria: ", True, False),
             ("eletromecânica, metrologia, caldeiraria, planejamento e controle da produção "
              "(PCP) e segurança do trabalho", False, False)])
bullet(doc, [("Gestão: ", True, False),
             ("rotinas administrativas, organização de documentos e apoio a relatórios "
              "gerenciais", False, False)])
bullet(doc, [("Comportamentais: ", True, False),
             ("disciplina, resiliência, aprendizado rápido e foco em resolução de problemas",
              False, False)], after=4)

# ── Informações adicionais ───────────────────────────────────────────────────
titulo_secao(doc, "Informações Adicionais")

bullet(doc, "CNH categorias A e B, com observação EAR (Exerce Atividade Remunerada)")
bullet(doc, "Disponibilidade para início imediato e para trabalho em turnos", after=0)

doc.save("Curriculo_Jeferson_dos_Santos_Cardoso.docx")
print("OK: Curriculo_Jeferson_dos_Santos_Cardoso.docx gerado.")
