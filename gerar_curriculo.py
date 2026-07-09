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
    bottom.set(qn("w:sz"), "6")
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
titulo_secao(doc, "Resumo Profissional", before=8)
texto(doc,
      "Profissional com dupla formação técnica pelo SENAI — Eletromecânica (1.500h) e "
      "Administração (1.200h) — em transição estratégica para Tecnologia da Informação, "
      "aprovado no Tecnólogo em Análise e Desenvolvimento de Sistemas (Facens). Na CBA, atuei "
      "no Suporte de Operações da Sala de Fornos 7 com governança documental (Docnix), "
      "monitoramento de indicadores em Power BI, suporte de TI (ServiceNow) e melhoria "
      "contínua — e desenvolvi, por iniciativa própria, um sistema web de gestão de "
      "ocorrências integrado a Power Automate e Excel Online, colocado em produção no setor. "
      "Antes, fui promovido de Operador de Armazém Júnior a Pleno em menos de 12 meses. "
      "Perfil analítico e autodidata, que transforma controles manuais em soluções digitais.",
      after=3)

# ── Experiência profissional ─────────────────────────────────────────────────
titulo_secao(doc, "Experiência Profissional")

emprego(
    doc,
    "Companhia Brasileira de Alumínio (CBA)", "Alumínio/SP",
    "Aprendiz – Suporte de Operações (Sala de Fornos 7 – ADM)", "jan/2025 – jun/2026",
    [
        "Governança documental no sistema Docnix: pesquisa avançada, verificação de vigência "
        "e distribuição de Procedimentos Operacionais, Padrões de Trabalho e FMEAs, com apoio "
        "à integração de novos funcionários.",
        "Key User local de TI: chamados no ServiceNow, recuperação de computadores "
        "bloqueados por BitLocker e suporte a impressoras e acessos corporativos.",
        "Monitoramento diário de indicadores de segurança (DDS e abrangências) em Power BI e "
        "ObraSoft/Power Apps, com geração de relatórios e baixa de pendências.",
        "Inventário completo do almoxarifado do setor, com planilhas de controle de entrada "
        "e saída em Excel, consultas de saldo no SAP e gestão da distribuição de EPIs.",
        "Melhoria de processo por iniciativa própria: substituição de anotações em lousa por "
        "cadernos de checklist padronizados, criando histórico auditável para auditorias.",
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
        "Operação de empilhadeiras (NR-11) e separação de cargas em câmaras frias a -25 °C "
        "com coletor de dados e WMS Blue Yonder (picking por LPN e conferência de estoque), "
        "além de expedição e carregamento de veículos em turno noturno de alto giro.",
        "Capacitação contínua em paralelo ao turno: PCP (40h), Metrologia (60h), Excel "
        "Avançado (40h), Power BI (32h) e Python (30h).",
    ],
)

emprego(
    doc,
    "Supermercado São Roque – Centro de Distribuição", "São Roque/SP",
    "Ajudante Operacional", "ago/2021 – jul/2022",
    [
        "Carregamento de caminhões em docas, paletização com filme stretch, organização do "
        "centro de distribuição e controle de devolução de ativos (caixas) a fornecedores.",
    ],
)

# ── Projetos de tecnologia ───────────────────────────────────────────────────
titulo_secao(doc, "Projetos de Tecnologia")

texto(doc, [("CBA-EXAUST – Sistema de Gestão de Ocorrências de Fornos (autoral)",
             True, False)], before=2, after=1)
bullet(doc, "Aplicação web (PWA) criada por iniciativa própria e colocada em produção para "
            "registrar e analisar ocorrências de exaustão dos fornos da Sala de Fornos 7: "
            "dashboards com Pareto, ranking de fornos críticos, alertas, modo TV e resumo "
            "executivo.")
bullet(doc, "Integração com Power Automate e Excel Online/OneDrive, exportação CSV/Excel e "
            "relatório mensal automático por e-mail; publicação via GitHub e Cloudflare.",
       after=4)

texto(doc, [("Sistema de Gestão de Estoque LIS (projeto SENAI)", True, False)], after=1)
bullet(doc, "Site em HTML/JavaScript publicado no GitHub Pages e integrado ao Excel Online "
            "via Power Automate: lançamentos de entrada e saída em tempo real, curva ABC, "
            "estoque de segurança, ponto de pedido e dashboards automáticos.", after=3)

# ── Formação acadêmica ───────────────────────────────────────────────────────
titulo_secao(doc, "Formação Acadêmica")

texto(doc, [("Tecnólogo em Análise e Desenvolvimento de Sistemas (EAD)", True, False),
            (" — Facens, Sorocaba/SP. Aprovado; início em agosto/2026.", False, False)],
      before=2, after=2)
texto(doc, [("Técnico em Administração", True, False),
            (" — SENAI (1.200h). Diploma em junho/2026 — Aprendiz CBA.", False, False)],
      after=2)
texto(doc, [("Técnico em Eletromecânica", True, False),
            (" — SENAI Mairinque (1.500h). Diploma em junho/2025.", False, False)], after=3)

# ── Cursos e certificações ───────────────────────────────────────────────────
titulo_secao(doc, "Cursos e Certificações")

bullet(doc, "Excel Avançado (40h), Power BI (32h) e Fundamentos do Python (30h) – "
            "SENAI (2024)")
bullet(doc, "Lógica de Programação (14h), Segurança Cibernética (4h) e Desvendando a "
            "Indústria 4.0 (20h) – SENAI (2025)")
bullet(doc, "Fundamentos da Inteligência Artificial (8h) e Ética na Inteligência "
            "Artificial (4h) – SENAI (2025)")
bullet(doc, "Caldeiraria Prática – AHCX Treinamentos / MJS Brasil (jan–jul/2026, em "
            "conclusão)")
bullet(doc, "NR-11 – Operação de Empilhadeira – SENAI (32h, 2023); reciclagens NR-11 – "
            "SuperFrio (2024)")
bullet(doc, "Metrologia Aplicada à Mecânica (60h) e PCP (40h) – SENAI (2024)")
bullet(doc, "Administração – Nova NETT (96h, 2024); Segurança no Trabalho – SENAI "
            "(14h, 2023)", after=3)

# ── Competências ─────────────────────────────────────────────────────────────
titulo_secao(doc, "Competências")

bullet(doc, [("Dados e BI: ", True, False),
             ("Excel Avançado (dashboards, tabelas dinâmicas, KPIs), Power BI e "
              "Microsoft Forms", False, False)])
bullet(doc, [("Automação e desenvolvimento: ", True, False),
             ("Power Automate, Power Apps, HTML/JavaScript, GitHub, fundamentos de Python "
              "e lógica de programação", False, False)])
bullet(doc, [("Sistemas corporativos: ", True, False),
             ("SAP (consultas), ServiceNow, Docnix e WMS Blue Yonder", False, False)])
bullet(doc, [("Melhoria contínua: ", True, False),
             ("5S, PDCA, Matriz GUT, gestão visual e noções de Scrum", False, False)])
bullet(doc, [("Logística e indústria: ", True, False),
             ("operação de empilhadeira (NR-11), inventários, PCP e metrologia",
              False, False)])
bullet(doc, [("Comportamentais: ", True, False),
             ("disciplina, autodidatismo, resiliência e foco em resolução de problemas",
              False, False)], after=3)

# ── Informações adicionais ───────────────────────────────────────────────────
titulo_secao(doc, "Informações Adicionais")

bullet(doc, "CNH categorias A e B (EAR); disponibilidade para início imediato e para "
            "trabalho em turnos", after=0)

doc.save("Curriculo_Jeferson_dos_Santos_Cardoso.docx")
print("OK: Curriculo_Jeferson_dos_Santos_Cardoso.docx gerado.")
