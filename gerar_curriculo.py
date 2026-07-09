#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador do currículo de Jeferson dos Santos Cardoso — versão final do usuário.

Especificação visual (padrão 10/10):
- Fonte: Arial em todo o documento
- Nome: 24 pt negrito | Títulos de seção: 14 pt negrito MAIÚSCULAS | Corpo: 11 pt
- Espaçamento entre linhas: 1,15 | Margens: 2,0 cm | Alinhado à esquerda
- Cores: fundo #FFFFFF, texto #333333, títulos #1B365D
- Linha horizontal fina (#1B365D) abaixo de cada título de seção
- Sem ícones, gráficos, colunas, tabelas ou caixas de texto
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TEXTO = RGBColor(0x33, 0x33, 0x33)
TITULO = RGBColor(0x1B, 0x36, 0x5D)
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
    p = paragrafo(doc, before=before, after=after, indent=indent)
    if isinstance(partes, str):
        partes = [(partes, False, False)]
    for t, b, i in partes:
        formatar_run(p.add_run(t), size, bold=b, italic=i)
    return p


def bullet(doc, conteudo, after=1):
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


def titulo_secao(doc, nome, before=7):
    p = paragrafo(doc, before=before, after=3)
    p.paragraph_format.keep_with_next = True
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
texto(doc, "linkedin.com/in/jeferson-santos35  |  github.com/j3ff-11  |  CNH: A/B (EAR)",
      after=1)
texto(doc, [("Suporte de Operações | Administração | BI | Tecnologia | Logística",
             True, False)], after=2)

# ── Resumo profissional ──────────────────────────────────────────────────────
titulo_secao(doc, "Resumo Profissional", before=8)
texto(doc,
      "Profissional técnico em transição para Tecnologia da Informação, com formação pelo "
      "SENAI em Eletromecânica e Administração, aprovado no Tecnólogo em Análise e "
      "Desenvolvimento de Sistemas pela Facens. Experiência em ambiente industrial, "
      "logística refrigerada, suporte operacional, governança documental, indicadores, "
      "inventários, sistemas corporativos e melhoria de processos.",
      after=3)
texto(doc,
      "Atuação na CBA com Docnix, ServiceNow, Power BI, Power Apps/ObraSoft, SAP, Excel e "
      "suporte técnico local, contribuindo para organização documental, controle de "
      "indicadores, gestão de materiais e digitalização de rotinas operacionais. Desenvolvi "
      "e testei em piloto um sistema web/PWA para gestão de ocorrências de fornos "
      "industriais, integrando Power Automate, Excel Online, GitHub e Cloudflare.",
      after=3)
texto(doc,
      "Perfil disciplinado, hands-on e autodidata, com facilidade para aprender "
      "ferramentas, resolver problemas práticos e transformar controles manuais em "
      "soluções mais organizadas, rastreáveis e eficientes.",
      after=3)

# ── Experiência profissional ─────────────────────────────────────────────────
titulo_secao(doc, "Experiência Profissional")

texto(doc, [("COMPANHIA BRASILEIRA DE ALUMÍNIO — CBA", True, False),
            ("  |  Alumínio/SP", False, False)], before=2, after=1)
texto(doc, [("Aprendiz – Suporte de Operações / Administrativo Técnico — Sala de Fornos 7",
             True, False)], after=1)
texto(doc, [("jan/2025 – jun/2026", False, True)], after=2)
bullet(doc, "Atuação em suporte administrativo e operacional na Sala de Fornos 7, apoiando "
            "rotinas de produção, segurança, manutenção, documentação, indicadores e "
            "controle de materiais.")
bullet(doc, "Gestão e governança documental no Docnix, com pesquisa avançada, verificação "
            "de vigência, organização e distribuição de Procedimentos Operacionais, "
            "Padrões de Trabalho e FMEAs.")
bullet(doc, "Atuação como apoio local de TI/Key User, com abertura e acompanhamento de "
            "chamados no ServiceNow, suporte a acessos, impressoras, equipamentos "
            "bloqueados por BitLocker e demandas técnicas do setor.")
bullet(doc, "Monitoramento diário de indicadores de segurança e operação, incluindo DDS, "
            "abrangências, pendências e registros em Power BI, Power Apps e ObraSoft.")
bullet(doc, "Elaboração e atualização de planilhas de controle em Excel, apoiando "
            "inventários, entrada e saída de materiais, distribuição de EPIs e consultas "
            "de saldo no SAP.")
bullet(doc, "Organização do almoxarifado do setor com aplicação de boas práticas de "
            "controle, 5S, rastreabilidade de materiais e padronização de registros.")
bullet(doc, "Identificação de falhas em controles manuais e criação de soluções digitais "
            "para aumentar rastreabilidade, reduzir perda de informação e melhorar a "
            "tomada de decisão.")
bullet(doc, "Desenvolvimento de checklists e controles digitais que evoluíram para um "
            "sistema web de gestão de ocorrências industriais, testado em piloto na área.",
       after=3)

texto(doc, [("CEFRI / SUPERFRIO — Logística e Armazenagem Frigorificada", True, False),
            ("  |  Mairinque/SP", False, False)], after=1)
texto(doc, [("Operador de Armazém Júnior → Pleno", True, False)], after=1)
texto(doc, [("abr/2023 – jan/2025", False, True)], after=2)
bullet(doc, "Promovido de Operador Júnior a Pleno em menos de 12 meses, em reconhecimento "
            "ao desempenho, confiabilidade, produtividade e adaptação à rotina operacional.")
bullet(doc, "Operação de empilhadeiras, conforme NR-11, em ambiente refrigerado de alta "
            "demanda, com movimentação, separação, organização e expedição de cargas.")
bullet(doc, "Atuação em câmaras frias de até -25 °C, realizando picking, conferência, "
            "armazenagem, carregamento e suporte à operação logística em turno noturno.")
bullet(doc, "Utilização de coletor de dados e sistema WMS Blue Yonder para controle de "
            "movimentações, separação de pedidos, localização de produtos e apoio ao "
            "fluxo de estoque.")
bullet(doc, "Apoio à organização do armazém, controle de produtividade, segurança "
            "operacional e cumprimento de procedimentos internos.")
bullet(doc, "Capacitação contínua em paralelo ao trabalho, com cursos em Excel Avançado, "
            "Power BI, Fundamentos de Python, PCP e Metrologia.", after=3)

texto(doc, [("SUPERMERCADO SÃO ROQUE — Centro de Distribuição", True, False),
            ("  |  São Roque/SP", False, False)], after=1)
texto(doc, [("Ajudante Operacional", True, False)], after=1)
texto(doc, [("ago/2021 – jul/2022", False, True)], after=2)
bullet(doc, "Apoio às rotinas de centro de distribuição, incluindo carregamento de "
            "caminhões, paletização, organização de estoque, movimentação de mercadorias "
            "e separação de cargas.")
bullet(doc, "Controle de devolução de ativos, como caixas e materiais retornáveis, "
            "auxiliando na organização e fluxo operacional do setor.", after=3)

# ── Projetos de tecnologia ───────────────────────────────────────────────────
titulo_secao(doc, "Projetos de Tecnologia")

texto(doc, [("Sistema de Gestão de Ocorrências de Fornos Industriais — Projeto Autoral",
             True, False)], before=2, after=2)
bullet(doc, "Identificação de uma dor real da operação: controles manuais em lousa, baixa "
            "rastreabilidade, ausência de histórico estruturado e dificuldade para análise "
            "de recorrências.")
bullet(doc, "Desenvolvimento de uma aplicação web PWA para registrar ocorrências, "
            "acompanhar histórico, visualizar fornos críticos, gerar rankings, aplicar "
            "análise de Pareto e apoiar decisões operacionais.")
bullet(doc, "Integração com Power Automate e Excel Online, permitindo armazenamento dos "
            "registros, automações e apoio à geração de relatórios.")
bullet(doc, "Deploy e testes utilizando GitHub e Cloudflare, com foco em acesso simples, "
            "visualização prática e uso em ambiente operacional.")
bullet(doc, "Uso de desenvolvimento assistido por IA, com apoio de ferramentas como "
            "Claude, Copilot e ChatGPT para estruturar telas, revisar lógica, melhorar "
            "código e acelerar prototipagem.")
bullet(doc, [("Resultado: ", True, False),
             ("substituição de controles soltos por um modelo digital mais rastreável, "
              "visual e organizado, facilitando a priorização de problemas e o "
              "acompanhamento de ocorrências.", False, False)], after=3)

texto(doc, [("Sistema de Gestão de Estoque LIS — Projeto Acadêmico SENAI", True, False)],
      after=2)
bullet(doc, "Desenvolvimento de site em HTML, JavaScript e GitHub Pages, integrado ao "
            "Excel Online via Power Automate.")
bullet(doc, "Criação de controles de entrada e saída, curva ABC, estoque de segurança, "
            "ponto de pedido e dashboards automáticos.")
bullet(doc, "Projeto voltado à melhoria da rastreabilidade, redução de controles manuais "
            "e apoio à tomada de decisão em rotinas de estoque.", after=3)

# ── Formação acadêmica ───────────────────────────────────────────────────────
titulo_secao(doc, "Formação Acadêmica")

texto(doc, [("Tecnólogo em Análise e Desenvolvimento de Sistemas — Facens", True, False),
            ("  |  Sorocaba/SP", False, False)], before=2, after=1)
texto(doc, "Aprovado — início em agosto/2026", after=3)

texto(doc, [("Técnico em Administração — SENAI", True, False)], after=1)
texto(doc, "Carga horária: 1.200h  |  Conclusão: junho/2026", after=3)

texto(doc, [("Técnico em Eletromecânica — SENAI Mairinque", True, False)], after=1)
texto(doc, "Carga horária: 1.500h  |  Conclusão: junho/2025", after=3)

# ── Cursos e certificações ───────────────────────────────────────────────────
titulo_secao(doc, "Cursos e Certificações")

bullet(doc, "Excel Avançado — SENAI | 40h")
bullet(doc, "Power BI — SENAI | 32h")
bullet(doc, "Fundamentos do Python — SENAI | 30h")
bullet(doc, "PCP — Planejamento e Controle da Produção — SENAI | 40h")
bullet(doc, "Metrologia Aplicada à Mecânica — SENAI | 60h")
bullet(doc, "Lógica de Programação — SENAI | 14h")
bullet(doc, "Segurança Cibernética — SENAI | 4h")
bullet(doc, "Desvendando a Indústria 4.0 — SENAI | 20h")
bullet(doc, "Fundamentos da Inteligência Artificial — SENAI | 8h")
bullet(doc, "Ética na Inteligência Artificial — SENAI | 4h")
bullet(doc, "NR-11 — Operação de Empilhadeira — SENAI | 32h")
bullet(doc, "Reciclagens NR-11 — SuperFrio")
bullet(doc, "Caldeiraria Prática — AHCX Treinamentos / MJS Brasil | jan/2026 – jul/2026",
       after=3)

# ── Competências técnicas ────────────────────────────────────────────────────
titulo_secao(doc, "Competências Técnicas")

texto(doc, [("Tecnologia, Dados e Automação:", True, False)], before=2, after=1)
texto(doc, "Excel Avançado, Power BI, Power Automate, Power Apps, HTML, JavaScript, "
           "GitHub, GitHub Pages, Excel Online, Microsoft Forms, fundamentos de Python e "
           "desenvolvimento assistido por IA.", after=3)

texto(doc, [("Sistemas Corporativos:", True, False)], after=1)
texto(doc, "ServiceNow, Docnix, SAP, WMS Blue Yonder e ObraSoft.", after=3)

texto(doc, [("Administração e Operações:", True, False)], after=1)
texto(doc, "Governança documental, relatórios, indicadores, inventários, controle de "
           "materiais, gestão de EPIs, organização de almoxarifado, rotinas "
           "administrativas e suporte operacional.", after=3)

texto(doc, [("Logística e Indústria:", True, False)], after=1)
texto(doc, "Operação de empilhadeira, NR-11, expedição, armazenagem, picking, "
           "conferência, câmaras frias, PCP, metrologia, 5S, PDCA, Matriz GUT e gestão "
           "visual.", after=3)

texto(doc, [("Competências Comportamentais:", True, False)], after=1)
texto(doc, "Disciplina, responsabilidade, foco em resolução de problemas, aprendizado "
           "rápido, proatividade, resiliência, organização e mentalidade de melhoria "
           "contínua.", after=3)

# ── Informações adicionais ───────────────────────────────────────────────────
titulo_secao(doc, "Informações Adicionais")

bullet(doc, "CNH categorias A e B, com EAR.")
bullet(doc, "Disponibilidade para início imediato.")
bullet(doc, "Disponibilidade para turnos ou horário comercial.")
bullet(doc, "Interesse em oportunidades nas áreas de suporte de TI, administração, BI, "
            "logística, operações, manutenção, melhoria de processos e tecnologia.",
       after=0)

doc.save("Curriculo_Jeferson_dos_Santos_Cardoso.docx")
print("OK: Curriculo_Jeferson_dos_Santos_Cardoso.docx gerado.")
