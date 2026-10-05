"""
ProvaOps v4 — Word -> LaTeX (Acelerador) e versões embaralhadas (Embaralhador).

Base: v3. Mudanças principais em relação à v3 estão marcadas com "v4:".
"""
import inspect
import os
import random
import re
import string
import unicodedata
import zipfile
from dataclasses import dataclass, field
from io import BytesIO

import docx
import pandas as pd
import streamlit as st
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run


# ==========================================
# UTILIDADES DE TEXTO / LATEX
# ==========================================

_PLACEHOLDER_MATH = "\ufffc"  # ocupa 1 caractere no texto "plano" para equações

_ESC = {
    '\\': r'\textbackslash{}', '{': r'\{', '}': r'\}', '$': r'\$', '&': r'\&',
    '#': r'\#', '_': r'\_', '%': r'\%', '^': r'\^{}', '~': r'\textasciitilde{}',
}
_RE_ESC = re.compile(r'[\\{}$&#_%^~]')

# Símbolos Unicode que o pdflatex/inputenc não conhece. Valor = comando em modo matemático.
_SIMBOLOS = {
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\varepsilon',
    'ζ': r'\zeta', 'η': r'\eta', 'θ': r'\theta', 'κ': r'\kappa', 'λ': r'\lambda',
    'μ': r'\mu', 'ν': r'\nu', 'ξ': r'\xi', 'π': r'\pi', 'ρ': r'\rho', 'σ': r'\sigma',
    'τ': r'\tau', 'φ': r'\varphi', 'χ': r'\chi', 'ψ': r'\psi', 'ω': r'\omega',
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Θ': r'\Theta', 'Λ': r'\Lambda', 'Ξ': r'\Xi',
    'Π': r'\Pi', 'Σ': r'\Sigma', 'Φ': r'\Phi', 'Ψ': r'\Psi', 'Ω': r'\Omega',
    '\u2126': r'\Omega', '\u2206': r'\Delta', '\u00b5': r'\mu',
    '≤': r'\leq', '≥': r'\geq', '≠': r'\neq', '≈': r'\approx', '±': r'\pm', '∓': r'\mp',
    '∞': r'\infty', '→': r'\rightarrow', '←': r'\leftarrow', '↔': r'\leftrightarrow',
    '⇒': r'\Rightarrow', '⇔': r'\Leftrightarrow', '√': r'\surd', '∑': r'\sum',
    '∫': r'\int', '∝': r'\propto', '∂': r'\partial', '∇': r'\nabla', '−': '-',
}
# Só em modo matemático (em texto o inputenc já cobre × · °)
_SIMBOLOS_MATH_EXTRA = {'×': r'\times', '·': r'\cdot', '°': r'^{\circ}'}

_SUP_DE = '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻'
_SUB_DE = '₀₁₂₃₄₅₆₇₈₉₊₋'
_SUP_TR = str.maketrans(_SUP_DE, '0123456789+-')
_SUB_TR = str.maketrans(_SUB_DE, '0123456789+-')
_RE_UNICODE = re.compile(
    f'([{_SUP_DE}]+)|([{_SUB_DE}]+)|([{re.escape("".join(_SIMBOLOS))}])'
)
_MENOS = r'\ensuremath{-}'


def _unicode_para_latex(t):
    """Expoentes/índices Unicode (10⁻¹⁵, H₂O) e símbolos gregos/relacionais."""
    def sub(m):
        if m.group(1):
            return r'\textsuperscript{' + m.group(1).translate(_SUP_TR).replace('-', _MENOS) + '}'
        if m.group(2):
            return r'\textsubscript{' + m.group(2).translate(_SUB_TR).replace('-', _MENOS) + '}'
        s = _SIMBOLOS[m.group(3)]
        return s if s == '-' and False else (r'\ensuremath{' + s + '}')
    return _RE_UNICODE.sub(sub, t)


def _texto_para_latex(texto, escapar, quebra=" \\\\\n"):
    """v4: escapa %, &, $, _, #, {, }, ^, ~, \\ (opcional) e converte símbolos Unicode."""
    t = texto.replace('\r', '')
    if escapar:
        t = _RE_ESC.sub(lambda m: _ESC[m.group()], t)
    t = _unicode_para_latex(t)
    t = t.replace('\t', ' ').replace('\x0b', '\n').replace('\n', quebra)
    return t


def _sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


# ---------- Equações do Word (OMML) -> LaTeX (subconjunto comum) ----------

def _m(tag):
    return qn('m:' + tag)


def _math_texto(t):
    t = re.sub(r'([{}$&#_%])', r'\\\1', t)
    out = []
    for ch in t:
        if ch in _SIMBOLOS:
            s = _SIMBOLOS[ch]
            out.append(s + ' ' if s.startswith('\\') else s)
        elif ch in _SIMBOLOS_MATH_EXTRA:
            out.append(_SIMBOLOS_MATH_EXTRA[ch] + ' ')
        elif ch in _SUP_DE:
            out.append('^{' + ch.translate(_SUP_TR) + '}')
        elif ch in _SUB_DE:
            out.append('_{' + ch.translate(_SUB_TR) + '}')
        else:
            out.append(ch)
    return ''.join(out)


def _omml(el):
    tag = el.tag

    def filhos(e):
        return ''.join(_omml(c) for c in e)

    def sub(e, nome):
        c = e.find(_m(nome))
        return filhos(c) if c is not None else ''

    if tag == _m('t'):
        return _math_texto(el.text or '')
    if tag.endswith('Pr') and tag.startswith('{' + _m('t').split('}')[0][1:] + '}'):
        return ''
    if tag == _m('f'):
        return r'\frac{' + sub(el, 'num') + '}{' + sub(el, 'den') + '}'
    if tag == _m('sSup'):
        return '{' + sub(el, 'e') + '}^{' + sub(el, 'sup') + '}'
    if tag == _m('sSub'):
        return '{' + sub(el, 'e') + '}_{' + sub(el, 'sub') + '}'
    if tag == _m('sSubSup'):
        return '{' + sub(el, 'e') + '}_{' + sub(el, 'sub') + '}^{' + sub(el, 'sup') + '}'
    if tag == _m('rad'):
        grau, base = sub(el, 'deg'), sub(el, 'e')
        return (r'\sqrt[' + grau + ']{' + base + '}') if grau.strip() else (r'\sqrt{' + base + '}')
    if tag == _m('d'):
        def chr_de(nome, padrao):
            pr = el.find(_m('dPr'))
            if pr is not None:
                c = pr.find(_m(nome))
                if c is not None and c.get(_m('val')) is not None:
                    return c.get(_m('val'))
            return padrao
        ab, fe = chr_de('begChr', '('), chr_de('endChr', ')')
        sep = chr_de('sepChr', '|') if el.find(_m('dPr')) is not None and \
            el.find(_m('dPr')).find(_m('sepChr')) is not None else ','
        mapa = {'{': r'\{', '}': r'\}', '': '.', '⟨': r'\langle ', '⟩': r'\rangle ', '|': '|'}
        corpo = sep.join(filhos(e) for e in el.findall(_m('e')))
        return r'\left' + mapa.get(ab, ab) + corpo + r'\right' + mapa.get(fe, fe)
    if tag == _m('nary'):
        pr = el.find(_m('naryPr'))
        simb = '∫'
        if pr is not None and pr.find(_m('chr')) is not None:
            simb = pr.find(_m('chr')).get(_m('val')) or simb
        cmd = {'∑': r'\sum', '∫': r'\int', '∏': r'\prod', '∮': r'\oint'}.get(simb, r'\int')
        s, p = sub(el, 'sub'), sub(el, 'sup')
        return cmd + ('_{' + s + '}' if s else '') + ('^{' + p + '}' if p else '') + '{' + sub(el, 'e') + '}'
    if tag == _m('func'):
        nome = sub(el, 'fName').strip()
        if nome in ('sin', 'cos', 'tan', 'log', 'ln', 'exp', 'lim', 'sen'):
            nome = r'\sin' if nome == 'sen' else '\\' + nome
        else:
            nome = r'\mathrm{' + nome + '}'
        return nome + ' ' + sub(el, 'e')
    return filhos(el)


def _omml_para_latex(el):
    try:
        return _omml(el).strip()
    except Exception:
        return ''.join(t.text or '' for t in el.iter(_m('t')))


# ==========================================
# FUNÇÕES DO ACELERADOR (WORD -> LATEX)
# ==========================================

_TAG_W_R = qn('w:r')
_TAG_W_P = qn('w:p')
_TAG_W_TBL = qn('w:tbl')
_TAG_W_DEL = qn('w:del')
_TAG_W_PPR = qn('w:pPr')
_TAG_M_OMATH = _m('oMath')
_TAG_MC_FALLBACK = '{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback'
_TAG_VML_IMAGEDATA = '{urn:schemas-microsoft-com:vml}imagedata'
_ATTR_VML_ID = qn('r:id')

_IMG_SUPORTADAS = {'png', 'jpg', 'pdf'}  # o que o pdflatex compila direto

_RE_HEADING = re.compile(r'^(heading|t[ií]tulo)\s*\d*$', re.IGNORECASE)
# v4: "Questão 7" aceita sem pontuação; número solto exige "." ou ")" e NÃO pode ser decimal (1.5)
_RE_Q_PREFIXO = re.compile(r'^\s*quest[aã]o\s*(\d+)\s*[\.\)\-–—:]?\s*', re.IGNORECASE)
_RE_Q_NUM = re.compile(r'^\s*(\d+)\s*[\.\)](?!\d)\s*')
# v4: aceita "(a)", "a)" e "a." — a SEQUÊNCIA é validada no código (começa em "a", crescente)
_RE_ALT = re.compile(r'^\s*(?:\(([A-Za-z])\)|([A-Za-z])[\.\)])\s*')

_CABECALHO_TEX = r"""\documentclass[a4paper,10pt]{exam}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[brazil]{babel}
\usepackage{graphicx}
\usepackage{amsmath,amssymb}
\usepackage[shortlabels]{enumitem}
\usepackage{multicol}

\begin{document}
"""


@dataclass
class ResultadoConversao:
    latex: str
    imagens: list = field(default_factory=list)   # [(bytes, ext)]
    avisos: list = field(default_factory=list)
    resumo: dict = field(default_factory=dict)


def _iterar_blocos(doc):
    """v4: percorre parágrafos E tabelas na ordem real do documento (a v3 ignorava tabelas)."""
    for ch in doc.element.body.iterchildren():
        if ch.tag == _TAG_W_P:
            yield Paragraph(ch, doc)
        elif ch.tag == _TAG_W_TBL:
            yield Table(ch, doc)


def _dentro_de_fallback(el):
    return any(a.tag == _TAG_MC_FALLBACK for a in el.iterancestors())


def _extrair_imagens(elemento, doc_part):
    """Imagens (bytes, ext) na ordem do XML. Cobre DrawingML (a:blip) e VML (v:imagedata)."""
    imagens = []
    candidatos = [(b, b.get(qn('r:embed'))) for b in elemento.iter(qn('a:blip'))]
    candidatos += [(v, v.get(_ATTR_VML_ID)) for v in elemento.iter(_TAG_VML_IMAGEDATA)]
    candidatos.sort(key=lambda par: _posicao_documento(elemento, par[0]))
    for node, r_id in candidatos:
        if not r_id or _dentro_de_fallback(node):
            continue
        try:
            part = doc_part.related_parts[r_id]
        except KeyError:
            continue
        # v4: extensão vem do nome real da parte (image/x-emf, svg+xml etc. davam nomes inválidos)
        ext = os.path.splitext(str(part.partname))[1].lstrip('.').lower() or \
            part.content_type.split('/')[-1]
        ext = {'jpeg': 'jpg', 'x-png': 'png'}.get(ext, ext)
        imagens.append((part.blob, ext))
    return imagens


def _posicao_documento(raiz, node):
    for i, n in enumerate(raiz.iter()):
        if n is node:
            return i
    return 0


def _segmentos(para, stats):
    """Lista de (texto, fmt). fmt = (negrito, itálico, sup, sub) ou 'math' (texto já em LaTeX)."""
    segs = []

    def visitar(el):
        for ch in el:
            tag = ch.tag
            if tag == _TAG_W_R:
                run = Run(ch, para)
                t = run.text
                if t:
                    f = run.font
                    segs.append((t, (bool(run.bold), bool(run.italic),
                                     bool(f.superscript), bool(f.subscript))))
            elif tag == _TAG_M_OMATH:
                stats['equacoes'] += 1
                segs.append((_omml_para_latex(ch), 'math'))
            elif tag in (_TAG_W_PPR, _TAG_W_DEL):
                continue
            else:  # hyperlink, ins, smartTag, sdt, oMathPara...
                visitar(ch)

    visitar(para._p)
    return segs


def _plano(segs):
    return ''.join(_PLACEHOLDER_MATH if fmt == 'math' else t for t, fmt in segs)


def _cortar(segs, ini, fim):
    out, pos = [], 0
    for t, fmt in segs:
        n = 1 if fmt == 'math' else len(t)
        a, b = max(ini, pos), min(fim, pos + n)
        if a < b:
            out.append((t, fmt) if fmt == 'math' else (t[a - pos:b - pos], fmt))
        pos += n
    return out


def _render(segs, escapar, quebra=" \\\\\n"):
    # junta segmentos vizinhos com a mesma formatação
    juntos = []
    for t, fmt in segs:
        if juntos and fmt != 'math' and juntos[-1][1] == fmt:
            juntos[-1] = (juntos[-1][0] + t, fmt)
        else:
            juntos.append((t, fmt))
    partes = []
    for t, fmt in juntos:
        if fmt == 'math':
            if t:
                partes.append('$' + t + '$')
            continue
        s = _texto_para_latex(t, escapar, quebra)
        core = s.strip()
        if not core:
            partes.append(s)
            continue
        lead, trail = s[:len(s) - len(s.lstrip())], s[len(s.rstrip()):]
        bold, ital, sup, sub = fmt
        if sup:
            core = r'\textsuperscript{' + core.replace('-', _MENOS) + '}'
        elif sub:
            core = r'\textsubscript{' + core.replace('-', _MENOS) + '}'
        if ital:
            core = r'\textit{' + core + '}'
        if bold:
            core = r'\textbf{' + core + '}'
        partes.append(lead + core + trail)
    return ''.join(partes).strip()


def _tabela_para_latex(tabela, escapar, avisos):
    linhas = []
    for row in tabela.rows:
        cels, vistos = [], []
        celulas = list(row.cells)
        for c in celulas:
            if any(c._tc is v for v in vistos):  # célula mesclada aparece repetida
                continue
            vistos.append(c._tc)
            stats = {'equacoes': 0}
            textos = [_render(_segmentos(p, stats), escapar, quebra=' ') for p in c.paragraphs]
            cels.append(' '.join(t for t in textos if t))
        linhas.append(cels)
    if not linhas:
        return ''
    ncols = max(len(l) for l in linhas)
    if len({len(l) for l in linhas}) > 1:
        avisos.append("Uma tabela tem células mescladas/linhas irregulares — confira o layout no Overleaf.")
    if tabela._tbl.xpath('.//a:blip'):
        avisos.append("Uma tabela contém imagens dentro das células; elas NÃO foram extraídas.")
    corpo = ''.join(' & '.join(l + [''] * (ncols - len(l))) + r' \\ \hline' + '\n' for l in linhas)
    return ("\\begin{center}\n\\begin{tabular}{|" + "c|" * ncols + "}\n\\hline\n" + corpo +
            "\\end{tabular}\n\\end{center}\n\n")


def converter_docx_para_latex(docx_file, escapar=True):
    """
    Converte o .docx em LaTeX. Retorna ResultadoConversao(latex, imagens, avisos, resumo).

    v4 (em relação à v3):
      - escapa %, &, $, _, # ... (antes "R$ 50" ou "20%" quebravam/comentavam o LaTeX) [opcional]
      - converte 10⁻¹⁵, H₂O, λ, Δ, ≤ ... e preserva sobrescrito/subscrito/negrito/itálico
      - converte equações do Word (subconjunto: fração, potência, índice, raiz, delimitadores...)
      - lê TABELAS (antes eram ignoradas) e imagens VML; extensão correta das imagens
      - imagens ficam junto da alternativa/questão a que pertencem (antes iam para o item anterior)
      - "Questão 7" sem ponto é reconhecida; "(a)" e "A." como alternativa; 1.5 não é questão;
        sequência das alternativas validada (a, b, c...) — "A. Einstein" não vira alternativa
      - gera relatório de avisos (lacunas de numeração, questões sem alternativas, etc.)
    """
    doc = docx.Document(docx_file)
    doc_part = doc.part
    stats = {'equacoes': 0}
    avisos, out = [], [_CABECALHO_TEX]
    imagens = []

    ultima_questao_numero = None
    alternativas_iniciadas = False
    ultimo_foi_rejeitado_como_questao = False
    dentro_enumerate = False
    ultima_alternativa_aberta = False
    ultima_letra = None
    disciplina_atual = "(sem disciplina)"
    questao_atual = None  # {'num', 'alts', 'disc'}
    questoes = []
    n_tabelas = n_numeracao_auto = 0
    continuacoes = []

    def fechar_enum():
        nonlocal dentro_enumerate, ultima_alternativa_aberta, ultima_letra
        if dentro_enumerate:
            out.append("\\end{enumerate}\n\n")
        dentro_enumerate = False
        ultima_alternativa_aberta = False
        ultima_letra = None

    def emitir_imagens(elemento):
        for blob, ext in _extrair_imagens(elemento, doc_part):
            idx = len(imagens) + 1
            if ext not in _IMG_SUPORTADAS:
                avisos.append(f"Imagem {idx} está em formato '{ext}', que o pdflatex não compila "
                              f"— converta para PNG/JPG antes de subir no Overleaf.")
            out.append("\\begin{center}\n"
                       f"    \\includegraphics[width=0.6\\linewidth,height=0.3\\textheight,"
                       f"keepaspectratio]{{images/image{idx}.{ext}}}\n\\end{center}\n")
            imagens.append((blob, ext))

    for bloco in _iterar_blocos(doc):
        # ---------- tabelas ----------
        if isinstance(bloco, Table):
            fechar_enum()
            n_tabelas += 1
            out.append(_tabela_para_latex(bloco, escapar, avisos))
            ultimo_foi_rejeitado_como_questao = False
            continue

        para = bloco
        segs = _segmentos(para, stats)
        plano = _plano(segs)
        texto = plano.strip()
        lead = len(plano) - len(plano.lstrip())
        segs_texto = _cortar(segs, lead, lead + len(texto))
        estilo = (para.style.name if para.style is not None else '') or ''

        # ---------- títulos / disciplinas ----------
        if _RE_HEADING.match(estilo.strip()):
            if texto:
                fechar_enum()
                if questao_atual:
                    questoes.append(questao_atual)
                    questao_atual = None
                nome = _render(segs_texto, escapar, quebra=' ')
                disciplina_atual = nome
                out.append("\n% ==========================================\n"
                           f"\\section*{{DISCIPLINA: {nome}}}\n"
                           "% ==========================================\n")
                ultima_questao_numero = None
                alternativas_iniciadas = False
                ultimo_foi_rejeitado_como_questao = False
            emitir_imagens(para._p)
            continue

        if not texto:
            emitir_imagens(para._p)
            continue

        pPr = para._p.pPr
        if pPr is not None and pPr.numPr is not None:
            n_numeracao_auto += 1

        m_pref = _RE_Q_PREFIXO.match(texto)
        m_num = None if m_pref else _RE_Q_NUM.match(texto)
        match_q = m_pref or m_num

        # alternativa: letra válida na sequência (a primeira tem que ser "a"; depois, crescente)
        match_alt, letra_alt = _RE_ALT.match(texto), None
        if match_alt:
            letra_alt = match_alt.group(1) or match_alt.group(2)
            if ultima_letra is None:
                if letra_alt.lower() != 'a':
                    match_alt = None
            elif letra_alt.lower() <= ultima_letra:
                match_alt = None

        numero_atual = int(match_q.group(1)) if match_q else None
        numero_esperado = (ultima_questao_numero + 1) if ultima_questao_numero is not None else None
        aceitar_como_nova_questao = bool(match_q) and (
            bool(m_pref)  # "Questão N" explícito é sinal forte
            or ultima_questao_numero is None
            or alternativas_iniciadas
            or (not ultimo_foi_rejeitado_como_questao and numero_atual == numero_esperado)
        )

        if aceitar_como_nova_questao:
            fechar_enum()
            if questao_atual:
                questoes.append(questao_atual)
            if numero_esperado is not None and numero_atual != numero_esperado:
                avisos.append(f"Numeração pulou de {ultima_questao_numero} para {numero_atual} "
                              f"({disciplina_atual}) — alguma questão pode ter sido perdida ou unida.")
            out.append(f"\\subsection*{{Questão {numero_atual}}}\n")
            resto = _render(_cortar(segs_texto, match_q.end(), len(plano)), escapar)
            if resto:
                out.append(f"{resto}\n\n")
            questao_atual = {'num': numero_atual, 'alts': 0, 'disc': disciplina_atual}
            ultima_questao_numero = numero_atual
            alternativas_iniciadas = False
            ultimo_foi_rejeitado_como_questao = False

        elif match_alt:
            if not dentro_enumerate:
                estilo_enum = '(A)' if letra_alt.isupper() else '(a)'
                out.append(f"\\begin{{enumerate}}[{estilo_enum}]\n")
                dentro_enumerate = True
            resto = _render(_cortar(segs_texto, match_alt.end(), len(plano)), escapar)
            out.append(f"\\item {resto}\n")
            ultima_alternativa_aberta = True
            ultima_letra = letra_alt.lower()
            alternativas_iniciadas = True
            ultimo_foi_rejeitado_como_questao = False
            if questao_atual:
                questao_atual['alts'] += 1

        elif dentro_enumerate and ultima_alternativa_aberta:
            # continuação da alternativa (fórmula, linha extra...)
            rendered = _render(segs_texto, escapar)
            out.append(f"{rendered}\n")
            ultimo_foi_rejeitado_como_questao = False
            if questao_atual:
                continuacoes.append((questao_atual['num'], ultima_letra, texto[:50]))

        else:
            fechar_enum()
            out.append(f"{_render(segs_texto, escapar)}\n\n")
            ultimo_foi_rejeitado_como_questao = bool(match_q)

        emitir_imagens(para._p)

    fechar_enum()
    if questao_atual:
        questoes.append(questao_atual)
    out.append("\\end{document}")

    # ---------- relatório ----------
    sem_alt = [q['num'] for q in questoes if q['alts'] == 0]
    if sem_alt:
        avisos.append("Questões sem alternativas reconhecidas: " + ", ".join(map(str, sem_alt)) +
                      ". (Normal se forem dissertativas; senão confira as letras a)/b)/c)...)")
    estranhas = [f"{q['num']} ({q['alts']})" for q in questoes if q['alts'] not in (0, 4, 5)]
    if estranhas:
        avisos.append("Questões com número incomum de alternativas: " + ", ".join(estranhas) + ".")
    for num, letra, trecho in continuacoes[:10]:
        avisos.append(f"Questão {num}: parágrafo anexado à alternativa ({letra}) como continuação: "
                      f"\"{trecho}…\" — se for texto de apoio da PRÓXIMA questão, mova no Overleaf.")
    if len(continuacoes) > 10:
        avisos.append(f"... e mais {len(continuacoes) - 10} parágrafos anexados como continuação.")
    if n_numeracao_auto:
        avisos.append(f"{n_numeracao_auto} parágrafo(s) usam numeração/lista automática do Word: "
                      f"o número/letra não vem no texto e esses itens podem não ter sido reconhecidos.")
    if stats['equacoes']:
        avisos.append(f"{stats['equacoes']} equação(ões) do Word convertida(s) automaticamente — "
                      f"confira no Overleaf.")
    resumo = {'questoes': len(questoes), 'com_alternativas': len(questoes) - len(sem_alt),
              'imagens': len(imagens), 'tabelas': n_tabelas, 'equacoes': stats['equacoes']}
    return ResultadoConversao("".join(out), imagens, avisos, resumo)


def processar_acelerador_zip(docx_file, escapar=True):
    if hasattr(docx_file, 'seek'):
        docx_file.seek(0)
    res = converter_docx_para_latex(docx_file, escapar=escapar)
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as out_zip:
        out_zip.writestr("base.tex", res.latex)
        for idx, (blob, ext) in enumerate(res.imagens, start=1):
            out_zip.writestr(f"images/image{idx}.{ext}", blob)
    return zip_buffer.getvalue(), res


# ==========================================
# FUNÇÕES DO EMBARALHADOR (LATEX -> PROVAS)
# ==========================================

class QuestaoObj:
    def __init__(self, titulo, corpo):
        self.titulo = titulo
        self.corpo = corpo
        self.alternativas = []
        self.gabarito_orig = -1
        self.estilo_alternativa = "(a)"


class BlocoObj:
    def __init__(self):
        self.texto_apoio = ""
        self.questoes = []


class DisciplinaObj:
    def __init__(self, nome):
        self.nome = nome
        self.itens = []


@dataclass
class ProvaParseada:
    preambulo: str = ""
    cabecalho: str = ""
    disciplinas: list = field(default_factory=list)
    rodape: str = r"\end{document}"
    avisos: list = field(default_factory=list)
    erro: str = ""


_RE_TAG = re.compile(
    r'(\\(?:sub)?section\*?\{(?:[^{}\n]|\{[^{}\n]*\})*\}'
    r'|%[ \t]*IN[IÍ]CIO[ \t_-]*BLOCO'
    r'|%[ \t]*FIM[ \t_-]*BLOCO)',
    re.IGNORECASE,
)
_RE_MARCA_GABARITO = re.compile(r'%[ \t]*(?:CORRETO|CORRETA|CERTA)\b[^\n]*', re.IGNORECASE)
_RE_LISTA_TOKEN = re.compile(r'\\(begin|end)\{(?:enumerate|itemize|description)\}|\\item(?![a-zA-Z])')


def _remover_comando_hl(texto):
    """v4: remove \\hl{...} respeitando chaves aninhadas (a regex antiga quebrava \\hl{$x^{2}$})."""
    padrao = re.compile(r'\\hl(?![a-zA-Z])\s*\{')
    while True:
        m = padrao.search(texto)
        if not m:
            return texto
        i, prof, j = m.end(), 1, m.end()
        while j < len(texto) and prof:
            c = texto[j]
            if c == '\\':
                j += 2
                continue
            if c == '{':
                prof += 1
            elif c == '}':
                prof -= 1
            j += 1
        if prof:  # chaves desbalanceadas: remove só o comando
            texto = texto[:m.start()] + texto[m.end():]
        else:
            texto = texto[:m.start()] + texto[i:j - 1] + texto[j:]


def _enumerates_toplevel(texto):
    """[(inicio, fim)] dos \\begin{enumerate}...\\end{enumerate} de nível 0 (aninhados são respeitados)."""
    blocos, prof, ini = [], 0, None
    for m in re.finditer(r'\\(begin|end)\{enumerate\}', texto):
        if m.group(1) == 'begin':
            if prof == 0:
                ini = m.start()
            prof += 1
        elif prof > 0:
            prof -= 1
            if prof == 0:
                blocos.append((ini, m.end()))
    return blocos


def _itens_nivel0(inner):
    """Posições de \\item que pertencem ao nível 0 do bloco (ignora listas aninhadas)."""
    prof, pos = 0, []
    for m in _RE_LISTA_TOKEN.finditer(inner):
        if m.group(0).startswith('\\item'):
            if prof == 0:
                pos.append((m.start(), m.end()))
        elif m.group(1) == 'begin':
            prof += 1
        else:
            prof = max(0, prof - 1)
    return pos


def _extrair_alternativas(q, disc_nome, avisos):
    blocos = _enumerates_toplevel(q.corpo)
    if not blocos:
        return
    ini, fim = blocos[-1]
    bloco = q.corpo[ini:fim]
    m_ini = re.match(r'\\begin\{enumerate\}\s*(?:\[([^\]]*)\])?', bloco)
    inner = bloco[m_ini.end():-len(r'\end{enumerate}')]
    pos = _itens_nivel0(inner)
    if not pos:
        avisos.append(f"Questão '{q.titulo}' ({disc_nome}): lista sem \\item — mantida como está.")
        return
    if m_ini.group(1):
        q.estilo_alternativa = m_ini.group(1)

    prefixo = inner[:pos[0][0]]
    alternativas, idx_correto, n_marcas = [], -1, 0
    for k, (a, b) in enumerate(pos):
        fim_item = pos[k + 1][0] if k + 1 < len(pos) else len(inner)
        it = inner[b:fim_item]
        if _RE_MARCA_GABARITO.search(it):
            n_marcas += 1
            idx_correto = k
            # v4: remove a marca até o fim da linha (ela é comentário em LaTeX; antes o
            # texto depois de "%CORRETO" passava a aparecer na prova impressa)
            it = _RE_MARCA_GABARITO.sub('', it)
        it = _remover_comando_hl(it)  # v4: com chaves aninhadas
        alternativas.append(it.strip())

    q.alternativas = alternativas
    q.gabarito_orig = idx_correto
    q.corpo = q.corpo[:ini] + prefixo + "[[ALTS]]" + q.corpo[fim:]

    if n_marcas == 0:
        avisos.append(f"Questão '{q.titulo}' ({disc_nome}) sem %CORRETO identificado.")
    elif n_marcas > 1:
        avisos.append(f"Questão '{q.titulo}' ({disc_nome}) tem {n_marcas} alternativas com %CORRETO — "
                      f"só a última foi considerada.")
    if len(alternativas) > 26:
        avisos.append(f"Questão '{q.titulo}' ({disc_nome}) tem mais de 26 alternativas.")


def parse_latex_para_objetos(latex_content):
    """Retorna ProvaParseada. v4: cabeçalho fixo, erro explícito, marcações com acento, etc."""
    res = ProvaParseada()
    pre, sep, resto = latex_content.partition(r'\begin{document}')
    if not sep:
        res.erro = "não encontrei \\begin{document} no texto colado."
        return res
    corpo, sep2, _ = resto.rpartition(r'\end{document}')[::-1][::-1] if False else (None, None, None)
    corpo, sep2, _pos = resto.rpartition(r'\end{document}')
    if not sep2:
        res.erro = "não encontrei \\end{document} no texto colado."
        return res
    res.preambulo = pre + r'\begin{document}'

    tokens = _RE_TAG.split(corpo)
    # v4: tudo antes da primeira marcação é CABEÇALHO FIXO. Na v3 ele era descartado (quando a
    # 1ª marcação era uma DISCIPLINA) ou colado na 1ª questão e embaralhado junto com ela.
    res.cabecalho = tokens[0]
    if (len(tokens) - 1) % 2 != 0:
        tokens.append("")

    disciplinas = [DisciplinaObj("Geral")]
    disc_atual = disciplinas[0]
    bloco_atual = None
    questao_atual = None
    texto_buffer = ""

    for i in range(1, len(tokens), 2):
        tag = tokens[i].strip()
        conteudo = tokens[i + 1] if i + 1 < len(tokens) else ""

        if tag.startswith('\\'):
            m_t = re.match(r'\\(?:sub)?section\*?\{(.*)\}\s*$', tag, re.DOTALL)
            titulo = (m_t.group(1) if m_t else tag).strip()
            m_disc = re.match(r'\s*DISCIPLINA\s*:\s*(.*)$', titulo, re.IGNORECASE | re.DOTALL)

            if m_disc:
                if bloco_atual is not None:
                    res.avisos.append(f"'% INICIO BLOCO' sem 'FIM BLOCO' antes de {m_disc.group(1).strip()!r}.")
                disc_atual = DisciplinaObj(m_disc.group(1).strip())
                disciplinas.append(disc_atual)
                bloco_atual = None
                questao_atual = None
                texto_buffer = conteudo
            elif _sem_acento(titulo).lower().startswith('questao'):
                q = QuestaoObj(titulo, texto_buffer + conteudo)
                texto_buffer = ""
                questao_atual = q
                (bloco_atual.questoes if bloco_atual is not None else disc_atual.itens).append(q)
            else:
                texto_buffer += f"\n{tag}\n{conteudo}"
                if questao_atual:
                    questao_atual.corpo += f"\n{tag}\n{conteudo}"
                    texto_buffer = ""
        else:
            up = _sem_acento(tag).upper()
            if 'INICIO' in up:
                if bloco_atual is not None:
                    res.avisos.append("Dois '% INICIO BLOCO' seguidos sem 'FIM BLOCO' no meio.")
                bloco_atual = BlocoObj()
                bloco_atual.texto_apoio = texto_buffer + conteudo
                disc_atual.itens.append(bloco_atual)
                texto_buffer = ""
                questao_atual = None
            else:  # FIM BLOCO
                if bloco_atual is None:
                    res.avisos.append("'% FIM BLOCO' sem 'INICIO BLOCO' correspondente.")
                bloco_atual = None
                texto_buffer = conteudo
                questao_atual = None

    if bloco_atual is not None:
        res.avisos.append("'% INICIO BLOCO' sem 'FIM BLOCO' no final do documento.")

    for disc in disciplinas:
        for item in disc.itens:
            for q in (item.questoes if isinstance(item, BlocoObj) else [item]):
                _extrair_alternativas(q, disc.nome, res.avisos)

    res.disciplinas = [d for d in disciplinas if d.itens]
    if not res.disciplinas:
        res.erro = ("não encontrei nenhuma questão. Confira se existem linhas "
                    "\\subsection*{Questão N} no código.")
    return res


def _letra(i):
    return string.ascii_uppercase[i] if 0 <= i < 26 else "?"


def _questoes_do_item(item):
    return item.questoes if isinstance(item, BlocoObj) else [item]


def gerar_gabarito_original(prova):
    """v4: gabarito da Prova A (ordem original) — útil para conferir as marcações %CORRETO."""
    linhas, n = [], 1
    for disc in prova.disciplinas:
        for item in disc.itens:
            for q in _questoes_do_item(item):
                linhas.append({
                    "Disciplina": disc.nome, "Questão Nova": n,
                    "Gabarito": _letra(q.gabarito_orig) if q.alternativas and q.gabarito_orig >= 0
                    else ("?" if q.alternativas else "—"),
                    "Origem": q.titulo, "Versão": "Prova A",
                })
                n += 1
    return linhas


def gerar_latex_embaralhado(prova, seed, sufixo="B"):
    rng = random.Random(seed)
    novo = prova.preambulo + "\n" + prova.cabecalho  # v4: cabeçalho fixo preservado
    gabarito, contador = [], 1

    for disc in prova.disciplinas:
        if disc.nome != "Geral":
            novo += f"\n\\section*{{DISCIPLINA: {disc.nome}}}\n"

        itens = disc.itens.copy()
        rng.shuffle(itens)

        for item in itens:
            if isinstance(item, BlocoObj):
                novo += f"\n{item.texto_apoio}\n"

            for q in _questoes_do_item(item):
                novo += f"\\subsection*{{Questão {contador}}}\n"

                if q.alternativas:
                    indices = list(range(len(q.alternativas)))
                    rng.shuffle(indices)
                    bloco_alts = f"\\begin{{enumerate}}[{q.estilo_alternativa}]\n"
                    resp = "?"
                    for novo_i, original_i in enumerate(indices):
                        bloco_alts += f"\\item {q.alternativas[original_i]}\n"
                        if original_i == q.gabarito_orig:
                            resp = _letra(novo_i)
                    bloco_alts += "\\end{enumerate}\n"
                    texto_final = q.corpo.replace("[[ALTS]]", bloco_alts)
                else:
                    # v4: questão sem alternativas (dissertativa) — a v3 inseria um
                    # \begin{enumerate}\end{enumerate} vazio, que não compila.
                    texto_final, resp = q.corpo, "—"

                novo += texto_final
                gabarito.append({
                    "Disciplina": disc.nome, "Questão Nova": contador, "Gabarito": resp,
                    "Origem": q.titulo, "Versão": f"Prova {sufixo}",
                })
                contador += 1

    novo += prova.rodape
    return novo, gabarito


def gerar_pacote_embaralhado(latex_input, seed):
    """Retorna (zip_bytes, {A,B,C: DataFrame}, avisos, erro)."""
    prova = parse_latex_para_objetos(latex_input)
    if prova.erro:
        return None, {}, prova.avisos, prova.erro
    tex_b, gab_b = gerar_latex_embaralhado(prova, seed, "B")
    tex_c, gab_c = gerar_latex_embaralhado(prova, seed + 10, "C")
    dfs = {"A": pd.DataFrame(gerar_gabarito_original(prova)),
           "B": pd.DataFrame(gab_b), "C": pd.DataFrame(gab_c)}
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("Prova_A/Gabarito_A.csv", dfs["A"].to_csv(index=False))
        z.writestr("Prova_B/main_B.tex", tex_b)
        z.writestr("Prova_B/Gabarito_B.csv", dfs["B"].to_csv(index=False))
        z.writestr("Prova_C/main_C.tex", tex_c)
        z.writestr("Prova_C/Gabarito_C.csv", dfs["C"].to_csv(index=False))
    return buf.getvalue(), dfs, prova.avisos, ""


# ==========================================
# INTERFACE (FRONTEND)
# ==========================================

def _largura_total():
    """v4: compatível com versões novas e antigas do Streamlit (use_container_width foi descontinuado)."""
    params = inspect.signature(st.download_button).parameters
    return {'width': 'stretch'} if 'width' in params else {'use_container_width': True}


def _ler_upload_texto(arquivo):
    """v4: getvalue() em vez de read() (read() devolvia vazio nas reexecuções do Streamlit)."""
    dados = arquivo.getvalue()
    try:
        return dados.decode('utf-8-sig')
    except UnicodeDecodeError:
        return dados.decode('latin-1')


def main():
    st.set_page_config(page_title="ProvaOps - Sistema de Provas", layout="wide", page_icon="📝")
    st.title("🚀 Escola Analítica: ProvaOps")

    tab_acelerador, tab_embaralhador = st.tabs(
        ["⚡ 1. Acelerador (Extrair do Word)", "🎲 2. Embaralhador (Gerar Provas B e C)"])

    # --- ABA 1: ACELERADOR ---
    with tab_acelerador:
        st.header("Conversor Inteligente de Word para LaTeX")
        st.markdown("""
        **Como usar:**
        1. Baixe o documento do Google Docs clicando em `Arquivo > Fazer download > Microsoft Word (.docx)`.
        2. Suba o ficheiro abaixo.
        3. O sistema extrairá **texto, tabelas, equações e imagens** para você subir no Overleaf!
        """)

        file_docx = st.file_uploader("📂 Faça o upload da Prova em .docx", type=['docx'])
        escapar = st.checkbox(
            "Escapar caracteres especiais do LaTeX (%, &, $, _, #, ^, ~, chaves)", value=True,
            help="Deixe marcado para textos comuns (ex.: 'R$ 50', '20%'). Desmarque só se você "
                 "digitou comandos LaTeX/fórmulas com $...$ diretamente no Word.")

        if file_docx and st.button("⚙️ Processar e Extrair Imagens", type="primary"):
            try:
                zip_bytes, res = processar_acelerador_zip(file_docx, escapar)
                stem = os.path.splitext(file_docx.name)[0]
                st.session_state['acel'] = {'zip': zip_bytes, 'res': res, 'nome': f"{stem}_LaTeX.zip"}
            except Exception as e:
                st.session_state.pop('acel', None)
                st.error(f"Ocorreu um erro ao processar: {e}")

        acel = st.session_state.get('acel')
        if acel:
            res = acel['res']
            r = res.resumo
            st.success(f"✅ Conversão concluída: {r['questoes']} questões "
                       f"({r['com_alternativas']} com alternativas), {r['imagens']} imagens, "
                       f"{r['tabelas']} tabelas, {r['equacoes']} equações.")
            if res.avisos:
                st.warning("⚠️ Confira antes de seguir:\n\n" + "\n".join(f"- {a}" for a in res.avisos))
            st.info("💡 **Dica de Ouro:** Extraia o ficheiro .zip abaixo e suba tudo para o seu projeto "
                    "no Overleaf. Antes de usar o **Embaralhador**, abra o `base.tex` e adicione as tags "
                    "`%CORRETO` nas alternativas e `% INICIO BLOCO` / `% FIM BLOCO` nos textos de apoio.")
            st.download_button(label="📥 Baixar Pacote Base (.zip com imagens)", data=acel['zip'],
                               file_name=acel['nome'], mime="application/zip", **_largura_total())
            with st.expander("👀 Ver Prévia do Código Gerado"):
                st.code(res.latex, language="latex")

    # --- ABA 2: EMBARALHADOR ---
    with tab_embaralhador:
        st.header("Gerador de Versões (B e C)")

        with st.expander("📖 INSTRUÇÕES PARA O EDITOR (Clique para expandir)", expanded=True):
            st.markdown("""
            ### O que fazer com o ficheiro final?
            1. Certifique-se de que a sua **Prova A** no Overleaf já possui as marcações essenciais:
                * `\\section*{DISCIPLINA: Nome}`
                * `% INICIO BLOCO` e `% FIM BLOCO` nos textos de apoio.
                * `%CORRETO` dentro da alternativa certa.
            2. Tudo que vem **antes da primeira marcação** (cabeçalho, instruções) é mantido fixo no topo.
            3. Cole o código completo dessa Prova A abaixo (ou suba o `.tex`) e clique em Gerar.
            4. Baixe o `.zip`, extraia, e suba `main_B.tex` e `main_C.tex` para a mesma pasta da Prova A.
               O `Gabarito_A.csv` serve para conferir se as marcações `%CORRETO` foram lidas certo.
            5. Recompile e pronto!
            """)

        tex_upload = st.file_uploader("📂 (Opcional) Suba o ficheiro .tex em vez de colar", type=['tex'])
        texto_default = _ler_upload_texto(tex_upload) if tex_upload else ""
        latex_input = st.text_area("Cole o Código LaTeX da Prova A Original aqui:",
                                   value=texto_default, height=300)

        col_seed, _vazio = st.columns([1, 3])
        with col_seed:
            seed_val = int(st.number_input("Semente Inicial (Seed)", value=42, step=1))

        if st.button("🎲 Gerar Provas B e C (.zip)", type="primary"):
            if not latex_input.strip():
                st.warning("⚠️ Por favor, cole o código LaTeX (ou suba o .tex) antes de gerar.")
            else:
                zip_bytes, dfs, avisos, erro = gerar_pacote_embaralhado(latex_input, seed_val)
                if erro:
                    st.session_state.pop('emb', None)
                    st.error(f"❌ Não foi possível ler a estrutura da prova: {erro}")
                else:
                    st.session_state['emb'] = {'zip': zip_bytes, 'dfs': dfs, 'avisos': avisos,
                                               'seed': seed_val}

        emb = st.session_state.get('emb')
        if emb:
            if emb['avisos']:
                st.warning("⚠️ Atenção antes de baixar:\n\n" + "\n".join(f"- {a}" for a in emb['avisos']))
            st.success("✅ Provas geradas com sucesso!")
            st.download_button(label="📥 Baixar Pacote Completo de Embaralhamento (.zip)",
                               data=emb['zip'], file_name=f"Provas_Embaralhadas_Seed{emb['seed']}.zip",
                               mime="application/zip", **_largura_total())
            st.caption("🔍 Pré-visualização dos gabaritos:")
            t_b, t_c, t_a = st.tabs(["Gabarito B", "Gabarito C", "Gabarito A (original)"])
            for tab, chave in ((t_b, "B"), (t_c, "C"), (t_a, "A")):
                with tab:
                    st.dataframe(emb['dfs'][chave], hide_index=True)


if __name__ == "__main__":
    main()
