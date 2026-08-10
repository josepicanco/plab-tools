# -*- coding: utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except
"""
================================================================
             EXPORTSHEETS P-LAB - VERSAO 2.4.0
================================================================

HISTORICO:
   v2.4.0 (2026-08-10):
   - REESCRITO: "Vincular imagens ao DWG" nunca funcionou. A versao antiga
     dirigia o AutoCAD com script LISP + area de transferencia, e tinha tres
     bloqueios: (1) o ssget do LISP nao enxerga dentro de definicoes de
     bloco, que e exatamente onde o Revit poe o carimbo; (2) objeto OLE nao
     pode ser aninhado em bloco; (3) o clipboard e instavel em processo de
     segundo plano. Alem disso o log nunca era gravado (bug de codigo morto),
     entao a funcao podia reportar sucesso sem ter feito nada.
     Agora quem faz o trabalho e o plab-dwg-embed.exe (C# + ACadSharp), que
     grava o objeto OLE direto no arquivo. NAO exige AutoCAD instalado.
   - CORRIGIDO: perfil salvo nao restaurava a composicao do nome.
     apply_config ignorava custom_fields e separator; agora tambem
     restaura qualidade, cores, margens e configuracao DWG
   - ADICIONADO: texto fixo como item posicionavel da sequencia do
     nome (token TXT:) - pode ir a qualquer posicao, nao so nas pontas
   - ADICIONADO: separador avulso por posicao (token SEP:)
   - ADICIONADO: separador de campo livre e opcao de nao usar separador
   - ADICIONADO: campos calculados de data (ano, mes, dia, AAAAMMDD)
   - ADICIONADO: filtro para incluir ou nao parametros de projeto
   - ADICIONADO: mover item para o inicio/fim e limpar a sequencia
   - MELHORADO: janela de montagem migrada para BuildNameWindow.xaml
     (padrao P-LAB: XAML em arquivo separado, sem arquivo temporario)
   - ALTERADO: prefixo e sufixo valem apenas no modo simples; no modo
     montado eles sao desabilitados, substituidos pelo texto fixo

   v2.3.0 (2026-03-25):
   - ADICIONADO: Janela "Montar Nome" para composicao personalizada
     com parametros de projeto e informacoes de folha
   - ADICIONADO: Seletor de separador (-, _, .)
   - MELHORADO: generate_filename aceita custom_fields e separator

   v2.2.3 (2026-03-18):
   - CORRIGIDO: AddRaster ao inves de InsertOLE (imagens agora embedam)
   - CORRIGIDO: Nomes DWG sem espacos (AutoCAD COM compativel)
   - MELHORADO: Debug messages completo
   
   v2.2.2 (2026-02-27):
   - CORRIGIDO: Substituir espacos por underscore em DWG
   
   v2.2.1 (2026-02-27):
   - ADICIONADO: Vincular imagens via AutoCAD COM
   
   v2.2.0 (2026-02-27):
   - ADICIONADO: Exportacao DWF
   - ADICIONADO: Contador de tempo

AUTOR: P-LAB Engenharia
CONTATO: (61) 98206-8746 | engpicanco@yahoo.com.br
"""

import os
import os.path as op
import json
import codecs
import time
import glob
import subprocess

from pyrevit import HOST_APP, framework, forms, revit, DB, script

# ==========================================
# INICIALIZACAO
# ==========================================

logger = script.get_logger()
output = script.get_output()
output.set_height(600)

doc = revit.doc
forms.check_modeldoc(exitscript=True)
revit.selection.get_selection().clear()

REVIT_VERSION          = int(HOST_APP.version)
IS_REVIT_2021_OR_OLDER = REVIT_VERSION <= 2021
IS_REVIT_2022_OR_NEWER = HOST_APP.is_newer_than(2021)

# ==========================================
# CLASSE: ProfileManager
# ==========================================

class ProfileManager:
    """Gerencia perfis JSON"""

    @staticmethod
    def save_profile(config):
        from System.Windows.Forms import SaveFileDialog, DialogResult
        dialog = SaveFileDialog()
        dialog.Filter   = "Perfil JSON (*.json)|*.json"
        dialog.Title    = "Salvar Perfil de Exportacao"
        dialog.FileName = "Perfil_ExportSheets.json"
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                with codecs.open(dialog.FileName, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                return dialog.FileName
            except Exception as e:
                logger.error("Erro ao salvar perfil: %s", e)
        return None

    @staticmethod
    def load_profile():
        from System.Windows.Forms import OpenFileDialog, DialogResult
        dialog = OpenFileDialog()
        dialog.Filter = "Perfil JSON (*.json)|*.json"
        dialog.Title  = "Carregar Perfil de Exportacao"
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                with codecs.open(dialog.FileName, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except UnicodeDecodeError:
                try:
                    with codecs.open(dialog.FileName, 'r', encoding='latin-1') as f:
                        return json.load(f)
                except Exception as e:
                    logger.error("Erro ao carregar perfil: %s", e)
            except Exception as e:
                logger.error("Erro ao carregar perfil: %s", e)
        return None


# ==========================================
# CLASSE: PrintUtils
# ==========================================

class PrintUtils:

    @staticmethod
    def ensure_dir(dir_path):
        if not op.exists(dir_path):
            os.makedirs(dir_path)
        return dir_path

    @staticmethod
    def open_dir(dir_path):
        try:
            os.startfile(dir_path)
        except:
            pass

    @staticmethod
    def pdf_opts(hide_crop=True, hide_scope=True, hide_ref=True,
                 zoom_fit=True, use_vector=True, quality_tag="high",
                 color_tag="color", use_center=True, margin_x=0, margin_y=0):
        if IS_REVIT_2022_OR_NEWER:
            opts = DB.PDFExportOptions()
            opts.HideCropBoundaries = hide_crop
            opts.HideScopeBoxes     = hide_scope
            opts.HideReferencePlane = hide_ref
            if zoom_fit:
                opts.ZoomType = DB.ZoomType.FitToPage
            else:
                opts.ZoomType       = DB.ZoomType.Zoom
                opts.ZoomPercentage = 100
            if not use_vector:
                quality_map = {
                    "low":          DB.RasterQualityType.Draft,
                    "medium":       DB.RasterQualityType.Presentation,
                    "high":         DB.RasterQualityType.High,
                    "presentation": DB.RasterQualityType.High,
                }
                opts.RasterQuality = quality_map.get(quality_tag, DB.RasterQualityType.High)
                color_map = {
                    "color":     DB.ColorDepthType.Color,
                    "grayscale": DB.ColorDepthType.GrayScale,
                    "blackline": DB.ColorDepthType.BlackLine,
                }
                opts.ColorDepth = color_map.get(color_tag, DB.ColorDepthType.Color)
            try:
                if use_center:
                    opts.PaperPlacement = DB.PaperPlacementType.Center
                else:
                    opts.PaperPlacement = DB.PaperPlacementType.Margins
                    opts.OriginOffsetX  = margin_x * 0.00328084
                    opts.OriginOffsetY  = margin_y * 0.00328084
            except:
                pass
            return opts
        else:
            return DB.ViewScheduleExportOptions()

    @staticmethod
    def dwg_opts():
        opts = DB.DWGExportOptions()
        opts.SharedCoords   = False
        opts.MergedViews    = True
        opts.ExportingAreas = False
        opts.FileVersion    = DB.ACADVersion.R2013
        return opts

    @staticmethod
    def dwf_opts():
        opts = DB.DWFExportOptions()
        opts.MergedViews    = True
        opts.ExportingAreas = False
        opts.ImageFormat    = DB.ImageFileType.PNG
        opts.ImageQuality   = DB.ImageResolution.DPI_300
        return opts

    @staticmethod
    def export_sheet_pdf(dir_path, sheet, options, doc, filename):
        if IS_REVIT_2021_OR_OLDER:
            pdf_path = op.join(dir_path, filename)
            try:
                uidoc = HOST_APP.uiapp.ActiveUIDocument
                uidoc.ActiveView = sheet
                pm = doc.PrintManager
                try:
                    pm.SelectNewPrintDriver("Microsoft Print to PDF")
                except:
                    try:
                        pm.SelectNewPrintDriver("Adobe PDF")
                    except:
                        raise Exception("Driver PDF nao encontrado")
                ps = pm.PrintSetup
                ps.CurrentPrintSetting = ps.InSession
                pm.PrintRange      = DB.PrintRange.Current
                pm.PrintToFile     = True
                pm.CombinedFile    = True
                pm.PrintToFileName = pdf_path
                temp_name = "TempPDF_{}".format(sheet.Id.IntegerValue)
                try:
                    ps.SaveAs(temp_name)
                except:
                    pass
                pm.Apply()
                pm.SubmitPrint()
                try:
                    ps.Delete(temp_name)
                except:
                    pass
                return True
            except Exception as e:
                raise Exception("Revit 2021 PDF: {}".format(str(e)))
        else:
            options.FileName = op.splitext(filename)[0]
            sheet_ids = framework.List[DB.ElementId]()
            sheet_ids.Add(sheet.Id)
            doc.Export(dir_path, sheet_ids, options)
            return True

    @staticmethod
    def export_sheet_dwg(dir_path, sheet, options, doc, filename):
        sheet_ids = framework.List[DB.ElementId]()
        sheet_ids.Add(sheet.Id)
        doc.Export(dir_path, op.splitext(filename)[0] + ".dwg", sheet_ids, options)
        return True

    @staticmethod
    def export_sheet_dwf(dir_path, sheet, options, doc, filename):
        sheet_ids = framework.List[DB.ElementId]()
        sheet_ids.Add(sheet.Id)
        doc.Export(dir_path, op.splitext(filename)[0] + ".dwf", sheet_ids, options)
        return True


# ==========================================
# CLASSE: DWGPostProcessor v2.4.0
# ==========================================

class DWGPostProcessor:
    u"""Embute as imagens referenciadas nos DWGs, deixando-os autocontidos.

    O Revit exporta logos e carimbos como REFERENCIA externa: o DWG guarda
    apenas o caminho de um PNG que fica ao lado. Quem recebe so o DWG nao ve
    a imagem.

    O formato DWG nao tem imagem embutida nativa - a unica forma de colocar
    os pixels dentro do arquivo e um objeto OLE. Ate a v2.3.0 isso era
    tentado dirigindo o AutoCAD com um script LISP e a area de transferencia
    do Windows, o que nunca funcionou: o `ssget` do LISP nao enxerga dentro
    de definicoes de bloco (e o carimbo do Revit e um bloco), OLE nao pode
    ser aninhado em bloco, e o clipboard e instavel em processo de segundo
    plano.

    A partir da v2.4.0 quem faz o trabalho e o `plab-dwg-embed.exe`, escrito
    em C# sobre a biblioteca ACadSharp. Ele le o DWG, encontra as imagens
    inclusive dentro dos blocos, converte a posicao pela insercao do bloco,
    grava um objeto OLE do tipo StaticDib e remove a referencia externa.

    Vantagens sobre a abordagem antiga:
      - NAO precisa de AutoCAD instalado, nem na nossa maquina nem na do cliente
      - roda em segundo plano de verdade, sem janela e sem clipboard
      - StaticDib e renderizado pelo proprio Windows, sem servidor OLE
    """

    EXECUTAVEL = "plab-dwg-embed.exe"

    # Maior lado da imagem em pixels. O AutoCAD nao renderiza OLE grande de
    # forma confiavel, e o DIB nao tem compressao: cada pixel custa 4 bytes.
    MAX_PIXELS = 1600

    @staticmethod
    def localizar_executavel():
        u"""Procura o executavel numa pasta 'bin'.

        Primeiro dentro do proprio botao, que e como a versao distribuida
        instala (o instalador copia paineis inteiros, entao o exe precisa
        estar dentro do painel). Depois sobe ate a raiz da extensao, que e
        o layout usado em desenvolvimento.
        """
        pasta = op.dirname(__file__)
        for _ in range(8):
            candidato = op.join(pasta, "bin", DWGPostProcessor.EXECUTAVEL)
            if op.exists(candidato):
                return candidato
            pai = op.dirname(pasta)
            if pai == pasta:
                break
            pasta = pai
        return None

    @staticmethod
    def bind_xref_images_folder(dwg_paths_or_folder, output=None):
        u"""Embute as imagens nos DWGs indicados.

        Retorna (quantidade_ok, quantidade_erro).
        """
        def log(msg):
            if output:
                output.print_md(msg)

        log("")
        log(u"## [BIND] Embutindo imagens nos DWGs")
        log("")

        exe = DWGPostProcessor.localizar_executavel()
        if not exe:
            log(u"[ERRO] `{}` nao encontrado na pasta bin da extensao.".format(
                DWGPostProcessor.EXECUTAVEL))
            return 0, 1

        if isinstance(dwg_paths_or_folder, list):
            arquivos = [p for p in dwg_paths_or_folder if op.exists(p)]
        else:
            arquivos = glob.glob(op.join(dwg_paths_or_folder, "*.dwg"))

        if not arquivos:
            log(u"[AVISO] Nenhum arquivo DWG para processar.")
            return 0, 0

        log(u"Processando {} DWG(s) sem abrir o AutoCAD...".format(len(arquivos)))

        comando = [exe, "--json", "--max-px", str(DWGPostProcessor.MAX_PIXELS)]
        comando.extend(arquivos)

        try:
            processo = subprocess.Popen(
                comando,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                shell=False
            )
            saida, erro = processo.communicate()
        except Exception as e:
            log(u"[ERRO] Falha ao executar: {}".format(e))
            return 0, len(arquivos)

        # O executavel depende do runtime .NET 8. Sem ele o Windows responde
        # com uma mensagem generica; traduzimos para algo acionavel.
        if erro and ".NET" in erro and "install" in erro.lower():
            log(u"[ERRO] O runtime .NET 8 nao esta instalado nesta maquina.")
            log(u"Baixe em: https://dotnet.microsoft.com/download/dotnet/8.0/runtime")
            log(u"(reinstalar o P-LAB Tools pelo instalador tambem resolve)")
            return 0, len(arquivos)

        ok = 0
        falhas = 0
        total_imagens = 0

        for linha in (saida or "").splitlines():
            linha = linha.strip()
            if not linha or not linha.startswith("{"):
                continue
            try:
                r = json.loads(linha)
            except Exception:
                continue

            nome = op.basename(r.get("arquivo", "?"))
            if not r.get("sucesso"):
                falhas += 1
                log(u"[ERRO] {}: {}".format(nome, r.get("erro")))
            elif r.get("encontradas", 0) == 0:
                ok += 1
                log(u"[--] {}: nenhuma imagem referenciada".format(nome))
            else:
                ok += 1
                total_imagens += r.get("embutidas", 0)
                log(u"[OK] {}: {}/{} imagem(ns) embutida(s)".format(
                    nome, r.get("embutidas"), r.get("encontradas")))

            for aviso in r.get("avisos") or []:
                log(u"   aviso: {}".format(aviso))

        if erro:
            for linha in erro.splitlines():
                if linha.strip():
                    log(u"   [stderr] {}".format(linha.strip()))

        log("")
        log(u"**Resumo: {} arquivo(s) processado(s), {} imagem(ns) embutida(s), {} erro(s)**".format(
            ok, total_imagens, falhas))

        return ok, falhas


# ==========================================
# TOKENS DA COMPOSICAO DO NOME
# ==========================================
#
# A sequencia do nome (custom_fields) e uma lista ordenada de strings.
# Cada string e um "token" que pode ser:
#
#   "Numero da Prancha"   -> campo interno da folha
#   "Nome da Prancha"     -> campo interno da folha
#   "[Projeto] X"         -> parametro X de Informacoes do Projeto
#   "Qualquer Parametro"  -> parametro de instancia da folha
#   "TXT:REV01"           -> texto fixo, posicionavel em qualquer lugar
#   "SEP:_"               -> separador avulso, usado so naquela posicao
#   "DATA:ano"            -> campo calculado na hora da exportacao
#
# Tokens sem prefixo sao os do formato antigo (v2.3.0) - perfis antigos
# continuam carregando normalmente.

TOKEN_TEXTO = u"TXT:"
TOKEN_SEP   = u"SEP:"
TOKEN_DATA  = u"DATA:"
TOKEN_PROJ  = u"[Projeto] "

CAMPO_NUMERO = u"Numero da Prancha"
CAMPO_NOME   = u"Nome da Prancha"

# token -> rotulo exibido nas listas
CAMPOS_DATA = [
    (TOKEN_DATA + u"ano",  u"Ano (atual)"),
    (TOKEN_DATA + u"mes",  u"Mes (atual)"),
    (TOKEN_DATA + u"dia",  u"Dia (atual)"),
    (TOKEN_DATA + u"data", u"Data completa (AAAAMMDD)"),
]

INVALID_CHARS = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']


def limpar_nome(texto):
    """Remove caracteres invalidos para nome de arquivo."""
    if not texto:
        return u""
    for c in INVALID_CHARS:
        texto = texto.replace(c, '_')
    return texto


def valor_data(chave):
    """Resolve um campo de data calculado na hora da exportacao."""
    tm = time.localtime()
    if chave == u"ano":
        return "{:04d}".format(tm.tm_year)
    if chave == u"mes":
        return "{:02d}".format(tm.tm_mon)
    if chave == u"dia":
        return "{:02d}".format(tm.tm_mday)
    if chave == u"data":
        return "{:04d}{:02d}{:02d}".format(tm.tm_year, tm.tm_mon, tm.tm_mday)
    return u""


def rotulo_token(token):
    """Texto amigavel exibido nas listas da janela de montagem."""
    if token.startswith(TOKEN_TEXTO):
        return u'Texto fixo:  "{}"'.format(token[len(TOKEN_TEXTO):])
    if token.startswith(TOKEN_SEP):
        return u'Separador:  "{}"'.format(token[len(TOKEN_SEP):])
    if token.startswith(TOKEN_DATA):
        for tok, rotulo in CAMPOS_DATA:
            if tok == token:
                return rotulo
        return token
    return token


def ler_parametro(elemento, nome):
    """Le um parametro pelo nome e devolve o valor como texto. '' se nao existir."""
    try:
        p = elemento.LookupParameter(nome)
    except:
        return u""
    if not p or not p.HasValue:
        return u""
    try:
        if p.StorageType == DB.StorageType.String:
            return p.AsString() or u""
        if p.StorageType == DB.StorageType.Integer:
            return str(p.AsInteger())
        if p.StorageType == DB.StorageType.Double:
            # AsValueString respeita as unidades do projeto
            return p.AsValueString() or str(round(p.AsDouble(), 4))
    except:
        pass
    return u""


def resolver_token(sheet, token):
    """Resolve um token para (tipo, valor).

    tipo 'sep' entra literal no nome, sem separador de campo ao redor.
    tipo 'val' e um campo comum, separado dos vizinhos pelo separador de campo.
    """
    if token.startswith(TOKEN_SEP):
        return ('sep', token[len(TOKEN_SEP):])

    if token.startswith(TOKEN_TEXTO):
        return ('val', token[len(TOKEN_TEXTO):])

    if token.startswith(TOKEN_DATA):
        return ('val', valor_data(token[len(TOKEN_DATA):]))

    if token == CAMPO_NUMERO:
        return ('val', sheet.SheetNumber or u"SEM_NUMERO")

    if token == CAMPO_NOME:
        return ('val', sheet.Name or u"SEM_NOME")

    if token.startswith(TOKEN_PROJ):
        try:
            return ('val', ler_parametro(doc.ProjectInformation,
                                         token[len(TOKEN_PROJ):]))
        except:
            return ('val', u"")

    return ('val', ler_parametro(sheet, token))


def juntar_partes(itens, sep, use_separator):
    """Junta os itens resolvidos aplicando o separador de campo entre valores.

    itens: lista de (tipo, texto) ja limpos, sem vazios.
    """
    partes  = []
    anterior_foi_valor = False
    for tipo, texto in itens:
        if tipo == 'sep':
            partes.append(texto)
            anterior_foi_valor = False
        else:
            if anterior_foi_valor and use_separator and sep:
                partes.append(sep)
            partes.append(texto)
            anterior_foi_valor = True
    return u"".join(partes)


# ==========================================
# FUNCAO AUXILIAR
# ==========================================

def get_sheet_params(sheet):
    """Parametros da folha disponiveis para compor o nome. {token: valor_exemplo}"""
    params = {}
    params[CAMPO_NUMERO] = sheet.SheetNumber or u""
    params[CAMPO_NOME]   = sheet.Name        or u""

    for p in sheet.Parameters:
        try:
            nome = p.Definition.Name
            if nome in params:
                continue
            if p.StorageType == DB.StorageType.String:
                params[nome] = p.AsString() or u""
            elif p.StorageType == DB.StorageType.Integer:
                params[nome] = str(p.AsInteger())
            elif p.StorageType == DB.StorageType.Double:
                params[nome] = p.AsValueString() or str(round(p.AsDouble(), 4))
        except:
            pass
    return params


def get_project_params():
    """Parametros de Informacoes do Projeto. {token: valor_exemplo}"""
    params = {}
    try:
        proj_info = doc.ProjectInformation
        for p in proj_info.Parameters:
            try:
                nome = TOKEN_PROJ + p.Definition.Name
                if nome in params:
                    continue
                if p.StorageType == DB.StorageType.String:
                    params[nome] = p.AsString() or u""
                elif p.StorageType == DB.StorageType.Integer:
                    params[nome] = str(p.AsInteger())
            except:
                pass
    except:
        pass
    return params


def generate_filename(sheet, prefix="", suffix="", inc_num=True, inc_name=True,
                      replace_spaces=False, separator="-", custom_fields=None,
                      use_separator=True):
    """Gera o nome do arquivo de uma prancha.

    Args:
        prefix/suffix:  usados APENAS no modo simples (sem custom_fields).
                        No modo montado, texto fixo entra como token TXT:.
        inc_num/inc_name: modo simples - incluir numero e/ou nome da prancha.
        replace_spaces: se True, troca espacos pelo separador (DWG/AutoCAD).
        separator:      separador de campo. Qualquer texto, nao so '-', '_', '.'.
        custom_fields:  sequencia de tokens montada pelo usuario.
        use_separator:  se False, os campos sao concatenados sem separador.
    """
    sep = limpar_nome(separator if separator is not None else u"")

    def preparar(texto):
        texto = limpar_nome(texto)
        if replace_spaces:
            texto = texto.replace(' ', sep if sep else '_')
        return texto

    itens = []

    if custom_fields:
        # Modo montado: prefixo e sufixo nao se aplicam - o usuario posiciona
        # o texto fixo onde quiser dentro da propria sequencia.
        for token in custom_fields:
            tipo, valor = resolver_token(sheet, token)
            if tipo == 'sep':
                valor = limpar_nome(valor)
            else:
                valor = preparar(valor)
            if valor:
                itens.append((tipo, valor))
    else:
        # Modo simples: prefixo + numero + nome + sufixo
        p = preparar(prefix)
        if p:
            itens.append(('val', p))
        if inc_num:
            itens.append(('val', preparar(sheet.SheetNumber or u"SEM_NUMERO")))
        if inc_name:
            itens.append(('val', preparar(sheet.Name or u"SEM_NOME")))
        s = preparar(suffix)
        if s:
            itens.append(('val', s))
        itens = [it for it in itens if it[1]]

    nome = juntar_partes(itens, sep, use_separator)
    if not nome:
        nome = limpar_nome(sheet.SheetNumber) or u"ARQUIVO"
    return nome


# ==========================================
# CLASSE: BuildNameWindow (WPF)
# ==========================================

class BuildNameWindow(forms.WPFWindow):
    """Janela de montagem do nome do arquivo.

    A sequencia (self.result_fields) e uma lista ordenada de tokens. O usuario
    pode intercalar parametros, textos fixos e separadores avulsos, e mover
    qualquer item para qualquer posicao - inclusive um texto fixo como prefixo
    ou sufixo.
    """

    SEPARADORES_SUGERIDOS = [u"-", u"_", u".", u" "]

    def __init__(self, sheet_params, project_params, current_fields, sample_sheet,
                 separator=u"-", use_separator=True):
        forms.WPFWindow.__init__(self, script.get_bundle_file('BuildNameWindow.xaml'))

        self.sheet_params   = sheet_params      # {token: valor_exemplo}
        self.project_params = project_params    # {token: valor_exemplo}
        self.sample_sheet   = sample_sheet
        self.result_fields  = list(current_fields) if current_fields else []

        # Se a sequencia carregada ja usa parametros de projeto, mostra a lista deles
        tem_proj = False
        for token in self.result_fields:
            if token.startswith(TOKEN_PROJ):
                tem_proj = True
                break
        self.incluir_projeto_cb.IsChecked = tem_proj

        self._separador_inicial = separator if separator is not None else u"-"
        self.sep_campo_cb.ItemsSource = list(self.SEPARADORES_SUGERIDOS)
        if self._separador_inicial in self.SEPARADORES_SUGERIDOS:
            self.sep_campo_cb.SelectedItem = self._separador_inicial
        else:
            self.sep_campo_cb.Text = self._separador_inicial

        self.usar_sep_cb.IsChecked  = bool(use_separator)
        self.sep_campo_cb.IsEnabled = bool(use_separator)

        self._refresh_lists()
        self._connect_events()
        # O template do ComboBox editavel so existe apos o Loaded - reaplicar o
        # texto aqui evita que um separador digitado a mao se perca na abertura.
        self.Loaded += self._on_loaded

    def _on_loaded(self, sender, args):
        if self.sep_campo_cb.Text != self._separador_inicial:
            self.sep_campo_cb.Text = self._separador_inicial
        self._update_preview()

    # ---------- estado ----------

    def get_separator(self):
        texto = self.sep_campo_cb.Text
        return texto if texto is not None else u""

    def get_use_separator(self):
        return bool(self.usar_sep_cb.IsChecked)

    def _tokens_disponiveis(self):
        """Tokens da lista da esquerda, ja ordenados e sem os que estao em uso."""
        tokens = list(self.sheet_params.keys())
        if self.incluir_projeto_cb.IsChecked:
            tokens += list(self.project_params.keys())
        tokens.sort()
        # campos calculados sempre no fim da lista
        tokens += [tok for tok, _ in CAMPOS_DATA]
        return [t for t in tokens if t not in self.result_fields]

    # ---------- eventos ----------

    def _connect_events(self):
        self.add_btn.Click       += self._on_add
        self.remove_btn.Click    += self._on_remove
        self.up_btn.Click        += self._on_up
        self.down_btn.Click      += self._on_down
        self.topo_btn.Click      += self._on_topo
        self.fundo_btn.Click     += self._on_fundo
        self.limpar_btn.Click    += self._on_limpar
        self.add_texto_btn.Click += self._on_add_texto
        self.add_sep_btn.Click   += self._on_add_sep
        self.ok_btn.Click        += self._on_ok
        self.cancel_btn.Click    += self._on_cancel

        self.params_lb.MouseDoubleClick   += self._on_add
        self.selected_lb.MouseDoubleClick += self._on_remove

        self.incluir_projeto_cb.Checked   += self._on_toggle_projeto
        self.incluir_projeto_cb.Unchecked += self._on_toggle_projeto

        self.usar_sep_cb.Checked          += self._on_sep_changed
        self.usar_sep_cb.Unchecked        += self._on_sep_changed
        self.sep_campo_cb.SelectionChanged += self._on_sep_changed
        self.sep_campo_cb.KeyUp            += self._on_sep_changed

        self.texto_fixo_tb.KeyUp += self._on_texto_key
        self.sep_avulso_tb.KeyUp += self._on_sep_avulso_key

    def _on_toggle_projeto(self, sender, args):
        self._refresh_lists()

    def _on_sep_changed(self, sender, args):
        self.sep_campo_cb.IsEnabled = self.usar_sep_cb.IsChecked
        self._update_preview()

    def _on_texto_key(self, sender, args):
        # Comparacao por texto evita depender do namespace System.Windows.Input
        try:
            if str(args.Key) == "Return":
                self._on_add_texto(sender, args)
        except:
            pass

    def _on_sep_avulso_key(self, sender, args):
        try:
            if str(args.Key) == "Return":
                self._on_add_sep(sender, args)
        except:
            pass

    # ---------- manipulacao da sequencia ----------

    def _refresh_lists(self, indice_selecionado=-1):
        self.params_lb.ItemsSource   = [rotulo_token(t) for t in self._tokens_disponiveis()]
        self.selected_lb.ItemsSource = [rotulo_token(t) for t in self.result_fields]
        if 0 <= indice_selecionado < len(self.result_fields):
            self.selected_lb.SelectedIndex = indice_selecionado
        self._update_preview()

    def _on_add(self, sender, args):
        idx = self.params_lb.SelectedIndex
        disponiveis = self._tokens_disponiveis()
        if 0 <= idx < len(disponiveis):
            self.result_fields.append(disponiveis[idx])
            self._refresh_lists(len(self.result_fields) - 1)

    def _on_remove(self, sender, args):
        idx = self.selected_lb.SelectedIndex
        if 0 <= idx < len(self.result_fields):
            self.result_fields.pop(idx)
            self._refresh_lists(min(idx, len(self.result_fields) - 1))

    def _on_add_texto(self, sender, args):
        texto = (self.texto_fixo_tb.Text or u"").strip()
        if not texto:
            return
        self.result_fields.append(TOKEN_TEXTO + texto)
        self.texto_fixo_tb.Text = u""
        self._refresh_lists(len(self.result_fields) - 1)

    def _on_add_sep(self, sender, args):
        texto = self.sep_avulso_tb.Text or u""
        if not texto:
            return
        self.result_fields.append(TOKEN_SEP + texto)
        self.sep_avulso_tb.Text = u""
        self._refresh_lists(len(self.result_fields) - 1)

    def _mover(self, origem, destino):
        token = self.result_fields.pop(origem)
        self.result_fields.insert(destino, token)
        self._refresh_lists(destino)

    def _on_up(self, sender, args):
        idx = self.selected_lb.SelectedIndex
        if idx > 0:
            self._mover(idx, idx - 1)

    def _on_down(self, sender, args):
        idx = self.selected_lb.SelectedIndex
        if 0 <= idx < len(self.result_fields) - 1:
            self._mover(idx, idx + 1)

    def _on_topo(self, sender, args):
        idx = self.selected_lb.SelectedIndex
        if idx > 0:
            self._mover(idx, 0)

    def _on_fundo(self, sender, args):
        idx = self.selected_lb.SelectedIndex
        if 0 <= idx < len(self.result_fields) - 1:
            self._mover(idx, len(self.result_fields) - 1)

    def _on_limpar(self, sender, args):
        self.result_fields = []
        self._refresh_lists()

    # ---------- preview ----------

    def _update_preview(self):
        try:
            if not self.sample_sheet:
                self.preview_tb.Text = u"(nenhuma prancha para exemplo)"
                return
            if not self.result_fields:
                self.preview_tb.Text = u"(sequencia vazia - sera usado Numero + Nome padrao)"
                return
            nome = generate_filename(
                self.sample_sheet,
                separator=self.get_separator(),
                custom_fields=self.result_fields,
                use_separator=self.get_use_separator()
            )
            self.preview_tb.Text = nome + u".pdf"
        except Exception as e:
            self.preview_tb.Text = u"Erro: {}".format(e)

    # ---------- fechamento ----------

    def _on_ok(self, sender, args):
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# ==========================================
# CLASSE: ExportSheetsWindow (WPF)
# ==========================================

class ExportSheetsWindow(forms.WPFWindow):

    def __init__(self, xaml_file, selected_sheets, dwg_settings_dict):
        forms.WPFWindow.__init__(self, xaml_file)
        self.selected_sheets   = selected_sheets
        self.dwg_settings_dict = dwg_settings_dict
        # Estado do construtor de nome
        self.custom_fields    = []     # [] = modo simples (numero + nome)
        self.custom_separator = None   # None = usa os radios da janela principal
        self.use_separator    = True
        self._setup_combos()
        self._connect_events()
        self._atualizar_modo_nome()

    def _setup_combos(self):
        dwg_names = list(self.dwg_settings_dict.keys()) if self.dwg_settings_dict else ["<Padrao>"]
        self.dwg_setup_cb.ItemsSource = dwg_names
        if dwg_names:
            self.dwg_setup_cb.SelectedIndex = 0

    def _connect_events(self):
        self.prefix_tb.TextChanged       += self.update_preview
        self.suffix_tb.TextChanged       += self.update_preview
        self.include_number_cb.Checked   += self.update_preview
        self.include_number_cb.Unchecked += self.update_preview
        self.include_name_cb.Checked     += self.update_preview
        self.include_name_cb.Unchecked   += self.update_preview
        self.sep_hifen_rb.Checked        += self.update_preview
        self.sep_underline_rb.Checked    += self.update_preview
        self.sep_ponto_rb.Checked        += self.update_preview
        self.raster_rb.Checked           += self.processing_changed
        self.vector_rb.Checked           += self.processing_changed
        self.position_offset_rb.Checked  += self.position_changed
        self.position_center_rb.Checked  += self.position_changed
        self.margins_cb.SelectionChanged += self.margins_changed
        self.bind_images_cb.Checked      += self.bind_images_changed
        self.bind_images_cb.Unchecked    += self.bind_images_changed
        self.build_name_btn.Click        += self.build_name_click
        self.export_btn.Click            += self.export_click
        self.cancel_btn.Click            += self.cancel_click
        self.save_profile_btn.Click      += self.save_profile_click
        self.load_profile_btn.Click      += self.load_profile_click

    def processing_changed(self, sender, args):
        if self.raster_rb.IsChecked:
            self.raster_options_border.Visibility = framework.Windows.Visibility.Visible
            self.vector_options_border.Visibility = framework.Windows.Visibility.Collapsed
        else:
            self.vector_options_border.Visibility = framework.Windows.Visibility.Visible
            self.raster_options_border.Visibility = framework.Windows.Visibility.Collapsed

    def position_changed(self, sender, args):
        self.margins_cb.IsEnabled = self.position_offset_rb.IsChecked
        if not self.position_offset_rb.IsChecked:
            self.custom_margins_border.Visibility = framework.Windows.Visibility.Collapsed

    def margins_changed(self, sender, args):
        if self.margins_cb.SelectedItem:
            if self.margins_cb.SelectedItem.Tag == "custom":
                self.custom_margins_border.Visibility = framework.Windows.Visibility.Visible
            else:
                self.custom_margins_border.Visibility = framework.Windows.Visibility.Collapsed

    def bind_images_changed(self, sender, args):
        if self.bind_images_cb.IsChecked:
            self.bind_images_info_border.Visibility = framework.Windows.Visibility.Visible
            self.bind_images_tip_border.Visibility  = framework.Windows.Visibility.Collapsed
        else:
            self.bind_images_info_border.Visibility = framework.Windows.Visibility.Collapsed
            self.bind_images_tip_border.Visibility  = framework.Windows.Visibility.Visible

    def _get_separator(self):
        # No modo montado, o separador escolhido na janela de montagem tem prioridade
        if self.custom_fields and self.custom_separator is not None:
            return self.custom_separator
        if self.sep_underline_rb.IsChecked:
            return "_"
        if self.sep_ponto_rb.IsChecked:
            return "."
        return "-"

    def _atualizar_modo_nome(self):
        """Reflete na interface se o nome esta no modo simples ou no modo montado.

        No modo montado, prefixo e sufixo nao sao usados - o texto fixo entra
        como item da sequencia, na posicao que o usuario escolher.
        """
        montado = bool(self.custom_fields)

        self.include_number_cb.IsEnabled = not montado
        self.include_name_cb.IsEnabled   = not montado
        self.prefix_tb.IsEnabled         = not montado
        self.suffix_tb.IsEnabled         = not montado
        self.sep_hifen_rb.IsEnabled      = not montado
        self.sep_underline_rb.IsEnabled  = not montado
        self.sep_ponto_rb.IsEnabled      = not montado

        if montado:
            sep = self._get_separator()
            elo = u" {} ".format(sep) if self.use_separator and sep else u" + "
            self.nome_formula_tb.Text = elo.join(
                [rotulo_token(t) for t in self.custom_fields])
        else:
            self.nome_formula_tb.Text = u"(usando Numero + Nome padrao)"

        self.update_preview(None, None)

    def update_preview(self, sender, args):
        try:
            sheet = self.selected_sheets[0] if self.selected_sheets else None
            if sheet:
                nome = generate_filename(
                    sheet,
                    prefix=self.prefix_tb.Text.strip(),
                    suffix=self.suffix_tb.Text.strip(),
                    inc_num=self.include_number_cb.IsChecked,
                    inc_name=self.include_name_cb.IsChecked,
                    separator=self._get_separator(),
                    custom_fields=self.custom_fields if self.custom_fields else None,
                    use_separator=self.use_separator
                )
            else:
                nome = "ARQUIVO"
            self.preview_tb.Text = nome + ".pdf"
        except Exception as e:
            self.preview_tb.Text = "Erro: {}".format(e)

    def build_name_click(self, sender, args):
        try:
            sheet = self.selected_sheets[0] if self.selected_sheets else None
            if not sheet:
                forms.alert(u"Nenhuma prancha disponivel para montar o nome.",
                            title="Aviso")
                return

            win = BuildNameWindow(
                get_sheet_params(sheet),
                get_project_params(),
                self.custom_fields,
                sheet,
                separator=self._get_separator(),
                use_separator=self.use_separator
            )
            win.Owner = self
            result = win.ShowDialog()

            if result:
                self.custom_fields    = list(win.result_fields)
                self.custom_separator = win.get_separator() if self.custom_fields else None
                self.use_separator    = win.get_use_separator()
                self._atualizar_modo_nome()
        except Exception as e:
            forms.alert(u"Erro ao abrir janela: {}".format(e), title="Erro")

    def save_profile_click(self, sender, args):
        try:
            filepath = ProfileManager.save_profile(self.get_config())
            if filepath:
                forms.alert("Perfil salvo!\n\n{}".format(filepath), title="Perfil Salvo")
        except Exception as e:
            forms.alert("Erro ao salvar perfil:\n{}".format(str(e)), title="Erro")

    def load_profile_click(self, sender, args):
        try:
            config = ProfileManager.load_profile()
            if config:
                self.apply_config(config)
                forms.alert("Perfil carregado!", title="Perfil Carregado")
        except Exception as e:
            forms.alert("Erro ao carregar perfil:\n{}".format(str(e)), title="Erro")

    def _selecionar_por_tag(self, combo, tag):
        """Seleciona o ComboBoxItem cujo Tag corresponde. Ignora se nao achar."""
        if not tag:
            return
        try:
            for item in combo.Items:
                if item.Tag == tag:
                    combo.SelectedItem = item
                    return
        except Exception:
            pass

    def apply_config(self, config):
        try:
            self.export_pdf_cb.IsChecked        = config.get('export_pdf', True)
            self.export_dwg_cb.IsChecked        = config.get('export_dwg', False)
            self.export_dwf_cb.IsChecked        = config.get('export_dwf', False)
            self.combine_pdf_cb.IsChecked       = config.get('combine_pdf', False)
            self.combined_name_tb.Text          = config.get('combined_name', 'Conjunto')
            self.create_subfolders_cb.IsChecked = config.get('create_subfolders', True)
            self.prefix_tb.Text                 = config.get('file_prefix', '')
            self.suffix_tb.Text                 = config.get('file_suffix', '')
            self.include_number_cb.IsChecked    = config.get('include_number', True)
            self.include_name_cb.IsChecked      = config.get('include_name', True)

            # --- separador de campo do modo simples ---
            sep_salvo = config.get('separator', '-')
            if sep_salvo == '_':
                self.sep_underline_rb.IsChecked = True
            elif sep_salvo == '.':
                self.sep_ponto_rb.IsChecked = True
            else:
                self.sep_hifen_rb.IsChecked = True

            if config.get('position_center', True):
                self.position_center_rb.IsChecked = True
            else:
                self.position_offset_rb.IsChecked = True
            if config.get('zoom_fit', True):
                self.zoom_fit_rb.IsChecked = True
            else:
                self.zoom_100_rb.IsChecked = True
            if config.get('use_vector', True):
                self.vector_rb.IsChecked = True
            else:
                self.raster_rb.IsChecked = True

            # --- qualidade e cores ---
            quality_tag = config.get('quality_tag', 'high')
            if config.get('use_vector', True):
                self._selecionar_por_tag(self.vector_quality_cb, quality_tag)
            else:
                self._selecionar_por_tag(self.raster_quality_cb, quality_tag)
            self._selecionar_por_tag(self.raster_colors_cb, config.get('color_tag', 'color'))

            # --- margens ---
            margin_x = config.get('margin_x', 0)
            margin_y = config.get('margin_y', 0)
            self.margin_x_tb.Text = str(margin_x)
            self.margin_y_tb.Text = str(margin_y)
            if not config.get('position_center', True) and (margin_x or margin_y):
                self._selecionar_por_tag(self.margins_cb, 'custom')
            else:
                self._selecionar_por_tag(self.margins_cb, 'none')

            # --- configuracao DWG salva no projeto ---
            dwg_setup = config.get('dwg_setup')
            if dwg_setup and dwg_setup in self.dwg_settings_dict:
                self.dwg_setup_cb.SelectedItem = dwg_setup

            self.hide_ref_cb.IsChecked    = config.get('hide_ref_planes', True)
            self.hide_scope_cb.IsChecked  = config.get('hide_scope_boxes', True)
            self.hide_crop_cb.IsChecked   = config.get('hide_crop_boundaries', True)
            self.bind_images_cb.IsChecked = config.get('bind_images', False)

            # --- composicao montada do nome (era o que nao voltava do perfil) ---
            self.custom_fields = list(config.get('custom_fields') or [])
            self.use_separator = bool(config.get('use_separator', True))
            if self.custom_fields:
                # Perfis da v2.3.0 nao gravavam separador proprio do modo montado
                self.custom_separator = config.get('custom_separator',
                                                   config.get('separator', '-'))
            else:
                self.custom_separator = None

            # Garante que os paineis condicionais acompanhem o perfil carregado
            self.processing_changed(None, None)
            self.position_changed(None, None)
            self.margins_changed(None, None)
            self.bind_images_changed(None, None)
            self._atualizar_modo_nome()
        except Exception as e:
            logger.error("Erro ao aplicar config: %s", e)

    def export_click(self, sender, args):
        if not self.export_pdf_cb.IsChecked and \
           not self.export_dwg_cb.IsChecked and \
           not self.export_dwf_cb.IsChecked:
            forms.alert("Selecione ao menos um formato!", title="Erro")
            return
        if self.combine_pdf_cb.IsChecked and not self.combined_name_tb.Text.strip():
            forms.alert("Digite o nome para o PDF combinado!", title="Erro")
            return
        self.DialogResult = True
        self.Close()

    def cancel_click(self, sender, args):
        self.DialogResult = False
        self.Close()

    def get_config(self):
        margin_x, margin_y = 0, 0
        if self.position_offset_rb.IsChecked:
            if self.margins_cb.SelectedItem and self.margins_cb.SelectedItem.Tag == "custom":
                try:
                    margin_x = float(self.margin_x_tb.Text.replace(',', '.'))
                    margin_y = float(self.margin_y_tb.Text.replace(',', '.'))
                except:
                    pass

        use_vector  = self.vector_rb.IsChecked
        quality_tag = "high"
        color_tag   = "color"
        if use_vector:
            if self.vector_quality_cb.SelectedItem:
                quality_tag = self.vector_quality_cb.SelectedItem.Tag
        else:
            if self.raster_quality_cb.SelectedItem:
                quality_tag = self.raster_quality_cb.SelectedItem.Tag
            if self.raster_colors_cb.SelectedItem:
                color_tag = self.raster_colors_cb.SelectedItem.Tag

        return {
            'export_pdf':           self.export_pdf_cb.IsChecked,
            'export_dwg':           self.export_dwg_cb.IsChecked,
            'export_dwf':           self.export_dwf_cb.IsChecked,
            'combine_pdf':          self.combine_pdf_cb.IsChecked,
            'combined_name':        self.combined_name_tb.Text.strip(),
            'file_prefix':          self.prefix_tb.Text.strip(),
            'file_suffix':          self.suffix_tb.Text.strip(),
            'include_number':       self.include_number_cb.IsChecked,
            'include_name':         self.include_name_cb.IsChecked,
            'separator':            self._get_separator(),
            'custom_separator':     self.custom_separator,
            'use_separator':        self.use_separator,
            'custom_fields':        list(self.custom_fields),
            'hide_ref_planes':      self.hide_ref_cb.IsChecked,
            'hide_scope_boxes':     self.hide_scope_cb.IsChecked,
            'hide_crop_boundaries': self.hide_crop_cb.IsChecked,
            'zoom_fit':             self.zoom_fit_rb.IsChecked,
            'use_vector':           use_vector,
            'quality_tag':          quality_tag,
            'color_tag':            color_tag,
            'position_center':      self.position_center_rb.IsChecked,
            'margin_x':             margin_x,
            'margin_y':             margin_y,
            'dwg_setup':            self.dwg_setup_cb.SelectedItem,
            'create_subfolders':    self.create_subfolders_cb.IsChecked,
            'bind_images':          self.bind_images_cb.IsChecked,
        }


# ==========================================
# MAIN
# ==========================================

output.print_md("# ExportSheets P-LAB v2.4.0")
output.print_md("---")

tempo_inicio = time.time()

if IS_REVIT_2021_OR_OLDER:
    output.print_md("**REVIT {} DETECTADO**".format(REVIT_VERSION))
    output.print_md("- PDF via PrintManager | DWG/DWF normalmente")
    output.print_md("---")

all_sheets = DB.FilteredElementCollector(doc)\
    .OfClass(framework.get_type(DB.ViewSheet))\
    .WhereElementIsNotElementType()\
    .ToElements()

if not all_sheets:
    forms.alert("Nenhuma prancha no projeto!", exitscript=True)

selected_sheets = forms.select_sheets(
    title='ExportSheets P-LAB - Selecionar Pranchas',
    button_name='Selecionar',
    multiple=True
)

if not selected_sheets:
    script.exit()

output.print_md("## Pranchas selecionadas: {}".format(len(selected_sheets)))
for sh in selected_sheets:
    output.print_md("- {} - {}".format(sh.SheetNumber, sh.Name))
output.print_md("")

dwg_settings = DB.FilteredElementCollector(doc)\
    .OfClass(DB.ExportDWGSettings)\
    .ToElements()

dwg_settings_dict = {}
for s in dwg_settings:
    dwg_settings_dict[s.Name] = s

xaml_file = script.get_bundle_file('ExportSheetsWindow.xaml')
window    = ExportSheetsWindow(xaml_file, selected_sheets, dwg_settings_dict)
result    = window.ShowDialog()

if not result:
    output.print_md("[AVISO] Cancelado.")
    script.exit()

cfg    = window.get_config()
folder = forms.pick_folder()
if not folder:
    script.exit()

PrintUtils.ensure_dir(folder)


def nome_arquivo(sheet):
    """Nome da prancha conforme a configuracao escolhida na janela."""
    return generate_filename(
        sheet,
        cfg['file_prefix'],
        cfg['file_suffix'],
        cfg['include_number'],
        cfg['include_name'],
        separator=cfg.get('separator', '-'),
        custom_fields=cfg.get('custom_fields') or None,
        use_separator=cfg.get('use_separator', True)
    )

output.print_md("---")
output.print_md("## Configuracoes")
output.print_md("- Pasta: `{}`".format(folder))
output.print_md("- PDF: {} | DWG: {} | DWF: {}".format(
    'Sim' if cfg['export_pdf'] else 'Nao',
    'Sim' if cfg['export_dwg'] else 'Nao',
    'Sim' if cfg['export_dwf'] else 'Nao',
))
output.print_md("- Vincular imagens: {}".format('Sim' if cfg['bind_images'] else 'Nao'))
output.print_md("---")

# PDF Revit 2021
if IS_REVIT_2021_OR_OLDER and cfg['export_pdf']:
    output.print_md("## [PDF] Exportando (Revit 2021)")
    pdf_folder = PrintUtils.ensure_dir(op.join(folder, "PDF")) if cfg['create_subfolders'] else folder
    success = errors = 0
    with forms.ProgressBar(title='Exportando PDFs {value}/{max_value}', cancellable=True) as pb:
        for idx, sheet in enumerate(selected_sheets, 1):
            if pb.cancelled:
                break
            pb.update_progress(idx, len(selected_sheets))
            filename = nome_arquivo(sheet)
            try:
                PrintUtils.export_sheet_pdf(pdf_folder, sheet, None, doc, filename + ".pdf")
                output.print_md("[OK] `{}.pdf`".format(filename))
                success += 1
            except Exception as e:
                output.print_md("[ERRO] {}: {}".format(filename, str(e)))
                errors += 1
    output.print_md("**Resumo: {} OK, {} erros**".format(success, errors))

# Transaction
t = DB.Transaction(doc, "ExportSheets P-LAB")
t.Start()

try:
    # PDF 2022+
    if cfg['export_pdf'] and not IS_REVIT_2021_OR_OLDER:
        output.print_md("## [PDF] Exportando")
        pdf_folder = PrintUtils.ensure_dir(op.join(folder, "PDF")) if cfg['create_subfolders'] else folder
        output.print_md("Subpasta: `{}`".format(pdf_folder)) if cfg['create_subfolders'] else None

        pdf_options = PrintUtils.pdf_opts(
            hide_crop=cfg['hide_crop_boundaries'], hide_scope=cfg['hide_scope_boxes'],
            hide_ref=cfg['hide_ref_planes'], zoom_fit=cfg['zoom_fit'],
            use_vector=cfg['use_vector'], quality_tag=cfg['quality_tag'],
            color_tag=cfg['color_tag'], use_center=cfg['position_center'],
            margin_x=cfg['margin_x'], margin_y=cfg['margin_y']
        )

        if cfg['combine_pdf']:
            pdf_options.Combine  = True
            pdf_options.FileName = cfg['combined_name']
            sheet_ids = framework.List[DB.ElementId]()
            for sh in selected_sheets:
                sheet_ids.Add(sh.Id)
            try:
                doc.Export(pdf_folder, sheet_ids, pdf_options)
                output.print_md("[OK] `{}.pdf` ({} pranchas)".format(
                    cfg['combined_name'], len(selected_sheets)))
            except Exception as e:
                output.print_md("[ERRO] {}".format(str(e)))
        else:
            success = errors = 0
            with forms.ProgressBar(title='Exportando PDFs {value}/{max_value}', cancellable=True) as pb:
                for idx, sheet in enumerate(selected_sheets, 1):
                    if pb.cancelled:
                        break
                    pb.update_progress(idx, len(selected_sheets))
                    filename = nome_arquivo(sheet)
                    try:
                        PrintUtils.export_sheet_pdf(pdf_folder, sheet, pdf_options, doc, filename + ".pdf")
                        output.print_md("[OK] `{}.pdf`".format(filename))
                        success += 1
                    except Exception as e:
                        output.print_md("[ERRO] {}: {}".format(filename, str(e)))
                        errors += 1
            output.print_md("**Resumo: {} OK, {} erros**".format(success, errors))

    # DWG
    if cfg['export_dwg']:
        output.print_md("")
        output.print_md("## [DWG] Exportando")
        dwg_folder = PrintUtils.ensure_dir(op.join(folder, "DWG")) if cfg['create_subfolders'] else folder
        output.print_md("Subpasta: `{}`".format(dwg_folder)) if cfg['create_subfolders'] else None

        selected_setup = cfg['dwg_setup']
        if selected_setup in dwg_settings_dict:
            dwg_options = dwg_settings_dict[selected_setup].GetDWGExportOptions()
            output.print_md("Config: {}".format(selected_setup))
        else:
            dwg_options = PrintUtils.dwg_opts()
            output.print_md("Config: Padrao")

        success = errors = 0
        exported_dwg_paths = []  # rastreia apenas os DWGs que exportamos agora
        with forms.ProgressBar(title='Exportando DWGs {value}/{max_value}', cancellable=True) as pb:
            for idx, sheet in enumerate(selected_sheets, 1):
                if pb.cancelled:
                    break
                pb.update_progress(idx, len(selected_sheets))
                
                filename = nome_arquivo(sheet)

                try:
                    PrintUtils.export_sheet_dwg(dwg_folder, sheet, dwg_options, doc, filename + ".dwg")
                    output.print_md("[OK] `{}.dwg`".format(filename))
                    exported_dwg_paths.append(op.join(dwg_folder, filename + ".dwg"))
                    success += 1
                except Exception as e:
                    output.print_md("[ERRO] {}: {}".format(filename, str(e)))
                    errors += 1
        output.print_md("**Resumo: {} OK, {} erros**".format(success, errors))

        # Vincular imagens - passa lista exata dos DWGs exportados agora
        if cfg['bind_images'] and exported_dwg_paths:
            DWGPostProcessor.bind_xref_images_folder(exported_dwg_paths, output)

    # DWF
    if cfg['export_dwf']:
        output.print_md("")
        output.print_md("## [DWF] Exportando")
        dwf_folder = PrintUtils.ensure_dir(op.join(folder, "DWF")) if cfg['create_subfolders'] else folder
        dwf_options = PrintUtils.dwf_opts()
        success = errors = 0
        with forms.ProgressBar(title='Exportando DWFs {value}/{max_value}', cancellable=True) as pb:
            for idx, sheet in enumerate(selected_sheets, 1):
                if pb.cancelled:
                    break
                pb.update_progress(idx, len(selected_sheets))
                filename = nome_arquivo(sheet)
                try:
                    PrintUtils.export_sheet_dwf(dwf_folder, sheet, dwf_options, doc, filename + ".dwf")
                    output.print_md("[OK] `{}.dwf`".format(filename))
                    success += 1
                except Exception as e:
                    output.print_md("[ERRO] {}: {}".format(filename, str(e)))
                    errors += 1
        output.print_md("**Resumo: {} OK, {} erros**".format(success, errors))

    t.Commit()

except Exception as e:
    t.RollBack()
    output.print_md("[ERRO CRITICO] {}".format(str(e)))
    logger.error("Erro critico: %s", e)

# Tempo total
tempo_total = time.time() - tempo_inicio
minutos  = int(tempo_total // 60)
segundos = int(tempo_total % 60)

output.print_md("")
output.print_md("---")
output.print_md("# Concluido!")
output.print_md("**Tempo total: {} min {} seg**".format(minutos, segundos) if minutos > 0
                else "**Tempo total: {} seg**".format(segundos))

if forms.alert("Abrir pasta de destino?", yes=True, no=True):
    PrintUtils.open_dir(folder)
