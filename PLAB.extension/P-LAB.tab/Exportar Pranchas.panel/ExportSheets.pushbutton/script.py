# -*- coding: utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except
"""
================================================================
             EXPORTSHEETS P-LAB - VERSAO 2.3.0
================================================================

HISTORICO:
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
import tempfile

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
# CLASSE: DWGPostProcessor v2.3.0
# ==========================================

class DWGPostProcessor:
    """
    Pos-processamento DWG: embutir imagens raster como OLE via AutoCAD.

    Estrategia:
      1. Gera script LISP temporario
      2. AutoCAD (accoreconsole ou acad.exe) executa o LISP
      3. LISP varre blocos procurando AcDbRasterImage
      4. Para cada imagem: le posicao local via entget, entra no BEDIT,
         copia arquivo para Clipboard via PowerShell, cola com PASTECLIP,
         apaga a imagem original
      5. QSAVE + QUIT
    """

    # ------------------------------------------------------------------
    # Script LISP embutido — sera salvo em arquivo temporario
    # ------------------------------------------------------------------
    LISP_SCRIPT = u"""
; ============================================================
; embed-images.lsp  -  P-LAB Engenharia  v2.3.0
; Embute imagens raster como OLE dentro dos blocos do DWG
; ============================================================

; --- utilidade: copiar arquivo de imagem para o Clipboard via PowerShell ---
(defun plab-copy-to-clipboard (filepath / cmd)
  (setq cmd (strcat
    "powershell -WindowStyle Hidden -Command \\"Add-Type -AssemblyName System.Windows.Forms;"
    "[System.Windows.Forms.Clipboard]::SetImage("
    "[System.Drawing.Image]::FromFile('"
    (vl-string-subst "/" "\\\\" filepath)
    "'))\\""
  ))
  (startapp "cmd.exe" (strcat "/c " cmd))
  ; aguarda o PowerShell terminar de copiar
  (command "._delay" 1500)
)

; --- utilidade: retorna nome do bloco pai de uma entidade ---
(defun plab-block-name (ent / owner blkname)
  (setq owner (cdr (assoc 330 (entget ent))))
  (if owner
    (setq blkname (cdr (assoc 2 (entget owner))))
    (setq blkname nil)
  )
  blkname
)

; --- utilidade: verifica se string termina com extensao de imagem ---
(defun plab-is-image-file (fname)
  (or
    (wcmatch (strcase fname) "*.PNG")
    (wcmatch (strcase fname) "*.JPG")
    (wcmatch (strcase fname) "*.JPEG")
    (wcmatch (strcase fname) "*.BMP")
    (wcmatch (strcase fname) "*.TIF")
    (wcmatch (strcase fname) "*.TIFF")
  )
)

; --- utilidade: resolve caminho absoluto da imagem ---
; tenta o caminho original, depois relativo ao DWG
(defun plab-resolve-path (imgpath dwgdir / candidate)
  (cond
    ((findfile imgpath) imgpath)
    (t
      (setq candidate (strcat dwgdir (vl-filename-base imgpath)
                              "." (vl-filename-extension imgpath)))
      (if (findfile candidate) candidate nil)
    )
  )
)

; --- funcao principal ---
(defun plab-embed-images (/ dwgdir ss idx ent edata imgpath resolved
                             blkname inspt blknames processed)

  ; diretorio do DWG atual
  (setq dwgdir (vl-filename-directory (getvar "DWGNAME")))
  (if (not (= (substr dwgdir (strlen dwgdir)) "\\\\"))
    (setq dwgdir (strcat dwgdir "\\\\"))
  )

  (princ "\\n[PLAB] Iniciando embed de imagens...\\n")

  ; coleta todos os objetos AcDbRasterImage no DWG inteiro
  (setq ss (ssget "_X" '((0 . "IMAGE"))))

  (if (not ss)
    (progn
      (princ "\\n[PLAB] Nenhuma imagem raster encontrada.\\n")
      (exit)
    )
  )

  (princ (strcat "\\n[PLAB] " (itoa (sslength ss)) " imagem(ns) encontrada(s).\\n"))

  ; dicionario para agrupar imagens por bloco
  ; processamos bloco a bloco
  (setq blknames '())
  (setq idx 0)

  ; primeira passagem: coleta lista de (bloco . entidade) unicos por bloco
  (repeat (sslength ss)
    (setq ent (ssname ss idx))
    (setq blkname (plab-block-name ent))
    (if blkname
      (if (not (assoc blkname blknames))
        (setq blknames (cons (list blkname) blknames))
      )
    )
    (setq idx (1+ idx))
  )

  ; segunda passagem: para cada bloco, processa todas as suas imagens
  (foreach blkentry blknames
    (setq blkname (car blkentry))

    (princ (strcat "\\n[PLAB] Processando bloco: " blkname "\\n"))

    ; entra no BEDIT do bloco
    (command "._-bedit" blkname)
    (command)

    ; coleta imagens dentro deste bloco (no contexto do bedit)
    (setq processed '())
    (setq idx 0)

    (repeat (sslength ss)
      (setq ent (ssname ss idx))

      (if (and ent
               (equal (plab-block-name ent) blkname)
               (not (member (cdr (assoc 5 (entget ent))) processed)))
        (progn
          (setq edata  (entget ent))
          (setq inspt  (cdr (assoc 10 edata)))   ; ponto de insercao local
          (setq imgpath (cdr (assoc 1 edata)))   ; caminho do arquivo

          (princ (strcat "\\n[PLAB]   Imagem: " imgpath "\\n"))
          (princ (strcat "[PLAB]   Insercao local: "
                         (rtos (car inspt) 2 6) ", "
                         (rtos (cadr inspt) 2 6) "\\n"))

          ; resolve o caminho do arquivo
          (setq resolved (plab-resolve-path imgpath dwgdir))

          (if resolved
            (progn
              ; copia imagem para o Clipboard
              (princ "[PLAB]   Copiando para Clipboard...\\n")
              (plab-copy-to-clipboard resolved)

              ; cola como OLE no ponto de insercao da imagem original
              ; (canto inferior esquerdo = ponto 10 da raster image)
              (command "._pasteclip"
                       (list (car inspt) (cadr inspt))  ; ponto de insercao
                       ""                                ; scale = 1 (Enter)
                       ""                               ; rotation = 0 (Enter)
              )
              (princ "[PLAB]   OLE inserido.\\n")

              ; apaga a imagem raster original
              (entdel ent)
              (princ "[PLAB]   Imagem raster removida.\\n")

              ; registra como processado
              (setq processed
                    (cons (cdr (assoc 5 (entget ent))) processed))
            )
            (progn
              (princ (strcat "[PLAB]   AVISO: arquivo nao encontrado: " imgpath "\\n"))
            )
          )
        )
      )
      (setq idx (1+ idx))
    )

    ; fecha o BEDIT salvando
    (command "._bclose" "_y")
    (princ (strcat "\\n[PLAB] Bloco " blkname " salvo.\\n"))
  )

  (princ "\\n[PLAB] Embed concluido. Salvando DWG...\\n")
  (command "._qsave")
  (princ "\\n[PLAB] Concluido!\\n")
)

; executa ao carregar
(vl-load-com)
(plab-embed-images)
"""

    # ------------------------------------------------------------------
    # Script de inicializacao .scr — carrega o LISP e sai
    # ------------------------------------------------------------------
    SCR_TEMPLATE = u"""(load "{lisp_path}")
_quit
y
"""

    # ------------------------------------------------------------------
    # Localizacao do executavel AutoCAD
    # ------------------------------------------------------------------
    AUTOCAD_SEARCH_PATHS = [
        # Core Console — versoes comuns
        r"C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2023\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2022\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2021\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2020\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2019\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2018\accoreconsole.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2017\accoreconsole.exe",
        # AutoCAD completo como fallback
        r"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2025\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2024\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2023\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2022\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2021\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2020\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2019\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2018\acad.exe",
        r"C:\Program Files\Autodesk\AutoCAD 2017\acad.exe",
    ]

    @staticmethod
    def find_autocad():
        """Procura o executavel do AutoCAD nas pastas padrao."""
        import glob as _glob
        # Tenta caminhos diretos
        for path in DWGPostProcessor.AUTOCAD_SEARCH_PATHS:
            if op.exists(path):
                return path
        # Busca generica por glob
        for pattern in [
            r"C:\Program Files\Autodesk\AutoCAD*\accoreconsole.exe",
            r"C:\Program Files\Autodesk\AutoCAD*\acad.exe",
            r"C:\Program Files (x86)\Autodesk\AutoCAD*\accoreconsole.exe",
        ]:
            found = _glob.glob(pattern)
            if found:
                # prefere a versao mais recente (maior numero no nome)
                found.sort(reverse=True)
                return found[0]
        return None

    @staticmethod
    def bind_xref_images_folder(dwg_paths_or_folder, output=None):
        """Embute imagens raster como OLE nos DWGs indicados.
        
        Args:
            dwg_paths_or_folder: lista de caminhos .dwg OU string de pasta (fallback)
        """

        if output:
            output.print_md("")
            output.print_md("## [BIND] Embutindo Imagens nos DWGs")
            output.print_md("")

        # Localiza AutoCAD — PRECISA ser acad.exe (Clipboard requer GUI)
        # O accoreconsole nao tem acesso ao Clipboard do Windows
        acad_exe = DWGPostProcessor._find_acad_gui()
        if not acad_exe:
            if output:
                output.print_md("[ERRO] AutoCAD nao encontrado. Instale o AutoCAD 2017+ para usar esta funcao.")
            return 0, 1

        if output:
            output.print_md("AutoCAD encontrado: `{}`".format(acad_exe))
            output.print_md("Modo: acad.exe em segundo plano (necessario para Clipboard)")

        # Aceita lista de paths ou pasta
        if isinstance(dwg_paths_or_folder, list):
            dwg_files = [p for p in dwg_paths_or_folder if op.exists(p)]
        else:
            dwg_files = glob.glob(op.join(dwg_paths_or_folder, "*.dwg"))

        if not dwg_files:
            if output:
                output.print_md("[AVISO] Nenhum arquivo DWG para processar.")
            return 0, 0

        if output:
            output.print_md("Processando {} DWG(s)...".format(len(dwg_files)))

        ok = 0
        erro = 0

        for dwg_path in dwg_files:
            dwg_path = op.normpath(op.abspath(dwg_path))
            nome = op.basename(dwg_path)

            if output:
                output.print_md("")
                output.print_md("**{}**".format(nome))

            result = DWGPostProcessor._process_single_dwg(dwg_path, acad_exe, output)
            if result:
                ok += 1
                if output:
                    output.print_md("[OK] {}".format(nome))
            else:
                erro += 1
                if output:
                    output.print_md("[ERRO] {}".format(nome))

        if output:
            output.print_md("")
            output.print_md("**Resumo: {} OK, {} erros**".format(ok, erro))

        return ok, erro

    @staticmethod
    def _find_acad_gui():
        """Procura acad.exe (versao completa com GUI — necessaria para Clipboard)."""
        import glob as _glob
        # Tenta caminhos diretos, versoes mais recentes primeiro
        for path in DWGPostProcessor.AUTOCAD_SEARCH_PATHS:
            if "acad.exe" in path.lower() and op.exists(path):
                return path
        # Busca generica
        for pattern in [
            r"C:\Program Files\Autodesk\AutoCAD*\acad.exe",
            r"C:\Program Files (x86)\Autodesk\AutoCAD*\acad.exe",
        ]:
            found = _glob.glob(pattern)
            if found:
                found.sort(reverse=True)
                return found[0]
        return None

    @staticmethod
    def _process_single_dwg(dwg_path, acad_exe, output=None):
        """Processa um unico DWG: gera LISP + SCR temporarios, lanca acad.exe /b, aguarda."""

        lisp_path = None
        scr_path  = None
        log_path  = None

        try:
            # --- arquivo de log para debug (o AutoCAD nao tem stdout capturavel) ---
            log_path = dwg_path.replace('.dwg', '_plab_embed.log')

            # --- script LISP ---
            lisp_fd   = tempfile.NamedTemporaryFile(suffix='.lsp', delete=False)
            lisp_path = lisp_fd.name
            lisp_fd.close()

            lisp_path_fwd = lisp_path.replace('\\', '/')
            log_path_fwd  = log_path.replace('\\', '/')

            # Injeta o caminho do log no LISP para redirecionar princ
            lisp_content = DWGPostProcessor.LISP_SCRIPT.replace(
                "(princ ",
                "(plab-log "
            )
            # Adiciona funcao de log no inicio do LISP
            log_func = u"""
; --- log para arquivo ---
(defun plab-log (msg / f)
  (setq f (open "{log}" "a"))
  (if f (progn (write-line msg f) (close f)))
  (princ msg)
)
""".format(log=log_path_fwd)

            with codecs.open(lisp_path, 'w', encoding='utf-8') as f:
                f.write(log_func + DWGPostProcessor.LISP_SCRIPT)

            # --- script .scr ---
            # acad.exe /b executa o .scr apos carregar o DWG
            # O .scr carrega o LISP e depois fecha o AutoCAD
            scr_fd   = tempfile.NamedTemporaryFile(suffix='.scr', delete=False)
            scr_path = scr_fd.name
            scr_fd.close()

            scr_content = u'(load "{lisp}")\n_quit\ny\n'.format(lisp=lisp_path_fwd)

            with codecs.open(scr_path, 'w', encoding='utf-8') as f:
                f.write(scr_content)

            # --- limpa log anterior ---
            try:
                if op.exists(log_path):
                    os.remove(log_path)
            except:
                pass

            # --- lanca acad.exe em segundo plano ---
            # /b = batch script executado apos abertura do DWG
            # /nologo = sem splash screen
            cmd = '"{acad}" "{dwg}" /b "{scr}" /nologo'.format(
                acad=acad_exe,
                dwg=dwg_path,
                scr=scr_path
            )

            if output:
                output.print_md("  Lancando AutoCAD (segundo plano)...")
                output.print_md("  Cmd: `{}`".format(cmd[:120]))

            proc = subprocess.Popen(
                cmd,
                shell=True,
                creationflags=0x00000020  # DETACHED_PROCESS — nao bloqueia o Revit
            )

            # Aguarda o AutoCAD terminar (timeout 5 minutos por DWG)
            timeout = 300
            intervalo = 5
            elapsed = 0
            while elapsed < timeout:
                time.sleep(intervalo)
                elapsed += intervalo
                ret = proc.poll()
                if ret is not None:
                    break
                if output and elapsed % 30 == 0:
                    output.print_md("  Aguardando... {}s".format(elapsed))

            if proc.poll() is None:
                proc.kill()
                if output:
                    output.print_md("  [TIMEOUT] AutoCAD nao terminou em {}s — processo encerrado.".format(timeout))
                return False

            # --- le o log gerado pelo LISP ---
            if output:
                if op.exists(log_path):
                    try:
                        with codecs.open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                            for linha in f.readlines():
                                linha = linha.strip()
                                if linha:
                                    output.print_md("  " + linha)
                    except:
                        pass
                else:
                    output.print_md("  [AVISO] Log nao gerado — o LISP pode nao ter executado.")
                    output.print_md("  Returncode AutoCAD: {}".format(proc.returncode))

            # Sucesso se o log existe e contem "Concluido"
            if op.exists(log_path):
                try:
                    with codecs.open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                        conteudo = f.read()
                    return "Concluido" in conteudo or "concluido" in conteudo
                except:
                    pass

            return proc.returncode == 0

        except Exception as e:
            if output:
                output.print_md("  [EXCECAO] {}: {}".format(
                    op.basename(dwg_path), str(e)))
            return False

        finally:
            for f in [lisp_path, scr_path]:
                if f:
                    try:
                        os.remove(f)
                    except:
                        pass
            # Mantem o log para debug — apaga na proxima execucao (feito acima)




# ==========================================
# FUNCAO AUXILIAR
# ==========================================

def get_param_value(sheet, param_name):
    """Tenta ler o valor de um parametro da folha ou do projeto pelo nome."""
    # Tenta primeiro nos parametros da folha
    param = sheet.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == DB.StorageType.String:
            return param.AsString() or ""
        elif param.StorageType == DB.StorageType.Integer:
            return str(param.AsInteger())
        elif param.StorageType == DB.StorageType.Double:
            return str(param.AsDouble())
    # Tenta nos parametros de informacoes do projeto
    proj_info = doc.ProjectInformation
    param = proj_info.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == DB.StorageType.String:
            return param.AsString() or ""
        elif param.StorageType == DB.StorageType.Integer:
            return str(param.AsInteger())
    return ""


def get_available_params(sheet):
    """Retorna lista de parametros disponiveis (folha + projeto), sem duplicatas."""
    params = {}

    # Parametros built-in uteis da folha
    builtins = [
        ("Numero da Prancha",  "SheetNumber"),
        ("Nome da Prancha",    "Name"),
        ("Emitido para revisao", None),
    ]
    params["Numero da Prancha"] = sheet.SheetNumber or ""
    params["Nome da Prancha"]   = sheet.Name        or ""

    # Parametros de instancia da folha
    for p in sheet.Parameters:
        try:
            nome = p.Definition.Name
            if p.HasValue and nome not in params:
                if p.StorageType == DB.StorageType.String:
                    params[nome] = p.AsString() or ""
                elif p.StorageType == DB.StorageType.Integer:
                    params[nome] = str(p.AsInteger())
                elif p.StorageType == DB.StorageType.Double:
                    params[nome] = str(round(p.AsDouble(), 4))
        except:
            pass

    # Parametros de informacoes do projeto
    try:
        proj_info = doc.ProjectInformation
        for p in proj_info.Parameters:
            try:
                nome = "[Projeto] " + p.Definition.Name
                if p.HasValue and nome not in params:
                    if p.StorageType == DB.StorageType.String:
                        params[nome] = p.AsString() or ""
                    elif p.StorageType == DB.StorageType.Integer:
                        params[nome] = str(p.AsInteger())
            except:
                pass
    except:
        pass

    return params


def generate_filename(sheet, prefix="", suffix="", inc_num=True, inc_name=True,
                      replace_spaces=False, separator="-", custom_fields=None):
    """Gera nome do arquivo.
    
    Args:
        replace_spaces: Se True, substitui espacos por separador (para DWG/AutoCAD)
        separator:      Separador entre partes do nome ('-', '_', '.')
        custom_fields:  Lista de nomes de parametros para compor o nome (substitui inc_num/inc_name)
    """
    sep = separator if separator in ("-", "_", ".") else "-"

    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

    def limpar(s):
        for c in invalid_chars:
            s = s.replace(c, '_')
        if replace_spaces:
            s = s.replace(' ', sep)
        return s

    prefix = limpar(prefix)
    suffix = limpar(suffix)

    parts = []
    if prefix:
        parts.append(prefix)

    if custom_fields:
        # Modo avancado: usa lista de parametros escolhidos pelo usuario
        for campo in custom_fields:
            if campo == "Numero da Prancha":
                val = sheet.SheetNumber or "SEM_NUMERO"
            elif campo == "Nome da Prancha":
                val = sheet.Name or "SEM_NOME"
            elif campo.startswith("[Projeto] "):
                nome_real = campo[len("[Projeto] "):]
                try:
                    p = doc.ProjectInformation.LookupParameter(nome_real)
                    val = (p.AsString() or "") if (p and p.HasValue and p.StorageType == DB.StorageType.String) else ""
                except:
                    val = ""
            else:
                p = sheet.LookupParameter(campo)
                if p and p.HasValue:
                    if p.StorageType == DB.StorageType.String:
                        val = p.AsString() or ""
                    elif p.StorageType == DB.StorageType.Integer:
                        val = str(p.AsInteger())
                    else:
                        val = ""
                else:
                    val = ""
            val = limpar(val)
            if val:
                parts.append(val)
    else:
        # Modo simples: numero e nome
        if inc_num:
            parts.append(limpar(sheet.SheetNumber or "SEM_NUMERO"))
        if inc_name:
            parts.append(limpar(sheet.Name or "SEM_NOME"))

    if suffix:
        parts.append(suffix)

    return sep.join(parts) if parts else (limpar(sheet.SheetNumber) or "ARQUIVO")


# ==========================================
# CLASSE: BuildNameWindow (WPF inline)
# ==========================================

BUILD_NAME_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Montar Nome da Prancha"
        Width="560" Height="520"
        ShowInTaskbar="False"
        ResizeMode="CanResize"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F5">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Cabecalho -->
        <Border Grid.Row="0" Background="#1976D2" Padding="12,10">
            <TextBlock Text="Montar Composicao do Nome"
                      Foreground="White" FontSize="14" FontWeight="Bold"/>
        </Border>

        <!-- Conteudo -->
        <Grid Grid.Row="1" Margin="10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Titulos das colunas -->
            <TextBlock Grid.Row="0" Grid.Column="0"
                      Text="Parametros disponiveis"
                      FontWeight="Bold" FontSize="11" Margin="0,0,0,5"/>
            <TextBlock Grid.Row="0" Grid.Column="2"
                      Text="Ordem no nome (arraste para reordenar)"
                      FontWeight="Bold" FontSize="11" Margin="0,0,0,5"/>

            <!-- Lista de parametros -->
            <Border Grid.Row="1" Grid.Column="0"
                   BorderBrush="#BDBDBD" BorderThickness="1" CornerRadius="2">
                <ListBox x:Name="params_lb"
                        SelectionMode="Single"
                        FontSize="11"/>
            </Border>

            <!-- Botoes do meio -->
            <StackPanel Grid.Row="1" Grid.Column="1"
                       VerticalAlignment="Center" Margin="8,0">
                <Button x:Name="add_btn"    Content="Adicionar &#x25BA;" Width="100" Margin="0,4"/>
                <Button x:Name="remove_btn" Content="&#x25C4; Remover"   Width="100" Margin="0,4"/>
                <Separator Margin="0,8"/>
                <Button x:Name="up_btn"     Content="&#x25B2; Subir"     Width="100" Margin="0,4"/>
                <Button x:Name="down_btn"   Content="&#x25BC; Descer"    Width="100" Margin="0,4"/>
            </StackPanel>

            <!-- Lista de campos selecionados -->
            <Border Grid.Row="1" Grid.Column="2"
                   BorderBrush="#BDBDBD" BorderThickness="1" CornerRadius="2">
                <ListBox x:Name="selected_lb"
                        SelectionMode="Single"
                        FontSize="11"/>
            </Border>
        </Grid>

        <!-- Preview -->
        <Border Grid.Row="2"
               Margin="10,0,10,5"
               Padding="10"
               Background="#E8F5E9"
               BorderBrush="#4CAF50"
               BorderThickness="1"
               CornerRadius="2">
            <StackPanel>
                <TextBlock Text="Preview:" FontWeight="Bold" FontSize="10" Foreground="#2E7D32"/>
                <TextBlock x:Name="preview_tb"
                          FontFamily="Consolas"
                          FontSize="11"
                          Foreground="#1B5E20"
                          Margin="0,3,0,0"
                          TextWrapping="Wrap"/>
            </StackPanel>
        </Border>

        <!-- Botoes OK/Cancelar -->
        <Border Grid.Row="3"
               Background="White"
               BorderBrush="#BDBDBD"
               BorderThickness="0,1,0,0"
               Padding="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="ok_btn"
                       Content="OK"
                       Background="#4CAF50"
                       Foreground="White"
                       FontWeight="Bold"
                       Width="100"
                       Margin="5"/>
                <Button x:Name="cancel_btn"
                       Content="Cancelar"
                       Width="100"
                       Margin="5"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"""


class BuildNameWindow(forms.WPFWindow):
    """Janela para o usuario montar a composicao do nome da prancha."""

    def __init__(self, available_params, current_fields, sample_sheet, separator="-"):
        # Salvar XAML em arquivo temporario
        import tempfile, codecs
        tmp = tempfile.NamedTemporaryFile(suffix='.xaml', delete=False, mode='wb')
        tmp.write(BUILD_NAME_XAML.encode('utf-8'))
        tmp.close()
        self._xaml_tmp = tmp.name

        forms.WPFWindow.__init__(self, self._xaml_tmp)

        self.available_params = available_params   # dict {nome: valor_exemplo}
        self.sample_sheet     = sample_sheet
        self.separator        = separator
        self.result_fields    = list(current_fields) if current_fields else []

        self._populate()
        self._connect_events()
        self._update_preview()

    def _populate(self):
        # Lista de disponiveis (excluindo os ja selecionados)
        self.params_lb.ItemsSource = [
            k for k in sorted(self.available_params.keys())
            if k not in self.result_fields
        ]
        self.selected_lb.ItemsSource = list(self.result_fields)

    def _connect_events(self):
        self.add_btn.Click     += self._on_add
        self.remove_btn.Click  += self._on_remove
        self.up_btn.Click      += self._on_up
        self.down_btn.Click    += self._on_down
        self.ok_btn.Click      += self._on_ok
        self.cancel_btn.Click  += self._on_cancel
        self.params_lb.MouseDoubleClick   += self._on_add
        self.selected_lb.MouseDoubleClick += self._on_remove

    def _refresh_lists(self):
        self.params_lb.ItemsSource = [
            k for k in sorted(self.available_params.keys())
            if k not in self.result_fields
        ]
        self.selected_lb.ItemsSource = list(self.result_fields)
        self._update_preview()

    def _on_add(self, sender, args):
        sel = self.params_lb.SelectedItem
        if sel and sel not in self.result_fields:
            self.result_fields.append(sel)
            self._refresh_lists()

    def _on_remove(self, sender, args):
        sel = self.selected_lb.SelectedItem
        if sel in self.result_fields:
            self.result_fields.remove(sel)
            self._refresh_lists()

    def _on_up(self, sender, args):
        sel = self.selected_lb.SelectedItem
        if sel and sel in self.result_fields:
            idx = self.result_fields.index(sel)
            if idx > 0:
                self.result_fields[idx], self.result_fields[idx-1] = \
                    self.result_fields[idx-1], self.result_fields[idx]
                self._refresh_lists()
                self.selected_lb.SelectedIndex = idx - 1

    def _on_down(self, sender, args):
        sel = self.selected_lb.SelectedItem
        if sel and sel in self.result_fields:
            idx = self.result_fields.index(sel)
            if idx < len(self.result_fields) - 1:
                self.result_fields[idx], self.result_fields[idx+1] = \
                    self.result_fields[idx+1], self.result_fields[idx]
                self._refresh_lists()
                self.selected_lb.SelectedIndex = idx + 1

    def _update_preview(self):
        try:
            if self.result_fields and self.sample_sheet:
                nome = generate_filename(
                    self.sample_sheet,
                    separator=self.separator,
                    custom_fields=self.result_fields
                )
                self.preview_tb.Text = nome + ".pdf"
            elif self.sample_sheet:
                self.preview_tb.Text = "(nenhum campo selecionado)"
        except Exception as e:
            self.preview_tb.Text = "Erro: {}".format(e)

    def _on_ok(self, sender, args):
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()

    def cleanup(self):
        try:
            os.remove(self._xaml_tmp)
        except:
            pass


# ==========================================
# CLASSE: ExportSheetsWindow (WPF)
# ==========================================

class ExportSheetsWindow(forms.WPFWindow):

    def __init__(self, xaml_file, selected_sheets, dwg_settings_dict):
        forms.WPFWindow.__init__(self, xaml_file)
        self.selected_sheets   = selected_sheets
        self.dwg_settings_dict = dwg_settings_dict
        # Estado do construtor de nome
        self.custom_fields     = []   # [] = modo simples (numero+nome)
        self._setup_combos()
        self._connect_events()
        self.update_preview(None, None)

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
        if self.sep_underline_rb.IsChecked:
            return "_"
        if self.sep_ponto_rb.IsChecked:
            return "."
        return "-"

    def update_preview(self, sender, args):
        try:
            sheet = self.selected_sheets[0] if self.selected_sheets else None
            sep   = self._get_separator()
            if sheet:
                nome = generate_filename(
                    sheet,
                    prefix=self.prefix_tb.Text.strip(),
                    suffix=self.suffix_tb.Text.strip(),
                    inc_num=self.include_number_cb.IsChecked,
                    inc_name=self.include_name_cb.IsChecked,
                    separator=sep,
                    custom_fields=self.custom_fields if self.custom_fields else None
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
                forms.alert("Nenhuma prancha disponivel para montar o nome.", title="Aviso")
                return
            available = get_available_params(sheet)
            sep = self._get_separator()
            win = BuildNameWindow(available, self.custom_fields, sheet, separator=sep)
            win.Owner = self
            result = win.ShowDialog()
            win.cleanup()
            if result:
                self.custom_fields = list(win.result_fields)
                # Atualiza label de formula
                if self.custom_fields:
                    self.nome_formula_tb.Text = " {} ".format(sep).join(self.custom_fields)
                    # Desabilita checkboxes simples pois modo avancado esta ativo
                    self.include_number_cb.IsEnabled = False
                    self.include_name_cb.IsEnabled   = False
                else:
                    self.nome_formula_tb.Text = "(usando Numero + Nome padrao)"
                    self.include_number_cb.IsEnabled = True
                    self.include_name_cb.IsEnabled   = True
                self.update_preview(None, None)
        except Exception as e:
            forms.alert("Erro ao abrir janela: {}".format(e), title="Erro")

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
            self.hide_ref_cb.IsChecked    = config.get('hide_ref_planes', True)
            self.hide_scope_cb.IsChecked  = config.get('hide_scope_boxes', True)
            self.hide_crop_cb.IsChecked   = config.get('hide_crop_boundaries', True)
            self.bind_images_cb.IsChecked = config.get('bind_images', False)
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

output.print_md("# ExportSheets P-LAB v2.3.0")
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
            filename = generate_filename(sheet, cfg['file_prefix'], cfg['file_suffix'],
                                         cfg['include_number'], cfg['include_name'],
                                         separator=cfg.get('separator', '-'),
                                         custom_fields=cfg.get('custom_fields') or None)
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
                    filename = generate_filename(sheet, cfg['file_prefix'], cfg['file_suffix'],
                                                 cfg['include_number'], cfg['include_name'],
                                                 separator=cfg.get('separator', '-'),
                                                 custom_fields=cfg.get('custom_fields') or None)
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
                
                filename = generate_filename(
                    sheet,
                    cfg['file_prefix'],
                    cfg['file_suffix'],
                    cfg['include_number'],
                    cfg['include_name'],
                    replace_spaces=True,
                    separator=cfg.get('separator', '-'),
                    custom_fields=cfg.get('custom_fields') or None
                )
                
                try:
                    PrintUtils.export_sheet_dwg(dwg_folder, sheet, dwg_options, doc, filename + ".dwg")
                    output.print_md("[OK] `{}.dwg`".format(filename))
                    exported_dwg_paths.append(op.join(dwg_folder, filename + ".dwg"))
                    success += 1
                except Exception as e:
                    output.print_md("[ERRO] {}: {}".format(filename, str(e)))
                    errors += 1
        output.print_md("**Resumo: {} OK, {} erros**".format(success, errors))

        # Vincular imagens — passa lista exata dos DWGs exportados agora
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
                filename = generate_filename(sheet, cfg['file_prefix'], cfg['file_suffix'],
                                             cfg['include_number'], cfg['include_name'],
                                             separator=cfg.get('separator', '-'),
                                             custom_fields=cfg.get('custom_fields') or None)
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