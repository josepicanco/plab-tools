# -*- coding: utf-8 -*-
__title__ = 'Export\nSchedules'
__author__ = 'P-LAB'

import os
import shutil
import tempfile
import datetime
import csv
import io
import subprocess
import clr

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System.Windows.Forms')

import System
from System.Windows import Window, Visibility, MessageBox
from System.Windows.Controls import CheckBox as WpfCheckBox
from System.Windows.Forms import OpenFileDialog, FolderBrowserDialog, DialogResult
from System.Windows.Media import SolidColorBrush, Color, Brushes

from pyrevit import HOST_APP, DB
from pyrevit.framework import wpf

# -----------------------------------------------------------------------------
# FAILURE PREPROCESSOR
# -----------------------------------------------------------------------------
class _IgnoreAllFailures(DB.IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        for f in failuresAccessor.GetFailureMessages():
            try:    failuresAccessor.DeleteWarning(f)
            except: pass
        return DB.FailureProcessingResult.Continue

# -----------------------------------------------------------------------------
# MODELOS DE DADOS — sem INotifyPropertyChanged (nao funciona bem no IronPython)
# -----------------------------------------------------------------------------
class ScheduleItem(object):
    """ViewSchedule simples. O estado do checkbox e gerenciado pela UI."""
    def __init__(self, name, element_id):
        self.Name      = name
        self.ElementId = element_id
        self.Selecionada = False   # controlado pelo script, nao por binding


class ModeloItem(object):
    def __init__(self, caminho):
        self.Caminho = caminho
        self.Nome    = os.path.splitext(os.path.basename(caminho))[0]
        self.Display = os.path.basename(caminho)
        self.Tabelas = []

    def tabelas_selecionadas(self):
        return [t for t in self.Tabelas if t.Selecionada]


# -----------------------------------------------------------------------------
# JANELA PRINCIPAL
# -----------------------------------------------------------------------------
class ExportSchedulesUI(Window):

    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'ui.xaml')
        wpf.LoadComponent(self, xaml_path)

        self.app           = HOST_APP.app
        self.pasta_destino = ''
        self._modelos      = []
        self._modelo_atual = None
        self._log_lines    = []
        self._ultimo_chk_idx = None   # indice do ultimo checkbox clicado (Shift+Click)

        self._log("OK", "Pronto. Selecione um ou mais modelos Revit.")

    # -------------------------------------------------------------------------
    # HELPERS
    # -------------------------------------------------------------------------
    def _log(self, level, msg):
        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        line = "[{}] {:5s}  {}".format(ts, level, msg)
        self._log_lines.append(line)
        self.log_tb.Text = "\n".join(self._log_lines)
        self.log_scroll.ScrollToEnd()
        errs  = sum(1 for l in self._log_lines if "ERROR" in l)
        warns = sum(1 for l in self._log_lines if "WARN"  in l)
        self.log_summary_tb.Text = "{} linhas | {} warn | {} erro".format(
            len(self._log_lines), warns, errs)

    def _set_progress(self, value, label="", pct=""):
        self.progress_bar.Value = value
        self.progress_bar.Foreground = (
            Brushes.Green if value >= 100 else
            Brushes.Gray  if value == 0  else
            Brushes.DodgerBlue)
        self.progress_label_tb.Text = label
        self.progress_pct_tb.Text   = pct
        self._forcar_render()

    def _forcar_render(self):
        from System.Windows.Threading import DispatcherPriority
        self.Dispatcher.Invoke(DispatcherPriority.Background,
                               System.Action(lambda: None))

    def _atualizar_contador_modelos(self):
        total     = len(self._modelos)
        carregados = sum(1 for m in self._modelos if m.Tabelas)
        self.modelos_info_tb.Text = "{} modelo(s) | {} carregado(s)".format(
            total, carregados)
        self.carregar_todos_btn.IsEnabled = total > 0
        self.remover_modelo_btn.IsEnabled = total > 0

    def _atualizar_contador_tabelas(self):
        if not self._modelo_atual:
            self.selection_info_tb.Text = "Selecione um modelo para ver as tabelas."
            return
        total = len(self._modelo_atual.Tabelas)
        sel   = len(self._modelo_atual.tabelas_selecionadas())
        # Total geral selecionado em todos os modelos
        total_geral = sum(len(m.tabelas_selecionadas()) for m in self._modelos)
        self.selection_info_tb.Text = \
            "{}/{} neste modelo  |  {} total selecionadas".format(
                sel, total, total_geral)

    def _verificar_exportar_habilitado(self):
        tem_sel   = any(m.tabelas_selecionadas() for m in self._modelos)
        tem_pasta = bool(self.pasta_destino and os.path.isdir(self.pasta_destino))
        self.export_btn.IsEnabled = tem_sel and tem_pasta

    def _refresh_modelos_lv(self):
        self.modelos_lv.ItemsSource = None
        self.modelos_lv.ItemsSource = self._modelos

    # -------------------------------------------------------------------------
    # BROWSE MODELOS
    # -------------------------------------------------------------------------
    def btnBrowseModelo_Click(self, sender, args):
        dlg             = OpenFileDialog()
        dlg.Filter      = "Revit Files (*.rvt)|*.rvt"
        dlg.Title       = "Selecionar modelos Revit"
        dlg.Multiselect = True
        if dlg.ShowDialog() != DialogResult.OK:
            return
        novos = 0
        for path in dlg.FileNames:
            if any(m.Caminho == path for m in self._modelos):
                self._log("WARN", "Ja adicionado: {}".format(os.path.basename(path)))
                continue
            self._modelos.append(ModeloItem(path))
            novos += 1
            self._log("INFO", "Adicionado: {}".format(os.path.basename(path)))
        self._refresh_modelos_lv()
        self._atualizar_contador_modelos()
        if novos:
            self._log("OK", "{} modelo(s) adicionado(s).".format(novos))

    def btnRemoverModelo_Click(self, sender, args):
        sel = list(self.modelos_lv.SelectedItems)
        if not sel:
            return
        for item in sel:
            self._modelos = [m for m in self._modelos if m != item]
            self._log("INFO", "Removido: {}".format(item.Nome))
        if self._modelo_atual in sel:
            self._modelo_atual = None
            self.schedules_lv.ItemsSource = None
        self._refresh_modelos_lv()
        self._atualizar_contador_modelos()
        self._atualizar_contador_tabelas()
        self._verificar_exportar_habilitado()

    def btnBrowsePasta_Click(self, sender, args):
        dlg             = FolderBrowserDialog()
        dlg.Description = "Selecionar pasta de destino"
        if dlg.ShowDialog() == DialogResult.OK:
            self.pasta_destino       = dlg.SelectedPath
            self.output_path_tb.Text = dlg.SelectedPath
            self._log("INFO", "Pasta de saida: {}".format(dlg.SelectedPath))
            self._verificar_exportar_habilitado()

    # -------------------------------------------------------------------------
    # CARREGAR MODELOS
    # -------------------------------------------------------------------------
    def btnCarregarTodos_Click(self, sender, args):
        nao_carregados = [m for m in self._modelos if not m.Tabelas]
        if not nao_carregados:
            self._log("INFO", "Todos os modelos ja foram carregados.")
            return
        total = len(nao_carregados)
        for idx, modelo in enumerate(nao_carregados, start=1):
            pct = int((idx / float(total)) * 100)
            self._set_progress(pct,
                "Carregando {}/{} — {}".format(idx, total, modelo.Nome),
                "{}%".format(pct))
            self._carregar_modelo(modelo)
        self._set_progress(100, "Modelos carregados.", "100%")
        self._atualizar_contador_modelos()

    def _carregar_modelo(self, modelo):
        doc_temp  = None
        temp_path = None
        try:
            doc_temp, temp_path = self._abrir_documento(modelo.Caminho)
            collector = DB.FilteredElementCollector(doc_temp).OfClass(DB.ViewSchedule)
            tabelas = []
            for sv in collector:
                if sv.IsTemplate:
                    continue
                if hasattr(sv, 'IsTitleblockRevisionSchedule') \
                        and sv.IsTitleblockRevisionSchedule:
                    continue
                tabelas.append(ScheduleItem(sv.Name, sv.Id))
            modelo.Tabelas = tabelas
            self._log("OK", "'{}' — {} tabelas.".format(modelo.Nome, len(tabelas)))
        except Exception as ex:
            self._log("ERROR", "Falha '{}': {}".format(modelo.Nome, str(ex)))
        finally:
            if doc_temp:
                try:    doc_temp.Close(False)
                except: pass
            if temp_path and os.path.exists(temp_path):
                try:    os.remove(temp_path)
                except: pass
        self._refresh_modelos_lv()

    def _abrir_documento(self, caminho_rvt):
        pasta_temp   = tempfile.gettempdir()
        nome_base    = os.path.splitext(os.path.basename(caminho_rvt))[0]
        caminho_temp = os.path.join(pasta_temp,
                                    '__temp_sch_{}.rvt'.format(nome_base))
        shutil.copy2(caminho_rvt, caminho_temp)
        path_temp = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(caminho_temp)
        open_opts = DB.OpenOptions()
        open_opts.DetachFromCentralOption = \
            DB.DetachFromCentralOption.DetachAndPreserveWorksets
        try:
            fho = open_opts.GetFailureHandlingOptions()
            fho.SetFailuresPreprocessor(_IgnoreAllFailures())
            fho.SetClearAfterRollback(True)
            open_opts.SetFailureHandlingOptions(fho)
        except Exception:
            pass
        doc = self.app.OpenDocumentFile(path_temp, open_opts)
        try:
            self.Activate()
            self.Focus()
        except Exception:
            pass
        return doc, caminho_temp

    # -------------------------------------------------------------------------
    # SELECAO DE MODELO — popula lista de tabelas com checkboxes manuais
    # -------------------------------------------------------------------------
    def modelosLv_SelectionChanged(self, sender, args):
        sel = self.modelos_lv.SelectedItem
        if sel is None:
            self._modelo_atual = None
            self.schedules_lv.Items.Clear()
            self.select_all_btn.IsEnabled  = False
            self.select_none_btn.IsEnabled = False
            self._atualizar_contador_tabelas()
            return

        self._modelo_atual = sel
        self._ultimo_chk_idx = None   # resetar referencia de Shift+Click ao trocar modelo

        if not sel.Tabelas:
            self._log("INFO", "Carregando '{}'...".format(sel.Nome))
            self._set_progress(20, "Carregando {}...".format(sel.Nome), "")
            self._carregar_modelo(sel)
            self._set_progress(100, "Carregado.", "100%")
            self._atualizar_contador_modelos()

        self._popular_tabelas_lv(sel)
        total = len(sel.Tabelas)
        self.select_all_btn.IsEnabled  = total > 0
        self.select_none_btn.IsEnabled = total > 0
        self._atualizar_contador_tabelas()
        self._verificar_exportar_habilitado()

    def _popular_tabelas_lv(self, modelo):
        """
        Reconstroi o ListView de tabelas usando CheckBox + TextBlock criados
        por codigo — sem DataTemplate nem binding, evitando o bug do IronPython
        com INotifyPropertyChanged.
        """
        from System.Windows.Controls import (
            ListViewItem, CheckBox as WpfCheckBox, StackPanel, TextBlock,
            Orientation)
        from System.Windows import Thickness

        self.schedules_lv.Items.Clear()

        for tab in modelo.Tabelas:
            # Checkbox
            chk = WpfCheckBox()
            chk.IsChecked   = tab.Selecionada
            chk.Margin      = Thickness(4, 0, 8, 0)
            chk.VerticalAlignment = System.Windows.VerticalAlignment.Center
            # Guardar referencia ao ScheduleItem no Tag
            chk.Tag = tab
            chk.Click += self._chk_tabela_click

            # Texto
            txt = TextBlock()
            txt.Text     = tab.Name
            txt.FontSize = 11
            txt.VerticalAlignment = System.Windows.VerticalAlignment.Center

            # Linha horizontal
            sp = StackPanel()
            sp.Orientation = Orientation.Horizontal
            sp.Children.Add(chk)
            sp.Children.Add(txt)

            lvi = ListViewItem()
            lvi.Content = sp
            lvi.Tag     = tab
            self._aplicar_cor_linha(lvi, tab.Selecionada)
            self.schedules_lv.Items.Add(lvi)

    def _aplicar_cor_linha(self, lvi, selecionada):
        from System.Windows.Media import SolidColorBrush, Color
        if selecionada:
            lvi.Background = SolidColorBrush(Color.FromRgb(13, 71, 161))   # #0D47A1
            lvi.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        else:
            lvi.Background = SolidColorBrush(Color.FromRgb(255, 255, 255))
            lvi.Foreground = SolidColorBrush(Color.FromRgb(33, 33, 33))

    def _chk_tabela_click(self, sender, args):
        """
        Clique no checkbox com suporte a Shift+Click para selecionar intervalo.
        """
        from System.Windows.Input import Keyboard, ModifierKeys

        chk = sender
        tab = chk.Tag

        # Descobrir o indice deste item na lista
        idx_atual = None
        items = list(self.schedules_lv.Items)
        for i, lvi in enumerate(items):
            if lvi.Tag is tab:
                idx_atual = i
                break

        if idx_atual is None:
            return

        shift_pressionado = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift

        if shift_pressionado and self._ultimo_chk_idx is not None:
            # Selecionar/desselecionar intervalo entre ultimo clique e atual
            # O estado alvo é o estado atual do checkbox clicado
            estado_alvo = (chk.IsChecked == True)
            inicio = min(self._ultimo_chk_idx, idx_atual)
            fim    = max(self._ultimo_chk_idx, idx_atual)
            for i in range(inicio, fim + 1):
                lvi_i  = items[i]
                tab_i  = lvi_i.Tag
                sp_i   = lvi_i.Content
                chk_i  = sp_i.Children[0]
                tab_i.Selecionada = estado_alvo
                chk_i.IsChecked   = estado_alvo
                self._aplicar_cor_linha(lvi_i, estado_alvo)
        else:
            # Clique simples — apenas este item
            tab.Selecionada = (chk.IsChecked == True)
            parent = chk.Parent
            if parent is not None:
                lvi = parent.Parent
                if lvi is not None:
                    self._aplicar_cor_linha(lvi, tab.Selecionada)

        self._ultimo_chk_idx = idx_atual
        self._atualizar_contador_tabelas()
        self._verificar_exportar_habilitado()

    # Manter compatibilidade — click na linha tb nao e necessario agora
    def schedulesLv_MouseLeftButtonUp(self, sender, args):
        pass

    def chkTabela_Click(self, sender, args):
        pass

    def btnSelecionarTodas_Click(self, sender, args):
        if not self._modelo_atual:
            return
        for tab in self._modelo_atual.Tabelas:
            tab.Selecionada = True
        # Atualizar checkboxes e cores de todas as linhas visiveis
        for lvi in self.schedules_lv.Items:
            tab = lvi.Tag
            sp  = lvi.Content
            chk = sp.Children[0]
            chk.IsChecked = True
            self._aplicar_cor_linha(lvi, True)
        self._atualizar_contador_tabelas()
        self._verificar_exportar_habilitado()

    def btnLimparSelecao_Click(self, sender, args):
        if not self._modelo_atual:
            return
        for tab in self._modelo_atual.Tabelas:
            tab.Selecionada = False
        for lvi in self.schedules_lv.Items:
            tab = lvi.Tag
            sp  = lvi.Content
            chk = sp.Children[0]
            chk.IsChecked = False
            self._aplicar_cor_linha(lvi, False)
        self._atualizar_contador_tabelas()
        self._verificar_exportar_habilitado()

    # -------------------------------------------------------------------------
    # EXPORTAR
    # -------------------------------------------------------------------------
    def btnExportar_Click(self, sender, args):
        modelos_para_exportar = [
            m for m in self._modelos if m.tabelas_selecionadas()
        ]
        if not modelos_para_exportar:
            self._log("WARN", "Nenhuma tabela selecionada.")
            return
        if not self.pasta_destino or not os.path.isdir(self.pasta_destino):
            self._log("ERROR", "Pasta de saida invalida.")
            return

        self.export_btn.IsEnabled = False
        subfolder = self.output_filename_tb.Text.strip() or "Tabelas_Revit"
        out_dir   = os.path.join(self.pasta_destino, subfolder)
        if not os.path.exists(out_dir):
            os.makedirs(out_dir)

        arquivos_gerados  = []
        total_modelos     = len(modelos_para_exportar)

        try:
            for idx_m, modelo in enumerate(modelos_para_exportar, start=1):
                self._log("INFO", "=== {}/{}: {} ===".format(
                    idx_m, total_modelos, modelo.Nome))
                pct_base = int(((idx_m - 1) / float(total_modelos)) * 85)

                doc_temp  = None
                temp_path = None
                dados     = []

                try:
                    self._set_progress(pct_base + 2,
                        "Abrindo {} ({}/{})...".format(
                            modelo.Nome, idx_m, total_modelos), "")
                    doc_temp, temp_path = self._abrir_documento(modelo.Caminho)

                    selecionadas = modelo.tabelas_selecionadas()
                    total_tab    = len(selecionadas)

                    for idx_t, item in enumerate(selecionadas, start=1):
                        pct = pct_base + int((idx_t / float(total_tab)) * 25)
                        self._set_progress(pct,
                            "Lendo {}/{} — {}".format(idx_t, total_tab, item.Name), "")
                        try:
                            sched  = doc_temp.GetElement(item.ElementId)
                            linhas = self._read_schedule(sched)
                            if (self.skip_empty_cb.IsChecked == True) \
                                    and len(linhas) <= 1:
                                self._log("WARN", "Vazia: {}".format(item.Name))
                                continue
                            dados.append((item.Name, linhas))
                            self._log("INFO", "Lida: {} — {} linhas".format(
                                item.Name, len(linhas)))
                        except Exception as ex:
                            self._log("ERROR", "Falha '{}': {}".format(
                                item.Name, str(ex)))
                finally:
                    if doc_temp:
                        try:    doc_temp.Close(False)
                        except: pass
                    if temp_path and os.path.exists(temp_path):
                        try:    os.remove(temp_path)
                        except: pass

                if not dados:
                    self._log("WARN", "Sem dados: '{}'.".format(modelo.Nome))
                    continue

                if self.log_csv_cb.IsChecked == True:
                    self._salvar_csvs_backup(dados, out_dir, modelo.Nome)

                nome_safe = modelo.Nome
                for ch in "/\\?*:[]<>|":
                    nome_safe = nome_safe.replace(ch, "-")
                xlsx_path = os.path.join(out_dir, nome_safe + ".xlsx")

                self._set_progress(pct_base + 30,
                    "Gerando Excel — {}...".format(modelo.Nome), "")
                self._exportar_excel_modelo(dados, xlsx_path)
                arquivos_gerados.append(xlsx_path)
                self._log("OK", "Excel: {}".format(os.path.basename(xlsx_path)))

            if (self.mesclar_todos_cb.IsChecked == True) \
                    and len(arquivos_gerados) > 1:
                self._set_progress(92, "Mesclando todos os Excel...", "92%")
                self._mesclar_excels(arquivos_gerados, out_dir, subfolder)

            self._set_progress(100, "Exportacao concluida!", "100%")
            self._log("OK", "{} arquivo(s) gerado(s).".format(len(arquivos_gerados)))

            MessageBox.Show(
                "Exportacao concluida!\n\n{} arquivo(s) Excel gerado(s)\n\nPasta:\n{}".format(
                    len(arquivos_gerados), out_dir),
                "Concluido")

            if self.abrir_pasta_cb.IsChecked == True and os.path.exists(out_dir):
                subprocess.Popen(['explorer', out_dir])
            if self.abrir_arquivo_cb.IsChecked == True and arquivos_gerados:
                consolidado = os.path.join(out_dir, subfolder + ".xlsx")
                alvo = consolidado if os.path.exists(consolidado) \
                    else arquivos_gerados[0]
                subprocess.Popen(['start', '', alvo], shell=True)

        except Exception as ex:
            self._log("ERROR", "Erro geral: {}".format(str(ex)))
            self.progress_bar.Foreground = Brushes.Red
            self._set_progress(0, "Erro. Veja o log.", "")
        finally:
            self.export_btn.IsEnabled = True

    # -------------------------------------------------------------------------
    # LER SCHEDULE
    # -------------------------------------------------------------------------
    def _read_schedule(self, sched):
        rows = []
        try:
            td = sched.GetTableData()
            for stype in [DB.SectionType.Header, DB.SectionType.Body,
                          DB.SectionType.Footer, DB.SectionType.Summary]:
                try:
                    sec = td.GetSectionData(stype)
                    if sec is None or sec.NumberOfRows == 0:
                        continue
                    for r in range(sec.NumberOfRows):
                        row = []
                        for c in range(sec.NumberOfColumns):
                            try:    val = sched.GetCellText(stype, r, c)
                            except: val = ""
                            row.append(val or "")
                        rows.append(row)
                except Exception:
                    continue
        except Exception as ex:
            rows.append(["ERRO: {}".format(str(ex))])
        return rows

    # -------------------------------------------------------------------------
    # BACKUP CSV
    # -------------------------------------------------------------------------
    def _salvar_csvs_backup(self, dados, out_dir, nome_modelo):
        nome_safe = nome_modelo
        for ch in "/\\?*:[]<>|":
            nome_safe = nome_safe.replace(ch, "-")
        backup_dir = os.path.join(out_dir, "backup_csv_{}".format(nome_safe))
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        for nome, linhas in dados:
            try:
                arq = nome
                for ch in "/\\?*:[]<>|":
                    arq = arq.replace(ch, "-")
                with open(os.path.join(backup_dir, arq + ".csv"), "wb") as f:
                    f.write("\xef\xbb\xbf")
                    w = csv.writer(f, delimiter=";")
                    for row in linhas:
                        w.writerow([(c.encode("utf-8") if c else "") for c in row])
                self._log("OK", "CSV: {}.csv".format(arq))
            except Exception as ex:
                self._log("ERROR", "CSV '{}': {}".format(nome, str(ex)))

    # -------------------------------------------------------------------------
    # EXCEL — modelo individual
    # -------------------------------------------------------------------------
    def _exportar_excel_modelo(self, dados, xlsx_path):
        import System.Runtime.InteropServices as Interop

        fazer_consolidado = (self.aba_consolidado_cb.IsChecked == True)
        fazer_separadas   = (self.abas_separadas_cb.IsChecked  == True)

        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise Exception("Excel nao encontrado.")

        excel = System.Activator.CreateInstance(excel_type)
        excel.Visible = excel.DisplayAlerts = excel.ScreenUpdating = False
        try:
            wb = excel.Workbooks.Add()
            while wb.Sheets.Count > 1:
                wb.Sheets(wb.Sheets.Count).Delete()
            if fazer_consolidado:
                self._escrever_consolidado(wb, dados)
            if fazer_separadas:
                for nome, linhas in dados:
                    self._escrever_aba(wb, nome, linhas)
            wb.SaveAs(xlsx_path, 51)
            wb.Close(False)
        finally:
            try:    excel.Quit()
            except: pass
            try:    Interop.Marshal.ReleaseComObject(excel)
            except: pass

    def _escrever_consolidado(self, wb, dados):
        ws = wb.Sheets(1)
        ws.Name = "CONSOLIDADO"
        linha = 1
        for nome, linhas in dados:
            if not linhas:
                continue
            n_cols = max(len(r) for r in linhas)
            ws.Cells(linha, 1).Value2 = nome
            tr = ws.Range(ws.Cells(linha, 1), ws.Cells(linha, max(n_cols, 1)))
            tr.Merge()
            tr.Font.Bold = True; tr.Font.Size = 11
            tr.Interior.Color = 0x503C2C; tr.Font.Color = 0xFFFFFF
            tr.HorizontalAlignment = -4108
            linha += 1
            for c_idx, val in enumerate(linhas[0], 1):
                ws.Cells(linha, c_idx).Value2 = val
            hr = ws.Range(ws.Cells(linha, 1), ws.Cells(linha, len(linhas[0])))
            hr.Font.Bold = True; hr.Interior.Color = 0xD0E4F7
            hr.HorizontalAlignment = -4108
            linha += 1
            for row in linhas[1:]:
                for c_idx, val in enumerate(row, 1):
                    ws.Cells(linha, c_idx).Value2 = val
                linha += 1
            linha += 1
        ws.Columns.AutoFit()

    def _escrever_aba(self, wb, nome, linhas):
        if not linhas:
            return
        aba = nome
        for ch in "/\\?*:[]":
            aba = aba.replace(ch, "-")
        base = aba[:27]; final = base; n = 2
        existentes = [wb.Sheets(k).Name for k in range(1, wb.Sheets.Count + 1)]
        while final in existentes:
            suf = "({})".format(n); final = base[:31-len(suf)] + suf; n += 1
        ws = wb.Sheets.Add(System.Reflection.Missing.Value,
                           wb.Sheets(wb.Sheets.Count))
        ws.Name = final
        for r_idx, row in enumerate(linhas, 1):
            for c_idx, val in enumerate(row, 1):
                ws.Cells(r_idx, c_idx).Value2 = val
        if linhas:
            hr = ws.Range(ws.Cells(1,1), ws.Cells(1, len(linhas[0])))
            hr.Font.Bold = True; hr.Interior.Color = 0xD0E4F7
            hr.HorizontalAlignment = -4108
        ws.Columns.AutoFit()

    # -------------------------------------------------------------------------
    # MESCLAR EXCELS
    # -------------------------------------------------------------------------
    def _mesclar_excels(self, arquivos, out_dir, nome_pasta):
        import System.Runtime.InteropServices as Interop
        nome_final = nome_pasta
        for ch in "/\\?*:[]<>|":
            nome_final = nome_final.replace(ch, "-")
        xlsx_final = os.path.join(out_dir, nome_final + ".xlsx")

        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise Exception("Excel nao encontrado para mesclagem.")

        excel = System.Activator.CreateInstance(excel_type)
        excel.Visible = excel.DisplayAlerts = excel.ScreenUpdating = False
        self._log("INFO", "Mesclando {} arquivo(s)...".format(len(arquivos)))
        try:
            wb_dest = excel.Workbooks.Add()
            while wb_dest.Sheets.Count > 1:
                wb_dest.Sheets(wb_dest.Sheets.Count).Delete()
            for xlsx_path in arquivos:
                nome_modelo = os.path.splitext(os.path.basename(xlsx_path))[0]
                try:
                    wb_src = excel.Workbooks.Open(xlsx_path)
                    for k in range(1, wb_src.Sheets.Count + 1):
                        ws_src   = wb_src.Sheets(k)
                        nome_aba = ws_src.Name   # ex: "CONSOLIDADO" ou nome da tabela

                        # Nome final = NomeModelo_NomeAba, truncado para 31 chars
                        # Reservar espaco suficiente para o sufixo de unicidade "(N)"
                        candidato = "{}_{}".format(nome_modelo, nome_aba)
                        candidato = candidato[:27]   # 27 + len("(99)") = 31

                        # Garantir unicidade
                        n = 2
                        existentes = [wb_dest.Sheets(j).Name
                                      for j in range(1, wb_dest.Sheets.Count + 1)]
                        aba_final = candidato
                        while aba_final in existentes:
                            suf       = "({})".format(n)
                            aba_final = candidato[:31 - len(suf)] + suf
                            n += 1

                        ws_src.Copy(System.Reflection.Missing.Value,
                                    wb_dest.Sheets(wb_dest.Sheets.Count))
                        wb_dest.Sheets(wb_dest.Sheets.Count).Name = aba_final
                    wb_src.Close(False)
                except Exception as ex:
                    self._log("ERROR", "Falha mesclar '{}': {}".format(
                        nome_modelo, str(ex)))
            if wb_dest.Sheets.Count > 1:
                wb_dest.Sheets(1).Delete()
            wb_dest.SaveAs(xlsx_final, 51)
            wb_dest.Close(False)
            self._log("OK", "Consolidado: {}".format(os.path.basename(xlsx_final)))
        finally:
            try:    excel.Quit()
            except: pass
            try:    Interop.Marshal.ReleaseComObject(excel)
            except: pass

    # -------------------------------------------------------------------------
    # LOG
    # -------------------------------------------------------------------------
    def btnToggleLog_Click(self, sender, args):
        if self.log_content_border.Visibility == Visibility.Visible:
            self.log_content_border.Visibility = Visibility.Collapsed
            self.log_toggle_tb.Text = ">  Log"
        else:
            self.log_content_border.Visibility = Visibility.Visible
            self.log_toggle_tb.Text = "v  Log"

    def btnCopyLog_Click(self, sender, args):
        from System.Windows import Clipboard
        Clipboard.SetText("\n".join(self._log_lines))
        self._log("INFO", "Log copiado.")

    def btnClearLog_Click(self, sender, args):
        self._log_lines = []
        self.log_tb.Text = self.log_summary_tb.Text = ""

    def btnFechar_Click(self, sender, args):
        self.Close()


# -----------------------------------------------------------------------------
# ENTRY POINT
# -----------------------------------------------------------------------------
ui = ExportSchedulesUI()
ui.ShowDialog()