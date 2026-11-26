using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using SnapIndenter.Core;

namespace SnapIndenter.AddIn
{
    /// <summary>VBE 上の整形コマンドの実装。</summary>
    internal sealed class FormatService
    {
        private const string Title = "SnapIndenter";

        private readonly VBE _vbe;
        private UndoBatch _lastBatch;

        public Settings Settings { get; private set; }

        public FormatService(VBE vbe)
        {
            _vbe = vbe;
            Settings = Settings.Load();
        }

        public event EventHandler SettingsChanged;

        // ---------------- コマンド ----------------

        /// <summary>複数行選択中なら選択範囲、そうでなければカーソル位置のプロシージャを整形する。</summary>
        public void FormatCurrent() => Guard(() =>
        {
            var pane = ActivePane();
            if (pane == null) return;
            var cm = pane.CodeModule;
            pane.GetSelection(out int sl, out int sc, out int el, out int ec);

            int from, to;
            if (el > sl)
            {
                from = sl;
                to = ec == 1 ? el - 1 : el;
            }
            else
            {
                (from, to) = ProcedureRange(cm, sl);
            }
            RunOnPane(pane, from, to);
        });

        public void FormatModule() => Guard(() =>
        {
            var pane = ActivePane();
            if (pane == null) return;
            RunOnPane(pane, 1, int.MaxValue);
        });

        public void FormatProject() => Guard(() =>
        {
            VBProject project = _vbe.ActiveVBProject;
            if (project == null)
            {
                Message("プロジェクトが選択されていません。", MessageBoxIcon.Information);
                return;
            }
            if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
            {
                Message("プロジェクトがロックされているため整形できません。", MessageBoxIcon.Warning);
                return;
            }

            CodePane activePane = null;
            try { activePane = _vbe.ActiveCodePane; } catch { }
            CodeModule activeModule = null;
            int sl = 0, sc = 0, el = 0, ec = 0, top = 0;
            if (activePane != null)
            {
                try
                {
                    activeModule = activePane.CodeModule;
                    activePane.GetSelection(out sl, out sc, out el, out ec);
                    top = activePane.TopLine;
                }
                catch { activePane = null; }
            }

            var batch = new UndoBatch();
            int modules = 0, changedLines = 0;
            var failed = new List<string>();
            foreach (VBComponent comp in project.VBComponents)
            {
                try
                {
                    var cm = comp.CodeModule;
                    if (cm == null || cm.CountOfLines == 0) continue;
                    var r = FormatRange(cm, 1, int.MaxValue, batch);
                    modules++;
                    changedLines += r.Changed;
                    if (activePane != null && activeModule != null && SameModule(cm, activeModule))
                        RestoreSelection(activePane, r, sl, sc, el, ec, top);
                }
                catch (Exception ex)
                {
                    Log.Error(ex);
                    failed.Add(SafeName(comp));
                }
            }
            Remember(batch);

            string msg = $"{modules} 個のモジュールを整形しました（変更 {changedLines} 行）。";
            if (failed.Count > 0) msg += "\n\n整形できなかったモジュール:\n" + string.Join("\n", failed);
            Message(msg, failed.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        });

        public void Undo() => Guard(() =>
        {
            var batch = _lastBatch;
            if (batch == null || batch.Items.Count == 0)
            {
                Message("元に戻せる整形はありません。", MessageBoxIcon.Information);
                return;
            }
            _lastBatch = null;

            int skipped = 0;
            for (int i = batch.Items.Count - 1; i >= 0; i--)
            {
                var item = batch.Items[i];
                try
                {
                    if (item.Module.CountOfLines >= item.Line && item.Module.Lines[item.Line, 1] == item.After)
                        item.Module.ReplaceLine(item.Line, item.Before);
                    else
                        skipped++;
                }
                catch
                {
                    skipped++;
                }
            }
            if (skipped > 0)
                Message($"整形後に編集された行など {skipped} 行は元に戻せませんでした。", MessageBoxIcon.Warning);
        });

        public void ShowSettings() => Guard(() =>
        {
            using (var form = new SettingsForm(Settings))
            {
                if (form.ShowDialog(Owner()) != DialogResult.OK) return;
                Settings = form.Result;
                Settings.Save();
                SettingsChanged?.Invoke(this, EventArgs.Empty);
            }
        });

        // ---------------- 実装 ----------------

        private void RunOnPane(CodePane pane, int from, int to)
        {
            var cm = pane.CodeModule;
            if (cm.CountOfLines == 0) return;
            pane.GetSelection(out int sl, out int sc, out int el, out int ec);
            int top = pane.TopLine;

            var batch = new UndoBatch();
            var r = FormatRange(cm, from, to, batch);            if (r.Changed > 0)
            {
                Remember(batch);
                RestoreSelection(pane, r, sl, sc, el, ec, top);
            }
        }

        private sealed class RangeResult
        {
            public string[] Before;
            public string[] After;
            public int Changed;
        }

        private RangeResult FormatRange(CodeModule cm, int from, int to, UndoBatch batch)
        {
            int n = cm.CountOfLines;
            var result = new RangeResult();
            if (n == 0) return result;

            string[] lines = VbaIndenter.SplitLines(cm.Lines[1, n]);
            if (lines.Length != n) Array.Resize(ref lines, n);
            for (int i = 0; i < n; i++) lines[i] = lines[i] ?? string.Empty;

            string[] formatted = VbaIndenter.Format(lines, Settings.Options);
            result.Before = lines;
            result.After = formatted;

            from = Math.Max(1, from);
            to = Math.Min(n, to);
            for (int line = from; line <= to; line++)
            {
                string before = lines[line - 1];
                string after = formatted[line - 1];
                if (before == after) continue;
                cm.ReplaceLine(line, after);
                string actual = cm.Lines[line, 1];
                batch.Items.Add(new UndoItem { Module = cm, Line = line, Before = before, After = actual });
                result.Changed++;
            }
            return result;
        }

        private static (int from, int to) ProcedureRange(CodeModule cm, int line)
        {
            try
            {
                string name = cm.get_ProcOfLine(line, out vbext_ProcKind kind);
                if (!string.IsNullOrEmpty(name))
                {
                    int start = cm.get_ProcStartLine(name, kind);
                    int count = cm.get_ProcCountLines(name, kind);
                    return (start, start + count - 1);
                }
            }
            catch (COMException)
            {
                // 宣言セクションなど
            }
            int decl = cm.CountOfDeclarationLines;
            return line <= decl ? (1, decl) : (line, line);
        }

        private static void RestoreSelection(CodePane pane, RangeResult r, int sl, int sc, int el, int ec, int top)
        {
            if (r.Changed == 0 || r.Before == null) return;
            try
            {
                int nsc = AdjustColumn(r, sl, sc);
                int nec = AdjustColumn(r, el, ec);
                pane.SetSelection(sl, nsc, el, nec);
                pane.TopLine = top;
            }
            catch
            {
                // 選択位置の復元は失敗しても問題ない
            }
        }

        private static int AdjustColumn(RangeResult r, int line, int col)
        {
            if (line < 1 || line > r.Before.Length) return col;
            int oldLead = Leading(r.Before[line - 1]);
            int newLead = Leading(r.After[line - 1]);
            if (col <= 1) return col;
            if (col - 1 < oldLead) return newLead + 1;
            return Math.Max(1, col + newLead - oldLead);
        }

        private static int Leading(string s)
        {
            int i = 0;
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t')) i++;
            return i;
        }

        private CodePane ActivePane()
        {
            CodePane pane = null;
            try { pane = _vbe.ActiveCodePane; } catch (Exception ex) { Log.Error(ex); }
            if (pane == null)
            {
                Log.Info("ActiveCodePane is null");
                Message("コード ウィンドウを開いてから実行してください。", MessageBoxIcon.Information);
                return null;
            }
            try
            {
                var project = pane.CodeModule.Parent.Collection.Parent;
                if (project.Protection == vbext_ProjectProtection.vbext_pp_locked)
                {
                    Message("プロジェクトがロックされているため整形できません。", MessageBoxIcon.Warning);
                    return null;
                }
            }
            catch { }
            return pane;
        }

        private void Remember(UndoBatch batch)
        {
            if (batch.Items.Count > 0) _lastBatch = batch;
        }

        private static bool SameModule(CodeModule a, CodeModule b)
        {
            try
            {
                return a.Parent.Name == b.Parent.Name
                    && a.Parent.Collection.Parent.Name == b.Parent.Collection.Parent.Name;
            }
            catch { return false; }
        }

        private static string SafeName(VBComponent comp)
        {
            try { return comp.Name; } catch { return "(不明)"; }
        }

        private void Guard(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
                Message("整形中にエラーが発生しました。\n\n" + ex.Message, MessageBoxIcon.Error);
            }
        }

        internal IWin32Window Owner()
        {
            try { return new WindowHandle(new IntPtr(_vbe.MainWindow.HWnd)); }
            catch { return null; }
        }

        private void Message(string text, MessageBoxIcon icon)
            => MessageBox.Show(Owner(), text, Title, MessageBoxButtons.OK, icon);

        private sealed class WindowHandle : IWin32Window
        {
            public WindowHandle(IntPtr h) { Handle = h; }
            public IntPtr Handle { get; }
        }

        private sealed class UndoItem
        {
            public CodeModule Module;
            public int Line;
            public string Before;
            public string After;
        }

        private sealed class UndoBatch
        {
            public readonly List<UndoItem> Items = new List<UndoItem>();
        }
    }
}
