using System;
using System.Drawing;
using System.Windows.Forms;
using VbaFormatter.Core;

namespace VbaFormatter.AddIn
{
    /// <summary>設定ダイアログ（プレビュー付き）。</summary>
    internal sealed class SettingsForm : Form
    {
        private const string Sample =
@"Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal ms As LongPtr)
#End If
Private Type Point
X As Long
Y As Long
End Type
Public Sub Sample(ByVal n As Long)
' コメント
Dim i As Long
On Error GoTo ErrHandler
For i = 1 To n
Select Case i
Case 1
Debug.Print ""one""
Case Else
If i Mod 2 = 0 Then Debug.Print ""even""
End Select
Next i
MsgBox ""長いメッセージ"" & _
""を連結"", vbOKOnly
Exit Sub
ErrHandler:
MsgBox Err.Description
End Sub";

        private readonly NumericUpDown _indentSize = new NumericUpDown { Minimum = 1, Maximum = 16, Width = 60 };
        private readonly CheckBox _useTabs = Check("インデントにタブを使う");
        private readonly CheckBox _procBody = Check("プロシージャの本体をインデントする");
        private readonly CheckBox _indentCase = Check("Select Case 内の Case をインデントする");
        private readonly CheckBox _typeEnum = Check("Type / Enum のメンバーをインデントする");
        private readonly CheckBox _directives = Check("#If ～ #End If の中をインデントする（オフ: 指令は 1 桁目）");
        private readonly CheckBox _comments = Check("コメント行もインデントする");
        private readonly CheckBox _contAlign = Check("継続行（ _）の相対位置を保持する");
        private readonly NumericUpDown _contLevels = new NumericUpDown { Minimum = 0, Maximum = 8, Width = 60 };
        private readonly TextBox _hkCurrent = new TextBox { Width = 160 };
        private readonly TextBox _hkModule = new TextBox { Width = 160 };
        private readonly TextBox _preview = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font("MS Gothic", 10f),
            Dock = DockStyle.Fill,
            BackColor = SystemColors.Window,
        };

        public Settings Result { get; private set; }

        public SettingsForm(Settings current)
        {
            Text = "VBA Formatter の設定";
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Yu Gothic UI", 9f);
            ClientSize = new Size(900, 560);
            MinimumSize = new Size(760, 480);

            var o = current.Options;
            _indentSize.Value = o.IndentSize;
            _useTabs.Checked = o.UseTabs;
            _procBody.Checked = o.IndentProcedureBody;
            _indentCase.Checked = o.IndentCase;
            _typeEnum.Checked = o.IndentTypeAndEnumMembers;
            _directives.Checked = o.IndentCompilerDirectives;
            _comments.Checked = o.IndentCommentLines;
            _contAlign.Checked = o.PreserveContinuationAlignment;
            _contLevels.Value = o.ContinuationIndentLevels;
            _hkCurrent.Text = current.FormatCurrentHotkey;
            _hkModule.Text = current.FormatModuleHotkey;

            var left = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(8),
            };
            left.Controls.Add(Row("インデント幅:", _indentSize));
            left.Controls.Add(_useTabs);
            left.Controls.Add(_procBody);
            left.Controls.Add(_indentCase);
            left.Controls.Add(_typeEnum);
            left.Controls.Add(_directives);
            left.Controls.Add(_comments);
            left.Controls.Add(_contAlign);
            left.Controls.Add(Row("継続行の追加インデント（段）:", _contLevels));
            left.Controls.Add(new Label { Text = "ショートカットキー（例: Ctrl+Shift+F、空欄で無効）", AutoSize = true, Margin = new Padding(3, 16, 3, 3) });
            left.Controls.Add(Row("プロシージャ/選択範囲:", _hkCurrent));
            left.Controls.Add(Row("モジュール:", _hkModule));
            left.Controls.Add(new Label
            {
                Text = "設定ファイル: " + Settings.FilePath,
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 16, 3, 3),
                MaximumSize = new Size(380, 0),
            });

            var ok = new Button { Text = "OK", DialogResult = DialogResult.None, Width = 90 };
            var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Width = 90 };
            var reset = new Button { Text = "既定値に戻す", Width = 110 };
            ok.Click += (s, e) => OnOk();
            reset.Click += (s, e) => LoadOptions(new FormatterOptions());
            AcceptButton = ok;
            CancelButton = cancel;

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6) };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            buttons.Controls.Add(reset);

            var previewBox = new GroupBox { Text = "プレビュー", Dock = DockStyle.Fill, Padding = new Padding(6) };
            previewBox.Controls.Add(_preview);

            var split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
            split.Panel1.Controls.Add(left);
            split.Panel2.Controls.Add(previewBox);

            Controls.Add(split);
            Controls.Add(buttons);
            Load += (s, e) => { split.SplitterDistance = 420; UpdatePreview(); };

            foreach (var cb in new[] { _useTabs, _procBody, _indentCase, _typeEnum, _directives, _comments, _contAlign })
                cb.CheckedChanged += (s, e) => UpdatePreview();
            _indentSize.ValueChanged += (s, e) => UpdatePreview();
            _contLevels.ValueChanged += (s, e) => UpdatePreview();
        }

        private void LoadOptions(FormatterOptions o)
        {
            _indentSize.Value = o.IndentSize;
            _useTabs.Checked = o.UseTabs;
            _procBody.Checked = o.IndentProcedureBody;
            _indentCase.Checked = o.IndentCase;
            _typeEnum.Checked = o.IndentTypeAndEnumMembers;
            _directives.Checked = o.IndentCompilerDirectives;
            _comments.Checked = o.IndentCommentLines;
            _contAlign.Checked = o.PreserveContinuationAlignment;
            _contLevels.Value = o.ContinuationIndentLevels;
        }

        private FormatterOptions ReadOptions() => new FormatterOptions
        {
            IndentSize = (int)_indentSize.Value,
            UseTabs = _useTabs.Checked,
            IndentProcedureBody = _procBody.Checked,
            IndentCase = _indentCase.Checked,
            IndentTypeAndEnumMembers = _typeEnum.Checked,
            IndentCompilerDirectives = _directives.Checked,
            IndentCommentLines = _comments.Checked,
            PreserveContinuationAlignment = _contAlign.Checked,
            ContinuationIndentLevels = (int)_contLevels.Value,
        };

        private void UpdatePreview()
        {
            string text = VbaIndenter.Format(Sample, ReadOptions());
            // タブはプレビュー上で設定幅の空白として表示
            _preview.Text = text.Replace("\t", new string(' ', (int)_indentSize.Value));
        }

        private void OnOk()
        {
            foreach (var tb in new[] { _hkCurrent, _hkModule })
            {
                string t = tb.Text.Trim();
                if (t.Length > 0 && !Hotkey.TryParse(t, out _))
                {
                    MessageBox.Show(this, $"ショートカットキー「{t}」を解釈できません。\n例: Ctrl+Shift+F、Ctrl+Alt+I、F12",
                        Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tb.Focus();
                    return;
                }
            }
            Result = new Settings
            {
                Options = ReadOptions(),
                FormatCurrentHotkey = _hkCurrent.Text.Trim(),
                FormatModuleHotkey = _hkModule.Text.Trim(),
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private static CheckBox Check(string text) => new CheckBox { Text = text, AutoSize = true, Margin = new Padding(3, 4, 3, 4) };

        private static Control Row(string label, Control control)
        {
            var p = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 2) };
            p.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(3, 6, 3, 3) });
            p.Controls.Add(control);
            return p;
        }
    }
}
