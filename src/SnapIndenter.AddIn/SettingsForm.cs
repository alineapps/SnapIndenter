using System;
using System.Drawing;
using System.Windows.Forms;
using SnapIndenter.Core;

namespace SnapIndenter.AddIn
{
    /// <summary>設定ダイアログ（プレビュー付き）。</summary>
    internal sealed class SettingsForm : Form
    {
        private static string Sample => L10n.IsJapanese
            ? SampleTemplate
            : SampleTemplate.Replace("' コメント", "' Comment").Replace("長いメッセージ", "A long message").Replace("を連結", " continued");

        private const string SampleTemplate =
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
        private readonly CheckBox _useTabs = Check(L10n.T("インデントにタブを使う", "Use tabs for indentation"));
        private readonly CheckBox _procBody = Check(L10n.T("プロシージャの本体をインデントする", "Indent procedure bodies"));
        private readonly CheckBox _indentCase = Check(L10n.T("Select Case 内の Case をインデントする", "Indent Case inside Select Case"));
        private readonly CheckBox _typeEnum = Check(L10n.T("Type / Enum のメンバーをインデントする", "Indent Type / Enum members"));
        private readonly CheckBox _directives = Check(L10n.T("#If ～ #End If の中をインデントする（オフ: 指令は 1 桁目）", "Indent inside #If ... #End If (off: directives at column 1)"));
        private readonly CheckBox _comments = Check(L10n.T("コメント行もインデントする", "Indent comment lines"));
        private readonly CheckBox _contAlign = Check(L10n.T("継続行（ _）の相対位置を保持する", "Keep relative position of continuation lines ( _)"));
        private readonly NumericUpDown _contLevels = new NumericUpDown { Minimum = 0, Maximum = 8, Width = 60 };
        private readonly HotkeyBox _hkCurrent = new HotkeyBox { Width = 160 };
        private readonly HotkeyBox _hkModule = new HotkeyBox { Width = 160 };
        private readonly Label _hkMessage = new Label { AutoSize = true, Margin = new Padding(3, 2, 3, 3), MaximumSize = new Size(380, 0) };
        private readonly ComboBox _language = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        private static readonly string[] LanguageCodes = { L10n.Auto, L10n.Japanese, L10n.English };
        private readonly TextBox _preview = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Font = new Font(L10n.IsJapanese ? "MS Gothic" : "Consolas", 10f),
            Dock = DockStyle.Fill,
            BackColor = SystemColors.Window,
        };

        public Settings Result { get; private set; }

        public SettingsForm(Settings current)
        {
            Text = L10n.T("SnapIndenter の設定", "SnapIndenter Settings");
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = L10n.IsJapanese ? new Font("Yu Gothic UI", 9f) : SystemFonts.MessageBoxFont;
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
            _language.Items.AddRange(new object[] { L10n.T("自動（Office の表示言語）", "Auto (Office display language)"), "日本語", "English" });
            _language.SelectedIndex = Math.Max(0, Array.IndexOf(LanguageCodes, L10n.Normalize(current.Language)));

            var left = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(8),
            };
            left.Controls.Add(Row(L10n.T("インデント幅:", "Indent width:"), _indentSize));
            left.Controls.Add(_useTabs);
            left.Controls.Add(_procBody);
            left.Controls.Add(_indentCase);
            left.Controls.Add(_typeEnum);
            left.Controls.Add(_directives);
            left.Controls.Add(_comments);
            left.Controls.Add(_contAlign);
            left.Controls.Add(Row(L10n.T("継続行の追加インデント（段）:", "Extra indent for continuation lines (levels):"), _contLevels));
            left.Controls.Add(new Label { Text = L10n.T("ショートカットキー", "Shortcut keys"), AutoSize = true, Margin = new Padding(3, 16, 3, 3) });
            left.Controls.Add(HotkeyRow(L10n.T("プロシージャ/選択範囲:", "Procedure / selection:"), _hkCurrent));
            left.Controls.Add(HotkeyRow(L10n.T("モジュール:", "Module:"), _hkModule));
            left.Controls.Add(_hkMessage);
            ShowHotkeyHint();
            foreach (var box in new[] { _hkCurrent, _hkModule })
            {
                box.Accepted += (s, e) => ShowHotkeyHint();
                box.Rejected += (s, text) => ShowHotkeyError(L10n.T(
                    $"「{text}」は登録できません。Ctrl か Alt を含む組み合わせか、F1～F24 を押してください。",
                    $"\"{text}\" can't be used. Press a combination with Ctrl or Alt, or F1-F24."));
            }
            var languageRow = Row("言語 / Language:", _language);
            languageRow.Margin = new Padding(0, 16, 0, 2);
            left.Controls.Add(languageRow);
            left.Controls.Add(new Label
            {
                Text = L10n.T("設定ファイル: ", "Settings file: ") + Settings.FilePath,
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 16, 3, 3),
                MaximumSize = new Size(380, 0),
            });

            var ok = new Button { Text = "OK", DialogResult = DialogResult.None, Width = 90 };
            var cancel = new Button { Text = L10n.T("キャンセル", "Cancel"), DialogResult = DialogResult.Cancel, Width = 90 };
            var reset = new Button { Text = L10n.T("既定値に戻す", "Reset to Defaults"), Width = 110 };
            ok.Click += (s, e) => OnOk();
            reset.Click += (s, e) => LoadOptions(new FormatterOptions());
            AcceptButton = ok;
            CancelButton = cancel;

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6) };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            buttons.Controls.Add(reset);

            var previewBox = new GroupBox { Text = L10n.T("プレビュー", "Preview"), Dock = DockStyle.Fill, Padding = new Padding(6) };
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
                    MessageBox.Show(this, L10n.T($"ショートカットキー「{t}」を解釈できません。\n例: Ctrl+Shift+F、Ctrl+Alt+I、F12", $"Can't recognize the shortcut \"{t}\".\nExamples: Ctrl+Shift+F, Ctrl+Alt+I, F12"),
                        Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tb.Focus();
                    return;
                }
            }
            if (_hkCurrent.Text.Length > 0 && string.Equals(_hkCurrent.Text, _hkModule.Text, StringComparison.OrdinalIgnoreCase))
            {
                ShowHotkeyError(L10n.T(
                    $"「{_hkModule.Text}」は両方に登録されています。どちらかを変更してください。",
                    $"\"{_hkModule.Text}\" is assigned to both commands. Change one of them."));
                _hkModule.Focus();
                return;
            }
            Result = new Settings
            {
                Options = ReadOptions(),
                FormatCurrentHotkey = _hkCurrent.Text.Trim(),
                FormatModuleHotkey = _hkModule.Text.Trim(),
                Language = LanguageCodes[Math.Max(0, _language.SelectedIndex)],
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private static CheckBox Check(string text) => new CheckBox { Text = text, AutoSize = true, Margin = new Padding(3, 4, 3, 4) };

        private static Control Row(string label, params Control[] controls)
        {
            var p = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Margin = new Padding(0, 2, 0, 2) };
            p.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(3, 6, 3, 3) });
            p.Controls.AddRange(controls);
            return p;
        }

        /// <summary>ラベルの幅をそろえた「ラベル + キー入力欄 + 解除ボタン」の行。</summary>
        private Control HotkeyRow(string label, HotkeyBox box)
        {
            var clear = new Button { Text = L10n.T("解除", "Clear"), AutoSize = true, Margin = new Padding(3, 1, 3, 1) };
            clear.Click += (s, e) => { box.Text = ""; ShowHotkeyHint(); box.Focus(); };
            var row = (FlowLayoutPanel)Row(label, box, clear);
            row.Controls[0].MinimumSize = new Size(140, 0);
            return row;
        }

        private void ShowHotkeyHint()
        {
            _hkMessage.ForeColor = SystemColors.GrayText;
            _hkMessage.Text = L10n.T(
                "欄をクリックして、登録したいキーを押してください（例: Ctrl+Shift+F）。Backspace または「解除」で無効にできます。",
                "Click a box and press the keys you want (e.g. Ctrl+Shift+F). Press Backspace or \"Clear\" to remove.");
        }

        private void ShowHotkeyError(string text)
        {
            _hkMessage.ForeColor = Color.Firebrick;
            _hkMessage.Text = text;
        }
    }
}
