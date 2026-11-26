using System;
using System.Drawing;
using System.Windows.Forms;

namespace SnapIndenter.AddIn
{
    /// <summary>
    /// 実際に押したキーの組み合わせをショートカットとして登録する入力欄。
    /// Backspace / Delete で解除。Tab / Shift+Tab / Esc / Enter は通常どおり（フォーカス移動・キャンセル・OK）。
    /// </summary>
    internal sealed class HotkeyBox : TextBox
    {
        /// <summary>登録できない組み合わせが押されたとき。引数は押された組み合わせの文字列。</summary>
        public event EventHandler<string> Rejected;

        /// <summary>登録できる組み合わせが押された、または解除されたとき。</summary>
        public event EventHandler Accepted;

        public HotkeyBox()
        {
            ReadOnly = true;
            BackColor = SystemColors.Window;
            ShortcutsEnabled = false;   // 右クリックメニューや Ctrl+C などの既定動作を止める
            ImeMode = ImeMode.Disable;  // 日本語入力が有効でもキーをそのまま受け取る
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            Keys mods = keyData & Keys.Modifiers;

            // フォーカス移動とダイアログのキャンセルは通常どおり
            if (key == Keys.Tab && (mods == Keys.None || mods == Keys.Shift)) return base.ProcessCmdKey(ref msg, keyData);
            if ((key == Keys.Escape || key == Keys.Enter) && mods == Keys.None) return base.ProcessCmdKey(ref msg, keyData);

            if ((key == Keys.Back || key == Keys.Delete) && mods == Keys.None)
            {
                Text = "";
                Accepted?.Invoke(this, EventArgs.Empty);
                return true;
            }

            // 修飾キーだけが押された段階では何もせず、続けて押されるキーを待つ。
            // NumLock / CapsLock や日本語入力の切り替えキーも、ショートカットにはしないので黙って無視する
            if (IsModifierKey(key) || IsIgnoredKey(key)) return true;

            string text = Format(mods, key);
            if (Hotkey.TryParse(text, out _))
            {
                Text = text;
                SelectionStart = Text.Length;
                Accepted?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                Rejected?.Invoke(this, text);
            }
            return true;
        }

        private static bool IsModifierKey(Keys key)
            => key == Keys.ShiftKey || key == Keys.ControlKey || key == Keys.Menu
            || key == Keys.LShiftKey || key == Keys.RShiftKey
            || key == Keys.LControlKey || key == Keys.RControlKey
            || key == Keys.LMenu || key == Keys.RMenu
            || key == Keys.LWin || key == Keys.RWin;

        private static bool IsIgnoredKey(Keys key)
            => key == Keys.NumLock || key == Keys.CapsLock || key == Keys.Scroll
            || key == Keys.KanaMode || key == Keys.KanjiMode || key == Keys.ProcessKey
            || key == Keys.IMEConvert || key == Keys.IMENonconvert || key == Keys.IMEModeChange
            || key == Keys.IMEAccept || key == Keys.Packet || key == Keys.None
            || ((int)key >= 0xF0 && (int)key <= 0xF6); // 英数・全角/半角など日本語キーボード固有のキー

        /// <summary>Hotkey.TryParse で読み戻せる "Ctrl+Shift+F" 形式にする。</summary>
        internal static string Format(Keys mods, Keys key)
        {
            string name = key >= Keys.D0 && key <= Keys.D9
                ? ((char)('0' + (key - Keys.D0))).ToString()
                : key.ToString();
            string s = "";
            if ((mods & Keys.Control) != 0) s += "Ctrl+";
            if ((mods & Keys.Shift) != 0) s += "Shift+";
            if ((mods & Keys.Alt) != 0) s += "Alt+";
            return s + name;
        }
    }
}
