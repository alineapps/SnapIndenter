using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using SnapIndenter.Core;

namespace SnapIndenter.AddIn
{
    /// <summary>%APPDATA%\SnapIndenter\settings.ini に保存される設定。</summary>
    internal sealed class Settings
    {
        public FormatterOptions Options { get; set; } = new FormatterOptions();
        public string FormatCurrentHotkey { get; set; } = "Ctrl+Shift+F";
        public string FormatModuleHotkey { get; set; } = "Ctrl+Shift+M";
        /// <summary>表示言語。auto（Windows の表示言語に従う）/ ja / en。</summary>
        public string Language { get; set; } = L10n.Auto;

        public static string FilePath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SnapIndenter", "settings.ini");

        /// <summary>旧名称（Kanna for VBA 1.1、VBA Formatter 1.0）の設定ファイル。新しい設定がなければ、新しい順に探して引き継ぐ。</summary>
        private static readonly string[] LegacyFilePaths =
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kanna", "settings.ini"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VbaFormatter", "settings.ini"),
        };

        public static Settings Load()
        {
            var s = new Settings();
            try
            {
                string path = File.Exists(FilePath) ? FilePath : Array.Find(LegacyFilePaths, File.Exists);
                if (path == null) return s;
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line.StartsWith(";") || line.StartsWith("#") || line.StartsWith("[")) continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    map[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
                }

                var o = s.Options;
                o.IndentSize = Clamp(GetInt(map, "IndentSize", o.IndentSize), 1, 16);
                o.UseTabs = GetBool(map, "UseTabs", o.UseTabs);
                o.IndentProcedureBody = GetBool(map, "IndentProcedureBody", o.IndentProcedureBody);
                o.IndentCase = GetBool(map, "IndentCase", o.IndentCase);
                o.IndentTypeAndEnumMembers = GetBool(map, "IndentTypeAndEnumMembers", o.IndentTypeAndEnumMembers);
                o.IndentCompilerDirectives = GetBool(map, "IndentCompilerDirectives", o.IndentCompilerDirectives);
                o.IndentCommentLines = GetBool(map, "IndentCommentLines", o.IndentCommentLines);
                o.PreserveContinuationAlignment = GetBool(map, "PreserveContinuationAlignment", o.PreserveContinuationAlignment);
                o.ContinuationIndentLevels = Clamp(GetInt(map, "ContinuationIndentLevels", o.ContinuationIndentLevels), 0, 8);
                if (map.TryGetValue("FormatCurrentHotkey", out var h1)) s.FormatCurrentHotkey = h1;
                if (map.TryGetValue("FormatModuleHotkey", out var h2)) s.FormatModuleHotkey = h2;
                if (map.TryGetValue("Language", out var lang)) s.Language = L10n.Normalize(lang);
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
            return s;
        }

        public void Save()
        {
            var o = Options;
            var sb = new StringBuilder();
            sb.AppendLine("[Format]");
            sb.AppendLine($"IndentSize={o.IndentSize}");
            sb.AppendLine($"UseTabs={o.UseTabs}");
            sb.AppendLine($"IndentProcedureBody={o.IndentProcedureBody}");
            sb.AppendLine($"IndentCase={o.IndentCase}");
            sb.AppendLine($"IndentTypeAndEnumMembers={o.IndentTypeAndEnumMembers}");
            sb.AppendLine($"IndentCompilerDirectives={o.IndentCompilerDirectives}");
            sb.AppendLine($"IndentCommentLines={o.IndentCommentLines}");
            sb.AppendLine($"PreserveContinuationAlignment={o.PreserveContinuationAlignment}");
            sb.AppendLine($"ContinuationIndentLevels={o.ContinuationIndentLevels}");
            sb.AppendLine("[Hotkeys]");
            sb.AppendLine($"FormatCurrentHotkey={FormatCurrentHotkey}");
            sb.AppendLine($"FormatModuleHotkey={FormatModuleHotkey}");
            sb.AppendLine("[UI]");
            sb.AppendLine($"Language={L10n.Normalize(Language)}");
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(false));
        }

        private static int Clamp(int v, int min, int max) => Math.Max(min, Math.Min(max, v));

        private static int GetInt(Dictionary<string, string> m, string key, int def)
            => m.TryGetValue(key, out var v) && int.TryParse(v, out var i) ? i : def;

        private static bool GetBool(Dictionary<string, string> m, string key, bool def)
        {
            if (!m.TryGetValue(key, out var v)) return def;
            if (bool.TryParse(v, out var b)) return b;
            return v == "1" ? true : v == "0" ? false : def;
        }
    }

    /// <summary>"Ctrl+Shift+F" 形式のショートカットキー。</summary>
    internal sealed class Hotkey
    {
        public bool Ctrl, Shift, Alt;
        public Keys Key;

        public static bool TryParse(string text, out Hotkey hotkey)
        {
            hotkey = null;
            if (string.IsNullOrWhiteSpace(text)) return false;
            var h = new Hotkey();
            foreach (var part in text.Split('+'))
            {
                string p = part.Trim();
                if (p.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || p.Equals("Control", StringComparison.OrdinalIgnoreCase)) h.Ctrl = true;
                else if (p.Equals("Shift", StringComparison.OrdinalIgnoreCase)) h.Shift = true;
                else if (p.Equals("Alt", StringComparison.OrdinalIgnoreCase)) h.Alt = true;
                else if (p.Length == 1 && char.IsDigit(p[0])) h.Key = Keys.D0 + (p[0] - '0');
                else if (Enum.TryParse(p, true, out Keys k) && k != Keys.None) h.Key = k;
                else return false;
            }
            if (h.Key == Keys.None) return false;
            // 修飾キーなしの英数字は通常入力と衝突するので許可しない
            if (!h.Ctrl && !h.Alt && !(h.Key >= Keys.F1 && h.Key <= Keys.F24)) return false;
            hotkey = h;
            return true;
        }
    }
}
