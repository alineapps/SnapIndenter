using System;
using System.Globalization;
using Microsoft.Win32;

namespace SnapIndenter.AddIn
{
    /// <summary>画面に表示する文字列の言語（日本語 / 英語）を切り替える。</summary>
    internal static class L10n
    {
        public const string Auto = "auto";
        public const string Japanese = "ja";
        public const string English = "en";

        /// <summary>現在の表示言語が日本語かどうか。</summary>
        public static bool IsJapanese { get; private set; } = Resolve(Auto);

        /// <summary>設定値（auto / ja / en）から表示言語を決める。auto は Office の表示言語（取れなければ Windows の表示言語）に従う。</summary>
        public static void Apply(string language) => IsJapanese = Resolve(language);

        /// <summary>日本語と英語の文字列から、現在の表示言語のものを返す。</summary>
        public static string T(string ja, string en) => IsJapanese ? ja : en;

        public static string Normalize(string language)
        {
            if (string.Equals(language, Japanese, StringComparison.OrdinalIgnoreCase)) return Japanese;
            if (string.Equals(language, English, StringComparison.OrdinalIgnoreCase)) return English;
            return Auto;
        }

        private static bool Resolve(string language)
        {
            switch (Normalize(language))
            {
                case Japanese: return true;
                case English: return false;
                default:
                    string office = OfficeUiLanguage();
                    if (office != null) return office.StartsWith("ja", StringComparison.OrdinalIgnoreCase);
                    try { return CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "ja"; }
                    catch { return false; }
            }
        }

        /// <summary>
        /// Office の表示言語（例: "ja-jp"）。VBE のメニューは Office の表示言語に従うので、Windows の表示言語より優先する。
        /// 英語版 Windows に日本語の言語パックを入れた環境では、.NET の CurrentUICulture が en-US になることがあるため。
        /// </summary>
        private static string OfficeUiLanguage()
        {
            foreach (string version in new[] { "16.0", "15.0", "14.0" })
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{version}\Common\LanguageResources"))
                    {
                        if (key == null) continue;
                        if (key.GetValue("UILanguageTag") is string tag && tag.Length > 0) return tag;
                        if (key.GetValue("UILanguage") is int lcid && lcid > 0) return new CultureInfo(lcid).Name;
                    }
                }
                catch
                {
                    // 読めなければ次の候補へ
                }
            }
            return null;
        }
    }
}
