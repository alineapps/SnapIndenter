namespace Kanna.Core
{
    /// <summary>整形オプション。</summary>
    public sealed class FormatterOptions
    {
        /// <summary>1 段あたりのインデント幅（スペース数）。タブ使用時はタブ幅として扱う。</summary>
        public int IndentSize { get; set; } = 4;

        /// <summary>インデントにタブ文字を使う。</summary>
        public bool UseTabs { get; set; }

        /// <summary>Sub / Function / Property の本体をインデントする。</summary>
        public bool IndentProcedureBody { get; set; } = true;

        /// <summary>Select Case 内の Case 行をインデントする。</summary>
        public bool IndentCase { get; set; } = true;

        /// <summary>Type / Enum のメンバーをインデントする。</summary>
        public bool IndentTypeAndEnumMembers { get; set; } = true;

        /// <summary>
        /// true: #If ～ #End If の中身をインデントする。
        /// false: コンパイラ指令行は常に 1 桁目に置き、中身のインデントには影響させない。
        /// </summary>
        public bool IndentCompilerDirectives { get; set; } = true;

        /// <summary>コメントだけの行もコードと同じ位置にインデントする（false なら元の位置のまま）。</summary>
        public bool IndentCommentLines { get; set; } = true;

        /// <summary>行継続（ _）の 2 行目以降について、1 行目からの相対位置を保持する。</summary>
        public bool PreserveContinuationAlignment { get; set; } = true;

        /// <summary>行継続の 2 行目以降に追加するインデント段数（相対位置を保持しない場合、または元の相対位置が 0 以下の場合）。</summary>
        public int ContinuationIndentLevels { get; set; } = 1;

        public FormatterOptions Clone() => (FormatterOptions)MemberwiseClone();
    }
}
