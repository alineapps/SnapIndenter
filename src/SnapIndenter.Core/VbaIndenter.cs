using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace SnapIndenter.Core
{
    /// <summary>
    /// VBA コードのインデント整形エンジン。
    /// 行数は変えない（入力 n 行 → 出力 n 行）ので、VBE 側で行単位に差し替えられる。
    /// </summary>
    public static class VbaIndenter
    {
        private const RegexOptions Opt = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

        private static readonly Regex ProcStartRe = new Regex(@"^((Public|Private|Friend|Global|Static)\s+)*(Sub|Function|Property\s+(Get|Let|Set))\b", Opt);
        private static readonly Regex ProcEndRe = new Regex(@"^End\s+(Sub|Function|Property)\b", Opt);
        private static readonly Regex TypeStartRe = new Regex(@"^((Public|Private|Global)\s+)?Type\s+\S", Opt);
        private static readonly Regex TypeEndRe = new Regex(@"^End\s+Type\b", Opt);
        private static readonly Regex EnumStartRe = new Regex(@"^((Public|Private|Global)\s+)?Enum\s+\S", Opt);
        private static readonly Regex EnumEndRe = new Regex(@"^End\s+Enum\b", Opt);
        private static readonly Regex IfRe = new Regex(@"^If\b", Opt);
        private static readonly Regex ThenRe = new Regex(@"\bThen\b", Opt);
        private static readonly Regex ElseIfRe = new Regex(@"^ElseIf\b", Opt);
        private static readonly Regex ElseRe = new Regex(@"^Else\b", Opt);
        private static readonly Regex EndIfRe = new Regex(@"^(End\s+If|EndIf)\b", Opt);
        private static readonly Regex SelectRe = new Regex(@"^Select\s+Case\b", Opt);
        private static readonly Regex CaseRe = new Regex(@"^Case\b", Opt);
        private static readonly Regex EndSelectRe = new Regex(@"^End\s+Select\b", Opt);
        private static readonly Regex ForRe = new Regex(@"^For\b", Opt);
        private static readonly Regex NextRe = new Regex(@"^Next\b(?<vars>.*)$", Opt);
        private static readonly Regex DoRe = new Regex(@"^Do\b", Opt);
        private static readonly Regex LoopRe = new Regex(@"^Loop\b", Opt);
        private static readonly Regex WhileRe = new Regex(@"^While\b", Opt);
        private static readonly Regex WendRe = new Regex(@"^Wend\b", Opt);
        private static readonly Regex WithRe = new Regex(@"^With\b", Opt);
        private static readonly Regex EndWithRe = new Regex(@"^End\s+With\b", Opt);
        private static readonly Regex RemRe = new Regex(@"^Rem(\s|$)", Opt);

        private static readonly Regex DirIfRe = new Regex(@"^#If\b", Opt);
        private static readonly Regex DirElseIfRe = new Regex(@"^#ElseIf\b", Opt);
        private static readonly Regex DirElseRe = new Regex(@"^#Else\b", Opt);
        private static readonly Regex DirEndIfRe = new Regex(@"^#End\s*If\b", Opt);

        private static readonly Regex LineNumberRe = new Regex(@"^\d+(?=\s|$)", Opt);
        private static readonly Regex LabelRe = new Regex(@"^[\p{L}_][\p{L}\p{Nd}_]*:(?!=)", Opt);

        private static readonly HashSet<string> NotLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Else", "End", "Exit", "Loop", "Next", "Wend", "Do", "Resume", "Return", "Stop", "Rem",
            "Call", "GoSub", "GoTo", "With", "Select", "Case", "If", "ElseIf", "For", "While",
        };

        /// <summary>テキスト全体を整形する。改行は CRLF で返す。</summary>
        public static string Format(string code, FormatterOptions options = null)
        {
            if (string.IsNullOrEmpty(code)) return code ?? string.Empty;
            return string.Join("\r\n", Format(SplitLines(code), options));
        }

        public static string[] SplitLines(string text)
            => text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

        /// <summary>行の配列を整形する。戻り値の行数は入力と同じ。</summary>
        public static string[] Format(IReadOnlyList<string> lines, FormatterOptions options = null)
        {
            options = options ?? new FormatterOptions();
            if (options.IndentSize < 1) options = WithIndent(options, 1);

            var result = new string[lines.Count];
            var state = new IndentState(options);
            int i = 0;
            while (i < lines.Count)
            {
                int start = i;
                while (i < lines.Count - 1 && EndsWithContinuation(lines[i] ?? string.Empty)) i++;
                int end = i;
                i++;
                FormatLogicalLine(lines, start, end, state, options, result);
            }
            return result;
        }

        private static FormatterOptions WithIndent(FormatterOptions o, int size)
        {
            var c = o.Clone();
            c.IndentSize = size;
            return c;
        }

        internal static bool EndsWithContinuation(string line)
        {
            string t = line.TrimEnd();
            if (!t.EndsWith("_", StringComparison.Ordinal)) return false;
            return t.Length == 1 || t[t.Length - 2] == ' ' || t[t.Length - 2] == '\t';
        }

        private static void FormatLogicalLine(IReadOnlyList<string> lines, int start, int end,
            IndentState state, FormatterOptions o, string[] result)
        {
            int count = end - start + 1;
            var content = new string[count];
            for (int k = 0; k < count; k++) content[k] = (lines[start + k] ?? string.Empty).Trim();

            if (count == 1 && content[0].Length == 0)
            {
                result[start] = string.Empty;
                return;
            }

            string code = ExtractCode(content).TrimEnd();
            string label = null;
            int level;

            if (code.Length == 0)
            {
                // コメントだけの行
                if (!o.IndentCommentLines)
                {
                    for (int k = 0; k < count; k++) result[start + k] = (lines[start + k] ?? string.Empty).TrimEnd();
                    return;
                }
                level = state.Level;
            }
            else if (code[0] == '#')
            {
                level = state.ProcessDirective(code);
            }
            else
            {
                label = MatchLabel(code);
                string statements = label == null ? code : code.Substring(label.Length);
                level = state.ProcessStatements(statements);
            }

            // ---- 1 行目の組み立て ----
            string first = content[0];
            string body = label == null ? first : first.Substring(label.Length).TrimStart();
            int col = level * o.IndentSize;
            int firstCol;
            string newFirst;
            if (label == null)
            {
                newFirst = Whitespace(0, col, o) + body;
                firstCol = col;
            }
            else if (body.Length == 0)
            {
                newFirst = label;
                firstCol = label.Length;
            }
            else if (label.Length < col)
            {
                newFirst = label + Whitespace(label.Length, col, o) + body;
                firstCol = col;
            }
            else
            {
                newFirst = label + " " + body;
                firstCol = label.Length + 1;
            }
            result[start] = newFirst;

            if (count == 1) return;

            // ---- 継続行 ----
            string origFirstLine = lines[start] ?? string.Empty;
            int bodyIndex = origFirstLine.Length - origFirstLine.TrimStart().Length;
            if (label != null)
            {
                bodyIndex += label.Length;
                while (bodyIndex < origFirstLine.Length && (origFirstLine[bodyIndex] == ' ' || origFirstLine[bodyIndex] == '\t')) bodyIndex++;
            }
            int origFirstCol = VisualColumn(origFirstLine, bodyIndex, o.IndentSize);

            for (int k = 1; k < count; k++)
            {
                string raw = lines[start + k] ?? string.Empty;
                if (content[k].Length == 0)
                {
                    result[start + k] = string.Empty;
                    continue;
                }
                int leading = raw.Length - raw.TrimStart().Length;
                int offset = VisualColumn(raw, leading, o.IndentSize) - origFirstCol;
                int newCol = o.PreserveContinuationAlignment && offset > 0
                    ? firstCol + offset
                    : firstCol + Math.Max(0, o.ContinuationIndentLevels) * o.IndentSize;
                result[start + k] = Whitespace(0, newCol, o) + content[k];
            }
        }

        /// <summary>
        /// 論理行からコード部分だけを取り出す。文字列リテラルの中身は 'x' に置き換え、
        /// コメント以降は捨て、継続記号は空白に置き換えて連結する。
        /// 1 行目については元の文字位置と 1 対 1 に対応する。
        /// </summary>
        private static string ExtractCode(string[] content)
        {
            var sb = new StringBuilder();
            for (int k = 0; k < content.Length; k++)
            {
                string c = content[k];
                bool isLast = k == content.Length - 1;
                int len = isLast ? c.Length : Math.Max(0, c.Length - 1); // 末尾の "_" を除く
                bool inString = false;
                for (int j = 0; j < len; j++)
                {
                    char ch = c[j];
                    if (inString)
                    {
                        if (ch == '"')
                        {
                            if (j + 1 < len && c[j + 1] == '"')
                            {
                                sb.Append("xx");
                                j++;
                                continue;
                            }
                            inString = false;
                            sb.Append('"');
                        }
                        else
                        {
                            sb.Append('x');
                        }
                        continue;
                    }
                    if (ch == '"')
                    {
                        inString = true;
                        sb.Append('"');
                        continue;
                    }
                    if (ch == '\'') return sb.ToString(); // 以降はコメント（継続行も含む）
                    sb.Append(ch);
                }
                if (!isLast) sb.Append(' ');
            }
            return sb.ToString();
        }

        private static string MatchLabel(string code)
        {
            var m = LineNumberRe.Match(code);
            if (m.Success) return m.Value;
            m = LabelRe.Match(code);
            if (m.Success)
            {
                string name = m.Value.Substring(0, m.Value.Length - 1);
                if (!NotLabels.Contains(name)) return m.Value;
            }
            return null;
        }

        private static int VisualColumn(string line, int index, int tabSize)
        {
            int col = 0;
            for (int i = 0; i < index && i < line.Length; i++)
            {
                if (line[i] == '\t') col = (col / tabSize + 1) * tabSize;
                else col++;
            }
            return col;
        }

        /// <summary>表示上の桁 from から to までを埋める空白文字列。</summary>
        private static string Whitespace(int from, int to, FormatterOptions o)
        {
            if (to <= from) return from == 0 ? string.Empty : " ";
            if (!o.UseTabs) return new string(' ', to - from);

            var sb = new StringBuilder();
            int pos = from;
            while (true)
            {
                int nextStop = (pos / o.IndentSize + 1) * o.IndentSize;
                if (nextStop > to) break;
                sb.Append('\t');
                pos = nextStop;
            }
            sb.Append(' ', to - pos);
            return sb.ToString();
        }

        // ------------------------------------------------------------------

        private enum BlockKind { Procedure, If, Select, CaseBody, For, Do, While, With, Type, Enum, Directive }

        private sealed class Block
        {
            public BlockKind Kind;
            public int Inc;
            public int BaseLevel;              // Directive 用
            public List<Block> Snapshot;       // Directive 用: #If 直後の状態
            public List<Block> FirstBranchEnd; // Directive 用: 最初の分岐終了時の状態
        }

        private sealed class IndentState
        {
            private readonly FormatterOptions _o;
            private List<Block> _stack = new List<Block>();
            private int? _lineLevel;

            public IndentState(FormatterOptions o) { _o = o; }

            public int Level
            {
                get
                {
                    int s = 0;
                    foreach (var b in _stack) s += b.Inc;
                    return s;
                }
            }

            private void Push(BlockKind kind, int inc = 1) => _stack.Add(new Block { Kind = kind, Inc = inc });

            private void Mark()
            {
                if (_lineLevel == null) _lineLevel = Level;
            }

            /// <summary>指定種別のブロックまで閉じる。プロシージャ・コンパイラ指令の境界は越えない。</summary>
            private bool PopTo(BlockKind kind)
            {
                for (int i = _stack.Count - 1; i >= 0; i--)
                {
                    var b = _stack[i];
                    if (b.Kind == kind)
                    {
                        _stack.RemoveRange(i, _stack.Count - i);
                        return true;
                    }
                    if (b.Kind == BlockKind.Procedure || b.Kind == BlockKind.Directive) return false;
                }
                return false;
            }

            private void ClearNonDirective() => _stack.RemoveAll(b => b.Kind != BlockKind.Directive);

            public int ProcessStatements(string code)
            {
                _lineLevel = null;
                int pos = 0;
                while (pos < code.Length)
                {
                    int sep = FindSeparator(code, pos);
                    string s = code.Substring(pos, sep - pos).Trim();
                    if (s.Length > 0 && ProcessStatement(s, code.Substring(pos))) break;
                    pos = sep + 1;
                }
                return _lineLevel ?? Level;
            }

            private static int FindSeparator(string code, int pos)
            {
                for (int i = pos; i < code.Length; i++)
                {
                    if (code[i] == '"')
                    {
                        // マスク済み文字列はスキップ
                        int close = code.IndexOf('"', i + 1);
                        if (close < 0) return code.Length;
                        i = close;
                        continue;
                    }
                    if (code[i] == ':' && (i + 1 >= code.Length || code[i + 1] != '=')) return i;
                }
                return code.Length;
            }

            /// <returns>true ならこの論理行の残りを処理しない。</returns>
            private bool ProcessStatement(string s, string rest)
            {
                if (RemRe.IsMatch(s))
                {
                    Mark();
                    return true;
                }

                if (IfRe.IsMatch(s))
                {
                    Mark();
                    var m = ThenRe.Match(rest);
                    if (!m.Success) return true;
                    if (rest.Substring(m.Index + m.Length).Trim().Length > 0) return true; // 1 行 If
                    Push(BlockKind.If);
                    return false;
                }
                if (ElseIfRe.IsMatch(s) || ElseRe.IsMatch(s))
                {
                    bool found = PopTo(BlockKind.If);
                    Mark();
                    if (found) Push(BlockKind.If);
                    return false;
                }
                if (EndIfRe.IsMatch(s)) { PopTo(BlockKind.If); Mark(); return false; }

                if (SelectRe.IsMatch(s)) { Mark(); Push(BlockKind.Select, _o.IndentCase ? 1 : 0); return false; }
                if (CaseRe.IsMatch(s))
                {
                    int sel = -1;
                    for (int i = _stack.Count - 1; i >= 0; i--)
                    {
                        var k = _stack[i].Kind;
                        if (k == BlockKind.Select) { sel = i; break; }
                        if (k == BlockKind.Procedure || k == BlockKind.Directive) break;
                    }
                    if (sel >= 0) _stack.RemoveRange(sel + 1, _stack.Count - sel - 1);
                    Mark();
                    if (sel >= 0) Push(BlockKind.CaseBody);
                    return false;
                }
                if (EndSelectRe.IsMatch(s)) { PopTo(BlockKind.Select); Mark(); return false; }

                if (ForRe.IsMatch(s)) { Mark(); Push(BlockKind.For); return false; }
                var nm = NextRe.Match(s);
                if (nm.Success)
                {
                    string vars = nm.Groups["vars"].Value.Trim();
                    int n = vars.Length == 0 ? 1 : vars.Split(',').Length;
                    for (int i = 0; i < n; i++) PopTo(BlockKind.For);
                    Mark();
                    return false;
                }
                if (DoRe.IsMatch(s)) { Mark(); Push(BlockKind.Do); return false; }
                if (LoopRe.IsMatch(s)) { PopTo(BlockKind.Do); Mark(); return false; }
                if (WhileRe.IsMatch(s)) { Mark(); Push(BlockKind.While); return false; }
                if (WendRe.IsMatch(s)) { PopTo(BlockKind.While); Mark(); return false; }
                if (WithRe.IsMatch(s)) { Mark(); Push(BlockKind.With); return false; }
                if (EndWithRe.IsMatch(s)) { PopTo(BlockKind.With); Mark(); return false; }

                if (ProcStartRe.IsMatch(s))
                {
                    ClearNonDirective();
                    Mark();
                    Push(BlockKind.Procedure, _o.IndentProcedureBody ? 1 : 0);
                    return false;
                }
                if (ProcEndRe.IsMatch(s)) { ClearNonDirective(); Mark(); return false; }

                int memberInc = _o.IndentTypeAndEnumMembers ? 1 : 0;
                if (TypeStartRe.IsMatch(s)) { Mark(); Push(BlockKind.Type, memberInc); return false; }
                if (TypeEndRe.IsMatch(s)) { PopTo(BlockKind.Type); Mark(); return false; }
                if (EnumStartRe.IsMatch(s)) { Mark(); Push(BlockKind.Enum, memberInc); return false; }
                if (EnumEndRe.IsMatch(s)) { PopTo(BlockKind.Enum); Mark(); return false; }

                Mark();
                return false;
            }

            public int ProcessDirective(string code)
            {
                if (DirIfRe.IsMatch(code))
                {
                    int baseLevel = Level;
                    var d = new Block
                    {
                        Kind = BlockKind.Directive,
                        Inc = _o.IndentCompilerDirectives ? 1 : 0,
                        BaseLevel = baseLevel,
                    };
                    _stack.Add(d);
                    d.Snapshot = new List<Block>(_stack);
                    return DirectiveColumn(baseLevel);
                }

                bool isElse = DirElseIfRe.IsMatch(code) || DirElseRe.IsMatch(code);
                bool isEnd = !isElse && DirEndIfRe.IsMatch(code);
                if (!isElse && !isEnd) return DirectiveColumn(Level); // #Const など

                Block dir = null;
                for (int i = _stack.Count - 1; i >= 0; i--)
                {
                    if (_stack[i].Kind == BlockKind.Directive) { dir = _stack[i]; break; }
                }
                if (dir == null) return DirectiveColumn(Level);

                if (isElse)
                {
                    // 各分岐は #If 直後の状態から始める（#If で囲まれたプロシージャ宣言などに対応）
                    if (dir.FirstBranchEnd == null) dir.FirstBranchEnd = new List<Block>(_stack);
                    _stack = new List<Block>(dir.Snapshot);
                    return DirectiveColumn(dir.BaseLevel);
                }

                if (dir.FirstBranchEnd != null) _stack = new List<Block>(dir.FirstBranchEnd);
                _stack.Remove(dir);
                return DirectiveColumn(dir.BaseLevel);
            }

            private int DirectiveColumn(int level) => _o.IndentCompilerDirectives ? level : 0;
        }
    }
}
