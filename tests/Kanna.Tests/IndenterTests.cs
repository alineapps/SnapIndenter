using System;
using System.Linq;
using Kanna.Core;
using Xunit;

namespace Kanna.Tests
{
    public class IndenterTests
    {
        /// <summary>期待値の各行から共通の左余白を取り除き、'|' を行頭の目印として使う。</summary>
        private static string Code(string s)
        {
            var lines = s.Replace("\r\n", "\n").Split('\n')
                .SkipWhile(string.IsNullOrWhiteSpace)
                .Reverse().SkipWhile(string.IsNullOrWhiteSpace).Reverse()
                .Select(l => l.TrimStart())
                .Select(l => l.StartsWith("|") ? l.Substring(1) : l);
            return string.Join("\r\n", lines);
        }

        private static void AssertFormat(string expected, FormatterOptions o = null)
        {
            string exp = Code(expected);
            // 全行のインデントを潰した入力から整形して期待値に戻ること
            string flattened = string.Join("\r\n", VbaIndenter.SplitLines(exp).Select(l => l.TrimStart()));
            Assert.Equal(exp, VbaIndenter.Format(flattened, o));
            // 冪等性
            Assert.Equal(exp, VbaIndenter.Format(exp, o));
        }

        [Fact]
        public void BasicProcedure()
        {
            AssertFormat(@"
                |Option Explicit
                |
                |Public Sub Test()
                |    Dim i As Long
                |    For i = 1 To 10
                |        If i Mod 2 = 0 Then
                |            Debug.Print i
                |        ElseIf i = 3 Then
                |            Debug.Print ""three""
                |        Else
                |            Debug.Print ""odd""
                |        End If
                |    Next i
                |End Sub");
        }

        [Fact]
        public void SingleLineIfDoesNotIndent()
        {
            AssertFormat(@"
                |Sub A()
                |    If x Then y = 1
                |    If x Then y = 1 Else y = 2
                |    If x Then: y = 1
                |    If x Then Exit Sub ' comment
                |    If x Then ' comment
                |        y = 1
                |    End If
                |End Sub");
        }

        [Fact]
        public void SelectCase()
        {
            AssertFormat(@"
                |Function F(x) As String
                |    Select Case x
                |        Case 1
                |            F = ""one""
                |        Case 2, 3: F = ""two or three""
                |        Case Else
                |            F = ""other""
                |    End Select
                |End Function");
        }

        [Fact]
        public void SelectCaseWithoutCaseIndent()
        {
            AssertFormat(@"
                |Sub A()
                |    Select Case x
                |    Case 1
                |        y = 1
                |    Case Else
                |        y = 2
                |    End Select
                |End Sub", new FormatterOptions { IndentCase = false });
        }

        [Fact]
        public void Loops()
        {
            AssertFormat(@"
                |Sub A()
                |    Do While x < 10
                |        x = x + 1
                |    Loop
                |    Do
                |        x = x - 1
                |    Loop Until x = 0
                |    While x < 3
                |        x = x + 1
                |    Wend
                |    For Each c In Range(""A1:B2"")
                |        With c
                |            .Value = 1
                |        End With
                |    Next
                |    Do: x = x + 1: Loop Until x > 5
                |End Sub");
        }

        [Fact]
        public void NextWithMultipleVariables()
        {
            AssertFormat(@"
                |Sub A()
                |    For i = 1 To 2
                |        For j = 1 To 2
                |            k = i * j
                |    Next j, i
                |    Beep
                |End Sub");
        }

        [Fact]
        public void StringsAndCommentsDoNotConfuse()
        {
            AssertFormat(@"
                |Sub A()
                |    s = ""If x Then: End Sub ' not comment""
                |    t = ""He said """"Loop"""""" ' End If
                |    ' For i = 1 To 10
                |    Rem Do While
                |    Debug.Print s: Debug.Print t
                |End Sub");
        }

        [Fact]
        public void LabelsAndLineNumbersStayAtColumnZero()
        {
            AssertFormat(@"
                |Sub A()
                |    On Error GoTo ErrHandler
                |    If x Then
                |Retry:
                |        y = 1
                |    End If
                |10  z = 1
                |100 z = 2
                |    Exit Sub
                |ErrHandler:
                |    MsgBox Err.Description
                |    Resume Next
                |End Sub");
        }

        [Fact]
        public void NamedArgumentIsNotLabel()
        {
            AssertFormat(@"
                |Sub A()
                |    MsgBox Prompt:=""x"", Buttons:=vbOKOnly
                |    Range(""A1"").Sort Key1:=Range(""A1""), Order1:=xlAscending
                |End Sub");
        }

        [Fact]
        public void TypeAndEnum()
        {
            AssertFormat(@"
                |Private Type Point
                |    X As Long
                |    Y As Long
                |End Type
                |
                |Public Enum Color
                |    Red = 1
                |    Green
                |End Enum
                |
                |Private Declare PtrSafe Function GetTickCount Lib ""kernel32"" () As Long
                |Private WithEvents app As Application
                |Static Function S() As Long
                |    S = 1
                |End Function");
        }

        [Fact]
        public void PropertiesAndModifiers()
        {
            AssertFormat(@"
                |Private m As Long
                |Public Property Get Value() As Long
                |    Value = m
                |End Property
                |Friend Property Let Value(ByVal v As Long)
                |    m = v
                |End Property");
        }

        [Fact]
        public void LineContinuationIndentsFollowingLines()
        {
            string input = Code(@"
                |Sub A()
                |s = ""a"" & _
                |""b"" & _
                |""c""
                |End Sub");
            string expected = Code(@"
                |Sub A()
                |    s = ""a"" & _
                |        ""b"" & _
                |        ""c""
                |End Sub");
            Assert.Equal(expected, VbaIndenter.Format(input));
            Assert.Equal(expected, VbaIndenter.Format(expected));
        }

        [Fact]
        public void LineContinuationPreservesRelativeAlignment()
        {
            string input = Code(@"
                |Sub A()
                |        Call Foo(a, _
                |                 b)
                |End Sub");
            string expected = Code(@"
                |Sub A()
                |    Call Foo(a, _
                |             b)
                |End Sub");
            Assert.Equal(expected, VbaIndenter.Format(input));
        }

        [Fact]
        public void ContinuedIfHeaderIsBlockIf()
        {
            AssertFormat(@"
                |Sub A()
                |    If a = 1 And _
                |        b = 2 Then
                |        c = 3
                |    End If
                |End Sub");
        }

        [Fact]
        public void CompilerDirectives()
        {
            AssertFormat(@"
                |#If VBA7 Then
                |    Private Declare PtrSafe Sub Sleep Lib ""kernel32"" (ByVal ms As LongPtr)
                |#Else
                |    Private Declare Sub Sleep Lib ""kernel32"" (ByVal ms As Long)
                |#End If
                |
                |Sub A()
                |    #If DEBUG_MODE Then
                |        Debug.Print 1
                |    #End If
                |    x = 1
                |End Sub");
        }

        [Fact]
        public void CompilerDirectivesAroundProcedureHeader()
        {
            AssertFormat(@"
                |#If VBA7 Then
                |    Private Function H() As LongPtr
                |#Else
                |    Private Function H() As Long
                |#End If
                |    H = 0
                |End Function
                |Sub Next1()
                |    x = 1
                |End Sub");
        }

        [Fact]
        public void CompilerDirectivesAtColumnZero()
        {
            AssertFormat(@"
                |Sub A()
                |#If DEBUG_MODE Then
                |    Debug.Print 1
                |#Else
                |    Debug.Print 2
                |#End If
                |End Sub", new FormatterOptions { IndentCompilerDirectives = false });
        }

        [Fact]
        public void TabsOption()
        {
            var o = new FormatterOptions { UseTabs = true };
            string expected = "Sub A()\r\n\tIf x Then\r\n\t\ty = 1\r\n\tEnd If\r\nEnd Sub";
            Assert.Equal(expected, VbaIndenter.Format("Sub A()\r\nIf x Then\r\n  y = 1\r\n      End If\r\nEnd Sub", o));
            Assert.Equal(expected, VbaIndenter.Format(expected, o));
        }

        [Fact]
        public void IndentSizeOption()
        {
            AssertFormat(@"
                |Sub A()
                |  If x Then
                |    y = 1
                |  End If
                |End Sub", new FormatterOptions { IndentSize = 2 });
        }

        [Fact]
        public void NoProcedureBodyIndent()
        {
            AssertFormat(@"
                |Sub A()
                |If x Then
                |    y = 1
                |End If
                |End Sub", new FormatterOptions { IndentProcedureBody = false });
        }

        [Fact]
        public void CommentLinesKeptWhenOptionOff()
        {
            string input = "Sub A()\r\n' top\r\n        ' deep\r\nx = 1\r\nEnd Sub";
            string expected = "Sub A()\r\n' top\r\n        ' deep\r\n    x = 1\r\nEnd Sub";
            Assert.Equal(expected, VbaIndenter.Format(input, new FormatterOptions { IndentCommentLines = false }));
        }

        [Fact]
        public void BlankLinesAndTrailingSpacesAreCleaned()
        {
            string input = "Sub A()   \r\n    \r\nx = 1  \r\nEnd Sub";
            Assert.Equal("Sub A()\r\n\r\n    x = 1\r\nEnd Sub", VbaIndenter.Format(input));
        }

        [Fact]
        public void UnbalancedCodeDoesNotThrowAndRecoversAtNextProcedure()
        {
            AssertFormat(@"
                |Sub A()
                |    If x Then
                |        For i = 1 To 2
                |End Sub
                |Sub B()
                |    y = 1
                |End Sub");
        }

        [Fact]
        public void JapaneseIdentifiersAndLabels()
        {
            AssertFormat(@"
                |Sub 集計()
                |    Dim 合計 As Long
                |    On Error GoTo エラー処理
                |    For Each 行 In 範囲.Rows
                |        合計 = 合計 + 行.Cells(1).Value
                |    Next
                |    Exit Sub
                |エラー処理:
                |    MsgBox ""エラー: "" & Err.Description
                |End Sub");
        }

        [Fact]
        public void LineCountIsPreserved()
        {
            var lines = new[] { "Sub A()", "If x Then _", "  y", "", "End Sub", "" };
            Assert.Equal(lines.Length, VbaIndenter.Format(lines).Length);
        }

        [Fact]
        public void KeywordPrefixedIdentifiersAreNotBlocks()
        {
            AssertFormat(@"
                |Sub A()
                |    DoEvents
                |    Forecast = 1
                |    Withdraw 5
                |    Next_Value = 2
                |    Ending = 3
                |    Exit Do
                |End Sub");
        }
    }
}
