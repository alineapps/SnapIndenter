using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace Kanna.AddIn
{
    /// <summary>
    /// VBE がアクティブなときだけショートカットキーを拾うスレッドローカルのキーボードフック。
    /// </summary>
    internal sealed class KeyboardHook : IDisposable
    {
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;
        private const int VK_SHIFT = 0x10, VK_CONTROL = 0x11, VK_MENU = 0x12;
        private const uint GA_ROOTOWNER = 3;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        private readonly VBE _vbe;
        private readonly FormatService _service;
        private HookProc _proc; // GC されないよう保持
        private IntPtr _hook;
        private MessageWindow _invoker;
        private IntPtr _vbeWindow;
        private Hotkey _currentKey;
        private Hotkey _moduleKey;

        public KeyboardHook(VBE vbe, FormatService service)
        {
            _vbe = vbe;
            _service = service;
            _service.SettingsChanged += (s, e) => LoadKeys();
        }

        public void Install()
        {
            LoadKeys();
            try { _vbeWindow = new IntPtr(_vbe.MainWindow.HWnd); } catch { }

            // フック内で直接整形せず、メッセージループ経由で実行する
            _invoker = new MessageWindow();

            _proc = Callback;
            _hook = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, GetCurrentThreadId());
            Log.Info($"Keyboard hook={(_hook != IntPtr.Zero)} err={Marshal.GetLastWin32Error()} vbeWindow={_vbeWindow}");
        }

        private void LoadKeys()
        {
            Hotkey.TryParse(_service.Settings.FormatCurrentHotkey, out _currentKey);
            Hotkey.TryParse(_service.Settings.FormatModuleHotkey, out _moduleKey);
        }

        private IntPtr Callback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code == HC_ACTION)
                {
                    uint flags = unchecked((uint)lParam.ToInt64());
                    bool keyDown = (flags & 0x80000000) == 0;
                    bool repeat = (flags & 0x40000000) != 0;
                    if (keyDown && !repeat)
                    {
                        var key = (Keys)wParam.ToInt32();
                        Action action = null;
                        if (Matches(_currentKey, key)) action = _service.FormatCurrent;
                        else if (Matches(_moduleKey, key)) action = _service.FormatModule;

                        if (action != null && IsVbeActive())
                        {
                            _invoker.Post(action);
                            return new IntPtr(1); // VBE にはキーを渡さない
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        private static bool Matches(Hotkey h, Keys key)
        {
            if (h == null || h.Key != key) return false;
            bool ctrl = GetKeyState(VK_CONTROL) < 0;
            bool shift = GetKeyState(VK_SHIFT) < 0;
            bool alt = GetKeyState(VK_MENU) < 0;
            return ctrl == h.Ctrl && shift == h.Shift && alt == h.Alt;
        }

        private bool IsVbeActive()
        {
            if (_vbeWindow == IntPtr.Zero) return false;
            IntPtr fg = GetForegroundWindow();
            return fg == _vbeWindow || GetAncestor(fg, GA_ROOTOWNER) == _vbeWindow;
        }

        public void Dispose()
        {
            if (_hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hook);
                _hook = IntPtr.Zero;
            }
            _invoker?.DestroyHandle();
            _invoker = null;
        }

        /// <summary>
        /// メッセージ専用ウィンドウ。PostMessage で処理を UI スレッドのメッセージループに後回しにする。
        /// （親のない WinForms Control の BeginInvoke は Office 内では届かないことがあった）
        /// </summary>
        private sealed class MessageWindow : NativeWindow
        {
            private const int WM_RUN = 0x8000 + 0x1F1; // WM_APP + α
            private static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);
            private readonly System.Collections.Generic.Queue<Action> _queue = new System.Collections.Generic.Queue<Action>();

            [DllImport("user32.dll")]
            private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            public MessageWindow()
            {
                CreateHandle(new CreateParams { Parent = HWND_MESSAGE });
            }

            public void Post(Action action)
            {
                _queue.Enqueue(action);
                if (!PostMessage(Handle, WM_RUN, IntPtr.Zero, IntPtr.Zero))
                    Log.Info("PostMessage failed: " + Marshal.GetLastWin32Error());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_RUN)
                {
                    while (_queue.Count > 0)
                    {
                        var action = _queue.Dequeue();
                        try { action(); }
                        catch (Exception ex) { Log.Error(ex); }
                    }
                    return;
                }
                base.WndProc(ref m);
            }
        }
    }
}
