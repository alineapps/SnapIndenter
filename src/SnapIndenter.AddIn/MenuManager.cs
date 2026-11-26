using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace SnapIndenter.AddIn
{
    /// <summary>VBE のメニューバーとコードウィンドウの右クリックメニューに項目を追加する。</summary>
    internal sealed class MenuManager : IDisposable
    {
        private const string Tag = "SnapIndenter.Menu";

        private readonly VBE _vbe;
        private readonly FormatService _service;
        private readonly List<CommandBarControl> _controls = new List<CommandBarControl>();
        // イベントオブジェクトは参照を保持しないと GC され、クリックが届かなくなる
        private readonly List<CommandBarEvents> _events = new List<CommandBarEvents>();
        private readonly Dictionary<string, Action> _actions = new Dictionary<string, Action>();
        // 設定変更（言語・ショートカット）を反映するため、表示文字列の取得方法をコントロールごとに覚えておく
        private readonly List<Action> _refreshers = new List<Action>();
        private int _seq;

        public MenuManager(VBE vbe, FormatService service)
        {
            _vbe = vbe;
            _service = service;
            _service.SettingsChanged += OnSettingsChanged;
        }

        public void Create()
        {
            RemoveLeftovers();

            CommandBar menuBar = FindMenuBar();
            if (menuBar != null)
            {
                int before = Math.Max(1, menuBar.Controls.Count - 1); // 「ウィンドウ」の前
                var popup = (CommandBarPopup)menuBar.Controls.Add(MsoControlType.msoControlPopup, Before: before, Temporary: true);
                popup.Tag = Tag;
                _controls.Add((CommandBarControl)popup);
                Bind(() => popup.Caption = L10n.T("整形(&J)", "&SnapIndenter"));

                AddButton(popup.Controls, () => L10n.T("プロシージャ/選択範囲を整形(&P)", "Indent &Procedure / Selection"), _service.FormatCurrent, false,
                    () => HotkeyText(_service.Settings.FormatCurrentHotkey));
                AddButton(popup.Controls, () => L10n.T("モジュールを整形(&M)", "Indent &Module"), _service.FormatModule, false,
                    () => HotkeyText(_service.Settings.FormatModuleHotkey));
                AddButton(popup.Controls, () => L10n.T("プロジェクト全体を整形(&A)", "Indent &All Modules in Project"), _service.FormatProject, false, null);
                AddButton(popup.Controls, () => L10n.T("直前の整形を元に戻す(&U)", "&Undo Last Indent"), _service.Undo, true, null);
                AddButton(popup.Controls, () => L10n.T("設定(&S)...", "&Settings..."), _service.ShowSettings, true, null);
            }

            CommandBar codeWindow = FindBar("Code Window");
            if (codeWindow != null)
            {
                var b1 = AddButton(codeWindow.Controls, () => L10n.T("プロシージャ/選択範囲を整形", "Indent Procedure / Selection"), _service.FormatCurrent, true, null);
                var b2 = AddButton(codeWindow.Controls, () => L10n.T("モジュールを整形", "Indent Module"), _service.FormatModule, false, null);
                b1.Tag = Tag;
                b2.Tag = Tag;
                _controls.Add(b1);
                _controls.Add(b2);
            }
        }

        private void Bind(Action refresh)
        {
            refresh();
            _refreshers.Add(refresh);
        }

        private void OnSettingsChanged(object sender, EventArgs e)
        {
            foreach (var refresh in _refreshers)
            {
                try { refresh(); }
                catch (Exception ex) { Log.Error(ex); }
            }
        }

        private CommandBarButton AddButton(CommandBarControls controls, Func<string> caption, Action action, bool beginGroup, Func<string> shortcut)
        {
            var button = (CommandBarButton)controls.Add(MsoControlType.msoControlButton, Temporary: true);
            button.Style = MsoButtonStyle.msoButtonCaption;
            button.BeginGroup = beginGroup;
            string id = Tag + "." + (++_seq);
            button.Tag = id;
            Bind(() =>
            {
                button.Caption = caption();
                if (shortcut != null) button.ShortcutText = shortcut() ?? "";
            });
            _actions[id] = action;

            CommandBarEvents ev = _vbe.Events.CommandBarEvents[button];
            ev.Click += OnClick;
            _events.Add(ev);
            return button;
        }

        private void OnClick(object commandBarControl, ref bool handled, ref bool cancelDefault)
        {
            try
            {
                var control = commandBarControl as CommandBarControl;
                if (control != null && _actions.TryGetValue(control.Tag, out var action))
                {
                    action();
                    handled = true;
                    cancelDefault = true;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }

        private static string HotkeyText(string hotkey)
            => Hotkey.TryParse(hotkey, out _) ? hotkey : null;

        private CommandBar FindMenuBar()
        {
            try
            {
                foreach (CommandBar bar in _vbe.CommandBars)
                {
                    if (bar.Type == MsoBarType.msoBarTypeMenuBar) return bar;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
            return FindBar("Menu Bar");
        }

        private CommandBar FindBar(string name)
        {
            try { return _vbe.CommandBars[name]; }
            catch { return null; }
        }

        /// <summary>前回異常終了したときなどに残ったメニューを削除する。</summary>
        private void RemoveLeftovers()
        {
            try
            {
                var found = _vbe.CommandBars.FindControls(Tag: Tag);
                if (found == null) return;
                var list = new List<CommandBarControl>();
                foreach (CommandBarControl c in found) list.Add(c);
                foreach (var c in list) c.Delete();
            }
            catch
            {
                // 見つからない場合は例外になることがある
            }
        }

        public void Dispose()
        {
            _service.SettingsChanged -= OnSettingsChanged;
            _refreshers.Clear();
            foreach (var ev in _events)
            {
                try { ev.Click -= OnClick; } catch { }
            }
            _events.Clear();
            foreach (var c in _controls)
            {
                try { c.Delete(); } catch { }
            }
            _controls.Clear();
            _actions.Clear();
        }
    }
}
