using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace SnapIndenter.AddIn
{
    /// <summary>VBE から読み込まれる COM アドインのエントリポイント。</summary>
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdName)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class Connect : IDTExtensibility2
    {
        public const string ClassId = "0E25551F-0D2E-42D6-A7AD-1F172CD3E5A7";
        public const string ProgIdName = "SnapIndenter.Connect";

        // VBE が COM 参照を保持し続ける保証がないため、自分で強参照を持って GC を防ぐ
        private static Connect _instance;

        private VBE _vbe;
        private FormatService _service;
        private MenuManager _menus;
        private KeyboardHook _hook;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _instance = this;
                _vbe = (VBE)Application;
                Log.Info($"OnConnection mode={ConnectMode} host={SafeHostName()}");
                // 起動時読み込みの場合は UI の準備が整う OnStartupComplete まで待つ
                if (ConnectMode != ext_ConnectMode.ext_cm_Startup) Initialize();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try { Initialize(); }
            catch (Exception ex) { Log.Error(ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                _hook?.Dispose();
                _menus?.Dispose();
            }
            catch (Exception ex)
            {
                Log.Error(ex);
            }
            finally
            {
                Log.Info($"OnDisconnection mode={RemoveMode}");
                _instance = null;
                _hook = null;
                _menus = null;
                _service = null;
                _vbe = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void Initialize()
        {
            if (_menus != null || _vbe == null) return;
            _service = new FormatService(_vbe);
            _menus = new MenuManager(_vbe, _service);
            _menus.Create();
            _hook = new KeyboardHook(_vbe, _service);
            _hook.Install();
            Log.Info($"Initialized thread={System.Threading.Thread.CurrentThread.ManagedThreadId} apartment={System.Threading.Thread.CurrentThread.GetApartmentState()}");
        }

        private string SafeHostName()
        {
            try { return System.Diagnostics.Process.GetCurrentProcess().ProcessName; }
            catch { return "?"; }
        }
    }
}
