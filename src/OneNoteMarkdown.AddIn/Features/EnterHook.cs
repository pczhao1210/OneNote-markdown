using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OneNoteMarkdown.AddIn;
using OneNoteMarkdown.Logging;

namespace OneNoteMarkdown.Features
{
    internal static class EnterHook
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int VK_RETURN = 0x0D;
        private const int RenderDelayMilliseconds = 450;

        private static readonly object Gate = new object();
        private static IntPtr _hook;
        private static Thread _pumpThread;
        private static volatile bool _running;
        private static LowLevelKeyboardProc _proc;
        private static Form _pumpForm;
        private static System.Threading.Timer _renderTimer;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        public static void Install()
        {
            lock (Gate)
            {
                if (_hook != IntPtr.Zero) return;
                if (_running) return;

                _running = true;
                _proc = Callback;
                _pumpThread = new Thread(PumpThreadProc)
                {
                    IsBackground = true,
                    Name = "EnterHook.Pump"
                };
                _pumpThread.SetApartmentState(ApartmentState.STA);
                _pumpThread.Start();
            }
        }

        public static void Uninstall()
        {
            lock (Gate)
            {
                if (_hook != IntPtr.Zero)
                {
                    try { UnhookWindowsHookEx(_hook); }
                    catch (Exception ex) { Logger.Error("EnterHook UnhookWindowsHookEx failed", ex); }
                    _hook = IntPtr.Zero;
                }
                _running = false;

                if (_renderTimer != null)
                {
                    try { _renderTimer.Dispose(); }
                    catch { }
                    _renderTimer = null;
                }

                if (_pumpForm != null && !_pumpForm.IsDisposed)
                {
                    try { _pumpForm.Invoke((Action)(() => _pumpForm.Close())); }
                    catch { }
                }

                if (_pumpThread != null && _pumpThread.IsAlive)
                {
                    try { _pumpThread.Join(2000); }
                    catch { }
                }

                _pumpThread = null;
                _pumpForm = null;
                _proc = null;

                Logger.Info("EnterHook uninstalled");
            }
        }

        private static IntPtr ResolveModuleHandle()
        {
            // 1. Try pulling the add-in DLL's own module handle.
            IntPtr h = GetModuleHandle("OneNoteMarkdown.AddIn.dll");
            if (h != IntPtr.Zero) return h;

            // 2. Fall back to the host EXE (ONENOTE.EXE) — valid for WH_KEYBOARD_LL
            //    because the hook procedure is called in-process, not injected.
            h = GetModuleHandle(null);
            if (h != IntPtr.Zero) return h;

            // 3. Try the CRT GetHINSTANCE (may return -1 for managed-only modules).
            try { h = Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().ManifestModule); }
            catch { }
            if (h != IntPtr.Zero && h != (IntPtr)(-1)) return h;

            // 4. Last resort: NULL. The OS will use the current process's address
            //    space for validation, which works for WH_KEYBOARD_LL.
            return IntPtr.Zero;
        }

        private static void PumpThreadProc()
        {
            try
            {
                using (_pumpForm = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    ShowInTaskbar = false,
                    StartPosition = FormStartPosition.Manual,
                    Location = new System.Drawing.Point(-32000, -32000),
                    Size = new System.Drawing.Size(1, 1),
                    Opacity = 0,
                    Visible = false
                })
                {
                    _pumpForm.Load += (_, _) =>
                    {
                        IntPtr hMod = ResolveModuleHandle();
                        _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, hMod, 0);
                        if (_hook == IntPtr.Zero)
                        {
                            int err = Marshal.GetLastWin32Error();
                            Logger.Error("EnterHook SetWindowsHookEx(WH_KEYBOARD_LL) failed, err=" + err +
                                ", hMod=" + hMod.ToString("X"), null);
                            _running = false;
                            _pumpForm.Close();
                            return;
                        }

                        Logger.Info("EnterHook installed (global LL hook, hMod=" + hMod.ToString("X") + ")");
                    };

                    Application.Run(_pumpForm);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("EnterHook pump thread crashed", ex);
            }
            finally
            {
                if (_hook != IntPtr.Zero)
                {
                    try { UnhookWindowsHookEx(_hook); }
                    catch { }
                    _hook = IntPtr.Zero;
                }
                Logger.Info("EnterHook pump thread exited");
            }
        }

        private static IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (int)wParam == WM_KEYDOWN)
            {
                try
                {
                    KBDLLHOOKSTRUCT ks = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                    if (ks.vkCode == VK_RETURN && LivePreviewService.IsEnabled && IsOneNoteForeground())
                    {
                        Logger.Info("EnterHook: Enter accepted; scheduling render");
                        ScheduleRenderAfterOneNoteCommit();
                    }
                }
                catch { }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        private static void ScheduleRenderAfterOneNoteCommit()
        {
            lock (Gate)
            {
                if (_renderTimer != null)
                {
                    try { _renderTimer.Dispose(); }
                    catch { }
                }

                _renderTimer = new System.Threading.Timer(delegate
                {
                    if (!LivePreviewService.IsEnabled) return;
                    Connect.PostToOneNoteThread(RenderCurrentLineCommand.Execute);
                }, null, RenderDelayMilliseconds, Timeout.Infinite);
            }
        }

        private static bool IsOneNoteForeground()
        {
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;

            uint processId;
            GetWindowThreadProcessId(foreground, out processId);
            if (IsOneNoteProcess(processId)) return true;

            IntPtr rootOwner = GetAncestor(foreground, GA_ROOTOWNER);
            if (rootOwner != IntPtr.Zero && rootOwner != foreground)
            {
                GetWindowThreadProcessId(rootOwner, out processId);
                if (IsOneNoteProcess(processId)) return true;
            }

            return false;
        }

        private static bool IsOneNoteProcess(uint processId)
        {
            if (processId == 0) return false;
            if (processId == (uint)System.Diagnostics.Process.GetCurrentProcess().Id) return true;

            try
            {
                string processName = System.Diagnostics.Process.GetProcessById((int)processId).ProcessName;
                return string.Equals(processName, "ONENOTE", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private const uint GA_ROOTOWNER = 3;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
    }
}
