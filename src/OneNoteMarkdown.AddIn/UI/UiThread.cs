using System;
using System.Threading;
using System.Windows.Forms;
using OneNoteMarkdown.Logging;

namespace OneNoteMarkdown.UI
{
    public static class UiThread
    {
        private static Thread _thread;
        private static UiAnchorWindow _anchor;
        private static volatile SynchronizationContext _syncContext;
        private static readonly object _gate = new object();
        private static ManualResetEventSlim _ready;

        /// <summary>
        /// Start the background STA pump thread WITHOUT blocking the caller.
        /// Safe to call from OnConnection (OneNote's STA thread).
        /// </summary>
        public static void EnsureStarted()
        {
            lock (_gate)
            {
                if (_thread != null && _thread.IsAlive) return;

                _ready = new ManualResetEventSlim(false);
                _thread = new Thread(ThreadProc)
                {
                    IsBackground = true,
                    Name = "OneNoteMarkdown.UiThread"
                };
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
                // Do NOT Wait() here — blocking OneNote's STA thread causes crashes.
            }
        }

        private static void ThreadProc()
        {
            UiAnchorWindow anchor = null;
            WindowsFormsSynchronizationContext context = null;
            ManualResetEventSlim ready = _ready;
            try
            {
                DpiAwareness.EnablePerMonitorV2ForCurrentThread();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                context = new WindowsFormsSynchronizationContext();
                SynchronizationContext.SetSynchronizationContext(context);
                anchor = new UiAnchorWindow();
                _anchor = anchor;
                _syncContext = context;
                if (ready != null) ready.Set();
                Application.Run();
            }
            catch (Exception ex)
            {
                Logger.Error("UiThread message pump failed", ex);
                if (ready != null && !ready.IsSet) ready.Set();
            }
            finally
            {
                if (anchor != null)
                {
                    try { anchor.Dispose(); }
                    catch { }
                }
                if (context != null)
                {
                    try { context.Dispose(); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Post an action to the UiThread.  Blocks the caller briefly (up to
        /// 5 s) only when the thread is still initialising — this never happens
        /// on OneNote's STA thread because EnsureStarted() is called well before
        /// the first Post().
        /// </summary>
        public static void Post(Action action)
        {
            if (action == null) return;
            EnsureStarted();
            // Wait until the pump is ready (happens in milliseconds normally).
            ManualResetEventSlim ready = _ready;
            if (ready != null) ready.Wait(5000);
            SynchronizationContext ctx = _syncContext;
            if (ctx == null)
            {
                // Pump failed to initialise; run inline as a last resort.
                try { action(); }
                catch (Exception ex) { Logger.Error("UiThread inline fallback failed", ex); }
                return;
            }
            ctx.Post(_ => action(), null);
        }

        public static IWin32Window Anchor
        {
            get
            {
                EnsureStarted();
                ManualResetEventSlim ready = _ready;
                if (ready != null) ready.Wait(5000);
                return _anchor;
            }
        }

        /// <summary>
        /// Signals the pump thread to exit and waits for it to finish.
        /// Called during add-in shutdown (OnDisconnection) so the background
        /// thread does not prevent the process from exiting cleanly.
        /// </summary>
        public static void Shutdown()
        {
            Thread thread;
            SynchronizationContext context;
            ManualResetEventSlim ready;
            lock (_gate)
            {
                thread = _thread;
                context = _syncContext;
                ready = _ready;
                _thread = null;
                _syncContext = null;
                _anchor = null;
                _ready = null;
            }

            if (context != null)
            {
                try { context.Post(_ => Application.ExitThread(), null); }
                catch (Exception ex) { Logger.Error("UiThread Shutdown: exit pump failed", ex); }
            }
            if (thread != null && thread.IsAlive && !ReferenceEquals(Thread.CurrentThread, thread))
            {
                try { thread.Join(3000); }
                catch (Exception ex) { Logger.Error("UiThread Shutdown: thread join failed", ex); }
            }
            if (ready != null)
            {
                try { ready.Dispose(); }
                catch { }
            }
            Logger.Info("UiThread shut down");
        }

        private sealed class UiAnchorWindow : NativeWindow, IWin32Window, IDisposable
        {
            private const int WsPopup = unchecked((int)0x80000000);
            private const int WsExToolWindow = 0x00000080;
            private const int WsExNoActivate = 0x08000000;

            public UiAnchorWindow()
            {
                CreateHandle(new CreateParams
                {
                    Caption = "OneNoteMarkdown.UiAnchor",
                    Style = WsPopup,
                    ExStyle = WsExToolWindow | WsExNoActivate,
                    X = -32000,
                    Y = -32000,
                    Width = 1,
                    Height = 1
                });
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }
    }
}
