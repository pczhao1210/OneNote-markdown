using System;
using System.Runtime.InteropServices;

namespace OneNoteMarkdown.UI
{
    internal static class DpiAwareness
    {
        private static readonly IntPtr PerMonitorAwareV2 = new IntPtr(-4);

        internal static bool EnablePerMonitorV2ForCurrentThread()
        {
            try
            {
                return SetThreadDpiAwarenessContext(PerMonitorAwareV2) != IntPtr.Zero;
            }
            catch (EntryPointNotFoundException)
            {
                return false;
            }
            catch (DllNotFoundException)
            {
                return false;
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);
    }
}
