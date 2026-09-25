using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OneNoteMarkdown.UI
{
    internal static class FileDialogHost
    {
        public static DialogResult ShowOpen(OpenFileDialog dialog)
        {
            return dialog.ShowDialog(GetOwner());
        }

        public static DialogResult ShowSave(SaveFileDialog dialog)
        {
            return dialog.ShowDialog(GetOwner());
        }

        private static IWin32Window GetOwner()
        {
            IntPtr foreground = GetForegroundWindow();
            return foreground == IntPtr.Zero
                ? UiThread.Anchor
                : new WindowWrapper(foreground);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }
}
