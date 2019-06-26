using System;
using System.Windows.Forms;

namespace ThumbnailedBOM
{
    public class AppWindow :IWin32Window
    {
        public IntPtr Handle { get; set; }
    }
}
