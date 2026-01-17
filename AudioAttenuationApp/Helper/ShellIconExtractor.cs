using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace AudioAttenuationApp.Helper
{
    [SupportedOSPlatform("windows")]
    static class ShellIconExtractor
    {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern uint ExtractIconEx(
            string szFileName,
            int nIconIndex,
            IntPtr[] phiconLarge,
            IntPtr[] phiconSmall,
            uint nIcons);

        [DllImport("user32.dll")]
        static extern bool DestroyIcon(IntPtr hIcon);

        public static Icon LoadIndirectIcon(string indirectPath)
        {
            if (string.IsNullOrWhiteSpace(indirectPath))
                return null;

            if (!indirectPath.StartsWith("@"))
                return null;

            // Example: "@%SystemRoot%\System32\AudioSrv.dll,-203"
            string raw = indirectPath.Substring(1);
            string[] parts = raw.Split(',');

            if (parts.Length != 2)
                return null;

            string modulePath = Environment.ExpandEnvironmentVariables(parts[0]);

            if (!int.TryParse(parts[1], out int iconIndex))
                return null;

            var large = new IntPtr[1];
            var small = new IntPtr[1];

            uint count = ExtractIconEx(
                modulePath,
                iconIndex,
                large,
                small,
                1);

            if (count == 0)
                return null;

            IntPtr hIcon = large[0] != IntPtr.Zero ? large[0] : small[0];
            if (hIcon == IntPtr.Zero)
                return null;

            // Clone to managed icon to avoid handle lifetime issues
            Icon icon = (Icon)Icon.FromHandle(hIcon).Clone();

            DestroyIcon(hIcon);

            return icon;
        }
    }
}