using System;
using System.Runtime.Versioning;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    [SupportedOSPlatform("windows")]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new AudioAttenuationApp.Form());
    }
}
