using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Woof.ProcessEx;

namespace Launcher {
    class Program {

        static int Main(string[] args) {
            if (args.Length < 1) return -1;
            switch (args[0]) {
                case "-c": Commit(); break;
                case "-i": AfterInstall(); break;
                case "-u1": BeforeUninstall(); break;
                case "-u2": AfterUninstall(); break;
                default: return -1;
            }
            return 0;
        }

        static void Commit() => Process.Start(new ProcessStartInfo {
            FileName = InstallerPath,
            Arguments = "-i",
            UseShellExecute = false,
            CreateNoWindow = true
        });


        static void AfterInstall() {
            SetUninstallString();
            using (var msiExecProcess = Process.GetProcessesByName("msiexec").SingleOrDefault(i => i.MainWindowTitle == "Installer Tests")) {
                if (msiExecProcess == null) return; // nothing to wait for.
                msiExecProcess.WaitForExit();
            }
            var demoPath = Path.Combine(InstallerDir, "Demo.exe");
            ProcessEx.Start(new ProcessStartInfo {
                FileName = demoPath,
                WorkingDirectory = InstallerDir,
                UseShellExecute = false,
                CreateNoWindow = false
            });
        }

        static void BeforeUninstall() {
            using (var process = Process.GetProcessesByName("Demo").FirstOrDefault()) { // try to close demo window...
                if (process != null) {
                    process.SendCloseRequest();
                    process.WaitForExit(2500);
                    if (!process.HasExited) {
                        process.Kill();
                        Thread.Sleep(250);
                    }
                }
            }
            Process.Start(new ProcessStartInfo("msiexec", $"/x{ProductCode} /passive"));
        }

        static void AfterUninstall() {
            // optional cleanup code...
        }

        /// <summary>
        /// Hacks the registry to make this installer start instead of Microsoft Windows Installer when Uninstall Program option is used.
        /// </summary>
        private static void SetUninstallString() {
            using (var k = Registry.LocalMachine.OpenSubKey(UninstallKey, true)) {
                if (k == null) return;
                k.SetValue("WindowsInstaller", 0);
                k.SetValue("UninstallString", $"{InstallerPath} -u1");
            }
        }

        #region Private data

        private const string ProductCode = "{FE35560B-E11C-4597-912B-3DDF3AC14F60}"; // See Setup project.
        private const string UninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + ProductCode;
        private const string CleanUpExe = "CleanUp.exe";
        private const uint WM_CLOSE = 0x0010;
        private readonly static string InstallerPath = Assembly.GetExecutingAssembly().Location;
        private readonly static string InstallerDir = Path.GetDirectoryName(InstallerPath);
        private readonly static string CleanUpPath = Path.Combine(InstallerDir, CleanUpExe);
        private readonly static string StartupDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Startup);

        #endregion


    }

}
