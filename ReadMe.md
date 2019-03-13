# Woof.Process

## How to use (basics):


```csharp
using Woof.ProcessEx;

// To run program as current user:

ProcessEx.Start(new ProcessStartInfo {
    FileName = programFileName,
    WorkingDirectory = programWorkingDir,
    UseShellExecute = false // THIS ONE IS MANDATORY
});

// To 'gracefully' close other program:

using (var process = Process.GetProcessesByName(processName).FirstOrDefault()) { // try to close demo window...
    if (process != null) process.SendCloseRequest();
}
```

When not in a SYSTEM account context, or `UseShellExecute` is set `true`, `ProcessEx.Start()` behaves exactly as `Process.Start()`.

## To test this thing:

1. Build Installer.
2. Run Installer (Install).
3. When Installer is closed, a demo app should be run. Do not close it.
4. Use "Add / remove programs" to uninstall "Installer Tests".
5. The installer should close demo app and uninstall all program files.

## What does it prove:

The installer works on SYSTEM account.
The demo application should be run on user account.
The demo application process will not be forcibly "killed" but closed with WM_CLOSE message, exactly as it was requested to close by a user.

## Other points of interest:

Look at Launcher app how it handles installer events and introduces ones inaccessible by MSVS Installer Projects addon alone.

## Notes:

It replaces ProcessEx class from Woof.IPC.
Latest version of Woof.IPC will have the class removed.