using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace Woof.ProcessEx {

    /// <summary>
    /// Special class allowing to create process as user from SYSTEM account context.
    /// </summary>
    public class ProcessEx : Process {

        /// <summary>
        /// Starts a process as user from SYSTEM account context, in user context behaves exactly as <see cref="Process.Start(ProcessStartInfo)"/>.
        /// IMPORTANT: UseShellExecute property of the <see cref="ProcessStartInfo"/> provided must be false!
        /// </summary>
        /// <param name="processStartInfo">
        /// The <see cref="System.Diagnostics.ProcessStartInfo"/> that contains the information that is 
        /// used to start the process, including the file name and any command-line arguments.</param>
        /// <returns>A new System.Diagnostics.Process that is associated with the process resource,
        /// or null if no process resource is started. Note that a new process that’s started
        /// alongside already running instances of the same process will be independent from
        /// the others. In addition, Start may return a non-null Process with its System.Diagnostics.Process.HasExited
        /// property already set to true. In this case, the started process may have activated
        /// an existing instance of itself and then exited.</returns>
        public static new Process Start(ProcessStartInfo processStartInfo) {
            if (IsSystemContext && !processStartInfo.UseShellExecute) {
                var process = new ProcessEx { StartInfo = processStartInfo };
                process.CreateProcessAsUser();
                return process._UserProcess;
            } else {
                return Process.Start(processStartInfo);
            }
        }

        #region Private properties

        /// <summary>
        /// Gets a value indicationg whether the current account is Windows System account.
        /// </summary>
        private static bool IsSystemContext {
            get {
                using (var identity = WindowsIdentity.GetCurrent()) return identity.IsSystem;
            }
        }

        #endregion

        #region Windows API part

        #region Constants and static readonly

        /// <summary>
        /// If this flag is set, the environment block pointed to by lpEnvironment uses Unicode characters. Otherwise, the environment block uses ANSI characters.
        /// </summary>
        private const int CREATE_UNICODE_ENVIRONMENT = 0x00000400;
        /// <summary>
        /// The process is a console application that is being run without a console window. Therefore, the console handle for the application is not set.
        /// This flag is ignored if the application is not a console application, or if it is used with either CREATE_NEW_CONSOLE or DETACHED_PROCESS.
        /// </summary>
        private const int CREATE_NO_WINDOW = 0x08000000;
        /// <summary>
        /// The new process has a new console, instead of inheriting its parent's console (the default).
        /// This flag cannot be used with DETACHED_PROCESS.
        /// </summary>
        private const int CREATE_NEW_CONSOLE = 0x00000010;
        /// <summary>
        /// Session identifier considered as invalid.
        /// </summary>
        private const uint INVALID_SESSION_ID = 0xFFFFFFFF;
        /// <summary>
        /// Handy null pointer.
        /// </summary>
        private static readonly IntPtr NULL = IntPtr.Zero;
        /// <summary>
        /// RD Session Host server on which the application is running.
        /// </summary>
        private static readonly IntPtr WTS_CURRENT_SERVER_HANDLE = IntPtr.Zero;

        #endregion

        #region Structs

        /// <summary>
        /// ShowWindow command enumeration.
        /// </summary>
        private enum SW {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            SW_HIDE = 0,
            /// <summary>
            /// Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
            /// </summary>
            SW_SHOWNORMAL = 1,
            /// <summary>
            /// Default window size and position.
            /// </summary>
            SW_NORMAL = 1,
            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            SW_SHOWMINIMIZED = 2,
            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>
            SW_SHOWMAXIMIZED = 3,
            /// <summary>
            /// Maximizes the specified window.
            /// </summary>
            SW_MAXIMIZE = 3,
            /// <summary>
            /// Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
            /// </summary>
            SW_SHOWNOACTIVATE = 4,
            /// <summary>
            /// Activates the window and displays it in its current size and position.
            /// </summary>
            SW_SHOW = 5,
            /// <summary>
            /// Minimizes the specified window and activates the next top-level window in the Z order.
            /// </summary>
            SW_MINIMIZE = 6,
            /// <summary>
            /// Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
            /// </summary>
            SW_SHOWMINNOACTIVE = 7,
            /// <summary>
            /// Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
            /// </summary>
            SW_SHOWNA = 8,
            /// <summary>
            /// Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
            /// </summary>
            SW_RESTORE = 9,
            /// <summary>
            /// Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
            /// </summary>
            SW_SHOWDEFAULT = 10,
            /// <summary>
            /// Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
            /// </summary>
            SW_MAX = 10
        }

        /// <summary>
        /// Specifies the connection state of a Remote Desktop Services session.
        /// </summary>
        private enum WTS_CONNECTSTATE_CLASS {
            /// <summary>
            /// A user is logged on to the WinStation.
            /// </summary>
            WTSActive,
            /// <summary>
            /// The WinStation is connected to the client.
            /// </summary>
            WTSConnected,
            /// <summary>
            /// The WinStation is in the process of connecting to the client.
            /// </summary>
            WTSConnectQuery,
            /// <summary>
            /// The WinStation is shadowing another WinStation.
            /// </summary>
            WTSShadow,
            /// <summary>
            /// The WinStation is active but the client is disconnected.
            /// </summary>
            WTSDisconnected,
            /// <summary>
            /// The WinStation is waiting for a client to connect.
            /// </summary>
            WTSIdle,
            /// <summary>
            /// The WinStation is listening for a connection. A listener session waits for requests for new client connections. No user is logged on a listener session. A listener session cannot be reset, shadowed, or changed to a regular client session.
            /// </summary>
            WTSListen,
            /// <summary>
            /// The WinStation is being reset.
            /// </summary>
            WTSReset,
            /// <summary>
            /// The WinStation is down due to an error.
            /// </summary>
            WTSDown,
            /// <summary>
            /// The WinStation is initializing.
            /// </summary>
            WTSInit
        }

        /// <summary>
        /// Contains information about a newly created process and its primary thread.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct PROCESS_INFORMATION {
            /// <summary>
            /// A handle to the newly created process. The handle is used to specify the process in all functions that perform operations on the process object.
            /// </summary>
            public IntPtr hProcess;
            /// <summary>
            /// A handle to the primary thread of the newly created process. The handle is used to specify the thread in all functions that perform operations on the thread object.
            /// </summary>
            public IntPtr hThread;
            /// <summary>
            /// A value that can be used to identify a process. The value is valid from the time the process is created until all handles to the process are closed and the process object is freed; at this point, the identifier may be reused.
            /// </summary>
            public uint dwProcessId;
            /// <summary>
            /// A value that can be used to identify a thread. The value is valid from the time the thread is created until all handles to the thread are closed and the thread object is freed; at this point, the identifier may be reused.
            /// </summary>
            public uint dwThreadId;
        }

        /// <summary>
        /// Contains values that specify security impersonation levels. Security impersonation levels govern the degree to which a server process can act on behalf of a client process.
        /// </summary>
        private enum SECURITY_IMPERSONATION_LEVEL {
            /// <summary>
            /// The server process cannot obtain identification information about the client, and it cannot impersonate the client. It is defined with no value given, and thus, by ANSI C rules, defaults to a value of zero.
            /// </summary>
            SecurityAnonymous = 0,
            /// <summary>
            /// The server process can obtain information about the client, such as security identifiers and privileges, but it cannot impersonate the client. This is useful for servers that export their own objects, for example, database products that export tables and views. Using the retrieved client-security information, the server can make access-validation decisions without being able to use other services that are using the client's security context.
            /// </summary>
            SecurityIdentification = 1,
            /// <summary>
            /// The server process can impersonate the client's security context on its local system. The server cannot impersonate the client on remote systems.
            /// </summary>
            SecurityImpersonation = 2,
            /// <summary>
            /// The server process can impersonate the client's security context on remote systems.
            /// </summary>
            SecurityDelegation = 3,
        }

        /// <summary>
        /// Specifies the window station, desktop, standard handles, and appearance of the main window for a process at creation time.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct STARTUPINFO {
            /// <summary>
            /// The size of the structure, in bytes.
            /// </summary>
            public int cb;
            /// <summary>
            /// Reserved; must be NULL.
            /// </summary>
            public String lpReserved;
            /// <summary>
            /// The name of the desktop, or the name of both the desktop and window station for this process. A backslash in the string indicates that the string includes both the desktop and window station names. For more information, see Thread Connection to a Desktop.
            /// </summary>
            public String lpDesktop;
            /// <summary>
            /// For console processes, this is the title displayed in the title bar if a new console window is created. If NULL, the name of the executable file is used as the window title instead. This parameter must be NULL for GUI or console processes that do not create a new console window.
            /// </summary>
            public String lpTitle;
            /// <summary>
            /// If dwFlags specifies STARTF_USEPOSITION, this member is the x offset of the upper left corner of a window if a new window is created, in pixels. Otherwise, this member is ignored.
            /// The offset is from the upper left corner of the screen. For GUI processes, the specified position is used the first time the new process calls CreateWindow to create an overlapped window if the x parameter of CreateWindow is CW_USEDEFAULT.
            /// </summary>
            public uint dwX;
            /// <summary>
            /// If dwFlags specifies STARTF_USEPOSITION, this member is the y offset of the upper left corner of a window if a new window is created, in pixels. Otherwise, this member is ignored.
            /// The offset is from the upper left corner of the screen. For GUI processes, the specified position is used the first time the new process calls CreateWindow to create an overlapped window if the y parameter of CreateWindow is CW_USEDEFAULT.
            /// </summary>
            public uint dwY;
            /// <summary>
            /// If dwFlags specifies STARTF_USESIZE, this member is the width of the window if a new window is created, in pixels. Otherwise, this member is ignored.
            /// For GUI processes, this is used only the first time the new process calls CreateWindow to create an overlapped window if the nWidth parameter of CreateWindow is CW_USEDEFAULT.
            /// </summary>
            public uint dwXSize;
            /// <summary>
            /// If dwFlags specifies STARTF_USESIZE, this member is the height of the window if a new window is created, in pixels. Otherwise, this member is ignored.
            /// For GUI processes, this is used only the first time the new process calls CreateWindow to create an overlapped window if the nHeight parameter of CreateWindow is CW_USEDEFAULT.
            /// </summary>
            public uint dwYSize;
            /// <summary>
            /// If dwFlags specifies STARTF_USECOUNTCHARS, if a new console window is created in a console process, this member specifies the screen buffer width, in character columns. Otherwise, this member is ignored.
            /// </summary>
            public uint dwXCountChars;
            /// <summary>
            /// If dwFlags specifies STARTF_USECOUNTCHARS, if a new console window is created in a console process, this member specifies the screen buffer height, in character rows. Otherwise, this member is ignored.
            /// </summary>
            public uint dwYCountChars;
            /// <summary>
            /// If dwFlags specifies STARTF_USEFILLATTRIBUTE, this member is the initial text and background colors if a new console window is created in a console application. Otherwise, this member is ignored.
            /// This value can be any combination of the following values: FOREGROUND_BLUE, FOREGROUND_GREEN, FOREGROUND_RED, FOREGROUND_INTENSITY, BACKGROUND_BLUE, BACKGROUND_GREEN, BACKGROUND_RED, and BACKGROUND_INTENSITY. For example, the following combination of values produces red text on a white background:
            /// FOREGROUND_RED | BACKGROUND_RED | BACKGROUND_GREEN | BACKGROUND_BLUE
            /// </summary>
            public uint dwFillAttribute;
            /// <summary>
            /// A bitfield that determines whether certain STARTUPINFO members are used when the process creates a window. This member can be one or more of the following values.
            /// </summary>
            public uint dwFlags;
            /// <summary>
            /// If dwFlags specifies STARTF_USESHOWWINDOW, this member can be any of the values that can be specified in the nCmdShow parameter for the ShowWindow function, except for SW_SHOWDEFAULT. Otherwise, this member is ignored.
            /// For GUI processes, the first time ShowWindow is called, its nCmdShow parameter is ignored wShowWindow specifies the default value. In subsequent calls to ShowWindow, the wShowWindow member is used if the nCmdShow parameter of ShowWindow is set to SW_SHOWDEFAULT.
            /// </summary>
            public short wShowWindow;
            /// <summary>
            /// Reserved for use by the C Run-time; must be zero.
            /// </summary>
            public short cbReserved2;
            /// <summary>
            /// Reserved for use by the C Run-time; must be NULL.
            /// </summary>
            public IntPtr lpReserved2;
            /// <summary>
            /// If dwFlags specifies STARTF_USESTDHANDLES, this member is the standard input handle for the process. If STARTF_USESTDHANDLES is not specified, the default for standard input is the keyboard buffer.
            /// If dwFlags specifies STARTF_USEHOTKEY, this member specifies a hotkey value that is sent as the wParam parameter of a WM_SETHOTKEY message to the first eligible top-level window created by the application that owns the process. If the window is created with the WS_POPUP window style, it is not eligible unless the WS_EX_APPWINDOW extended window style is also set. For more information, see CreateWindowEx.
            /// Otherwise, this member is ignored.
            /// </summary>
            public IntPtr hStdInput;
            /// <summary>
            /// If dwFlags specifies STARTF_USESTDHANDLES, this member is the standard output handle for the process. Otherwise, this member is ignored and the default for standard output is the console window's buffer.
            /// If a process is launched from the taskbar or jump list, the system sets hStdOutput to a handle to the monitor that contains the taskbar or jump list used to launch the process.For more information, see Remarks.
            /// Windows 7, Windows Server 2008 R2, Windows Vista, Windows Server 2008, Windows XP, and Windows Server 2003:  This behavior was introduced in Windows 8 and Windows Server 2012.
            /// </summary>
            public IntPtr hStdOutput;
            /// <summary>
            /// If dwFlags specifies STARTF_USESTDHANDLES, this member is the standard error handle for the process. Otherwise, this member is ignored and the default for standard error is the console window's buffer.
            /// </summary>
            public IntPtr hStdError;
        }

        /// <summary>
        /// Contains values that differentiate between a primary token and an impersonation token.
        /// </summary>
        private enum TOKEN_TYPE {
            /// <summary>
            /// Indicates a primary token.
            /// </summary>
            TokenPrimary = 1,
            /// <summary>
            /// Indicates an impersonation token.
            /// </summary>
            TokenImpersonation = 2
        }

        /// <summary>
        /// Contains information about a client session on a Remote Desktop Session Host (RD Session Host) server.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct WTS_SESSION_INFO {
            /// <summary>
            /// Session identifier of the session.
            /// </summary>
            public readonly UInt32 SessionID;
            /// <summary>
            /// Pointer to a null-terminated string that contains the WinStation name of this session. The WinStation name is a name that Windows associates with the session, for example, "services", "console", or "RDP-Tcp#0".
            /// </summary>
            [MarshalAs(UnmanagedType.LPStr)]
            public readonly String pWinStationName;
            /// <summary>
            /// A value from the <see cref="WTS_CONNECTSTATE_CLASS"/> enumeration type that indicates the session's current connection state.
            /// </summary>
            public readonly WTS_CONNECTSTATE_CLASS State;
        }

        #endregion

        #region Private data

        /// <summary>
        /// Process started in user context.
        /// </summary>
        private Process _UserProcess;

        /// <summary>
        /// <see cref="STARTUPINFO"/> instance.
        /// </summary>
        private STARTUPINFO _StartupInfo;

        /// <summary>
        /// <see cref="PROCESS_INFORMATION"/> instance.
        /// </summary>
        private PROCESS_INFORMATION _ProcessInformation;

        #endregion

        #region DLL Imports

        /// <summary>
        /// Windows API calls.
        /// </summary>
        private class NativeMethods {

            /// <summary>
            /// Creates a new process and its primary thread. The new process runs in the security context of the user represented by the specified token.
            /// </summary>
            /// <param name="hToken">A handle to the primary token that represents a user.</param>
            /// <param name="lpApplicationName">The name of the module to be executed.</param>
            /// <param name="lpCommandLine">The command line to be executed. The maximum length of this string is 32K characters. If lpApplicationName is NULL, the module name portion of lpCommandLine is limited to MAX_PATH characters.</param>
            /// <param name="lpProcessAttributes">A pointer to a SECURITY_ATTRIBUTES structure that specifies a security descriptor for the new process object and determines whether child processes can inherit the returned handle to the process. If lpProcessAttributes is NULL or lpSecurityDescriptor is NULL, the process gets a default security descriptor and the handle cannot be inherited. The default security descriptor is that of the user referenced in the hToken parameter. This security descriptor may not allow access for the caller, in which case the process may not be opened again after it is run. The process handle is valid and will continue to have full access rights.</param>
            /// <param name="lpThreadAttributes">A pointer to a SECURITY_ATTRIBUTES structure that specifies a security descriptor for the new thread object and determines whether child processes can inherit the returned handle to the thread. If lpThreadAttributes is NULL or lpSecurityDescriptor is NULL, the thread gets a default security descriptor and the handle cannot be inherited. The default security descriptor is that of the user referenced in the hToken parameter. This security descriptor may not allow access for the caller.</param>
            /// <param name="bInheritHandle">If this parameter is TRUE, each inheritable handle in the calling process is inherited by the new process. If the parameter is FALSE, the handles are not inherited. Note that inherited handles have the same value and access rights as the original handles. Terminal Services:  You cannot inherit handles across sessions. Additionally, if this parameter is TRUE, you must create the process in the same session as the caller.</param>
            /// <param name="dwCreationFlags">The flags that control the priority class and the creation of the process. For a list of values, see Process Creation Flags. This parameter also controls the new process's priority class, which is used to determine the scheduling priorities of the process's threads. For a list of values, see GetPriorityClass. If none of the priority class flags is specified, the priority class defaults to NORMAL_PRIORITY_CLASS unless the priority class of the creating process is IDLE_PRIORITY_CLASS or BELOW_NORMAL_PRIORITY_CLASS. In this case, the child process receives the default priority class of the calling process.</param>
            /// <param name="lpEnvironment">A pointer to an environment block for the new process. If this parameter is NULL, the new process uses the environment of the calling process.</param>
            /// <param name="lpCurrentDirectory">The full path to the current directory for the process. The string can also specify a UNC path. If this parameter is NULL, the new process will have the same current drive and directory as the calling process. (This feature is provided primarily for shells that need to start an application and specify its initial drive and working directory.)</param>
            /// <param name="lpStartupInfo">A pointer to a <see cref="STARTUPINFO"/> or STARTUPINFOEX structure.</param>
            /// <param name="lpProcessInformation">A pointer to a PROCESS_INFORMATION structure that receives identification information about the new process. Handles in PROCESS_INFORMATION must be closed with CloseHandle when they are no longer needed.</param>
            /// <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero.To get extended error information, call GetLastError.</returns>
            [DllImport("advapi32.dll", EntryPoint = "CreateProcessAsUser", SetLastError = true, CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
            public static extern bool CreateProcessAsUser(
                IntPtr hToken,
                String lpApplicationName,
                String lpCommandLine,
                IntPtr lpProcessAttributes,
                IntPtr lpThreadAttributes,
                bool bInheritHandle,
                uint dwCreationFlags,
                IntPtr lpEnvironment,
                String lpCurrentDirectory,
                ref STARTUPINFO lpStartupInfo,
                out PROCESS_INFORMATION lpProcessInformation);

            /// <summary>
            /// Creates a new access token that duplicates an existing token. This function can create either a primary token or an impersonation token.
            /// </summary>
            /// <param name="ExistingTokenHandle">A handle to an access token opened with TOKEN_DUPLICATE access.</param>
            /// <param name="dwDesiredAccess">Specifies the requested access rights for the new token. The DuplicateTokenEx function compares the requested access rights with the existing token's discretionary access control list (DACL) to determine which rights are granted or denied. To request the same access rights as the existing token, specify zero. To request all access rights that are valid for the caller, specify MAXIMUM_ALLOWED. For a list of access rights for access tokens, see Access Rights for Access-Token Objects.</param>
            /// <param name="lpThreadAttributes">A pointer to a SECURITY_ATTRIBUTES structure that specifies a security descriptor for the new token and determines whether child processes can inherit the token. If lpTokenAttributes is NULL, the token gets a default security descriptor and the handle cannot be inherited. If the security descriptor contains a system access control list (SACL), the token gets ACCESS_SYSTEM_SECURITY access right, even if it was not requested in dwDesiredAccess. To set the owner in the security descriptor for the new token, the caller's process token must have the SE_RESTORE_NAME privilege set.</param>
            /// <param name="TokenType">Specifies one of the following values from the TOKEN_TYPE enumeration.</param>
            /// <param name="ImpersonationLevel">Specifies a value from the <see cref="SECURITY_IMPERSONATION_LEVEL"/> enumeration that indicates the impersonation level of the new token.</param>
            /// <param name="DuplicateTokenHandle">A pointer to a variable that receives a handle to the duplicate token. This handle has TOKEN_IMPERSONATE and TOKEN_QUERY access to the new token. When you have finished using the new token, call the CloseHandle function to close the token handle.</param>
            /// <returns>If the function succeeds, the function returns a nonzero value. If the function fails, it returns zero.To get extended error information, call GetLastError.</returns>
            [DllImport("advapi32.dll", EntryPoint = "DuplicateTokenEx")]
            public static extern bool DuplicateTokenEx(
                IntPtr ExistingTokenHandle,
                uint dwDesiredAccess,
                IntPtr lpThreadAttributes,
                int TokenType,
                int ImpersonationLevel,
                ref IntPtr DuplicateTokenHandle);

            /// <summary>
            /// Retrieves the environment variables for the specified user. This block can then be passed to the CreateProcessAsUser function.
            /// </summary>
            /// <param name="lpEnvironment">When this function returns, receives a pointer to the new environment block. The environment block is an array of null-terminated Unicode strings. The list ends with two nulls (\0\0).</param>
            /// <param name="hToken">Token for the user, returned from the LogonUser function. If this is a primary token, the token must have TOKEN_QUERY and TOKEN_DUPLICATE access. If the token is an impersonation token, it must have TOKEN_QUERY access. For more information, see Access Rights for Access-Token Objects. If this parameter is NULL, the returned environment block contains system variables only.</param>
            /// <param name="bInherit">Specifies whether to inherit from the current process' environment. If this value is TRUE, the process inherits the current process' environment. If this value is FALSE, the process does not inherit the current process' environment.</param>
            /// <returns>TRUE if successful; otherwise, FALSE. To get extended error information, call GetLastError.</returns>
            [DllImport("userenv.dll", SetLastError = true)]
            public static extern bool CreateEnvironmentBlock(ref IntPtr lpEnvironment, IntPtr hToken, bool bInherit);

            /// <summary>
            /// Frees environment variables created by the CreateEnvironmentBlock function.
            /// </summary>
            /// <param name="lpEnvironment">Pointer to the environment block created by CreateEnvironmentBlock. The environment block is an array of null-terminated Unicode strings. The list ends with two nulls (\0\0).</param>
            /// <returns>TRUE if successful; otherwise, FALSE. To get extended error information, call GetLastError.</returns>
            [DllImport("userenv.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DestroyEnvironmentBlock(IntPtr lpEnvironment);

            /// <summary>
            /// Closes an open object handle.
            /// </summary>
            /// <param name="hObject">A valid handle to an open object.</param>
            /// <returns>TRUE if successful; otherwise, FALSE. To get extended error information, call GetLastError.</returns>
            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool CloseHandle(IntPtr hObject);

            /// <summary>
            /// Retrieves the session identifier of the console session. The console session is the session that is currently attached to the physical console. Note that it is not necessary that Remote Desktop Services be running for this function to succeed.
            /// </summary>
            /// <returns>The session identifier of the session that is attached to the physical console. If there is no session attached to the physical console, (for example, if the physical console session is in the process of being attached or detached), this function returns 0xFFFFFFFF.</returns>
            [DllImport("kernel32.dll")]
            public static extern uint WTSGetActiveConsoleSessionId();

            /// <summary>
            /// Obtains the primary access token of the logged-on user specified by the session ID. To call this function successfully, the calling application must be running within the context of the LocalSystem account and have the SE_TCB_NAME privilege.
            /// </summary>
            /// <param name="SessionId">A Remote Desktop Services session identifier. Any program running in the context of a service will have a session identifier of zero (0). You can use the WTSEnumerateSessions function to retrieve the identifiers of all sessions on a specified RD Session Host server. To be able to query information for another user's session, you need to have the Query Information permission. For more information, see Remote Desktop Services Permissions. To modify permissions on a session, use the Remote Desktop Services Configuration administrative tool.</param>
            /// <param name="phToken">If the function succeeds, receives a pointer to the token handle for the logged-on user. Note that you must call the CloseHandle function to close this handle.</param>
            /// <returns></returns>
            [DllImport("Wtsapi32.dll")]
            public static extern uint WTSQueryUserToken(uint SessionId, ref IntPtr phToken);

            /// <summary>
            /// Retrieves a list of sessions on a Remote Desktop Session Host (RD Session Host) server.
            /// </summary>
            /// <param name="hServer">A handle to the RD Session Host server.</param>
            /// <param name="Reserved">This parameter is reserved. It must be zero.</param>
            /// <param name="Version">The version of the enumeration request. This parameter must be 1.</param>
            /// <param name="ppSessionInfo">A pointer to an array of WTS_SESSION_INFO structures that represent the retrieved sessions. To free the returned buffer, call the WTSFreeMemory function.</param>
            /// <param name="pCount">A pointer to the number of WTS_SESSION_INFO structures returned in the ppSessionInfo parameter.</param>
            /// <returns></returns>
            [DllImport("wtsapi32.dll", SetLastError = true)]
            public static extern int WTSEnumerateSessions(
                IntPtr hServer,
                int Reserved,
                int Version,
                ref IntPtr ppSessionInfo,
                ref int pCount);

        }

        #endregion

        #region IDisposable implementaion

        /// <summary>
        /// Disposed unmanaged data.
        /// </summary>
        /// <param name="disposing">True if called from public <see cref="IDisposable.Dispose"/> method.</param>
        protected override void Dispose(bool disposing) {
            if (_UserProcess != null) {
                NativeMethods.CloseHandle(_ProcessInformation.hThread);
                NativeMethods.CloseHandle(_ProcessInformation.hThread);
                _UserProcess.Dispose();
            }
        }

        #endregion

        #region Win32 helpers

        /// <summary>
        /// Uses Win32 API CreateProcessAsUser to start defined process in current user's context.
        /// </summary>
        /// <returns>True if new process has started.</returns>
        private bool CreateProcessAsUser() {
            if (_UserProcess != null && !_UserProcess.HasExited) return false;
            var command = Path.GetFileNameWithoutExtension(StartInfo.FileName);
            var cmdLine = String.Join(" ", command, StartInfo.Arguments);
            var hUserToken = IntPtr.Zero;
            var pEnv = IntPtr.Zero;
            _StartupInfo = new STARTUPINFO() {
                cb = Marshal.SizeOf(typeof(STARTUPINFO)),
                wShowWindow = (short)(SW.SW_HIDE),
                lpDesktop = null // default active, "winsta0\\default" didn't work from Windows Installer.
            };
            _ProcessInformation = new PROCESS_INFORMATION();
            int iResultOfCreateProcessAsUser;
            try {
                if (!GetSessionUserToken(ref hUserToken)) throw new UnauthorizedAccessException("GetSessionUserToken failed.");
                uint dwCreationFlags = CREATE_UNICODE_ENVIRONMENT;
                if (StartInfo.CreateNoWindow) dwCreationFlags |= (uint)(CREATE_NO_WINDOW);
                if (!NativeMethods.CreateEnvironmentBlock(ref pEnv, hUserToken, false))
                    throw new InvalidOperationException("CreateEnvironmentBlock failed.");
                var ok = NativeMethods.CreateProcessAsUser(
                    hUserToken, // A handle to the primary token that represents a user
                    Path.Combine(StartInfo.WorkingDirectory, StartInfo.FileName), // The name of the module to be executed
                    cmdLine, // The command line to be executed
                    IntPtr.Zero, // A pointer to a SECURITY_ATTRIBUTES structure that specifies a security descriptor for the new process object and determines whether child processes can inherit the returned handle to the process
                    IntPtr.Zero, // A pointer to a SECURITY_ATTRIBUTES structure that specifies a security descriptor for the new thread object and determines whether child processes can inherit the returned handle to the thread
                    true, // If this parameter is TRUE, each inheritable handle in the calling process is inherited by the new process.
                    dwCreationFlags, // The flags that control the priority class and the creation of the process
                    pEnv, // A pointer to an environment block for the new process.
                    StartInfo.WorkingDirectory, // The full path to the current directory for the process
                    ref _StartupInfo, // A pointer to a STARTUPINFO or STARTUPINFOEX structure
                    out _ProcessInformation // A pointer to a PROCESS_INFORMATION structure that receives identification information about the new process
                );
                iResultOfCreateProcessAsUser = Marshal.GetLastWin32Error();
                if (!ok) throw new InvalidOperationException($"CreateProcessAsUser failed, Win32 error: {iResultOfCreateProcessAsUser}.");
                _UserProcess = GetProcessById((int)_ProcessInformation.dwProcessId);
            }
            finally {
                NativeMethods.CloseHandle(hUserToken);
                if (pEnv != IntPtr.Zero) NativeMethods.DestroyEnvironmentBlock(pEnv);
            }
            return true;
        }

        /// <summary>
        /// Gets the user token from the currently active session
        /// </summary>
        /// <param name="phUserToken">A pointer to user token structure.</param>
        /// <returns>True if successfull.</returns>
        private static bool GetSessionUserToken(ref IntPtr phUserToken) {
            var bResult = false;
            var hImpersonationToken = IntPtr.Zero;
            var activeSessionId = INVALID_SESSION_ID;
            var pSessionInfo = IntPtr.Zero;
            var sessionCount = 0;
            // Get a handle to the user access token for the current active session.
            if (NativeMethods.WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, ref pSessionInfo, ref sessionCount) != 0) {
                var arrayElementSize = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
                var current = (long)pSessionInfo;
                for (var i = 0; i < sessionCount; i++) {
                    var si = (WTS_SESSION_INFO)Marshal.PtrToStructure((IntPtr)current, typeof(WTS_SESSION_INFO));
                    current += (long)arrayElementSize;
                    if (si.State == WTS_CONNECTSTATE_CLASS.WTSActive) {
                        activeSessionId = si.SessionID;
                    }
                }
            }
            // If enumerating did not work, fall back to the old method
            if (activeSessionId == INVALID_SESSION_ID) {
                activeSessionId = NativeMethods.WTSGetActiveConsoleSessionId();
            }
            if (NativeMethods.WTSQueryUserToken(activeSessionId, ref hImpersonationToken) != 0) {
                // Convert the impersonation token to a primary token
                bResult = NativeMethods.DuplicateTokenEx(hImpersonationToken, 0, IntPtr.Zero,
                    (int)SECURITY_IMPERSONATION_LEVEL.SecurityImpersonation, (int)TOKEN_TYPE.TokenPrimary,
                    ref phUserToken);
                NativeMethods.CloseHandle(hImpersonationToken);
            }
            return bResult;
        }

        #endregion

        #endregion

    }

}