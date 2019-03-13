using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace Woof.ProcessEx {

    /// <summary>
    /// Extends the Process class providing methods to communicate with processes.
    /// </summary>
    public static class ProcessExtensions {

        /// <summary>
        /// Sends a WM_CLOSE message to the process.
        /// This asks the process politely to shut down properly.
        /// </summary>
        /// <param name="process">Process.</param>
        /// <returns>True if successful.</returns>
        public static bool SendCloseRequest(this Process process) {
            const uint WM_CLOSE = 0x0010;
            return process.PostThreadMessage(WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        /// <summary>
        /// Sends a message to the first enumerated window in the first enumerated thread with at least one window, and returns the handle of that window through the hwnd output parameter if such a window was enumerated.  If a window was enumerated, the return value is the return value of the SendMessage call, otherwise the return value is zero.
        /// </summary>
        /// <param name="p">Process.</param>
        /// <param name="msg">The message to be sent.</param>
        /// <param name="wParam">Additional message-specific information.</param>
        /// <param name="lParam">Additional message-specific information.</param>
        /// <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
        public static IntPtr SendMessage(this Process p, UInt32 msg, IntPtr wParam, IntPtr lParam) {
            var hwnd = p.GetWindowHandles().FirstOrDefault();
            if (hwnd != IntPtr.Zero)
                return NativeMethods.SendMessage(hwnd, msg, wParam, lParam);
            else
                return IntPtr.Zero;
        }

        /// <summary>
        /// Posts a message to the first enumerated window in the first enumerated thread with at least one window, and returns the handle of that window through the hwnd output parameter if such a window was enumerated.  If a window was enumerated, the return value is the return value of the PostMessage call, otherwise the return value is false.
        /// </summary>
        /// <param name="p">Process.</param>
        /// <param name="msg">The message to be sent.</param>
        /// <param name="wParam">Additional message-specific information.</param>
        /// <param name="lParam">Additional message-specific information.</param>
        /// <returns>True if successfull.</returns>
        public static bool PostMessage(this Process p, UInt32 msg, IntPtr wParam, IntPtr lParam) {
            var hwnd = p.GetWindowHandles().FirstOrDefault();
            if (hwnd != IntPtr.Zero)
                return NativeMethods.PostMessage(hwnd, msg, wParam, lParam);
            else
                return false;
        }

        /// <summary>
        /// Posts a thread message to the first enumerated thread (when ensureTargetThreadHasWindow is false), or posts a thread message to the first enumerated thread with a window, unless no windows are found in which case the call fails.  If an appropriate thread was found, the return value is the return value of PostThreadMessage call, otherwise the return value is false.
        /// </summary>
        /// <param name="p">Process.</param>
        /// <param name="msg">The message to be sent.</param>
        /// <param name="wParam">Additional message-specific information.</param>
        /// <param name="lParam">Additional message-specific information.</param>
        /// <param name="ensureTargetThreadHasWindow">Set true for threads having windows.</param>
        /// <returns>True if successfull.</returns>
        public static bool PostThreadMessage(this Process p, UInt32 msg, IntPtr wParam, IntPtr lParam, bool ensureTargetThreadHasWindow = false) {
            uint targetThreadId = 0;
            if (ensureTargetThreadHasWindow) {
                IntPtr hwnd = p.GetWindowHandles().FirstOrDefault();
                if (hwnd != IntPtr.Zero) targetThreadId = NativeMethods.GetWindowThreadProcessId(hwnd, out _);
            }
            else {
                targetThreadId = (uint)p.Threads[0].Id;
            }
            if (targetThreadId != 0)
                return NativeMethods.PostThreadMessage(targetThreadId, msg, wParam, lParam);
            else
                return false;
        }

        /// <summary>
        /// Enumerates window handles of the process.
        /// </summary>
        /// <param name="process">Process.</param>
        /// <returns>Handles enumeration.</returns>
        public static IEnumerable<IntPtr> GetWindowHandles(this Process process) {
            var handles = new List<IntPtr>();
            foreach (ProcessThread thread in process.Threads)
                NativeMethods.EnumThreadWindows((uint)thread.Id, (hWnd, lParam) => { handles.Add(hWnd); return true; }, IntPtr.Zero);
            return handles;
        }

        /// <summary>
        /// Windows API (user32.dll) calls.
        /// </summary>
        private static class NativeMethods {

            /// <summary>
            /// Sends the specified message to a window or windows. The SendMessage function calls the window procedure for the specified window and does not return until the window procedure has processed the message.
            /// </summary>
            /// <param name="hWnd">A handle to the window whose window procedure will receive the message. If this parameter is HWND_BROADCAST ((HWND)0xffff), the message is sent to all top-level windows in the system, including disabled or invisible unowned windows, overlapped windows, and pop-up windows; but the message is not sent to child windows.</param>
            /// <param name="Msg">The message to be sent.</param>
            /// <param name="wParam">Additional message-specific information.</param>
            /// <param name="lParam">Additional message-specific information.</param>
            /// <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

            /// <summary>
            /// Places (posts) a message in the message queue associated with the thread that created the specified window and returns without waiting for the thread to process the message.
            /// </summary>
            /// <param name="hWnd">A handle to the window whose window procedure will receive the message. If this parameter is HWND_BROADCAST ((HWND)0xffff), the message is sent to all top-level windows in the system, including disabled or invisible unowned windows, overlapped windows, and pop-up windows; but the message is not sent to child windows.</param>
            /// <param name="Msg">The message to be sent.</param>
            /// <param name="wParam">Additional message-specific information.</param>
            /// <param name="lParam">Additional message-specific information.</param>
            /// <returns>If the function succeeds, the return value is nonzero.</returns>
            [return: MarshalAs(UnmanagedType.Bool)]
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

            /// <summary>
            /// Posts a message to the message queue of the specified thread. It returns without waiting for the thread to process the message.
            /// </summary>
            /// <param name="threadId">The identifier of the thread to which the message is to be posted.</param>
            /// <param name="msg">The type of message to be posted.</param>
            /// <param name="wParam">Additional message-specific information.</param>
            /// <param name="lParam">Additional message-specific information.</param>
            /// <returns></returns>
            [return: MarshalAs(UnmanagedType.Bool)]
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

            /// <summary>
            /// Enumerates all nonchild windows associated with a thread by passing the handle to each window, in turn, to an application-defined callback function. EnumThreadWindows continues until the last window is enumerated or the callback function returns FALSE. To enumerate child windows of a particular window, use the EnumChildWindows function.
            /// </summary>
            /// <param name="dwThreadId">The identifier of the thread whose windows are to be enumerated.</param>
            /// <param name="lpfn">A pointer to an application-defined callback function. For more information, see EnumThreadWndProc.</param>
            /// <param name="lParam">An application-defined value to be passed to the callback function.</param>
            /// <returns></returns>
            [DllImport("user32.dll")]
            public static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

            /// <summary>
            /// Retrieves the identifier of the thread that created the specified window and, optionally, the identifier of the process that created the window.
            /// </summary>
            /// <param name="hWnd">A handle to the window.</param>
            /// <param name="lpdwProcessId">A pointer to a variable that receives the process identifier. If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process to the variable; otherwise, it does not.</param>
            /// <returns>The return value is the identifier of the thread that created the window.</returns>
            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

            /// <summary>
            /// An application-defined callback function used with the EnumThreadWindows function. It receives the window handles associated with a thread. The WNDENUMPROC type defines a pointer to this callback function. EnumThreadWndProc is a placeholder for the application-defined function name.
            /// </summary>
            /// <param name="hWnd">A handle to a window associated with the thread specified in the EnumThreadWindows function.</param>
            /// <param name="lParam">The application-defined value given in the EnumThreadWindows function.</param>
            /// <returns></returns>
            public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        }

    }

}