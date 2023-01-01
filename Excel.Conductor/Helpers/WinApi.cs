using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using System.Text;

namespace ChEJunkie.Office.Excel
{
    /// <summary> 
    /// Encapsulates P/Invoke methods.
    /// </summary>
    internal static class WinApi
    {
        /// <summary>
        /// The retrieved handle identifies the window above the specified window in the Z order.
        /// If the specified window is a topmost window, the handle identifies a topmost window.
        /// If the specified window is a top-level window, the handle identifies a top-level window.
        /// If the specified window is a child window, the handle identifies a sibling window.
        /// </summary>
        private const int GW_HWNDPREV = 3;

        #region Win32 API Automation

        // https://learn.microsoft.com/en-us/windows/win32/api/_automat/

        /// <summary>Given a ProgID, tetrieves a running object that has been registered with OLE.</summary>
        /// <param name="progID">The prog identifier.</param>
        /// <returns>Object.</returns>
        [System.Security.SecurityCritical]  // auto-generated_required
        public static Object GetActiveObject(String progID)
        {
            Guid clsid;

            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if CLSIDFromProgIDEx doesn't exist.
            try
            {
                WinApi.CLSIDFromProgIDEx(progID, out clsid);
            }
            catch (Exception)
            {
                WinApi.CLSIDFromProgID(progID, out clsid);
            }

            WinApi.GetActiveObject(ref clsid, IntPtr.Zero, out Object obj);
            return obj;
        }

        //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
        [DllImport("oleaut32.dll", PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        public static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

        #endregion Win32 API Automation

        #region Win32 API Component Object Model (COM)

        // https://learn.microsoft.com/en-us/windows/win32/api/_com/

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport("ole32.dll", PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        public static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        /// <summary>Looks up a CLSID in the registry, given a ProgID.</summary>
        /// <param name="progId">The prog identifier.</param>
        /// <param name="clsid">The CLSID.</param>
        [DllImport("ole32.dll", PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        public static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        #endregion Win32 API Component Object Model (COM)

        #region Win32 API Windows and Messages

        // https://learn.microsoft.com/en-us/windows/win32/api/_winmsg/

        /// <summary>Gets the name of the COM class to which the specified window belongs.</summary>
        /// <param name="windowHandle">The window handle.</param>
        /// <returns>System.String.</returns>
        public static string GetClassName(int windowHandle)
        {
            var buffer = new StringBuilder(128);
            GetClassName(windowHandle, buffer, 128);
            return buffer.ToString();
        }

        public static bool BringToFront(Process process)
        {
            if (process == null) throw new ArgumentNullException(nameof(process));

            var handle = process.MainWindowHandle;
            if (handle == IntPtr.Zero) return false;
            try
            {
                WinApi.SetForegroundWindow(handle);
                return true;
            }
            catch { return false; }
        }

        public static int ProcessIdFromWindowHandle(int windowHandle)
        {
            if (windowHandle == 0) throw new ArgumentOutOfRangeException("Window handle cannot be 0.", nameof(windowHandle));

            int processId;
            GetWindowThreadProcessId(windowHandle, out processId);
            return processId;
        }

        public static int GetWindowZ(int windowHandle)
        {
            var z = 0;
            //Count all windows above the starting window
            for (var h = new IntPtr(windowHandle);
                h != IntPtr.Zero;
                h = GetWindow(h, GW_HWNDPREV))
            {
                z++;
            }
            return z;
        }

        /// <summary>An application-defined callback function used with the EnumChildWindows function.
        /// It receives the child window handles.  The WNDENUMPROC type defines a pointer to this callback function.
        /// EnumChildProc is a placeholder for the application-defined function name.</summary>
        /// <param name="hwnd">A handle to the child window of the parent window specified in EnumChildWindows.</param>
        /// <param name="lParam">The application-defined value given in EnumChildWindows.</param>
        /// <returns>To continue enumeration, the callback function must return TRUE; to stop enumeration it must return FALSE.</returns>
        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        /// <summary>Retrieves the identifier of the thread that created the specified window and optionally,
        /// the identifier of the process that created the window.</summary>
        /// <param name="hWnd">A handle to the window.</param>
        /// <param name="lpdwProcessId">A pointer to a variable that receives the process identifier.
        /// If this parameter is not NULL, GetWindowThreadProcessId copies the identifier of the process
        /// to the variable; otherwise it does not.</param>
        /// <returns>The identifier of the thread that created the window.</returns>
        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        /// <summary>Enumerates the child windows that belong to the specified parent window by passing the handle to each child window, in turn,
        /// to an application-defined callback function. EnumChildWindows continues until the last child window is enumerated or
        /// the callback function returns false.</summary>
        /// <param name="hWndParent">A handle to the parent window whose child windows are to be enumerated. If this parameter is NULL,
        /// this function is equivalent to EnumWindows.</param>
        /// <param name="lpEnumFunc">A point to an application-defined callback function.</param>
        /// <param name="lParam">An application-defined value to be passed tot he callback function.</param>
        /// <returns>The return value is not used.</returns>
        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        /// <summary>Retrieves the name of the class to which the specified window belongs.</summary>
        /// <param name="hWnd">A handle to the window and, indirectly, the class to which the window belongs.</param>
        /// <param name="lpClassName">The class name string.</param>
        /// <param name="nMaxCount">The length of the lpClassName buffer, in characters. The buffer must be large enough to include
        /// the terminating null character; otherwise, the class name string is truncated to nMaxCount-1 characters.</param>
        /// <returns>If the function succeeds, the number of characters copied to the buffer not including the terminating null character;
        /// otherwise 0.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>Retrieves a handle to a window that has the specified relationship
        /// (Z-Order or owner) to the specified window.</summary>
        /// <param name="hWnd">A handle to a window. The window handle retrieved is relative to this window,
        /// based on the value of the uCmd parameter.</param>
        /// <param name="uCmd">The relationship between the specified window and the window whose handle is to be
        /// retrieved. This parameter can be one of the following values.
        /// (GW_CHILD, GW_ENABLEDPOPUP, GW_HWNDFIRST, GW_HWNDLAST, GW_HWNDNEXT, GWHWNDPREV, GW_OWNER)</param>
        /// <returns>If the function succeeds, a window handle; if no window exists with the specified relationship
        /// to the specified window, NULL.</returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        #endregion Win32 API Windows and Messages
    }
}