using System;
using System.Runtime.InteropServices;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using xlWindow = Microsoft.Office.Interop.Excel.Window;

namespace ChEJunkie.Office.Excel
{
    internal static class SessionUtil
    {
        private const uint WINDOW_OBJECT_ID = 0xFFFFFFF0;
        private static readonly byte[] WINDOW_INTERFACE_ID = new Guid("{00020400-0000-0000-C000-000000000046}").ToByteArray();

        public static xlApp AppFromMainWindowHandle(int mainWindowHandle)
        {
            if (mainWindowHandle == 0)
            {
                throw new ArgumentException("Window handle cannot be 0.", nameof(mainWindowHandle));
            }

            int childHandle = 0;
            WinApi.EnumChildWindows(mainWindowHandle, NextChildWindowHandle, ref childHandle);

            var win = ExcelWindowFromHandle(childHandle);

            return win.Application;
        }

        public static xlWindow ExcelWindowFromHandle(int handle)
        {
            AccessibleObjectFromWindow(handle, WINDOW_OBJECT_ID, WINDOW_INTERFACE_ID, out xlWindow result);
            return result;
        }

        /// <summary>Retrieves the address of the specified interface for the object associated with the specified window.</summary>
        /// <param name="hwnd">Specifies the handle of a window for which an object is to be retrieved.
        /// To retrieve an interface pointer to the cursor or caret object, specify NULL and use the appropriate ID in dwObjectID.</param>
        /// <param name="dwObjectID">Specifies the object ID. This value is one of the standard object identifier constants or a custom object ID
        /// such as OBJID_NATIVEOM, which is the object ID for the Office native object model.</param>
        /// <param name="riid">Specifies the reference identifier of the requested interface. This value is either IID_IAccessible or IID_Dispatch,
        /// but it can also be IID_IUnknown, or the IID of any interface that the object is expected to support.</param>
        /// <param name="ppvObject">Address of a pointer variable that receives the address of the specified interface.</param>
        /// <returns>If successful, returns S_OK; otherwise returns E_INVALIDARG, E_NOINTERFACE, or another standard COM error code.</returns>
        [DllImport("Oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out xlWindow ppvObject);

        private static bool NextChildWindowHandle(int currentChildHandle, ref int nextChildHandle)
        {
            const string excelClassName = "EXCEL7";
            //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ")");

            var result = true;

            var className = WinApi.GetClassName(currentChildHandle);
            // Debug.WriteLine(currentChildHandle + " ClassName: " + className);

            if (className == excelClassName)
            {
                nextChildHandle = currentChildHandle;
                result = false;
            }
            //  Debug.WriteLine(DateTime.Now.ToString("hh:mm:ss.fff") + " - NextChildWindowHandle(" + currentChildHandle + ", ref " + nextChildHandle + ") => " + result);
            return result;
        }
    }
}