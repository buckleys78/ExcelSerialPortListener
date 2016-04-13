using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerialPortListener {
    public class ExcelComms {
        private Excel.Workbook WkBook { get; }
        public string WkSheetName { get; }
        public string RngName { get; }

        [DllImport("Oleacc.dll")]
        static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out Excel.Window ptr);

        [DllImport("User32.dll")]
        public static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("User32.dll")]
        public static extern int GetClassName( int hWnd, StringBuilder lpClassName, int nMaxCount);

        public delegate bool EnumChildCallback(int hwnd, ref int lParam);

        public ExcelComms(string wkBookName, string wkSheetName, string rngName) {
            WkBook = WorkbookByName(wkBookName);
            WkSheetName = wkSheetName;
            RngName = rngName;
        }

        /// <summary>
        /// A function that returns the Excel.Workbook object that matches the passed Excel workbook file name.
        /// This function is substantially based on open-source code, not authored by me.
        /// However, none of the several sources that had this code clearly claimed original
        /// authorship, though I believe the author is Andrew Whitechapel. 
        /// @https://www.linkedin.com/in/andrew-whitechapel-083b75
        /// </summary>
        /// <param name="callingWkbkName"></param>
        /// <returns>Excel.Workbook</returns>
        public Excel.Workbook WorkbookByName(string callingWkbkName) {
            List<Process> processes = new List<Process>();
            processes.AddRange(Process.GetProcessesByName("excel"));

            foreach (Process p in processes) {
                int winHandle = (int)p.MainWindowHandle;
                //Console.WriteLine($"winHandle = {winHandle}");
                // We need to enumerate the child windows to find one that
                // supports accessibility. To do this, instantiate the
                // delegate and wrap the callback method in it, then call
                // EnumChildWindows, passing the delegate as the 2nd arg.
                if (winHandle != 0) {
                    int hwndChild = 0;
                    var cb = new EnumChildCallback(EnumChildProc);
                    EnumChildWindows(winHandle, cb, ref hwndChild);

                    // If we found an accessible child window, call
                    // AccessibleObjectFromWindow, passing the constant
                    // OBJID_NATIVEOM (defined in winuser.h) and
                    // IID_IDispatch - we want an IDispatch pointer
                    // into the native object model.
                    //Console.WriteLine($"hwndChild = {hwndChild}");
                    if (hwndChild != 0) {
                        const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                        Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

                        Excel.Window ptr = null;
                        int hr = AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), out ptr);
                        //Console.WriteLine($"hr ptr = {hr}");
                        if (hr >= 0) {
                            // If we successfully got a native OM
                            // IDispatch pointer, we can QI this for
                            // an Excel Application (using the implicit
                            // cast operator supplied in the PIA).
                            Excel.Application app = ptr.Application;
                            foreach (Excel.Workbook wkbk in app.Workbooks) {
                                if (wkbk.Name == callingWkbkName) {
                                    //Console.WriteLine($"Workbook name = {wkbk.Name}");
                                    return wkbk;
                                }
                            }
                        }
                    }
                }
            }
            //Console.WriteLine($"Failed to find Workbook named '{callingWkbkName}'");
            return null;
        }

        public bool WriteValueToWks(string valueToWrite) {
            if (WkBook == null) return false;
            try {
                WkBook.Worksheets[WkSheetName].Range[RngName].Value = valueToWrite;
                return true;
            }
            catch (Exception e) {
                //Console.WriteLine($"Failed to write value to Excel spreadsheet {WkBook?.Name}.{WkSheetName}.{RngName}, {e.Message}");
                return false;
            }
        }

        public static bool EnumChildProc(int hwndChild, ref int lParam) {
            StringBuilder buf = new StringBuilder(128);
            GetClassName(hwndChild, buf, 128);
            if (buf.ToString() == "EXCEL7") {
                lParam = hwndChild;
                return false;
            }
            return true;
        }
    }
}
