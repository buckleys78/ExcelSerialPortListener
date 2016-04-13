using System;
//using ExcelSerialPortListener = Microsoft.Office.Interop.Excel;
//using System.IO.Ports;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    class Program {
        static string response { get; set; }

        static void Main(string[] args) {
            Params parameters = new Params(args);
            // args: WkbookName, WkSheetName, Range, CommPort, BaudRate
            //for (var i=0; i < args.Length; i++) {
                //Console.WriteLine($"    Here is parameter{i}: {args[i]}");
            //}
            var scaleComms = new CommChannel(parameters.CommPort, parameters.Baudrate);
            bool commsAreOpen = scaleComms.OpenPort();
            //bool commsAreOpen = true;

            //Console.WriteLine($"Comms are open = {scaleComms.IsOpen}");

            if (commsAreOpen) {
                response = scaleComms.ReadData();
                //response = "new";
                //Console.WriteLine($"Scale response: {response}");
                ExcelComms excel = new ExcelComms(parameters.WorkbookName, parameters.WorksheetName, parameters.RangeName);
                excel.WriteValueToWks(response);
            }

            scaleComms.ClosePort();
            //Console.WriteLine($"and now Comms are open = {scaleComms.IsOpen}");
            //Console.ReadKey();
        }
    }
}
