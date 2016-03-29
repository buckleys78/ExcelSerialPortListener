using System;
using System.IO.Ports;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    class Program {
        static void Main(string[] args) {
            CommChannel scaleComms = new CommChannel("COM3", "19200", "8", "One", "None");
            bool commsAreOpen = scaleComms.OpenPort();
            Console.WriteLine($"Comms are open = {scaleComms.IsOpen}");
            Console.ReadKey();
            scaleComms.ClosePort();
            Console.WriteLine($"and now Comms are open = {scaleComms.IsOpen}");
            Console.ReadKey();
        }
    }
}
