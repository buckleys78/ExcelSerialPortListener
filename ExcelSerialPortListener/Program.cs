using System;
using System.IO.Ports;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerialPortListener {
    class Program {
        static string response { get; set; }

        static void Main(string[] args) {
            CommChannel scaleComms = new CommChannel("COM3", "19200", "8", "One", "None");
            bool commsAreOpen = scaleComms.OpenPort();
            Console.WriteLine($"Comms are open = {scaleComms.IsOpen}");
            //Console.ReadKey();
            if (commsAreOpen)
                Console.WriteLine($"Scale response: {scaleComms.ReadData()}");
            scaleComms.ClosePort();
            Console.WriteLine($"and now Comms are open = {scaleComms.IsOpen}");
            Console.ReadKey();
        }

        //public static void Main() {

        //    SerialPort mySerialPort = new SerialPort("COM3");

        //    mySerialPort.BaudRate = 19200;
        //    mySerialPort.Parity = Parity.None;
        //    mySerialPort.StopBits = StopBits.One;
        //    mySerialPort.DataBits = 8;
        //    //mySerialPort.Handshake = Handshake.None;
        //    //mySerialPort.RtsEnable = true;

        //    mySerialPort.DataReceived += new SerialDataReceivedEventHandler(DataReceivedHandler);

        //    mySerialPort.Open();

        //    Console.WriteLine("Press the scale's Print button...");
        //    Console.WriteLine();
        //    Console.ReadKey();

        //    Console.WriteLine($".DataReceived = {response}");
        //    mySerialPort.Close();
        //    Console.ReadKey();
        //}

        //private static void DataReceivedHandler(
        //                    object sender,
        //                    SerialDataReceivedEventArgs e) {
        //    SerialPort sp = (SerialPort)sender;
        //    response = sp.ReadExisting();
        //    Console.WriteLine("Data Received:");
        //    Console.Write(response);
        //}
    }
}
