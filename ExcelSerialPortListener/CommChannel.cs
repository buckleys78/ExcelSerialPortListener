using System;
using System.IO.Ports;
using System.Threading;

namespace ExcelSerialPortListener {
    public class CommChannel {
        public string PortName { get; set; }
        public string BaudRate { get; set; }
        public string Parity { get; set; }
        public string DataBits { get; set; }
        public string StopBits { get; set; }
        public SerialPort CommPort = new SerialPort();
        public bool IsOpen => CommPort.IsOpen;
        public string Response { get; set; } = string.Empty;

    //=== Constructor(s) ===
        public CommChannel(string portName = "COM1", string baudRate = "9600", 
                           string dataBits = "8", string stopBits = "One", string parity = "None") {
            PortName = portName;
            BaudRate = baudRate;
            DataBits = dataBits;
            StopBits = stopBits;        //None, One, OnePointFive, Two
            Parity = parity;            //Even, Mark, None, Odd, Space
            ConfigurePort();
        }

        // === Methods ===
        private void ConfigurePort() {
            if(CommPort.IsOpen) CommPort.Close();
            CommPort.PortName = PortName;
            CommPort.BaudRate = int.Parse(BaudRate);
            CommPort.DataBits = int.Parse(DataBits);
            CommPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), StopBits, ignoreCase: true);
            CommPort.Parity = (Parity)Enum.Parse(typeof(Parity), Parity, ignoreCase: true);
            CommPort.ReceivedBytesThreshold = 11;
            //CommPort.Handshake = Handshake.None;
            //CommPort.RtsEnable = true;
            // add listener event handler
            CommPort.DataReceived += new SerialDataReceivedEventHandler(SerialDeviceDataReceivedHandler);
        }

        public void ClosePort() {
            if (IsOpen) CommPort.Close();
        }

        public bool OpenPort() {
            try {
                CommPort.Open();
                return true;
            } catch {
                return false;
            }
        }

        public string ReadData(double timeOutInSeconds = 30) {
            DateTime timeOut = DateTime.Now.AddSeconds(timeOutInSeconds);
            bool isTimedOut = false;
            do {
                if (Response.Length > 0)
                    break;
                Thread.Sleep(200);
                isTimedOut = DateTime.Now > timeOut;
            } while (!isTimedOut);

            if (isTimedOut) {
                return "Timed Out";
            } else {
                return OnlyDigits(Response);
            }
        }

        public void WriteData(string dataString) {
            if (!IsOpen)
                CommPort.Open();
            CommPort.Write(dataString);
        }

        private string OnlyDigits(string s) {
            string onlyDigits = s.Trim();
            int indexOfSpaceG = onlyDigits.IndexOf(" g");
            if (indexOfSpaceG > 0)
                onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
            double tester;
            if (double.TryParse(onlyDigits, out tester)) {
                return onlyDigits;
            } else {
                return string.Empty;
            }
        }

        private void SerialDeviceDataReceivedHandler(object sender, SerialDataReceivedEventArgs e) {
            SerialPort sp = (SerialPort)sender;
            Response = sp.ReadExisting();
        }
    }
}
