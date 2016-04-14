using System;
using System.Threading;

namespace ExcelSerialPortListener {
    class Program {
        public static string Response { get; set; } = string.Empty;
        static bool _gotResponse;
        private static CommChannel ScaleComms { get; set; }
        private static bool CommsAreOpen { get; set; }

        static void Main(string[] args) {
            Params parameters = new Params(args);
            // args: WkbookName, WkSheetName, Range, CommPort, BaudRate

            ScaleComms = new CommChannel(parameters.CommPort, parameters.Baudrate);

            CommsAreOpen = ScaleComms.OpenPort();
            if (CommsAreOpen) {
                var mainThread = new Thread(() => ListenToScale());
                var consoleKeyListener = new Thread(ListenerKeyBoardEvent);
                
                consoleKeyListener.Start();
                mainThread.Start();

                while (true) {
                    if (_gotResponse) {
                        mainThread?.Abort();
                        consoleKeyListener?.Abort();
                        break;
                    } 
                }

                ExcelComms excel = new ExcelComms(parameters.WorkbookName, parameters.WorksheetName, parameters.RangeName);
                excel.WriteValueToWks(Response);
            }

            ScaleComms.ClosePort();
        }

        public static void ListenerKeyBoardEvent() {
            do {
                if (Console.ReadKey(true).Key == ConsoleKey.Spacebar) {
                    Console.WriteLine("Saw pressed key!");
                    const string printCmd = "P\r";
                    ScaleComms.WriteData(printCmd);
                }
            } while (true);
        }

        public static void ListenToScale(double timeOutInSeconds = 30) {
            var timeOut = DateTime.Now.AddSeconds(timeOutInSeconds);
            var isTimedOut = false;
            do {
                if (Response.Length > 0)
                    break;
                Thread.Sleep(200);
                isTimedOut = DateTime.Now > timeOut;
            } while (!isTimedOut);

            Response = isTimedOut ? "Timed Out" : OnlyDigits(Response);
            _gotResponse = true;
        }

        private static string OnlyDigits(string s) {
            var onlyDigits = s.Trim();
            var indexOfSpaceG = onlyDigits.IndexOf(" g");
            if (indexOfSpaceG > 0)
                onlyDigits = onlyDigits.Substring(0, indexOfSpaceG);
            double tester;
            return double.TryParse(onlyDigits, out tester) ? onlyDigits : string.Empty;
        }
    }
}
