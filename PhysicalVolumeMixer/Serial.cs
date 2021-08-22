using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Management;
using System.IO.Ports;
using System.Threading;
using Microsoft.Toolkit.Uwp.Notifications;

namespace PhysicalVolumeMixer
{
    class Serial
    {
        string _arduinoLine = "";
        string _arduPort = "";
        readonly SerialPort _serialPort = new();
        public string Line = "";

        public async Task GetArduinoThread()
        {
            await Task.Run(() =>
            {
                while (_arduinoLine == "" && _arduPort == "")
                {
                    GetArduino();
                    Thread.Sleep(20);
                }

                _serialPort.PortName = _arduPort;
                _serialPort.Open();
            });
            new ToastContentBuilder()
                .AddText("Arduino card connected!")
                .AddText($"Port: {_serialPort.PortName}")
                .SetToastDuration(ToastDuration.Short)
                .Show();
        }

        private void GetArduino()
        {
            List<string> arduinoElement = new();
            using var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort");
            string[] portnames = SerialPort.GetPortNames();
            var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
            var tList = (from n in portnames
                join p in ports on n equals p["DeviceID"].ToString()
                select n + " - " + p["Caption"]).ToList();

            foreach (string s in tList)
            {
                arduinoElement.Add(s);
                if (s.Contains("Arduino"))
                {
                    _arduinoLine = s;
                    _arduPort = s.Substring(0, 5).Replace(" ", string.Empty);
                }
            }
        }

        public void Read()
        {
            while (true)
            {
                if (_serialPort.IsOpen)
                {
                    Line = _serialPort.ReadLine();
                }
            }
        }
    }
}