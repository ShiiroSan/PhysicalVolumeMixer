using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.IO.Ports;
using System.Threading;

namespace PhysicalVolumeMixer
{
    class Serial
    {
        string arduinoLine = "";
        string arduPort = "";
        SerialPort serialPort = new();
        public string line = "";
        public Serial()
        {
        }

        public async Task GetArduinoThread()
        {
            await Task.Run(() =>
            {
                while (arduinoLine == "" && arduPort == "")
                {
                    GetArduino();
                    Thread.Sleep(20);
                }
                serialPort.PortName = arduPort;
                serialPort.Open();
            });
        }

        public void GetArduino()
        {
            List<string> arduinoElement = new();
            int arduinoItem = -1;
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
            {
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
                        arduinoLine = s;
                        arduinoItem = arduinoElement.IndexOf(s);
                        arduPort = s.Substring(0, 5).Replace(" ", string.Empty);
                    }
                }
            }
        }

        public void read()
        {
            while (true)
            {
                if (serialPort.IsOpen)
                {
                    line = serialPort.ReadLine();
                }
            }
        }
    }
}
