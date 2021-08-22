using System.Diagnostics;

namespace PhysicalVolumeMixer
{
    class AppAudio
    {
        Process process;

        //AudioSessionControl sessionControl;
        string name;
        string displayName;

        public Process Process
        {
            get => process;
            set => process = value;
        }

        //public AudioSessionControl SessionControl { get => sessionControl; set => sessionControl = value; }
        public string Name
        {
            get => name;
            set => name = value;
        }

        public string DisplayName
        {
            get => displayName;
            set => displayName = value;
        }

        public AppAudio(Process process /*, AudioSessionControl sessionControl*/)
        {
            Process = process;
            //SessionControl = sessionControl;
            if (process.ProcessName != "")
            {
                Name = process.ProcessName;
            }
            else
            {
                Name = "undefined";
            }

            if (displayName is null)
            {
                displayName = Name;
            }
        }

        public AppAudio()
        {
        }
    }
}