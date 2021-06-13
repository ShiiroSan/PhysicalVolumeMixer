using CSCore.CoreAudioAPI;
using PhysicalVolumeMixer.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PhysicalVolumeMixer
{
    static class Program
    {
        static void Main(string[] args)
        {
            /*
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            */
            // dump all audio devices
            /*
            foreach (AudioDevice device in AudioUtilities.GetAllDevices())
            {
                AudioUtilities.GetSpeakersDevice();
                Console.WriteLine("Device: " + device.FriendlyName);
            }

            List<AudioSession> audioSessions = new List<AudioSession>();
            // dump all audio sessions and only get unique one
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null)
                {
                    if (audioSessions.Count == 0)
                    {
                        audioSessions.Add(session);
                    }
                    foreach (AudioSession alreadyGroupedSession in audioSessions)
                    {
                        if (alreadyGroupedSession.GroupingParam == session.GroupingParam)
                        {
                            break;
                        }
                        else
                        {
                            audioSessions.Add(session);
                            break;
                        }
                    }
                }
            }
            foreach (AudioSession session in audioSessions)
            {
                // only the one associated with a defined process
                try
                {
                    if (session.DisplayName != "")
                    {
                        Console.WriteLine("Process: " + session.DisplayName);
                    }
                    else
                    {
                        Console.WriteLine("Process: " + session.Process.ProcessName);
                    }
                    Console.WriteLine(session.State);
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            Console.ReadKey();
            */

            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    
                    foreach (var session in sessionEnumerator)
                    {
                        var session2 = session.QueryInterface<AudioSessionControl2>();
                        if (string.IsNullOrWhiteSpace(session2.DisplayName))
                        {
                            //Debug.WriteLine(session2.Process.ProcessName);
                            using (var proc = Process.GetProcessById(session2.ProcessID))
                            {
                                session2.DisplayName = proc.MainWindowTitle;
                            }
                            var shellItem = Shell32.SHCreateItemInKnownFolder(FolderIds.AppsFolder, Shell32.KF_FLAG_DONT_VERIFY, appId, typeof(IShellItem2).GUID);
                            DisplayName = shellItem.GetString(ref PropertyKeys.PKEY_ItemNameDisplay);
                        }
                        Debug.WriteLine("Process: " + session2.DisplayName);
                        using (var audioMeterInformation = session2.QueryInterface<AudioMeterInformation>())
                        {
                            Console.WriteLine(audioMeterInformation.GetPeakValue());
                        }
                    }
                }
            }
        }

        private static AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    Debug.WriteLine("DefaultDevice: " + device.FriendlyName);
                    var sessionManager = AudioSessionManager2.FromMMDevice(device);
                    return sessionManager;
                }
            }
        }
    }
}