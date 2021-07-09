using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

namespace PhysicalVolumeMixer
{

    public partial class Form1 : Form, IAudioSessionEventsHandler
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        List<AppAudio> audioSessions = new();
        MMDeviceEnumerator deviceEnumerator;
        MMDevice device;
        SessionCollection sessions;

        readonly int ActiveWindowCooldown = 350;
        long lastGetCurrentWindowCall = 0;
        AppAudio activeWindowAudio = new();
        public Form1()
        {
            Application.ApplicationExit += Application_ApplicationExit;
            InitializeComponent();
            activeWindowAudio.Process = Process.GetProcessById(0);
            deviceEnumerator = new MMDeviceEnumerator();
            device = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            device.AudioSessionManager.OnSessionCreated += AudioSessionManager_OnSessionCreated;

            sessions = device.AudioSessionManager.Sessions;
            for (int i = 0; i < sessions.Count; i++)
            {
                sessions[i].RegisterEventClient(this);
                Process process = Process.GetProcessById((int)sessions[i].GetProcessID);
                audioSessions.Add(new AppAudio(process, sessions[i]));
                Debug.WriteLine(process.MainWindowTitle != "" ? process.MainWindowTitle : process.ProcessName + "    " + (int)sessions[i].GetProcessID);
            }

            Thread foregroundThread = new Thread(GetActiveWindow);
            foregroundThread.IsBackground = true;
            foregroundThread.Start();
            lastGetCurrentWindowCall = TimeSinceEpoch();
            UpdateComboBox();
        }

        private void Application_ApplicationExit(object sender, EventArgs e)
        {
            
        }

        private void AudioSessionManager_OnSessionCreated(object sender, IAudioSessionControl newSession)
        {
            Debug.WriteLine("New AudioSession");
            UpdateSessions();
        }

        private void UpdateSessions()
        {
            Debug.WriteLine("Reloading AudioSession");
            List<AppAudio> audioSessions2 = new();
            device.AudioSessionManager.RefreshSessions();
            sessions = device.AudioSessionManager.Sessions;
            audioSessions.Clear();
            for (int i = 0; i < sessions.Count; i++)
            {
                sessions[i].RegisterEventClient(this);
                Process process = Process.GetProcessById((int)sessions[i].GetProcessID);
                audioSessions.Add(new AppAudio(process, sessions[i]));
            }
            var ulist = audioSessions.Union(audioSessions2).ToList();
            audioSessions = ulist;
            UpdateComboBox();
        }

        private void UpdateComboBox()
        {
            audioSessions.Sort((e1, e2) =>
            {
                return -e2.Name.CompareTo(e1.Name);
            });
            audioSessions.Insert(0, activeWindowAudio);
            audioSessions[0].DisplayName = "Foreground window";
            if (comboBox1.InvokeRequired)
            {
                comboBox1.Invoke(new MethodInvoker(delegate
                {
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(audioSessions.ToArray());
                    comboBox1.DisplayMember = "DisplayName";
                }));
            }
            else
            {
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(audioSessions.ToArray());
                comboBox1.DisplayMember = "DisplayName";
            }
        }

        private void GetActiveWindow()
        {
            while (true)
            {
                if (lastGetCurrentWindowCall + ActiveWindowCooldown < TimeSinceEpoch())
                {
                    lastGetCurrentWindowCall = TimeSinceEpoch();
                    IntPtr hWnd = GetForegroundWindow(); // Get foreground window handle
                    uint activeWindowPtr;
                    GetWindowThreadProcessId(hWnd, out activeWindowPtr); // Get PID from window handle
                    int pID = Convert.ToInt32(activeWindowPtr);
                    if (pID != activeWindowAudio.Process.Id)
                    {
                        activeWindowAudio.SessionControl = GetActiveWindowAudioSession(Process.GetProcessById(pID).ProcessName);
                        if (activeWindowAudio.SessionControl is not null)
                        {
                            Debug.WriteLine(activeWindowAudio.SessionControl.SimpleAudioVolume.Volume);
                        }
                        activeWindowAudio.Process = Process.GetProcessById(pID);
                    }
                }
                else
                {
                }
                Thread.Sleep(10);
            }
        }

        private AudioSessionControl GetActiveWindowAudioSession(string processname)
        {
            try
            {
                var sessions = device.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    Process process = Process.GetProcessById((int)sessions[i].GetProcessID);
                    if (process.ProcessName == processname)
                    {
                        activeWindowAudio.Process = process;
                        return sessions[i];
                    }
                }
            }
            catch (Exception)
            {
            }
            return null;
        }

        private long TimeSinceEpoch()
        {
            return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds;
        }

        public void OnVolumeChanged(float volume, bool isMuted)
        {

        }

        public void OnDisplayNameChanged(string displayName)
        {

        }

        public void OnIconPathChanged(string iconPath)
        {
        }

        public void OnChannelVolumeChanged(uint channelCount, IntPtr newVolumes, uint channelIndex)
        {
        }

        public void OnGroupingParamChanged(ref Guid groupingId)
        {
        }

        public void OnStateChanged(AudioSessionState state)
        {
            if (state == AudioSessionState.AudioSessionStateExpired)
            {
                UpdateSessions();
            }
        }

        public void OnSessionDisconnected(AudioSessionDisconnectReason disconnectReason)
        {
            Debug.WriteLine(disconnectReason);
        }
    }
}
