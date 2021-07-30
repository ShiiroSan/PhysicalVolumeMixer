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

        Serial serial = new();

        List<ComboBox> comboBoxes = new();
        public Form1()
        {
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
                audioSessions.Add(new AppAudio(process));
                if (sessions[i].IsSystemSoundsSession)
                {
                    audioSessions[i].DisplayName = "System Sounds";
                }
                Debug.WriteLine(process.MainWindowTitle != "" ? process.MainWindowTitle : process.ProcessName + "    " + (int)sessions[i].GetProcessID);
            }
            Task task = Task.Run(async () => await serial.GetArduinoThread());

            Thread serialRead = new Thread(serial.read);
            serialRead.IsBackground = true;
            serialRead.Start();
            Thread foregroundThread = new Thread(GetActiveWindow);
            foregroundThread.IsBackground = true;
            foregroundThread.Start();
            Thread updateAudio = new Thread(UpdateAudio);
            updateAudio.IsBackground = true;
            updateAudio.Start();
            lastGetCurrentWindowCall = TimeSinceEpoch();
            BindingSource bindingSource = new();
            bindingSource.DataSource = audioSessions;
            comboBoxes.Add(comboBox1);
            comboBoxes.Add(comboBox2);
            comboBoxes.Add(comboBox3);
            comboBoxes.Add(comboBox4);
            comboBoxes.Add(comboBox5);

            UpdateComboBox();
        }

        string prevLine = "";
        //line ref: {SliderVolume1}|{SliderVolume2}|{SliderVolume3}|{SliderVolume4}|{SliderVolume5}
        private void UpdateAudio()
        {
            while (true)
            {
                if (serial.line != "" && serial.line != prevLine)
                {
                    var volumeSplit = serial.line.Split("|");
                    Debug.WriteLine(serial.line);
                    for (int i = 0; i < volumeSplit.Length; i++)
                    {
                        AppAudio selected;
                        if (comboBoxes[i].InvokeRequired)
                        {
                            comboBoxes[i].Invoke(new MethodInvoker(delegate
                            {
                                selected = (AppAudio)comboBoxes[i].SelectedItem;
                                if (selected is not null)
                                {
                                    AudioSessionControl selectedAudioSession = GetAudioSession(selected.Process.ProcessName);
                                    if (selectedAudioSession is not null)
                                    {
                                        Debug.WriteLine($"Application volume: {selectedAudioSession.SimpleAudioVolume.Volume}");
                                        Debug.WriteLine($"Desired volume: {(float)int.Parse(volumeSplit[i]) / 100}");
                                        selectedAudioSession.SimpleAudioVolume.Volume = (float)int.Parse(volumeSplit[i]) / 100;
                                    }
                                }
                            }));
                        }
                    }
                    prevLine = serial.line;
                }
                Thread.Sleep(200);
            }
        }

        private void AudioSessionManager_OnSessionCreated(object sender, IAudioSessionControl newSession)
        {
            Debug.WriteLine("New AudioSession");
            UpdateSessions();
        }

        private void UpdateSessions()
        {
            Debug.WriteLine("Reloading AudioSession");
            device.AudioSessionManager.RefreshSessions();
            sessions = device.AudioSessionManager.Sessions;
            audioSessions.Clear();
            for (int i = 0; i < sessions.Count; i++)
            {
                sessions[i].RegisterEventClient(this);
                Process process = Process.GetProcessById((int)sessions[i].GetProcessID);
                audioSessions.Add(new AppAudio(process/*, sessions[i]*/));
                if (sessions[i].IsSystemSoundsSession)
                {
                    audioSessions[i].DisplayName = "System Sounds";
                }
                Debug.WriteLine(process.MainWindowTitle != "" ? process.MainWindowTitle : process.ProcessName + "    " + (int)sessions[i].GetProcessID);
            }
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
            foreach (ComboBox combo in comboBoxes)
            {
                if (combo.InvokeRequired)
                {
                    combo.Invoke(new MethodInvoker(delegate
                    {
                        combo.Items.Clear();
                        combo.Items.AddRange(audioSessions.ToArray());
                        combo.DisplayMember = "DisplayName";
                    }));
                }
                else
                {
                    combo.Items.Clear();
                    combo.Items.AddRange(audioSessions.ToArray());
                    combo.DisplayMember = "DisplayName";
                }
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
                        AudioSessionControl activeWindowSession = GetAudioSession(Process.GetProcessById(pID).ProcessName);

                        if (activeWindowSession is not null)
                        {
                            Debug.WriteLine(activeWindowSession.SimpleAudioVolume.Volume);
                        }
                        activeWindowAudio.Process = Process.GetProcessById(pID);
                        Debug.WriteLine($"Active process: {activeWindowAudio.Process.ProcessName}");
                    }
                }
                else
                {
                }
                Thread.Sleep(10);
            }
        }

        private AudioSessionControl GetAudioSession(string processname)
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                notifyIcon1.Visible = true;
                Hide();
                e.Cancel = true;
            }
        }

        private void button_Click(object sender, EventArgs e)
        {
            switch ((sender as Button).Name)
            {
                case "button1":
                    try
                    {
                        textBox1.Enabled = true;
                        textBox1.Text = "";
                        comboBox1.SelectedItem = null;
                        comboBox1.Enabled = true;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "button2":
                    try
                    {
                        textBox2.Enabled = true;
                        textBox2.Text = "";
                        comboBox2.SelectedItem = null;
                        comboBox2.Enabled = true;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "button3":
                    try
                    {
                        textBox3.Enabled = true;
                        textBox3.Text = "";
                        comboBox3.SelectedItem = null;
                        comboBox3.Enabled = true;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "button4":
                    try
                    {
                        textBox4.Enabled = true;
                        textBox4.Text = "";
                        comboBox4.SelectedItem = null;
                        comboBox4.Enabled = true;
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                    break;
                case "button5":
                    try
                    {
                        textBox5.Enabled = true;
                        textBox5.Text = "";
                        comboBox5.SelectedItem = null;
                        comboBox5.Enabled = true;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;
                default:
                    break;
            }
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            switch ((sender as TextBox).Name)
            {
                case "textBox1":
                    try
                    {
                        if (textBox1.Text != "")
                        {
                            comboBox1.SelectedItem = null;
                            comboBox1.Enabled = false;
                        }
                        else
                        {
                            comboBox1.Enabled = true;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "textBox2":
                    try
                    {
                        if (textBox2.Text != "")
                        {
                            comboBox2.SelectedItem = null;
                            comboBox2.Enabled = false;
                        }
                        else
                        {
                            comboBox2.Enabled = true;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "textBox3":
                    try
                    {
                        if (textBox3.Text != "")
                        {
                            comboBox3.SelectedItem = null;
                            comboBox3.Enabled = false;
                        }
                        else
                        {
                            comboBox3.Enabled = true;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;

                case "textBox4":
                    try
                    {
                        if (textBox4.Text != "")
                        {
                            comboBox4.SelectedItem = null;
                            comboBox4.Enabled = false;
                        }
                        else
                        {
                            comboBox4.Enabled = true;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                    break;
                case "textBox5":
                    try
                    {
                        if (textBox5.Text != "")
                        {
                            comboBox5.SelectedItem = null;
                            comboBox5.Enabled = false;
                        }
                        else
                        {
                            comboBox5.Enabled = true;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    break;
                default:
                    break;
            }
        }

        private void comboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var selected = (AppAudio)(sender as ComboBox).SelectedItem;
            if (selected is not null)
            {
                AudioSessionControl selectedAudioSession = GetAudioSession(selected.Process.ProcessName);
                switch ((sender as ComboBox).Name)
                {
                    case "comboBox1":
                        try
                        {
                            if (selectedAudioSession is not null)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                {
                                    pictureBox1.Image = Icon.ExtractAssociatedIcon(selected.Process.MainModule.FileName).ToBitmap();
                                }
                            }
                            if (comboBox1.SelectedItem is "" || comboBox1.SelectedItem is null)
                            {
                                textBox1.Enabled = true;
                            }
                            else
                            {
                                textBox1.Enabled = false;
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        break;

                    case "comboBox2":
                        try
                        {
                            if (selectedAudioSession is not null)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                    pictureBox2.Image = Icon.ExtractAssociatedIcon(selected.Process.MainModule.FileName).ToBitmap();
                            }
                            if (comboBox2.SelectedItem is "" || comboBox2.SelectedItem is null)
                            {
                                textBox2.Enabled = true;
                            }
                            else
                            {
                                textBox2.Enabled = false;
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        break;

                    case "comboBox3":
                        try
                        {
                            if (selectedAudioSession is not null)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                    pictureBox3.Image = Icon.ExtractAssociatedIcon(selected.Process.MainModule.FileName).ToBitmap();
                            }
                            if (comboBox3.SelectedItem is "" || comboBox3.SelectedItem is null)
                            {
                                textBox3.Enabled = true;
                            }
                            else
                            {
                                textBox3.Enabled = false;
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        break;

                    case "comboBox4":
                        try
                        {
                            if (selectedAudioSession is not null)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                    pictureBox4.Image = Icon.ExtractAssociatedIcon(selected.Process.MainModule.FileName).ToBitmap();
                            }
                            if (comboBox4.SelectedItem is "" || comboBox4.SelectedItem is null)
                            {
                                textBox4.Enabled = true;
                            }
                            else
                            {
                                textBox4.Enabled = false;
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }

                        break;
                    case "comboBox5":
                        try
                        {
                            if (selectedAudioSession is not null)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                    pictureBox5.Image = Icon.ExtractAssociatedIcon(selected.Process.MainModule.FileName).ToBitmap();
                            }
                            if (comboBox5.SelectedItem is "" || comboBox5.SelectedItem is null)
                            {
                                textBox5.Enabled = true;
                            }
                            else
                            {
                                textBox5.Enabled = false;
                            }
                        }
                        catch (Exception)
                        {

                            throw;
                        }
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
