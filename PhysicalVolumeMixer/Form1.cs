using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.CoreAudioApi.Interfaces;
using Tommy;

namespace PhysicalVolumeMixer
{
    public partial class Form1 : Form, IAudioSessionEventsHandler
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        readonly List<AppAudio> _audioSessions = new();
        MMDevice _device;
        SessionCollection _sessions;

        readonly int _activeWindowCooldown = 350;
        long _lastGetCurrentWindowCall;
        readonly AppAudio _activeWindowAudio = new();

        readonly Serial _serial = new();

        readonly List<ComboBox> _comboBoxes = new();

        public Form1()
        {
            InitializeComponent();

            _activeWindowAudio.Process = Process.GetProcessById(0);
            var deviceEnumerator = new MMDeviceEnumerator();
            _device = deviceEnumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _device.AudioSessionManager.OnSessionCreated += AudioSessionManager_OnSessionCreated;

            _sessions = _device.AudioSessionManager.Sessions;
            for (int i = 0; i < _sessions.Count; i++)
            {
                _sessions[i].RegisterEventClient(this);
                Process process = Process.GetProcessById((int) _sessions[i].GetProcessID);
                _audioSessions.Add(new AppAudio(process));
                if (_sessions[i].IsSystemSoundsSession)
                {
                    _audioSessions[i].DisplayName = "System Sound";
                }

                Debug.WriteLine(process.MainWindowTitle != ""
                    ? process.MainWindowTitle
                    : process.ProcessName + "    " + (int) _sessions[i].GetProcessID);
            }

            Task.Run(async () => await _serial.GetArduinoThread());

            Thread serialRead = new(_serial.Read)
            {
                IsBackground = true
            };
            serialRead.Start();
            Thread foregroundThread = new(GetActiveWindow)
            {
                IsBackground = true
            };
            foregroundThread.Start();
            Thread updateAudio = new(UpdateAudio)
            {
                IsBackground = true
            };
            updateAudio.Start();
            _lastGetCurrentWindowCall = TimeSinceEpoch();
            BindingSource bindingSource = new();
            bindingSource.DataSource = _audioSessions;
            _comboBoxes.Add(comboBox1);
            _comboBoxes.Add(comboBox2);
            _comboBoxes.Add(comboBox3);
            _comboBoxes.Add(comboBox4);
            _comboBoxes.Add(comboBox5);

            UpdateComboBox();

            List<TextBox> textBoxes = new();
            textBoxes.Add(textBox1);
            textBoxes.Add(textBox2);
            textBoxes.Add(textBox3);
            textBoxes.Add(textBox4);
            textBoxes.Add(textBox5);


            if (File.Exists("configuration.toml"))
            {
                using (StreamReader reader = File.OpenText("configuration.toml"))
                {
                    // Parse the table
                    TomlTable table = TOML.Parse(reader);
                    for (int i = 0; i < table["slidersExe"].AsArray.ChildrenCount; i++)
                    {
                        if (table["slidersExe"][i] == "System Sound")
                        {
                            _comboBoxes[i].SelectedItem = _audioSessions.Find(x => x.DisplayName == "System Sound");
                        }
                        else if (table["slidersExe"][i] == "Foreground")
                        {
                            _comboBoxes[i].SelectedItem = _audioSessions.Find(x => x.DisplayName == "Foreground window");
                        }
                        else
                        {
                            
                            var tryToGetAudioSession = _audioSessions.Find(x =>
                                x.DisplayName != "System Sound" && x.DisplayName != "Foreground window" && x.Process.MainModule.ModuleName == table["slidersExe"][i]);
                            if (tryToGetAudioSession is not null)
                            {
                                _comboBoxes[i].SelectedItem = tryToGetAudioSession;
                            }
                            else
                            {
                                textBoxes[i].Text = table["slidersExe"][i];
                            }
                        }
                    }
                }
            }
            else
            {
                TomlTable toml = new TomlTable
                {
                    ["slidersExe"] = new TomlNode[] {"","","","",""}
                };

                using(StreamWriter writer = File.CreateText("configuration.toml"))
                {
                    toml.WriteTo(writer);
                    // Remember to flush the data if needed!
                    writer.Flush();
                }
            }
        }

        string _prevLine = "";

        //line ref: {SliderVolume1}|{SliderVolume2}|{SliderVolume3}|{SliderVolume4}|{SliderVolume5}
        private void UpdateAudio()
        {
            while (true)
            {
                if (_serial.Line != "" && _serial.Line != _prevLine)
                {
                    var volumeSplit = _serial.Line.Split("|");
                    Debug.WriteLine(_serial.Line);
                    for (int i = 0; i < volumeSplit.Length; i++)
                    {
                        AppAudio selected;
                        if (_comboBoxes[i].InvokeRequired)
                        {
                            _comboBoxes[i].Invoke(new MethodInvoker(delegate
                            {
                                selected = (AppAudio) _comboBoxes[i].SelectedItem;
                                if (selected is not null)
                                {
                                    AudioSessionControl selectedAudioSession =
                                        GetAudioSession(selected.Process.ProcessName);
                                    if (selectedAudioSession is not null)
                                    {
                                        Debug.WriteLine(
                                            $"Application volume: {selectedAudioSession.SimpleAudioVolume.Volume}");
                                        Debug.WriteLine($"Desired volume: {(float) int.Parse(volumeSplit[i]) / 100}");
                                        selectedAudioSession.SimpleAudioVolume.Volume =
                                            (float) int.Parse(volumeSplit[i]) / 100;
                                    }
                                }
                            }));
                        }
                    }

                    _prevLine = _serial.Line;
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
            _device.AudioSessionManager.RefreshSessions();
            _sessions = _device.AudioSessionManager.Sessions;
            _audioSessions.Clear();
            for (int i = 0; i < _sessions.Count; i++)
            {
                _sessions[i].RegisterEventClient(this);
                Process process = Process.GetProcessById((int) _sessions[i].GetProcessID);
                _audioSessions.Add(new AppAudio(process /*, sessions[i]*/));
                if (_sessions[i].IsSystemSoundsSession)
                {
                    _audioSessions[i].DisplayName = "System Sound";
                }

                Debug.WriteLine(process.MainWindowTitle != ""
                    ? process.MainWindowTitle
                    : process.ProcessName + "    " + (int) _sessions[i].GetProcessID);
            }

            UpdateComboBox();
        }

        private void UpdateComboBox()
        {
            _audioSessions.Sort((e1, e2) => { return -e2.Name.CompareTo(e1.Name); });
            _audioSessions.Insert(0, _activeWindowAudio);
            _audioSessions[0].DisplayName = "Foreground window";
            foreach (ComboBox combo in _comboBoxes)
            {
                if (combo.InvokeRequired)
                {
                    combo.Invoke(new MethodInvoker(delegate
                    {
                        combo.Items.Clear();
                        combo.Items.AddRange(_audioSessions.ToArray());
                        combo.DisplayMember = "DisplayName";
                    }));
                }
                else
                {
                    combo.Items.Clear();
                    combo.Items.AddRange(_audioSessions.ToArray());
                    combo.DisplayMember = "DisplayName";
                }
            }
        }

        private void GetActiveWindow()
        {
            while (true)
            {
                if (_lastGetCurrentWindowCall + _activeWindowCooldown < TimeSinceEpoch())
                {
                    _lastGetCurrentWindowCall = TimeSinceEpoch();
                    IntPtr hWnd = GetForegroundWindow(); // Get foreground window handle
                    uint activeWindowPtr;
                    GetWindowThreadProcessId(hWnd, out activeWindowPtr); // Get PID from window handle
                    int pId = Convert.ToInt32(activeWindowPtr);
                    if (pId != _activeWindowAudio.Process.Id)
                    {
                        AudioSessionControl activeWindowSession =
                            GetAudioSession(Process.GetProcessById(pId).ProcessName);

                        if (activeWindowSession is not null)
                        {
                            Debug.WriteLine(activeWindowSession.SimpleAudioVolume.Volume);
                        }

                        _activeWindowAudio.Process = Process.GetProcessById(pId);
                        Debug.WriteLine($"Active process: {_activeWindowAudio.Process.ProcessName}");
                    }
                }

                Thread.Sleep(10);
            }
        }

        private AudioSessionControl GetAudioSession(string processname)
        {
            try
            {
                var sessions = _device.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    Process process = Process.GetProcessById((int) sessions[i].GetProcessID);
                    if (process.ProcessName == processname)
                    {
                        _activeWindowAudio.Process = process;
                        return sessions[i];
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return null;
        }

        private long TimeSinceEpoch()
        {
            return (long) (DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds;
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
            switch ((sender as Button)?.Name)
            {
                case "button1":
                    textBox1.Enabled = true;
                    textBox1.Text = "";
                    comboBox1.SelectedItem = null;
                    comboBox1.Enabled = true;
                    AddExeNameToConf(1, "");

                    break;

                case "button2":
                    textBox2.Enabled = true;
                    textBox2.Text = "";
                    comboBox2.SelectedItem = null;
                    comboBox2.Enabled = true;
                    AddExeNameToConf(2, "");

                    break;

                case "button3":
                    textBox3.Enabled = true;
                    textBox3.Text = "";
                    comboBox3.SelectedItem = null;
                    comboBox3.Enabled = true;
                    AddExeNameToConf(3, "");

                    break;

                case "button4":
                    textBox4.Enabled = true;
                    textBox4.Text = "";
                    comboBox4.SelectedItem = null;
                    comboBox4.Enabled = true;
                    AddExeNameToConf(4, "");

                    break;
                case "button5":
                    textBox5.Enabled = true;
                    textBox5.Text = "";
                    comboBox5.SelectedItem = null;
                    comboBox5.Enabled = true;
                    AddExeNameToConf(5, "");

                    break;
            }
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            switch ((sender as TextBox)?.Name)
            {
                case "textBox1":
                    if (textBox1.Text != "")
                    {
                        comboBox1.SelectedItem = null;
                        comboBox1.Enabled = false;
                        if (textBox1.Text.EndsWith(".exe"))
                        {
                            AddExeNameToConf(1, textBox1.Text);
                        }
                    }
                    else
                    {
                        comboBox1.Enabled = true;
                    }

                    break;

                case "textBox2":
                    if (textBox2.Text != "")
                    {
                        comboBox2.SelectedItem = null;
                        comboBox2.Enabled = false;
                        if (textBox2.Text.EndsWith(".exe"))
                        {
                            AddExeNameToConf(2, textBox2.Text);
                        }
                    }
                    else
                    {
                        comboBox2.Enabled = true;
                    }

                    break;

                case "textBox3":
                    if (textBox3.Text != "")
                    {
                        comboBox3.SelectedItem = null;
                        comboBox3.Enabled = false;
                        if (textBox3.Text.EndsWith(".exe"))
                        {
                            AddExeNameToConf(3, textBox3.Text);
                        }
                    }
                    else
                    {
                        comboBox3.Enabled = true;
                    }

                    break;

                case "textBox4":
                    if (textBox4.Text != "")
                    {
                        comboBox4.SelectedItem = null;
                        comboBox4.Enabled = false;
                        if (textBox4.Text.EndsWith(".exe"))
                        {
                            AddExeNameToConf(4, textBox4.Text);
                        }
                    }
                    else
                    {
                        comboBox4.Enabled = true;
                    }

                    break;
                case "textBox5":
                    if (textBox5.Text != "")
                    {
                        comboBox5.SelectedItem = null;
                        comboBox5.Enabled = false;
                        if (textBox5.Text.EndsWith(".exe"))
                        {
                            AddExeNameToConf(5, textBox5.Text);
                        }
                    }
                    else
                    {
                        comboBox5.Enabled = true;
                    }

                    break;
            }
        }

        private void comboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var selected = (AppAudio) (sender as ComboBox)?.SelectedItem;
            if (selected is not null)
            {
                AudioSessionControl selectedAudioSession = GetAudioSession(selected.Process.ProcessName);
                switch ((sender as ComboBox).Name)
                {
                    case "comboBox1":
                        if (comboBox1.Text == "Foreground window")
                        {
                            AddExeNameToConf(1, "Foreground");
                        }
                        if (selectedAudioSession is not null)
                        {
                            if (!selectedAudioSession.IsSystemSoundsSession)
                            {
                                if (!selectedAudioSession.IsSystemSoundsSession)
                                {
                                    if (selected.Process.MainModule != null)
                                        if (selected.Process.MainModule.FileName != null)
                                        {
                                            pictureBox1.Image = Icon
                                                .ExtractAssociatedIcon(selected.Process.MainModule.FileName)
                                                ?.ToBitmap();
                                            AddExeNameToConf(1, selected.Process.MainModule.ModuleName);
                                        }
                                }
                                else
                                {
                                    AddExeNameToConf(1, "System Sound");
                                }
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

                        break;

                    case "comboBox2":
                        if (comboBox2.Text == "Foreground window")
                        {
                            AddExeNameToConf(2, "Foreground");
                        }
                        if (selectedAudioSession is not null)
                        {
                            if (!selectedAudioSession.IsSystemSoundsSession)
                            {
                                if (selected.Process.MainModule != null)
                                    if (selected.Process.MainModule.FileName != null)
                                    {
                                        pictureBox2.Image = Icon
                                            .ExtractAssociatedIcon(selected.Process.MainModule.FileName)
                                            ?.ToBitmap();
                                        AddExeNameToConf(2, selected.Process.MainModule.ModuleName);
                                    }
                            }
                            else
                            {
                                AddExeNameToConf(2, "System Sound");
                            }
                        }

                        if (comboBox2.SelectedItem is "" || comboBox2.SelectedItem is null)
                        {
                            textBox2.Enabled = true;
                        }
                        else
                        {
                            textBox2.Enabled = false;
                        }

                        break;

                    case "comboBox3":
                        if (comboBox3.Text == "Foreground window")
                        {
                            AddExeNameToConf(3, "Foreground");
                        }
                        if (selectedAudioSession is not null)
                        {
                            if (!selectedAudioSession.IsSystemSoundsSession)
                            {
                                if (selected.Process.MainModule != null)
                                    if (selected.Process.MainModule.FileName != null)
                                    {
                                        pictureBox3.Image = Icon
                                            .ExtractAssociatedIcon(selected.Process.MainModule.FileName)
                                            ?.ToBitmap();
                                        AddExeNameToConf(3, selected.Process.MainModule.ModuleName);
                                    }
                            }
                            else
                            {
                                AddExeNameToConf(3, "System Sound");
                            }
                        }

                        if (comboBox3.SelectedItem is "" || comboBox3.SelectedItem is null)
                        {
                            textBox3.Enabled = true;
                        }
                        else
                        {
                            textBox3.Enabled = false;
                        }

                        break;

                    case "comboBox4":
                        if (comboBox4.Text == "Foreground window")
                        {
                            AddExeNameToConf(4, "Foreground");
                        }
                        if (selectedAudioSession is not null)
                        {
                            if (!selectedAudioSession.IsSystemSoundsSession)
                            {
                                if (selected.Process.MainModule != null)
                                    if (selected.Process.MainModule.FileName != null)
                                    {
                                        pictureBox4.Image = Icon
                                            .ExtractAssociatedIcon(selected.Process.MainModule.FileName)
                                            ?.ToBitmap();
                                        AddExeNameToConf(4, selected.Process.MainModule.ModuleName);
                                    }
                            }
                            else
                            {
                                AddExeNameToConf(4, "System Sound");
                            }
                        }

                        if (comboBox4.SelectedItem is "" || comboBox4.SelectedItem is null)
                        {
                            textBox4.Enabled = true;
                        }
                        else
                        {
                            textBox4.Enabled = false;
                        }

                        break;
                    case "comboBox5":
                        if (comboBox5.Text == "Foreground window")
                        {
                            AddExeNameToConf(5, "Foreground");
                        }
                        if (selectedAudioSession is not null)
                        {
                            if (!selectedAudioSession.IsSystemSoundsSession)
                            {
                                if (selected.Process.MainModule != null)
                                    if (selected.Process.MainModule.FileName != null)
                                    {
                                        pictureBox5.Image = Icon
                                            .ExtractAssociatedIcon(selected.Process.MainModule.FileName)
                                            ?.ToBitmap();
                                        AddExeNameToConf(5, selected.Process.MainModule.ModuleName);
                                    }
                            }
                            else
                            {
                                AddExeNameToConf(5, "System Sound");
                            }
                        }

                        if (comboBox5.SelectedItem is "" || comboBox5.SelectedItem is null)
                        {
                            textBox5.Enabled = true;
                        }
                        else
                        {
                            textBox5.Enabled = false;
                        }

                        break;
                }
            }
        }

        private void AddExeNameToConf(int port, string exename)
        {
            TomlTable table;
            using (StreamReader reader = File.OpenText("configuration.toml"))
            {
                // Parse the table
                table = TOML.Parse(reader);
                table["slidersExe"][port-1] = exename;
            }
            using(StreamWriter writer = File.CreateText("configuration.toml"))
            {
                table.WriteTo(writer);
                // Remember to flush the data if needed!
                writer.Flush();
            }
        }

        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void notifyIcon1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) {
                MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                mi.Invoke(notifyIcon1, null);
            }
        }
    }
}