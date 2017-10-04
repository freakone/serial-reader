using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.ComponentModel;
using System.Management;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<string> tensoValues { get; set; } = new ObservableCollection<string>();
        public string calibrationFactor { get; set; } = "79";
        public ObservableCollection<string> boardID { get; set; } = new ObservableCollection<string>();
        System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer();

        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();
      
            BackgroundWorker work = new BackgroundWorker();
            work.DoWork += worker_DoWork;
            work.RunWorkerAsync();

            
            dispatcherTimer.Tick += dispatcherTimer_Tick;
            dispatcherTimer.Interval = new TimeSpan(0, 1, 0);
            dispatcherTimer.Start();

        }

        struct ComPort // custom struct with our desired values
        {
           public string name;
           public string vid;
           public string pid;
           public string description;
        }

        private const string vidPattern = @"VID_([0-9A-F]{4})";
        private const string pidPattern = @"PID_([0-9A-F]{4})";
        private List<ComPort> GetSerialPorts()
        {
            using (var searcher = new ManagementObjectSearcher
                ("SELECT * FROM WIN32_SerialPort"))
            {
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                return ports.Select(p =>
                {
                    ComPort c = new ComPort();
                    c.name = p.GetPropertyValue("DeviceID").ToString();
                    c.vid = p.GetPropertyValue("PNPDeviceID").ToString();
                    c.description = p.GetPropertyValue("Caption").ToString();

                    Match mVID = Regex.Match(c.vid, vidPattern, RegexOptions.IgnoreCase);
                    Match mPID = Regex.Match(c.vid, pidPattern, RegexOptions.IgnoreCase);

                    if (mVID.Success)
                        c.vid = mVID.Groups[1].Value;
                    if (mPID.Success)
                        c.pid = mPID.Groups[1].Value;

                    return c;

                }).ToList();
            }
        }
        SerialPort sp;

        private async void worker_DoWork(object sender, DoWorkEventArgs ea)
        {
            Begin:

            List<ComPort> ports = GetSerialPorts();
            
            try
            {
                if (ports.Count > 0)
                {
                    ComPort com = ports.Single(c => c.vid.Equals("0483") && c.pid.Equals("5740")); //filter ports to specific PID and VID
                    sp = new SerialPort(com.name);

                    if (sp.IsOpen)
                        sp.Close();

                    sp.Open();
                    sp.DataReceived += Sp_DataReceived;

                    sp.WriteLine("|id");

                    while (true)
                    {
                        sp.WriteLine("|val"); //periodically send command to device for returning actual values
                        await Task.Delay(100);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                boardID.Clear();
                await Task.Delay(1000);
                goto Begin;
            }

            await Task.Delay(1000);
            goto Begin;

        }

        private string incomingBuffer = "";

        private List<string> searchForCommands()
        {
            List<string> cmds = new List<string>();

            int start = incomingBuffer.IndexOf("|");
            int end = incomingBuffer.IndexOf("\n");

            if (start >= 0 && end >= 0)
            {
                if( end < start )
                {
                    incomingBuffer = incomingBuffer.Substring(end + 1);
                    end = incomingBuffer.IndexOf("\n");
                }

                while (start < end && start >= 0 && end >= 0)
                {
                    cmds.Add(incomingBuffer.Substring(start, end - start).Trim(new char[] { '|' }));
                    incomingBuffer = incomingBuffer.Substring(end + 1);

                    start = incomingBuffer.IndexOf("|");
                    end = incomingBuffer.IndexOf("\n");
                }
            }

            return cmds;
        }

        Regex r = new Regex(@"([A-F0-9]+)", RegexOptions.IgnoreCase);
        private void Sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            incomingBuffer += sp.ReadExisting();

            List<string> cmds = searchForCommands();

            foreach (string command in cmds)
            {
                if (command.Contains(":"))
                {
                    tensoValues.Clear();
                    List<string> lData = new List<string>(command.Split(new char[] { ':' }));
                    lData.RemoveAt(0);
                    lData.ForEach(it => tensoValues.Add(it));
                    tensoValues.Add(tensoValues.Sum(t => int.Parse(t)).ToString());
                }
                else
                {
                    if(r.Match(command).Success)
                    {
                        boardID.Clear();
                        boardID.Add("ID: " + command.Substring(command.Length - 5));
                    }
                    else
                    {
                        MessageBox.Show(command, "Command returns:");
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (sp != null && sp.IsOpen)
            {
                sp.WriteLine("|cal" + calibrationFactor);
            }
        }

        private void TextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                Button_Click(this, null);
            }
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            if (tensoValues.Count > 4)
            {
                textBoxLog.Text += DateTime.Now.ToLongTimeString() + ";" + tensoValues[4] + "\r\n" ;
            }
        }
    }
}
