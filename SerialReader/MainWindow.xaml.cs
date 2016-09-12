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

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<string> tensoValues { get; set; } = new ObservableCollection<string>();
        public MainWindow()
        {
            DataContext = this;
            InitializeComponent();
      
            BackgroundWorker work = new BackgroundWorker();
            work.DoWork += worker_DoWork;
            work.RunWorkerAsync();
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
        private async void worker_DoWork(object sender, DoWorkEventArgs ea)
        {
            Begin:

            List<ComPort> ports = GetSerialPorts();
            
            try
            {
                if (ports.Count > 0)
                {
                    ComPort com = ports.Single(c => c.vid.Equals("0483") && c.pid.Equals("5740")); //filter ports to specific PID and VID

                    SerialPort sp = new SerialPort(com.name);
                    if (sp.IsOpen)
                        sp.Close();

                    sp.Open();
                    sp.DataReceived += Sp_DataReceived;


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
                await Task.Delay(1000);
                goto Begin;
            }

            await Task.Delay(1000);
            goto Begin;

        }
        private void Sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            indata = indata.TrimStart(new char[]{ '|' });
            tensoValues.Clear();
            List<string> lData = new List<string>(indata.Split(new char[] { ':' }));
            lData.RemoveAt(0);
            lData.RemoveAt(lData.Count - 1);
            lData.ForEach(it => tensoValues.Add(it));
            tensoValues.Add(tensoValues.Sum(t => int.Parse(t)).ToString());
        }
    }
}
