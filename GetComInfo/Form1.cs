using System.Data;
using System.IO.Ports;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Timers;
using GetComInfo.Helpers;

namespace GetComInfo
{
    public partial class Form1 : Form
    {
        // all serial ports
        private List<SerialPort> serialPorts = new List<SerialPort>();
        // data table object
        public DataTable dataTableCom = new DataTable();
        // dictionary for (port name; port number)
        private Dictionary<string, int> numberPort = new Dictionary<string, int>();
        // DataTableCom class
        private DataTableCom classDataTableCom = new DataTableCom();
        // dictionary for (port name; port data)
        private Dictionary<string, string> recport = new Dictionary<string, string>();
        // dictionary for (port name; port discription)
        private Dictionary<string, string> portIdsDescription = new Dictionary<string, string>();
        // exe name
        public string exeName = AppDomain.CurrentDomain.FriendlyName;
        // exe path
        public string exePath = Assembly.GetExecutingAssembly().Location;
        // current version
        public Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
        // portsDescription file path
        private string portsDescriptionPath = Path.GetDirectoryName(Application.StartupPath) + @"\GetComInfo.Properties.PortsDescription.txt";

        public Form1()
        {
            InitializeComponent();

            // Set current direction for project
            Directory.SetCurrentDirectory(AppContext.BaseDirectory);

            // Set timer to check for updates every 60 minutes
            var aTimer = new System.Timers.Timer(60 * 60 * 1000);
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.Start();
            // Manually check for update
            Task.Run(() =>
            {
                OnTimedEvent(null, null);
            });

            using (var managementObjectSearcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                managementObjectSearcher.Get();
                // Get all ports
                string[] portNames = SerialPort.GetPortNames();

                // Sort port names
                Array.Sort(portNames, (a, b) => int.Parse(Regex.Replace(a, "[^0-9]", "")) - int.Parse(Regex.Replace(b, "[^0-9]", "")));

                // Init dataGridView
                dataTableCom = classDataTableCom.dataTableCom();
                dataGridView2.DataSource = dataTableCom;

                // Create PortDescription.txt file 
                var files = Directory.GetFiles(Application.StartupPath);
                if (files.FirstOrDefault(f => f == portsDescriptionPath) == null)
                {
                    using (FileStream fs = File.Create(portsDescriptionPath))
                    {
                    }
                }

                // Read and add all descriptions in dictionary
                using (StreamReader sr = new StreamReader(portsDescriptionPath))
                {
                    string? line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        foreach (var portName in portNames)
                        {
                            if (line.Split(" ", portName.Length)[0] == portName)
                            {
                                portIdsDescription.Add(portName, line.Substring(line.IndexOf(" ")));
                            }
                        }
                    }
                }

                #region DataGridView2 settings
                dataGridView2.Columns[0].Width = 100;
                dataGridView2.Columns[0].MinimumWidth = 60;
                dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView2.Columns[0].ReadOnly = true;
                dataGridView2.Columns[1].Width = 75;
                dataGridView2.Columns[1].MinimumWidth = 50;
                dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView2.Columns[1].ReadOnly = true;
                dataGridView2.Columns[2].Width = 125;
                dataGridView2.Columns[2].MinimumWidth = 60;
                dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView2.Columns[2].ReadOnly = true;
                dataGridView2.Columns[3].Width = 150;
                dataGridView2.Columns[3].MinimumWidth = 100;
                dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                dataGridView2.Columns[3].ReadOnly = true;
                dataGridView2.Columns[3].Visible = false;
                dataGridView2.Columns[4].Width = 350;
                dataGridView2.Columns[4].MinimumWidth = 100;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[4].ReadOnly = true;
                dataGridView2.Columns[5].Width = 350;
                dataGridView2.Columns[5].MinimumWidth = 100;
                dataGridView2.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                #endregion

                int counter = 0;
                recport.Clear();
                numberPort.Clear();
                foreach (var portName in portNames)
                {
                    // set port settings
                    var serialPort = new SerialPort(portName, 115200, Parity.None, 8, StopBits.One);
                    serialPort.Handshake = Handshake.None;
                    serialPort.DataReceived += new SerialDataReceivedEventHandler(this.SerialPort_DataReceived);
                    serialPort.ReadTimeout = 3000;
                    serialPort.WriteTimeout = 3000;
                    recport.Add(serialPort.PortName, "");
                    serialPorts.Add(serialPort);
                    numberPort.Add(serialPort.PortName, counter);

                    // get description for the port
                    var portDescription = String.Empty;
                    portIdsDescription.TryGetValue(serialPort.PortName, out portDescription);
                    
                    // set description for the port
                    dataTableCom.Rows.Add(serialPort.PortName, true, "", "", "", portDescription);

                    ++counter;
                }
            }
        }

        /// <summary>
        /// Change DataGrid cell value event
        /// </summary>
        private void DataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // get row and column indexes
            var columnIndex = e.ColumnIndex;
            var rowIndex = e.RowIndex;

            // get current com port
            var comPort = dataTableCom.Rows[rowIndex][0];

            // get new content for current port
            var newContent = dataTableCom.Rows[rowIndex][columnIndex];

            var tContent = string.Empty;

            // read all descriptions
            using (StreamReader sr = new StreamReader(portsDescriptionPath))
            {
                var content = sr.ReadToEnd();
                tContent = content;
            }

            if (tContent.Contains(comPort.ToString() + " "))
            {
                if (!portIdsDescription.ContainsKey(comPort.ToString()))
                {
                    // add new port description in dictionary
                    using (StreamReader sr = new StreamReader(portsDescriptionPath))
                    {
                        string? line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Split(" ", comPort.ToString().Length)[0] == comPort.ToString())
                            {
                                portIdsDescription.Add(comPort.ToString(), line.Substring(line.IndexOf(" ")));
                            }
                        }
                    }
                }
                else
                {
                    // change port description in dictionary
                    using (StreamReader sr = new StreamReader(portsDescriptionPath))
                    {
                        string? line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Split(" ", comPort.ToString().Length)[0] == comPort.ToString())
                            {
                                portIdsDescription[comPort.ToString()] = line.Substring(line.IndexOf(" "));
                            }
                        }
                    }
                }

                // build changeFrom string
                var changeFrom = comPort.ToString();
                if (portIdsDescription[comPort.ToString()][0] != ' ')
                {
                    changeFrom = changeFrom + " ";
                }
                changeFrom = changeFrom + portIdsDescription[comPort.ToString()];
                // build changeTo string
                var changeTO = comPort.ToString() + " " + newContent;
                // edit file content
                tContent = tContent.Replace(changeFrom, changeTO);
                
                // replace file content
                using (StreamWriter sw = new StreamWriter(portsDescriptionPath, false))
                {
                    sw.WriteLine(tContent);
                    sw.Close();
                }
            }
            else
            {
                // add new port description to file content
                using (StreamWriter sw = new StreamWriter(portsDescriptionPath, true))
                {
                    sw.WriteLine(comPort.ToString() + " " + newContent);
                    sw.Close();
                }
            }
        }

        /// <summary>
        /// Start work button event
        /// </summary>
        private async void button1_Click(object sender, EventArgs e)
        {
            new Thread(() =>
            {
                DataTable dt = dataTableCom;
                foreach (SerialPort serialPort in serialPorts)
                {
                    SerialPort sp = serialPort;
                    int count = numberPort[sp.PortName];
                    new Thread(() =>
                    {
                        if (!(bool)dt.Rows[count][1])
                            return;
                        USSD_Processing(sp);
                    })
                    {
                        IsBackground = true
                    }.Start();
                    Thread.Sleep(50);
                }
            })
            {
                IsBackground = true
            }.Start();
        }

        /// <summary>
        /// Ussd processing
        /// </summary>
        public void USSD_Processing(SerialPort sp)
        {
            var code = textBox1.Text;
            var column = numberPort[sp.PortName];
            var regex1 = new Regex("\\d{9,11}");
            var regex2 = new Regex(",\"([\\s\\S]*?)\",");
            try
            {
                // open serial port 
                if (!sp.IsOpen)
                {
                    sp.Open();
                }
            }
            catch
            {
                return;
            }
            ChangeGridCom(column, "Processing...", "", "");
            recport[sp.PortName] = "";
            // Get state 
            sp.Write("AT+CPIN?; \r");
            Thread.Sleep(500);
            if (!recport[sp.PortName].Contains("CPIN: READY"))
            {
                if (recport[sp.PortName].Contains("CME ERROR: 10"))
                {
                    ChangeGridCom(column, "Error : no sim", "", "");
                }
                else
                {
                    ChangeGridCom(column, "NOT READY", "", "");
                }
            }
            else
            {
                // set character set
                sp.Write("AT+CSCS=\"GSM\"; \r");
                Thread.Sleep(1000);
                // set sms format 
                sp.Write("AT+CMGF=1; \r");
                Thread.Sleep(1000);
                if (code.Contains("**21*") || code.Contains("##002#") || code.Contains("##21#"))
                {
                    sp.Write("ATD" + code + ";\r");
                    Thread.Sleep(9000);
                    var Status = recport[sp.PortName].Contains("ERROR") || !recport[sp.PortName].Contains("OK") ? "CME ERROR" : "Success";
                    ChangeGridCom(column, Status, "", "");
                    recport[sp.PortName] = "";
                }
                else
                {
                    // ussd request
                    sp.Write("AT+CUSD=1," + code + "; \r");
                    Thread.Sleep(15000);
                    // build response
                    var input = regex2.Match(recport[sp.PortName]).Groups[1].Value;
                    Thread.Sleep(1000);
                    var Content = input.Replace("\n", "");
                    Thread.Sleep(1000);
                    string Status;
                    if (Content != "")
                        Status = "Success";
                    else if (recport[sp.PortName].Contains("CME ERROR: 103"))
                        Status = "CME ERROR: 103";
                    else if (recport[sp.PortName].Contains("CME ERROR: 100"))
                    {
                        Status = "CME ERROR: 100";
                        ChangeGridCom(column, Status, regex1.Match(input).ToString(), Content);
                        Thread.Sleep(3000);
                        // restart ussd request
                        USSD_Processing(sp);
                        return;
                    }
                    else if (recport[sp.PortName].Contains("CME ERROR: 30"))
                    {
                        Status = "CME ERROR: 30";
                    }
                    else
                    {
                        Status = "ERROR";
                        Content = recport[sp.PortName];
                    }
                    //sp.Write("AT+CUSD=2; \r");
                    ChangeGridCom(column, Status, regex1.Match(input).ToString(), Content);
                    recport[sp.PortName] = "";
                }
            }
        }

        /// <summary>
        /// DataReceived event for serial port
        /// </summary>
        public void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            lock (this)
            {
                SerialPort sp = (SerialPort)sender;
                byte[] numArray = new byte[sp.ReadBufferSize];
                int count = 0;
                try
                {
                    count = sp.Read(numArray, 0, numArray.Length);
                }
                catch (Exception ex)
                {
                }
                // Convert data byte array to string
                string data = Encoding.ASCII.GetString(numArray, 0, count);
                recport[sp.PortName] += data;
                if (data.Contains("CPIN: NOT READY"))
                {
                    data = "";
                    int column = numberPort[sp.PortName];
                    recport[sp.PortName] = "";
                    ChangeGridCom(column, "NOT READY", "", "");
                }
            }
        }

        /// <summary>
        /// Change Grid 
        /// </summary>
        private void ChangeGridCom(int column, string Status, string Numberphone, string Content) => Invoke(() =>
        {
            DataTable tableCom = dataTableCom;
            tableCom.Rows[column][2] = Status;
            tableCom.Rows[column][3] = Numberphone;
            if (!string.IsNullOrEmpty(Content))
            {
                var byteContext = ConvertHelper.HexStr2HexBytes(Content);
                if (byteContext != null)
                {
                    Content = ConvertHelper.HexBytes2UnicodeStr(byteContext);
                }
            }
            tableCom.Rows[column][4] = Content;
        });

        /// <summary>
        /// Update button event click
        /// </summary>
        async void button2_Click(object sender, EventArgs e)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Get latest version
                    var latestVersionResponse = await HttpHelper.HttpResponse("https://analytics.otp-service.online/otp/GetComInfoAllVersions/latest_version.txt");
                    
                    // Create new version object
                    var latestVersion = new Version(latestVersionResponse);

                    // Compare versions
                    var versionsComparer = currentVersion.CompareTo(latestVersion);
                    if (versionsComparer == -1)
                    {
                        var latestVersionFileName = $"GetComInfo_v{latestVersion.ToString(1)}.zip";
                        // Delete archive if exist
                        if (File.Exists(Environment.CurrentDirectory + @$"\{latestVersionFileName}"))
                        {
                            File.Delete(Environment.CurrentDirectory + $@"\{latestVersionFileName}");
                        }

                        // Download new version archive
                        using (var stream = await client.GetStreamAsync($"https://analytics.otp-service.online/otp/GetComInfoAllVersions/{latestVersionFileName}"))
                        using (var fileStream = new FileStream(Environment.CurrentDirectory + @$"\{latestVersionFileName}", FileMode.CreateNew))
                        await stream.CopyToAsync(fileStream);

                        // Update application
                        UpdateHelper.Cmd($"tasklist && taskkill /f /im \"{exeName}.exe\" && timeout /t 1 && tar -xf {latestVersionFileName} --exclude GetComInfo.Properties.PortsDescription.txt && timeout /t 3 && del {latestVersionFileName} && \"{exeName}.exe\" & exit");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Download new version error");
            }
        }

        /// <summary>
        /// Check for new version
        /// </summary>
        private async Task<string> CheckForNewVersion()
        {
            // Get latest version 
            var latestVersionResponse = await HttpHelper.HttpResponse("https://analytics.otp-service.online/otp/GetComInfoAllVersions/latest_version.txt");
            
            // Create new version object
            var latestVersion = new Version(latestVersionResponse);

            // Compare versions
            var versionsComparer = currentVersion.CompareTo(latestVersion);
            if (versionsComparer == -1)
            {
                return latestVersion.ToString(1);
            }

            return null;
        }

        /// <summary>
        /// Update timer event
        /// </summary>
        private async void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            var version = await CheckForNewVersion();
            if (version != null)
            {
                Invoke(() => button2.Visible = true);
            }
        }
    }
}