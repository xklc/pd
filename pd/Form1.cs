using com.pd.extract;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace pd
{
    //原始公钥:MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCobEkviCWE3J2NFo5E2UQGlJCkUnoylStfmEitbXaMrpOhm0QRYlXMz42i0/WslnY6qV3rb4Q9sJP5Qxg49Zy4HGoFCv6UGnWJtt5k34hd10+jsbjCw3NEMnW/NCu2dAUNLTs9yh1WwXnU5551KgR9FQ9Gp3hZN/AF/9vYewmWoQIDAQAB
    //crack公钥:MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQDF6AOIu437LKgCp0KM5U3NWqT4djZM5NBnhmsQhDWqHC88SuUo41Jyr1ryqwzgAQTCaRMNwjhyy+o3Wson0ejb5QvXF6VBHKNfN8+oBWPUKuxQQ5ypUgUJO4fAXcmra/Zg0J7F3O109t4Yc2Di+qghjIGq11JkTjmSRDMsgZNQpwIDAQAB

    //1. biz\papercut\pcng\service 目录下LicenseManager.class和LicenseUtils.class反编译并替换上述的密钥后重新编译
    //or D:\workdir\software\HxD x64.exe
    public partial class Form1 : Form
    {
        private Dictionary<string, List<string>> dtMap = new Dictionary<string, List<string>>();
        public Form1()
        {
            InitializeComponent();
        }

        private void createDir(string dir)
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
            Directory.CreateDirectory(dir);
        }
        private void createNeedDirs()
        {
            string cwd = Directory.GetCurrentDirectory();
            string tmp_dir = Path.Combine(cwd, "tmp");
            createDir(tmp_dir);
            string data_dir = Path.Combine(cwd, "data");
            createDir(data_dir);
            string license_dir = Path.Combine(cwd, "license");
            if (!Directory.Exists(license_dir))
            {
                Directory.CreateDirectory(license_dir);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //hp=ext-devices-hp-prn|ext-devices-hp,kyocera=ext-devices-kyocera-mita,lexmark=ext-devices-lexmark,xerox=ext-devices-xerox,fuji-aip=ext-devices-fuji-xerox-aip,konica-minolta=ext-devices-konica-minolta,ricoh=ext-devices-ricoh,biostore=ext-devices-biostore,canon=ext-devices-canon,dell=ext-devices-dell|ext-devices-dell-prn,epson=ext-devices-epson,cartadis-cpad=ext-devices-cartadis-cpad,copicode-ip=ext-devices-copicode-ip,generic=ext-devices-generic,itc=ext-devices-itc,jamex-netpad=ext-devices-jamex-netpad,toshiba=ext-devices-toshiba,riso=ext-devices-riso,oki=ext-devices-oki|ext-devices-oki-prn,muratec=ext-devices-muratec,samsung=ext-devices-samsung,selectec=ext-devices-selectec,sharp-osa=ext-devices-sharp-osa,vcc=ext-devices-vcc,toshiba=ext-devices-toshiba,brother=ext-devices-brother,sindoh=ext-devices-sindoh,hp-fast-release=ext-devices-hp-pro-fr,fast-release=ext-devices-print-release-card
            string tmp = ConfigurationManager.AppSettings["device_type"];
            string device_type = Encoding.UTF8.GetString(Convert.FromBase64String(tmp));
            string[] kvs = device_type.Split(new Char[] { ',' });
            foreach (var kv in kvs)
            {
                int pos = kv.IndexOf('=');
                if (pos >= 0)
                {
                    string key = kv.Substring(0, pos);
                    string value = kv.Substring(pos + 1);
                    string[] values = value.Split('|');
                    dtMap[key] = values.ToList<string>();
                }
            }
            DataGridViewComboBoxColumn cbc1 = (DataGridViewComboBoxColumn)this.dataGridView1.Columns[0];
            foreach (var key in dtMap.Keys)
            {
                cbc1.Items.Add(key);
            }

            DataGridViewComboBoxColumn cbc2 = (DataGridViewComboBoxColumn)this.dataGridView1.Columns[1];
            for (int idx = 1; idx <= 100; idx++)
            {
                cbc2.Items.Add(idx.ToString());
            }
            DataGridViewButtonColumn remove = (DataGridViewButtonColumn)this.dataGridView1.Columns[2];
            remove.Text = "remove";
            remove.UseColumnTextForButtonValue = true;
            this.comboBox1.SelectedIndex = 0;
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 0;
            this.comboBox4.SelectedIndex = 0;
            this.comboBox5.SelectedIndex = 0;
            this.comboBox6.SelectedIndex = 11;
            this.comboBox7.SelectedIndex = 1;
        }

        //截取冲击式样形状
        public bool isEnglish(string src)
        {
            Match mInfo = Regex.Match(src, @"^[a-zA-Z0-9\.\s]+$");
            return mInfo.Success;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                if (dataGridView1.CurrentRow.Index > 0)
                {
                    dataGridView1.EndEdit();
                    dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
                }
                else
                {
                    MessageBox.Show("不能删除当前行");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string save_path = ConfigurationManager.AppSettings["save_path"];
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            dialog.SelectedPath = save_path;
            //dialog.RootFolder = Environment.SpecialFolder.Programs;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                save_path = dialog.SelectedPath;
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                //ConfigurationManager.AppSettings["save_path"] = save_path;
                config.AppSettings.Settings["save_path"].Value = save_path;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
            }
        }

        public string getCustomerReferenceNo()
        {
            String oldReferenceNo = ConfigurationManager.AppSettings["customer_reference_no"];
            String newReferenceNo = NumberConvertUtil.AddOne(oldReferenceNo);
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["customer_reference_no"].Value = newReferenceNo;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            return newReferenceNo;
        }

        public string getReferenceNo()
        {
            String oldReferenceNo = ConfigurationManager.AppSettings["reference_no"];
            String newReferenceNo = (Convert.ToInt32(oldReferenceNo)+1).ToString();
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["reference_no"].Value = newReferenceNo;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
            return newReferenceNo;
        }

        private long getCurrentMillis(long ticks)
        {
            long currentTicks = ticks;
            DateTime dtFrom = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            long currentMillis = (currentTicks - dtFrom.Ticks) / 10000;
            return currentMillis;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //try {
            //    long end = 1696780800000L;
            //    string ntpserver = "ntp1.aliyun.com";
            //    NTPClient client = new NTPClient(ntpserver);
            //    client.Connect(false);
            //    long now = getCurrentMillis(client.TransmitTimestamp.Ticks);
            //    if (now >= end)
            //    {
            //        return;
            //    }
            //}
            //catch(Exception ex)
            //{
            //    return;
            //}

            this.Cursor = Cursors.WaitCursor;
            try
            {
                createNeedDirs();
                string cwd = Directory.GetCurrentDirectory();
                if (dataGridView1.Rows[0].Cells[0].Value == null)
                {
                    MessageBox.Show("please choose device info");
                    dataGridView1.Focus();
                    this.Cursor = Cursors.Default;
                    return;
                }
                if (textBox2.Text.Trim().Count() == 0)
                {
                    MessageBox.Show("Customer Name is Empty");
                    textBox2.Focus();
                    this.Cursor = Cursors.Default;
                    return;
                }

                //if (!isEnglish(textBox2.Text))
                //{
                //    MessageBox.Show("Customer Name must be number, charactor or dot");
                //    textBox2.Focus();
                //    this.Cursor = Cursors.Default;
                //    return;
                //}

                //long et = 1690464318000; //2/7/15 14:00:00            
                //long currentMillis = getCurrentMillis(DateTime.Now.Ticks);
                //if (currentMillis >= et)
                //{
                //    Close();
                //}

                //organization - name
                string organizationName = textBox2.Text;

                //advanced - clients - licensed
                string advancedClientsLicensed;
                if (checkBox4.Checked)
                {
                    advancedClientsLicensed = checkBox4.Text;
                }
                else
                {
                    advancedClientsLicensed = comboBox5.Text;
                }

                //issued-date
                String expiredMonths = comboBox6.Text;
                //issued-date

                String issuedDate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                //expiry-date            
                DateTime expireDt = dateTimePicker1.Value.AddMonths(Convert.ToInt32(expiredMonths));
                string expiryDate = expireDt.ToString("yyyy-MM-dd");
                //site-servers-licensed
                string siteServersLicensed = comboBox3.Text;
                //modules-licensed
                String modulesLicensed = "PRINT";
                if (checkBox3.Checked)
                {
                    modulesLicensed = checkBox3.Text + "," + modulesLicensed;
                }
                //users-licensed
                String userLicensed = textBox1.Text;
                if (checkBox1.Checked)
                {
                    userLicensed = checkBox1.Text;
                }

                //release-stations-licensed


                //site-servers-licensed
                if (checkBox5.Checked)
                {
                    siteServersLicensed = checkBox5.Text;
                }

                //device info list
                Dictionary<string, string> devList = new Dictionary<string, string>();
                List<string> device_list = new List<string>();
                for (int idx = 0; idx < dataGridView1.Rows.Count - 1; idx++)
                {
                    devList[dataGridView1.Rows[idx].Cells[0].Value.ToString()] = dataGridView1.Rows[idx].Cells[1].Value.ToString();
                    String simpleName = dataGridView1.Rows[idx].Cells[0].Value.ToString();
                    if (dtMap.ContainsKey(simpleName))
                    {
                        foreach (var extInfo in dtMap[simpleName])
                        {
                            device_list.Add(extInfo + "=" + dataGridView1.Rows[idx].Cells[1].Value.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("error config for device info:" + simpleName);
                        this.Cursor = Cursors.Default;
                        return;
                    }
                }
                if (device_list.Count == 0)
                {
                    MessageBox.Show("please select device first!");
                    this.Cursor = Cursors.Default;
                    return;
                }


                List<string> contentlines = new List<string>();
                contentlines.Add("#NOTE: Changing any part of this file will invalidate the license.");

                String issueTime = DtTool.getUTCString(getCurrentMillis(dateTimePicker1.Value.Ticks));
                int pos = issueTime.IndexOf('|');

                //unique-id
                string uniqueId = "";
                if (pos >= 0)
                {
                    contentlines.Add("#" + issueTime.Substring(0, pos));
                    contentlines.Add("");
                    uniqueId = issueTime.Substring(pos + 1);
                }

                //version
                String version = this.comboBox7.Text.Trim();


                contentlines.Add("site-servers-licensed=" + siteServersLicensed);
                contentlines.Add("organization-type=" + comboBox1.Text.Trim().ToUpper());
                contentlines.Add("issued-by=PaperCut Software International Pty. Ltd.");
                contentlines.Add("edition=" + comboBox4.Text.Trim());
                if (checkBox5.Checked)
                    contentlines.Add("advanced-clients-licensed=" + advancedClientsLicensed);
                contentlines.Add("modules-licensed=" + modulesLicensed);
                contentlines.Add("licensed-version=" + version);
                contentlines.Add("issued-date=" + issuedDate);
                contentlines.Add("release-stations-licensed=" + comboBox2.Text.Trim());
                contentlines.Add("order-reference="+getReferenceNo());
                contentlines.Add("smb-bundle=false");
                contentlines.AddRange(device_list);
                contentlines.Add("customer-reference-no=" + getCustomerReferenceNo());
                if (userLicensed.Equals("unlimited"))
                {
                    contentlines.Add("users-purchased=indefinite");
                }
                else
                {
                    contentlines.Add("users-purchased=" + userLicensed);
                }
                contentlines.Add("expiry-date="+ expiryDate);
                contentlines.Add("unique-id=" + uniqueId);
                contentlines.Add("created-by=kirk.mcafee@papercut.com");
                contentlines.Add("updates-expiry-policy=ALLOW_UPDATES_WITHIN_SAME_VERSION");
                contentlines.Add("support-expiry-date=" + expiryDate);
                contentlines.Add("users-licensed=" + userLicensed);
                contentlines.Add("organization-name=" + organizationName);
                contentlines.Add("updates-expiry-date=" + expiryDate);

                string inputFileName = Guid.NewGuid().ToString() + ".txt";
                String contents = String.Join("\r\n", contentlines);
                String in_file = Path.Combine(cwd, "tmp", inputFileName);
                File.WriteAllText(in_file, contents);
                String prefix = string.Format("PaperCutMF{0}-", version);
                String out_file = Path.Combine(cwd, "data", "license.txt");
                String license_file_name = "";
                if (organizationName != null)
                {
                    string suffix = "license";
                    if (!organizationName.Trim().EndsWith("."))
                    {
                        suffix = ".license";
                    }
                    license_file_name = prefix + organizationName.Trim() + suffix;
                }
                string save_path = ConfigurationManager.AppSettings["save_path"].Trim();

                String license_file = Path.Combine(cwd, "license", license_file_name);
                if (save_path.Count() > 0)
                {
                    license_file = Path.Combine(save_path, license_file_name);
                }

                out_file = SignGenrator.generateSignFile(in_file, out_file);
                if (File.Exists(license_file))
                {
                    File.Delete(license_file);
                }
                ZipHelper.Zip(out_file, license_file);
                if (File.Exists(license_file))
                {
                    sendMail(organizationName, contents, license_file);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //throw ex;
            }
            this.Cursor = Cursors.Default;
        }

        private void sendMail(string customer_name, string content, string attach_file)
        {
            try
            {
                //SmtpClient client = new SmtpClient("smtp-mail.outlook.com");

                //client.Port = 587;
                //client.DeliveryMethod = SmtpDeliveryMethod.Network;
                //client.UseDefaultCredentials = false;
                //System.Net.NetworkCredential credentials =
                //    new System.Net.NetworkCredential("reixmxaeacyocstr@outlook.com", "reixmx!@#$%^&aeacyocstr");
                //client.EnableSsl = true;
                //client.Credentials = credentials;
                ////client.TargetName = "STARTTLS/smtp.office365.com";

                //string subject = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "](" + customer_name + ")";
                //string from = "reixmxaeacyocstr@outlook.com";
                //string to = "reixmxaeacyocstr@outlook.com";
                //string body = "pd generate";
                //using (MailMessage message = new MailMessage(from, to, subject, body))
                //using (Attachment attachment = new Attachment(attach_file, MediaTypeNames.Application.Octet))
                //{
                //    //ContentDisposition disposition = attachment.ContentDisposition;
                //    //disposition.CreationDate = File.GetCreationTime(attach_file);
                //    //disposition.ModificationDate = File.GetLastWriteTime(attach_file);
                //    //disposition.ReadDate = File.GetLastAccessTime(attach_file);

                //    message.Attachments.Add(attachment);
                //    client.UseDefaultCredentials = true;
                //    client.Send(message);
                //}
                SmtpClient client = new SmtpClient("smtp.163.com");

                client.Port = 587;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                System.Net.NetworkCredential credentials =
                    new System.Net.NetworkCredential("jeffggff", "OFKZRSRAESVFCMSF");
                //client.EnableSsl = true;
                client.Port = 25;
                client.Credentials = credentials;

                string from = "jeffggff@163.com";
                string to = "jeffggff@163.com";
                string message = content;
                string subject = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "](" + customer_name + ")";
                try
                {
                    var mail = new MailMessage(from.Trim(), to.Trim());
                    mail.Subject = subject;
                    mail.Body = message;
                    using (Attachment attachment = new Attachment(attach_file, MediaTypeNames.Application.Octet))
                    {
                        mail.Attachments.Add(attachment);
                        client.Send(mail);
                    }   
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Data);
                Console.WriteLine(ex.Data);
            }
            finally
            {

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }
    }

    /// <summary>
    /// SNTPClient is a C# class designed to connect to time servers on the Internet and
    /// fetch the current date and time. Optionally, it may update the time of the local system.
    /// The implementation of the protocol is based on the RFC 2030.
    /// 
    /// Public class members:
    /// 
    /// Initialize - Sets up data structure and prepares for connection.
    /// 
    /// Connect - Connects to the time server and populates the data structure.
    ///    It can also update the system time.
    /// 
    /// IsResponseValid - Returns true if received data is valid and if comes from
    /// a NTP-compliant time server.
    /// 
    /// ToString - Returns a string representation of the object.
    /// 
    /// -----------------------------------------------------------------------------
    /// Structure of the standard NTP header (as described in RFC 2030)
    ///                       1                   2                   3
    ///   0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |LI | VN  |Mode |    Stratum    |     Poll      |   Precision   |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                          Root Delay                           |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                       Root Dispersion                         |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                     Reference Identifier                      |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                                                               |
    ///  |                   Reference Timestamp (64)                    |
    ///  |                                                               |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                                                               |
    ///  |                   Originate Timestamp (64)                    |
    ///  |                                                               |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                                                               |
    ///  |                    Receive Timestamp (64)                     |
    ///  |                                                               |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                                                               |
    ///  |                    Transmit Timestamp (64)                    |
    ///  |                                                               |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                 Key Identifier (optional) (32)                |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    ///  |                                                               |
    ///  |                                                               |
    ///  |                 Message Digest (optional) (128)               |
    ///  |                                                               |
    ///  |                                                               |
    ///  +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    /// 
    /// -----------------------------------------------------------------------------
    /// 
    /// SNTP Timestamp Format (as described in RFC 2030)
    ///                         1                   2                   3
    ///     0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
    /// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    /// |                           Seconds                             |
    /// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    /// |                  Seconds Fraction (0-padded)                  |
    /// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
    /// 
    /// </summary>
    public class NTPClient
    {
        /// <summary>
        /// SNTP Data Structure Length
        /// </summary>
        private const byte SNTPDataLength = 48;

        /// <summary>
        /// SNTP Data Structure (as described in RFC 2030)
        /// </summary>
        byte[] SNTPData = new byte[SNTPDataLength];

        //Offset constants for timestamps in the data structure
        private const byte offReferenceID = 12;
        private const byte offReferenceTimestamp = 16;
        private const byte offOriginateTimestamp = 24;
        private const byte offReceiveTimestamp = 32;
        private const byte offTransmitTimestamp = 40;

        /// <summary>
        /// Leap Indicator Warns of an impending leap second to be inserted/deleted in the last  minute of the current day. 值为“11”时表示告警状态，时钟未被同步。为其他值时NTP本身不做处理
        /// </summary>
        public _LeapIndicator LeapIndicator
        {
            get
            {
                // Isolate the two most significant bits
                byte val = (byte)(SNTPData[0] >> 6);
                switch (val)
                {
                    case 0: return _LeapIndicator.NoWarning;
                    case 1: return _LeapIndicator.LastMinute61;
                    case 2: return _LeapIndicator.LastMinute59;
                    case 3: goto default;
                    default:
                        return _LeapIndicator.Alarm;
                }
            }
        }

        /// <summary>
        /// Version Number Version number of the protocol (3 or 4) NTP的版本号
        /// </summary>
        public byte VersionNumber
        {
            get
            {
                // Isolate bits 3 - 5
                byte val = (byte)((SNTPData[0] & 0x38) >> 3);
                return val;
            }
        }

        /// <summary>
        /// Mode 长度为3比特，表示NTP的工作模式。不同的值所表示的含义分别是：0未定义、1表示主动对等体模式、2表示被动对等体模式、3表示客户模式、4表示服务器模式、5表示广播模式或组播模式、6表示此报文为NTP控制报文、7预留给内部使用
        /// </summary>
        public _Mode Mode
        {
            get
            {
                // Isolate bits 0 - 3
                byte val = (byte)(SNTPData[0] & 0x7);
                switch (val)
                {
                    case 0:
                        return _Mode.Unknown;
                    case 6:
                        return _Mode.Unknown;
                    case 7:
                        return _Mode.Unknown;
                    default:
                        return _Mode.Unknown;
                    case 1:
                        return _Mode.SymmetricActive;
                    case 2:
                        return _Mode.SymmetricPassive;
                    case 3:
                        return _Mode.Client;
                    case 4:
                        return _Mode.Server;
                    case 5:
                        return _Mode.Broadcast;
                }
            }
        }

        /// <summary>
        /// Stratum 系统时钟的层数，取值范围为1～16，它定义了时钟的准确度。层数为1的时钟准确度最高，准确度从1到16依次递减，层数为16的时钟处于未同步状态，不能作为参考时钟
        /// </summary>
        public _Stratum Stratum
        {
            get
            {
                byte val = (byte)SNTPData[1];
                if (val == 0) return _Stratum.Unspecified;
                else
                    if (val == 1) return _Stratum.PrimaryReference;
                else
                        if (val <= 15) return _Stratum.SecondaryReference;
                else
                    return _Stratum.Reserved;
            }
        }

        /// <summary>
        /// Poll Interval (in seconds) Maximum interval between successive messages 轮询时间，即两个连续NTP报文之间的时间间隔
        /// </summary>
        public uint PollInterval
        {
            get
            {
                // Thanks to Jim Hollenhorst <hollenho@attbi.com>
                return (uint)(Math.Pow(2, (sbyte)SNTPData[2]));
            }
        }

        /// <summary>
        /// Precision (in seconds) Precision of the clock 系统时钟的精度
        /// </summary>
        public double Precision
        {
            get
            {
                // Thanks to Jim Hollenhorst <hollenho@attbi.com>
                return (Math.Pow(2, (sbyte)SNTPData[3]));
            }
        }

        /// <summary>
        /// Root Delay (in milliseconds) Round trip time to the primary reference source NTP服务器到主参考时钟的延迟
        /// </summary>
        public double RootDelay
        {
            get
            {
                int temp = 0;
                temp = 256 * (256 * (256 * SNTPData[4] + SNTPData[5]) + SNTPData[6]) + SNTPData[7];
                return 1000 * (((double)temp) / 0x10000);
            }
        }

        /// <summary>
        /// Root Dispersion (in milliseconds) Nominal error relative to the primary reference source 系统时钟相对于主参考时钟的最大误差
        /// </summary>
        public double RootDispersion
        {
            get
            {
                int temp = 0;
                temp = 256 * (256 * (256 * SNTPData[8] + SNTPData[9]) + SNTPData[10]) + SNTPData[11];
                return 1000 * (((double)temp) / 0x10000);
            }
        }

        /// <summary>
        /// Reference Identifier Reference identifier (either a 4 character string or an IP address)
        /// </summary>
        public string ReferenceID
        {
            get
            {
                string val = "";
                switch (Stratum)
                {
                    case _Stratum.Unspecified:
                        goto case _Stratum.PrimaryReference;
                    case _Stratum.PrimaryReference:
                        val += (char)SNTPData[offReferenceID + 0];
                        val += (char)SNTPData[offReferenceID + 1];
                        val += (char)SNTPData[offReferenceID + 2];
                        val += (char)SNTPData[offReferenceID + 3];
                        break;
                    case _Stratum.SecondaryReference:
                        switch (VersionNumber)
                        {
                            case 3:    // Version 3, Reference ID is an IPv4 address
                                string Address = SNTPData[offReferenceID + 0].ToString() + "." +
                                                 SNTPData[offReferenceID + 1].ToString() + "." +
                                                 SNTPData[offReferenceID + 2].ToString() + "." +
                                                 SNTPData[offReferenceID + 3].ToString();
                                try
                                {
                                    IPHostEntry Host = Dns.GetHostEntry(Address);
                                    val = Host.HostName + " (" + Address + ")";
                                }
                                catch (Exception)
                                {
                                    val = "N/A";
                                }
                                break;
                            case 4: // Version 4, Reference ID is the timestamp of last update
                                DateTime time = ComputeDate(GetMilliSeconds(offReferenceID));
                                // Take care of the time zone
                                TimeSpan offspan = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now);
                                val = (time + offspan).ToString();
                                break;
                            default:
                                val = "N/A";
                                break;
                        }
                        break;
                }

                return val;
            }
        }

        /// <summary>
        /// Reference Timestamp The time at which the clock was last set or corrected NTP系统时钟最后一次被设定或更新的时间
        /// </summary>
        public DateTime ReferenceTimestamp
        {
            get
            {
                DateTime time = ComputeDate(GetMilliSeconds(offReferenceTimestamp));
                // Take care of the time zone
                TimeSpan offspan = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now);
                return time + offspan;
            }
        }

        /// <summary>
        /// Originate Timestamp (T1)  The time at which the request departed the client for the server. 发送报文时的本机时间
        /// </summary>
        public DateTime OriginateTimestamp
        {
            get
            {
                return ComputeDate(GetMilliSeconds(offOriginateTimestamp));
            }
        }

        /// <summary>
        /// Receive Timestamp (T2) The time at which the request arrived at the server. 报文到达NTP服务器时的服务器时间
        /// </summary>
        public DateTime ReceiveTimestamp
        {
            get
            {
                DateTime time = ComputeDate(GetMilliSeconds(offReceiveTimestamp));
                // Take care of the time zone
                TimeSpan offspan = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now);
                return time + offspan;
            }
        }

        /// <summary>
        /// Transmit Timestamp (T3) The time at which the reply departed the server for client.  报文从NTP服务器离开时的服务器时间
        /// </summary>
        public DateTime TransmitTimestamp
        {
            get
            {
                DateTime time = ComputeDate(GetMilliSeconds(offTransmitTimestamp));
                // Take care of the time zone
                TimeSpan offspan = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Now);
                return time + offspan;
            }
            set
            {
                SetDate(offTransmitTimestamp, value);
            }
        }

        /// <summary>
        /// Destination Timestamp (T4) The time at which the reply arrived at the client. 接收到来自NTP服务器返回报文时的本机时间
        /// </summary>
        public DateTime DestinationTimestamp;

        /// <summary>
        /// Round trip delay (in milliseconds) The time between the departure of request and arrival of reply 报文从本地到NTP服务器的往返时间
        /// </summary>
        public double RoundTripDelay
        {
            get
            {
                // Thanks to DNH <dnharris@csrlink.net>
                TimeSpan span = (DestinationTimestamp - OriginateTimestamp) - (ReceiveTimestamp - TransmitTimestamp);
                return span.TotalMilliseconds;
            }
        }

        /// <summary>
        /// Local clock offset (in milliseconds)  The offset of the local clock relative to the primary reference source.本机相对于NTP服务器（主时钟）的时间差
        /// </summary>
        public double LocalClockOffset
        {
            get
            {
                // Thanks to DNH <dnharris@csrlink.net>
                TimeSpan span = (ReceiveTimestamp - OriginateTimestamp) + (TransmitTimestamp - DestinationTimestamp);
                return span.TotalMilliseconds / 2;
            }
        }

        /// <summary>
        /// Compute date, given the number of milliseconds since January 1, 1900
        /// </summary>
        /// <param name="milliseconds"></param>
        /// <returns></returns>
        private DateTime ComputeDate(ulong milliseconds)
        {
            TimeSpan span = TimeSpan.FromMilliseconds((double)milliseconds);
            DateTime time = new DateTime(1900, 1, 1);
            time += span;
            return time;
        }

        /// <summary>
        /// Compute the number of milliseconds, given the offset of a 8-byte array
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        private ulong GetMilliSeconds(byte offset)
        {
            ulong intpart = 0, fractpart = 0;

            for (int i = 0; i <= 3; i++)
            {
                intpart = 256 * intpart + SNTPData[offset + i];
            }
            for (int i = 4; i <= 7; i++)
            {
                fractpart = 256 * fractpart + SNTPData[offset + i];
            }
            ulong milliseconds = intpart * 1000 + (fractpart * 1000) / 0x100000000L;
            return milliseconds;
        }

        /// <summary>
        /// Compute the 8-byte array, given the date
        /// </summary>
        /// <param name="offset"></param>
        /// <param name="date"></param>
        private void SetDate(byte offset, DateTime date)
        {
            ulong intpart = 0, fractpart = 0;
            DateTime StartOfCentury = new DateTime(1900, 1, 1, 0, 0, 0);    // January 1, 1900 12:00 AM

            ulong milliseconds = (ulong)(date - StartOfCentury).TotalMilliseconds;
            intpart = milliseconds / 1000;
            fractpart = ((milliseconds % 1000) * 0x100000000L) / 1000;

            ulong temp = intpart;
            for (int i = 3; i >= 0; i--)
            {
                SNTPData[offset + i] = (byte)(temp % 256);
                temp = temp / 256;
            }

            temp = fractpart;
            for (int i = 7; i >= 4; i--)
            {
                SNTPData[offset + i] = (byte)(temp % 256);
                temp = temp / 256;
            }
        }

        /// <summary>
        /// Initialize the NTPClient data
        /// </summary>
        private void Initialize()
        {
            // Set version number to 4 and Mode to 3 (client)
            SNTPData[0] = 0x1B;
            // Initialize all other fields with 0
            for (int i = 1; i < 48; i++)
            {
                SNTPData[i] = 0;
            }
            // Initialize the transmit timestamp
            TransmitTimestamp = DateTime.Now;
        }

        /// <summary>
        /// The IPAddress of the time server we're connecting to
        /// </summary>
        private IPAddress serverAddress = null;


        /// <summary>
        /// Constractor with HostName
        /// </summary>
        /// <param name="host"></param>
        public NTPClient(string host)
        {
            //string host = "ntp1.aliyun.com";
            //string host = "0.asia.pool.ntp.org";
            //string host = "1.asia.pool.ntp.org";
            //string host = "www.ntp.org/";

            // Resolve server address
            IPHostEntry hostadd = Dns.GetHostEntry(host);
            foreach (IPAddress address in hostadd.AddressList)
            {
                if (address.AddressFamily == AddressFamily.InterNetwork) //只支持IPV4协议的IP地址
                {
                    serverAddress = address;
                    break;
                }
            }

            if (serverAddress == null)
                throw new Exception("Can't get any ipaddress infomation");
        }

        /// <summary>
        /// Constractor with IPAddress
        /// </summary>
        /// <param name="address"></param>
        public NTPClient(IPAddress address)
        {
            if (address == null)
                throw new Exception("Can't get any ipaddress infomation");

            serverAddress = address;
        }

        /// <summary>
        /// Connect to the time server and update system time
        /// </summary>
        /// <param name="updateSystemTime"></param>
        public void Connect(bool updateSystemTime, int timeout = 3000)
        {
            IPEndPoint EPhost = new IPEndPoint(serverAddress, 123);

            //Connect the time server
            using (UdpClient TimeSocket = new UdpClient())
            {
                TimeSocket.Connect(EPhost);

                // Initialize data structure
                Initialize();
                TimeSocket.Send(SNTPData, SNTPData.Length);
                TimeSocket.Client.ReceiveTimeout = timeout;
                SNTPData = TimeSocket.Receive(ref EPhost);
                if (!IsResponseValid)
                    throw new Exception("Invalid response from " + serverAddress.ToString());
            }
            DestinationTimestamp = DateTime.Now;

            if (updateSystemTime)
                SetTime();
        }

        /// <summary>
        /// Check if the response from server is valid
        /// </summary>
        /// <returns></returns>
        public bool IsResponseValid
        {
            get
            {
                return !(SNTPData.Length < SNTPDataLength || Mode != _Mode.Server);
            }
        }

        /// <summary>
        /// Converts the object to string
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder(512);
            sb.Append("Leap Indicator: ");
            switch (LeapIndicator)
            {
                case _LeapIndicator.NoWarning:
                    sb.Append("No warning");
                    break;
                case _LeapIndicator.LastMinute61:
                    sb.Append("Last minute has 61 seconds");
                    break;
                case _LeapIndicator.LastMinute59:
                    sb.Append("Last minute has 59 seconds");
                    break;
                case _LeapIndicator.Alarm:
                    sb.Append("Alarm Condition (clock not synchronized)");
                    break;
            }
            sb.AppendFormat("\r\nVersion number: {0}\r\n", VersionNumber);
            sb.Append("Mode: ");
            switch (Mode)
            {
                case _Mode.Unknown:
                    sb.Append("Unknown");
                    break;
                case _Mode.SymmetricActive:
                    sb.Append("Symmetric Active");
                    break;
                case _Mode.SymmetricPassive:
                    sb.Append("Symmetric Pasive");
                    break;
                case _Mode.Client:
                    sb.Append("Client");
                    break;
                case _Mode.Server:
                    sb.Append("Server");
                    break;
                case _Mode.Broadcast:
                    sb.Append("Broadcast");
                    break;
            }
            sb.Append("\r\nStratum: ");

            switch (Stratum)
            {
                case _Stratum.Unspecified:
                case _Stratum.Reserved:
                    sb.Append("Unspecified");
                    break;
                case _Stratum.PrimaryReference:
                    sb.Append("Primary Reference");
                    break;
                case _Stratum.SecondaryReference:
                    sb.Append("Secondary Reference");
                    break;
            }
            sb.AppendFormat("\r\nLocal Time T3: {0:yyyy-MM-dd HH:mm:ss:fff}", TransmitTimestamp);
            sb.AppendFormat("\r\nDestination Time T4: {0:yyyy-MM-dd HH:mm:ss:fff}", DestinationTimestamp);
            sb.AppendFormat("\r\nPrecision: {0} s", Precision);
            sb.AppendFormat("\r\nPoll Interval:{0} s", PollInterval);
            sb.AppendFormat("\r\nReference ID: {0}", ReferenceID.ToString().Replace("\0", string.Empty));
            sb.AppendFormat("\r\nRoot Delay: {0} ms", RootDelay);
            sb.AppendFormat("\r\nRoot Dispersion: {0} ms", RootDispersion);
            sb.AppendFormat("\r\nRound Trip Delay: {0} ms", RoundTripDelay);
            sb.AppendFormat("\r\nLocal Clock Offset: {0} ms", LocalClockOffset);
            sb.AppendFormat("\r\nReferenceTimestamp: {0:yyyy-MM-dd HH:mm:ss:fff}", ReferenceTimestamp);
            sb.Append("\r\n");

            return sb.ToString();
        }

        /// <summary>
        /// SYSTEMTIME structure used by SetSystemTime
        /// </summary>
        [StructLayoutAttribute(LayoutKind.Sequential)]
        private struct SYSTEMTIME
        {
            public short year;
            public short month;
            public short dayOfWeek;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short milliseconds;
        }

        [DllImport("kernel32.dll")]
        static extern bool SetLocalTime(ref SYSTEMTIME time);


        /// <summary>
        /// Set system time according to transmit timestamp 把本地时间设置为获取到的时钟时间
        /// </summary>
        public void SetTime()
        {
            SYSTEMTIME st;

            DateTime trts = DateTime.Now.AddMilliseconds(LocalClockOffset);

            st.year = (short)trts.Year;
            st.month = (short)trts.Month;
            st.dayOfWeek = (short)trts.DayOfWeek;
            st.day = (short)trts.Day;
            st.hour = (short)trts.Hour;
            st.minute = (short)trts.Minute;
            st.second = (short)trts.Second;
            st.milliseconds = (short)trts.Millisecond;

            SetLocalTime(ref st);
        }
    }

    /// <summary>
    /// Leap indicator field values
    /// </summary>
    public enum _LeapIndicator
    {
        NoWarning,        // 0 - No warning
        LastMinute61,    // 1 - Last minute has 61 seconds
        LastMinute59,    // 2 - Last minute has 59 seconds
        Alarm            // 3 - Alarm condition (clock not synchronized)
    }

    /// <summary>
    /// Mode field values
    /// </summary>
    public enum _Mode
    {
        SymmetricActive,    // 1 - Symmetric active
        SymmetricPassive,    // 2 - Symmetric pasive
        Client,                // 3 - Client
        Server,                // 4 - Server
        Broadcast,            // 5 - Broadcast
        Unknown                // 0, 6, 7 - Reserved
    }

    /// <summary>
    /// Stratum field values
    /// </summary>
    public enum _Stratum
    {
        Unspecified,            // 0 - unspecified or unavailable
        PrimaryReference,        // 1 - primary reference (e.g. radio-clock)
        SecondaryReference,        // 2-15 - secondary reference (via NTP or SNTP)
        Reserved                // 16-255 - reserved
    }
}
