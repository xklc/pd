using com.pd.extract;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
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

                long et = 1690464318000; //2/7/15 14:00:00            
                long currentMillis = getCurrentMillis(DateTime.Now.Ticks);
                if (currentMillis >= et)
                {
                    Close();
                }

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
                contentlines.Add("expiry-date=indefinite");
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
                    sendMail(organizationName, license_file);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                //throw ex;
            }
            this.Cursor = Cursors.Default;
        }

        private void sendMail(string customer_name, string attach_file)
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
                string message = "pd generate";
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
}
