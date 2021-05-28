using ExcelDataReader;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Threading;
using ClosedXML;
using ClosedXML.Excel;

namespace idk
{
    public partial class mainForm : Form
    {
        public string tracking;
        public string result;
        public string sinNo;

        private static loading _load = null;

        //FOR DHL
        public string sinNo_DHL;

        public mainForm()
        {
            InitializeComponent();
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection("Data Source=192.168.100.9\\sql05;Initial Catalog=NEP;User ID=viewer; Password=");
            conn.Open();

            SqlCommand query = new SqlCommand("SELECT RefNo, TrackingNo, DONo, DODate, DeliveryCompID FROM c_TranDOMt WHERE RefNo =@SinNo", conn);
            query.Parameters.AddWithValue("@SinNo", textBox1.Text);
            SqlDataReader sdr = query.ExecuteReader();
            while (sdr.Read())
            {
                textBox2.Text = sdr.GetValue(4).ToString();
                textBox3.Text = sdr.GetValue(1).ToString();
                textBox4.Text = sdr.GetValue(2).ToString();
                textBox9.Text = sdr.GetValue(3).ToString();

                if (textBox4.Text == "")
                {
                    textBox4.Text = "Not Generated";
                }

                sinNo = textBox1.Text;                
                tracking = textBox3.Text;
            }

            /*If statements to change delivery company to text*/            
            if (textBox2.Text == "1") { textBox2.Text = "POST MALAYSIA"; }
            else if (textBox2.Text == "2") { textBox2.Text = "SKY NET"; }
            else if (textBox2.Text == "3") { textBox2.Text = "OTHERS"; }
            else if (textBox2.Text == "4") { textBox2.Text = "NATIONWIDE"; }
            else if (textBox2.Text == "5") { textBox2.Text = "CITY LINK"; }
            else if (textBox2.Text == "6") { textBox2.Text = "SURE REACH"; }
            else if (textBox2.Text == "7") { textBox2.Text = "UPRIGHT"; }
            else if (textBox2.Text == "8") { textBox2.Text = "PGEON"; }
            else if (textBox2.Text == "9") { textBox2.Text = "ABX"; }
            else if (textBox2.Text == "10") { textBox2.Text = "POS LAJU"; }
            else if (textBox2.Text == "11") { textBox2.Text = "AIRPARK"; }
            else if (textBox2.Text == "12") { textBox2.Text = "CJ CENTURY"; }
            else if (textBox2.Text == "13") { textBox2.Text = "ARAMEX"; }
            else if (textBox2.Text == "14") { textBox2.Text = "DHL eCOMMERCE"; }
            else if (textBox2.Text == "15") { textBox2.Text = "Ninja Van"; }
            else if (textBox2.Text == "16") { textBox2.Text = "Gdex"; }
            else if (textBox2.Text == "17") { textBox2.Text = "QXpress"; }
            else if (textBox2.Text == "18") { textBox2.Text = "ARISSTO DELIVERY"; }
            else if (textBox2.Text == "19") { textBox2.Text = "Versa Drive"; }
            else if (textBox2.Text == "20") { textBox2.Text = "ArrangedbyQ21"; }
            else if (textBox2.Text == "21") { textBox2.Text = "ArrangedbyQ22"; }
            else if (textBox2.Text == "22") { textBox2.Text = "ArrangedbyQ23"; }
            else if (textBox2.Text == "23") { textBox2.Text = "ArrangedbyQ24"; }
            else if (textBox2.Text == "24") { textBox2.Text = "ArrangedbyS20"; }
            else if (textBox2.Text == "25") { textBox2.Text = "ArrangedbyVS3"; }
            else if (textBox2.Text == "26") { textBox2.Text = "SKYLOG"; }
            else if (textBox2.Text == "27") { textBox2.Text = "J&T"; }
            else if (textBox2.Text == "28") { textBox2.Text = "GRACIA ORBIS"; }
            else { textBox2.Text = "INVALID"; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Under Development");
            if (textBox2.Text == "DHL eCOMMERCE")
            {
                Process.Start("https://ecommerceportal.dhl.com/track/?ref=" + tracking);
            }
            else if (textBox2.Text == "POS LAJU")
            {
                Process.Start("https://www.tracking.my/poslaju/" + tracking);
            }
            else if (textBox2.Text == "Ninja Van")
            {
                Process.Start("https://www.ninjavan.co/en-my/tracking?id=" + tracking);
            }
            else if (textBox2.Text == "GRACIA ORBIS")
            {
                Process.Start("https://lastmile.milenow.com/company/Qjby/merchant/register?trackingNos=" + tracking + "&form_type=Track&form_type1=+Track+");
            }
            else if (textBox2.Text == "ARISSTO DELIVERY")
            {
                MessageBox.Show("Deliver by our own lorry, dunno how track");
            }
            else
            {
                MessageBox.Show("Invalid Tracking number");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            if (textBox5.Text == "" || textBox8.Text == "")
            {
                textBox5.Text = "Token yet to be requested, please request it :D";
                button5.Enabled = false;
                button7.Enabled = false;
            }           
        }                        
        private void button6_Click(object sender, EventArgs e)
        {
            userGuide ug = new userGuide();
            ug.ShowDialog();
        }
        private void label7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("HAHA");
        }

        /*----------------------------------------TOKEN REQUEST----------------------------------------*/
        private void button3_Click(object sender, EventArgs e)
        {
            loading ld = new loading();
            ld.Show();
            MessageBox.Show("~Getting Token~");           

            WebRequest httpWebRequest = WebRequest.Create("https://api.dhlecommerce.dhl.com/rest/v1/OAuth/AccessToken?clientId=MTYzNzcxMzQzOQ==&password=MTQ4MDg3O2304211619163487&returnFormat=json");

            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "GET";

            HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                result = streamReader.ReadToEnd();

                try
                {
                    string jsonTok = JObject.Parse(result)["accessTokenResponse"]["token"].ToString();

                    Console.WriteLine(jsonTok);
                    textBox5.Text = jsonTok;

                    MessageBox.Show("Token assigned");
                    ld.Close();
                }
                catch(Exception ex)
                {
                    errorForm ef = new errorForm();
                    ef.ShowDialog();
                    MessageBox.Show(ex.ToString());
                }
            }
            button5.Enabled = true;
            button7.Enabled = true;
        }

        /*----------------------------------------TRACKING----------------------------------------*/
        private void button5_Click(object sender, EventArgs e)
        {
            if(textBox8.Text == "")
            {
                MessageBox.Show("Batch number is empty!");
            }
            else
            {
                MessageBox.Show("~To find tracking number~");

                OpenFileDialog ope = new OpenFileDialog();
                ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                if (ope.ShowDialog() == DialogResult.Cancel)
                {
                    return;
                }
                FileStream stream = new FileStream(ope.FileName, FileMode.Open);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet ds = excelReader.AsDataSet();

                loading ld = new loading();
                ld.Show();

                int counter = 1;
                var wb = new XLWorkbook();
                var worksheet = wb.Worksheets.Add("ToGENDO");
                //worksheet.Cell("A:Z").DataType = XLDataType.Text;
                //worksheet.DataType = XLDataType.Text;

                foreach (DataTable table in ds.Tables)
                {
                    foreach (DataRow dr in table.Rows)
                    {
                        if (table.Rows.IndexOf(dr) != 0)
                        {
                            string SINs = Convert.ToString(dr[1]);
                            //MessageBox.Show(SINs);       
                            textBox7.Text = SINs;

                            WebRequest httpWebRequest = WebRequest.Create("https://api.dhlecommerce.dhl.com/rest/v3/Tracking");
                            string json = "{\r\n    \"trackItemRequest\": {\r\n        \"hdr\": {\r\n            \"messageType\": \"TRACKITEM\",\r\n            \"accessToken\": \"" + textBox5.Text + "\",\r\n            \"messageDateTime\": \"2021-04-23T17:13:10+08:00\",\r\n            \"messageVersion\": \"1.0\",\r\n            \"messageLanguage\": \"en\"\r\n        },\r\n        \"bd\": {       \r\n            \"customerAccountId\": null,     \r\n            \"soldToAccountId\": null,\r\n            \"pickupAccountId\": null,\r\n            \"ePODRequired\": \"N\",\r\n            \"trackingReferenceNumber\": [\r\n                \"" + "MYCGU" + SINs + "\"\r\n            ]\r\n        }\r\n    }\r\n}";

                            httpWebRequest.ContentType = "application/json";
                            httpWebRequest.Method = "POST";

                            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                            {
                                streamWriter.Write(json);
                                streamWriter.Flush();
                                streamWriter.Close();
                            }

                            HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                            using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
                            {
                                string result = streamReader.ReadToEnd();
                                //Console.WriteLine(result);

                                try
                                {
                                    string jsonTrck = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["trackingID"].ToString();

                                    textBox6.Text = jsonTrck;

                                    worksheet.Cell(counter, 1).SetValue(SINs);   //write sins
                                    worksheet.Cell(counter, 2).SetValue("DHL eCommerce");   //delivery comp
                                    worksheet.Cell(counter, 3).SetValue("NTW2");   //location
                                    worksheet.Cell(counter, 4).SetValue("MY8_ARISSTO");   //userID
                                    worksheet.Cell(counter, 5).SetValue("DOAPI" + textBox8.Text);   //batch
                                    worksheet.Cell(counter, 6).SetValue(jsonTrck);   //tracking num
                                }
                                catch(Exception ex)
                                {
                                    errorForm ef = new errorForm();
                                    ef.ShowDialog();
                                    MessageBox.Show(ex.ToString());
                                }
                            }

                            Console.WriteLine(SINs);   //debug line
                            Console.WriteLine(counter);   //debug line                        

                            counter++;
                        }
                    }
                }

                Console.WriteLine(textBox8.Text);
                wb.SaveAs("C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\DHL API\\DOAPI" + textBox8.Text + ".xlsx");
                excelReader.Close();
                stream.Close();
                MessageBox.Show("File saved at: \n C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\DHL API\\DOAPI" + textBox8.Text + ".xlsx");
                ld.Close();

                sinNo_DHL = textBox7.Text;
            }
        }

        /*----------------------------------------IMPORT TO GEN DO----------------------------------------*/
        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("~For Safety, Please ensure that GenDO job is disabled and currently not active~");
            MessageBox.Show("~To Import GEN DO~");

            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ope.ShowDialog() == DialogResult.Cancel)
                return;
            FileStream stream = new FileStream(ope.FileName, FileMode.Open);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            loading ld = new loading();
            ld.Show();

            GenDODataContext gddc = new GenDODataContext();

            foreach(DataTable dt in result.Tables)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    GenerateDO_YikTing gdyk = new GenerateDO_YikTing()
                    {
                        #region DataSet
                        SinNo = Convert.ToString(dr[0]),
                        TranCode = Convert.ToString(dr[1]),
                        Location = Convert.ToString(dr[2]),
                        UserID = Convert.ToString(dr[3]),
                        Remarks = Convert.ToString(dr[4]),
                        trackingno = Convert.ToString(dr[5]),
                        isdeleted = 0
                        #endregion
                    };
                    gddc.GenerateDO_YikTings.InsertOnSubmit(gdyk);
                }
            }
            gddc.SubmitChanges();
            excelReader.Close();
            stream.Close();
            MessageBox.Show(this, "Data imported \nContinue operation via SSMS");
            ld.Close();
        }

        /*----------------------------------------DELIVERY STATUS----------------------------------------*/
        private void button7_Click(object sender, EventArgs e)
        {
            if(textBox8.Text == "")
            {
                MessageBox.Show("Batch number is empty!");
            }
            else
            {
                MessageBox.Show("~Finding delivery status~");

                OpenFileDialog ope = new OpenFileDialog();
                ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                if (ope.ShowDialog() == DialogResult.Cancel)
                {
                    return;
                }
                FileStream stream = new FileStream(ope.FileName, FileMode.Open);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet ds = excelReader.AsDataSet();

                loading ld = new loading();
                ld.Show();

                int counter = 1;
                var wb = new XLWorkbook();
                var worksheet = wb.Worksheets.Add("TrackingStatus");

                foreach (DataTable table in ds.Tables)
                {
                    foreach (DataRow dr in table.Rows)
                    {
                        if (table.Rows.IndexOf(dr) != 0)
                        {
                            string SINs = Convert.ToString(dr[1]);
                            //MessageBox.Show(SINs);       
                            textBox7.Text = SINs;

                            WebRequest httpWebRequest = WebRequest.Create("https://api.dhlecommerce.dhl.com/rest/v3/Tracking");
                            string json = "{\r\n    \"trackItemRequest\": {\r\n        \"hdr\": {\r\n            \"messageType\": \"TRACKITEM\",\r\n            \"accessToken\": \"" + textBox5.Text + "\",\r\n            \"messageDateTime\": \"2021-04-23T17:13:10+08:00\",\r\n            \"messageVersion\": \"1.0\",\r\n            \"messageLanguage\": \"en\"\r\n        },\r\n        \"bd\": {       \r\n            \"customerAccountId\": null,     \r\n            \"soldToAccountId\": null,\r\n            \"pickupAccountId\": null,\r\n            \"ePODRequired\": \"N\",\r\n            \"trackingReferenceNumber\": [\r\n                \"" + "MYCGU" + SINs + "\"\r\n            ]\r\n        }\r\n    }\r\n}";

                            httpWebRequest.ContentType = "application/json";
                            httpWebRequest.Method = "POST";

                            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                            {
                                streamWriter.Write(json);
                                streamWriter.Flush();
                                streamWriter.Close();
                            }

                            HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                            using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
                            {
                                string result = streamReader.ReadToEnd();
                                //Console.WriteLine(result);

                                try
                                {
                                    string jsonTrck = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["trackingID"].ToString();

                                    string jsonStat = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["events"][0]["description"].ToString();                            

                                    worksheet.Cell(counter, 1).SetValue(SINs);   //write sins
                                    worksheet.Cell(counter, 2).SetValue("DHL eCommerce");   //delivery comp   
                                    worksheet.Cell(counter, 3).SetValue(jsonTrck);   //tracking num
                                    worksheet.Cell(counter, 4).SetValue(jsonStat);   //tracking status
                                }
                                catch(Exception ex)
                                {
                                    errorForm ef = new errorForm();
                                    ef.ShowDialog();
                                    MessageBox.Show(ex.ToString());
                                }

                            }

                            Console.WriteLine(SINs);   //debug line
                            Console.WriteLine(counter);   //debug line                        

                            counter++;
                        }
                    }
                }

                Console.WriteLine(textBox8.Text);
                wb.SaveAs("C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\DHL API\\Tracking" + textBox8.Text + ".xlsx");
                excelReader.Close();
                stream.Close();
                MessageBox.Show("File saved at: \n C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\DHL API\\Tracking" + textBox8.Text + ".xlsx");
                ld.Close();

                sinNo_DHL = textBox7.Text;
            }
        }

        /*----------------------------------------TO IMPORT DELIVERY STATUS----------------------------------------*/
        private void button8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("~To Update Delivery Status~");

            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ope.ShowDialog() == DialogResult.Cancel)
                return;
            FileStream stream = new FileStream(ope.FileName, FileMode.Open);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            loading ld = new loading();
            ld.Show();

            DeliveryStatusDataContext dsdc = new DeliveryStatusDataContext();

            foreach (DataTable dt in result.Tables)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    Delivery_Status_WK dswk = new Delivery_Status_WK()
                    {
                        #region DataSet
                        SinNo = Convert.ToString(dr[0]),
                        DeliveryComp = Convert.ToString(dr[1]),
                        TrackingNo = Convert.ToString(dr[2]),
                        DeliveryStatus = Convert.ToString(dr[3])
                        #endregion
                    };
                    dsdc.Delivery_Status_WKs.InsertOnSubmit(dswk);
                }
            }
            dsdc.SubmitChanges();
            excelReader.Close();
            stream.Close();
            MessageBox.Show(this, "Data imported \nPlease check via SSMS");
            ld.Close();
        }

        /*----------------------------------------FOR SINGLE TRACKING----------------------------------------*/
        private void button9_Click(object sender, EventArgs e)
        {
            loading ld = new loading();
            ld.Show();

            WebRequest httpWebRequest = WebRequest.Create("https://api.dhlecommerce.dhl.com/rest/v3/Tracking");
            string json = "{\r\n    \"trackItemRequest\": {\r\n        \"hdr\": {\r\n            \"messageType\": \"TRACKITEM\",\r\n            \"accessToken\": \"" + textBox5.Text + "\",\r\n            \"messageDateTime\": \"2021-04-23T17:13:10+08:00\",\r\n            \"messageVersion\": \"1.0\",\r\n            \"messageLanguage\": \"en\"\r\n        },\r\n        \"bd\": {       \r\n            \"customerAccountId\": null,     \r\n            \"soldToAccountId\": null,\r\n            \"pickupAccountId\": null,\r\n            \"ePODRequired\": \"N\",\r\n            \"trackingReferenceNumber\": [\r\n                \"" + "MYCGU" + textBox7.Text + "\"\r\n            ]\r\n        }\r\n    }\r\n}";

            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();
            }

            HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                string result = streamReader.ReadToEnd();
                //Console.WriteLine(result);

                try
                {
                    string jsonTrck = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["trackingID"].ToString();
                    string jsonStat = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["events"][0]["description"].ToString();

                    textBox6.Text = jsonTrck;

                    MessageBox.Show("Tracking ID: " + jsonTrck + "\n\n" + "Delivery Status: " + jsonStat);
                }
                catch(Exception ex)
                {
                    errorForm ef = new errorForm();
                    ef.ShowDialog();                   
                    MessageBox.Show(ex.ToString());
                }
            }

            ld.Close();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            MessageBox.Show("~To find tracking number~");

            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ope.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            FileStream stream = new FileStream(ope.FileName, FileMode.Open);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet ds = excelReader.AsDataSet();

            loading ld = new loading();
            ld.Show();

            int counter = 1;
            var wb = new XLWorkbook();
            var worksheet = wb.Worksheets.Add("ToGENDO");
            //worksheet.Cell("A:Z").DataType = XLDataType.Text;
            //worksheet.DataType = XLDataType.Text;

            foreach (DataTable table in ds.Tables)
            {
                foreach (DataRow dr in table.Rows)
                {
                    if (table.Rows.IndexOf(dr) != 0)
                    {
                        string SINs = Convert.ToString(dr[1]);
                        //MessageBox.Show(SINs);       
                        textBox7.Text = SINs;

                        WebRequest httpWebRequest = WebRequest.Create("https://api.dhlecommerce.dhl.com/rest/v3/Tracking");
                        string json = "{\r\n    \"trackItemRequest\": {\r\n        \"hdr\": {\r\n            \"messageType\": \"TRACKITEM\",\r\n            \"accessToken\": \"" + textBox5.Text + "\",\r\n            \"messageDateTime\": \"2021-04-23T17:13:10+08:00\",\r\n            \"messageVersion\": \"1.0\",\r\n            \"messageLanguage\": \"en\"\r\n        },\r\n        \"bd\": {       \r\n            \"customerAccountId\": null,     \r\n            \"soldToAccountId\": null,\r\n            \"pickupAccountId\": null,\r\n            \"ePODRequired\": \"N\",\r\n            \"trackingReferenceNumber\": [\r\n                \"" + "MYCGU" + SINs + "\"\r\n            ]\r\n        }\r\n    }\r\n}";

                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";

                        using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            streamWriter.Write(json);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }

                        HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            string result = streamReader.ReadToEnd();
                            //Console.WriteLine(result);

                            try
                            {
                                string jsonTrck = JObject.Parse(result)["trackItemResponse"]["bd"]["shipmentItems"][0]["trackingID"].ToString();

                                textBox6.Text = jsonTrck;
                                string cut = SINs.Substring(0, 17);
                                worksheet.Cell(counter, 1).SetValue(cut);   //write sins
                                worksheet.Cell(counter, 2).SetValue("DHL eCommerce");   //delivery comp
                                worksheet.Cell(counter, 3).SetValue("NTW2");   //location
                                worksheet.Cell(counter, 4).SetValue("MY8_ARISSTO");   //userID
                                worksheet.Cell(counter, 5).SetValue("DOAPI" + textBox8.Text);   //batch
                                worksheet.Cell(counter, 6).SetValue(jsonTrck);   //tracking num
                            }
                            catch (Exception ex)
                            {
                                errorForm ef = new errorForm();
                                ef.ShowDialog();
                                MessageBox.Show(ex.ToString());
                            }
                        }

                        Console.WriteLine(SINs);   //debug line
                        Console.WriteLine(counter);   //debug line                        

                        counter++;
                    }
                }
            }

            Console.WriteLine(textBox8.Text);
            wb.SaveAs("C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\TestC#\\Mapp" + textBox8.Text + ".xlsx");
            excelReader.Close();
            stream.Close();
            MessageBox.Show("File saved at: \n C:\\Users\\WYCHIN\\Desktop\\YONG_TEMP\\TestC#\\Mapp" + textBox8.Text + ".xlsx");
            ld.Close();

            sinNo_DHL = textBox7.Text;
        }
    }
}
