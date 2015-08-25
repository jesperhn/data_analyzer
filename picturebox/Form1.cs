using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows.Media.Media3D;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace picturebox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            btnNext.Enabled = false;
            btnPrev.Enabled = false;
            btnChart.Enabled = false;
            trackBar1.Value = trackBar1.Maximum;
        }

        Bitmap drawBitmap;
        public System.Drawing.Graphics g;
        public System.Drawing.Graphics g2;
        private System.Drawing.Pen pen1 = new System.Drawing.Pen(Color.Red, 80F);
        private System.Drawing.Pen pen2 = new System.Drawing.Pen(Color.White, 80F);

        public static int cellID = 0;
        public static int[] iArray = new int[50];
        public static int arrayID = 0;
        public static Vector3D[] xyValues = new Vector3D[200];
        public static int k = 0;
        public static int color = 0;
        public static int totalProcessTime = 0;
        public static int totalNotPlanned = 0;
        public static int totalPlanned = 0;
        public static int totalAutoMode = 0;

        public static class Globals
        {
            public static String s1 = null;
            public static Int32 VALUE = 40;
            public static Int32 xTemp = 0;
            public static Int32 i = 0;
            public static Int32 j = 0;
            public static string readText = null;            
        }

        public string toTime(int time)
        {
            int seconds = 0;
            int minutes = 0;
            int hours = 0;
            string result;
            //timer
            if (time >= 60)
            {
                minutes = time / 60;
                seconds = time % 60;
            }
            else
            {
                seconds = time;
            }

            if (minutes >= 60)
            {
                hours = minutes / 60;
                minutes = minutes % 60;
            }
            else
            {
                minutes = minutes;
            }

            if (seconds < 10 && minutes < 10 && hours < 10)
            {
                result = ("0" + hours.ToString() + ":0" + minutes.ToString()) + ":0" + seconds.ToString();
                return result;
            }
            else if (hours < 10 && minutes < 10)
            {
                result = ("0" + hours.ToString()) + ":0" + minutes.ToString() + ":" + seconds.ToString();
                return result;
            }
            else if (hours < 10 && seconds < 10)
            {
                result = ("0" + hours.ToString()) + ":" + minutes.ToString() + ":0" + seconds.ToString();
                return result;
            }
            else if (hours < 10)
            {
                result = ("0" + hours.ToString()) + ":" + minutes.ToString() + ":" + seconds.ToString();
                return result;
            }
            else if (minutes < 10)
            {
                result = (hours.ToString() + ":0" + minutes.ToString()) + ":" + seconds.ToString();
                return result;
            }
            else if (seconds < 10)
            {
                result = (hours.ToString() + ":" + minutes.ToString()) + ":0" + seconds.ToString();
                return result;
            }
            else
            {
                result = (hours.ToString() + ":" + minutes.ToString()) + ":" + seconds.ToString();
                return result;
            }
        }

        public static Bitmap resizeImage(Bitmap imgToResize, Size size)
        {
              return (Bitmap)(new Bitmap(imgToResize, size));
        }

        private void draw_invalidData()
        {
            drawBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            g = Graphics.FromImage(drawBitmap);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            
            g.DrawString("Invalid Data", new Font("Tahoma", 60), Brushes.Black, 0, 0);
            g.Flush();
            pictureBox1.Image = drawBitmap;
            pictureBox1.Image.Tag = "jhn";
        }

        private void draw(int xStart, int xEnd, int y)
        {
            if ((xEnd/Globals.VALUE) > drawBitmap.Size.Width)
            {
                pictureBox1.Refresh();
                drawBitmap = resizeImage(drawBitmap, new Size((pictureBox1.Size.Width + xEnd / Globals.VALUE), 290));
            }
            g = Graphics.FromImage(drawBitmap);
            g.DrawLine(pen1, xStart / Globals.VALUE, y, xEnd / Globals.VALUE, y); //x1,y1,x2,y2
            pictureBox1.Image = drawBitmap;
            g.Dispose();
            xyValues[k] = new Vector3D(xStart,xEnd,color);
            k++;
        }
        private void iterateText()
        {
            drawBitmap.Dispose();
            drawBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height); // create a bitmap with the same size as the picturebox;
            int y = 60;
            int x1 = 0;
            int x2 = 0;
            int plannedProcent = 0;
            int notPlannedProcent = 0;
            int autoProcent = 0;
            int[] number = new int[10];
            for (int i = Globals.j; i < Globals.readText.Length; i++)
            {
                //auto-mode
                if (Globals.readText[i] == 'A' && Globals.readText[i + 1] == 'u')
                {
                    //hours
                    Int32.TryParse(Globals.readText[i + 11].ToString() + Globals.readText[i + 12].ToString(), out number[0]);
                    number[0] = number[0] * 3600;

                    //minutes
                    Int32.TryParse(Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString(), out number[1]);
                    number[1] = number[1] * 60;

                    //seconds
                    Int32.TryParse(Globals.readText[i + 17].ToString() + Globals.readText[i + 18].ToString(), out number[2]);

                    number[6] = number[0] + number[1] + number[2] + x2;

                    pen1.Color = Color.Green;
                    color = 1;
                    x2 = number[6];
                    autoProcent = number[0] + number[1] + number[2] + autoProcent;
                    draw(x1, x2, y);
                    x1 = x2;
                }

                //planned
                if (Globals.readText[i] == 'P' && Globals.readText[i - 1] == '\n')
                {
                    //hours
                    Int32.TryParse(Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString(), out number[0]);
                    number[0] = number[0] * 3600;

                    //minutes
                    Int32.TryParse(Globals.readText[i + 17].ToString() + Globals.readText[i + 18].ToString(), out number[1]);
                    number[1] = number[1] * 60;

                    //seconds
                    Int32.TryParse(Globals.readText[i + 20].ToString() + Globals.readText[i + 21].ToString(), out number[2]);

                    number[6] = number[0] + number[1] + number[2] + x2;

                    pen1.Color = Color.Yellow;
                    color = 2;
                    x2 = number[6];
                    plannedProcent = number[0] + number[1] + number[2] + plannedProcent;
                    draw(x1, x2, y);
                    x1 = x2;
                }

                //not planned - skal opdateres med det nye format "unplanned"
                if (Globals.readText[i] == 'U' && Globals.readText[i + 2] == 'p')
                {
                    //hours
                    Int32.TryParse(Globals.readText[i + 16].ToString() + Globals.readText[i + 17].ToString(), out number[0]);
                    number[0] = number[0] * 3600;
                    //minutes
                    Int32.TryParse(Globals.readText[i + 19].ToString() + Globals.readText[i + 20].ToString(), out number[1]);
                    number[1] = number[1] * 60;

                    //seconds
                    Int32.TryParse(Globals.readText[i + 22].ToString() + Globals.readText[i + 23].ToString(), out number[2]);

                    number[6] = number[0] + number[1] + number[2] + x2;

                    pen1.Color = Color.Red;
                    color = 3;

                    x2 = number[6];
                    notPlannedProcent = number[0] + number[1] + number[2] + notPlannedProcent;
                    draw(x1, x2, y);
                    x1 = x2;
                }

                //start time
                if (Globals.readText[i] == 'S' && Globals.readText[i + 1] == 't' && Globals.readText[i + 2] == 'a')
                {
                    //date
                    tbStartDate.Text = (Globals.readText[i + 12].ToString() + Globals.readText[i + 13].ToString()) + Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString() + Globals.readText[i + 16].ToString() + Globals.readText[i + 17].ToString() + Globals.readText[i + 18].ToString() + Globals.readText[i + 19].ToString() + Globals.readText[i + 20].ToString() + Globals.readText[i + 21].ToString();

                    //start time
                    tbStartTime.Text = (Globals.readText[i + 23].ToString() + Globals.readText[i + 24].ToString() + Globals.readText[i + 25].ToString() + Globals.readText[i + 26].ToString() + Globals.readText[i + 27].ToString() + Globals.readText[i + 28].ToString() + Globals.readText[i + 29].ToString() + Globals.readText[i + 30].ToString());
                }

                //end time
                if (Globals.readText[i] == 'p' && Globals.readText[i + 2] == 't')
                {
                    //date
                    tbEndDate.Text = (Globals.readText[i + 8].ToString() + Globals.readText[i + 9].ToString()) + Globals.readText[i + 10].ToString() + Globals.readText[i + 11].ToString() + Globals.readText[i + 12].ToString() + Globals.readText[i + 13].ToString() + Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString() + Globals.readText[i + 16].ToString() + Globals.readText[i + 17].ToString();

                    //start time
                    tbEndTime.Text = (Globals.readText[i + 19].ToString() + Globals.readText[i + 20].ToString() + Globals.readText[i + 21].ToString() + Globals.readText[i + 22].ToString() + Globals.readText[i + 23].ToString() + Globals.readText[i + 24].ToString() + Globals.readText[i + 25].ToString() + Globals.readText[i + 26].ToString());
                }

                //total process tid
                if (Globals.readText[i] == 'P' && Globals.readText[i + 1] == 'r')
                {
                    x1 = 0;
                    x2 = 0;
                    y = y + 110;
                    //hours
                    Int32.TryParse(Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString(), out number[0]);
                    number[0] = number[0] * 3600;
                    //minutes
                    Int32.TryParse(Globals.readText[i + 17].ToString() + Globals.readText[i + 18].ToString(), out number[1]);
                    number[1] = number[1] * 60;
                    //seconds
                    Int32.TryParse(Globals.readText[i + 20].ToString() + Globals.readText[i + 21].ToString(), out number[2]);                  
                    //total
                    number[6] = number[0] + number[1] + number[2] + x2;
                    if(number[6] == 0 )
                    {
                        tbEndDate.Text = "N/A";
                        tbStartDate.Text = "N/A";
                        tbStartTime.Text = "N/A";
                        tbEndTime.Text = "N/A";
                        tbTime.Text = "";
                        tbArea.Text = "";
                        tbSeconds.Text = "N/A";
                        tbTotalNotPlanned.Text = "N/A";
                        tbTotalPlanned.Text = "N/A";
                        tbAutoP.Text = "N/A";
                        Globals.i = i + 21;
                        pictureBox1.Image = null;
                        pictureBox1.Invalidate();
                        draw_invalidData();
                        return;
                    }
                    totalProcessTime = number[6];
                    pen1.Color = Color.DarkGray;
                    color = 4;
                    x2 = number[6];
                    draw(x1, x2, y);

                    //Write to global integers
                    totalAutoMode = autoProcent;
                    totalPlanned = plannedProcent;
                    totalNotPlanned = notPlannedProcent;

                    //autoprocent writeout
					tbAutoP.Text = (autoProcent / (decimal)(x2/100.00)).ToString("0.##") + "%";//+ autoProcent;
					tbTotalPlanned.Text = (plannedProcent / (decimal)(x2/100.00)).ToString("0.##") + "%";
					tbTotalNotPlanned.Text = (notPlannedProcent / (decimal)(x2/100.00)).ToString("0.##") + "%";
                    //total time writeout
                    tbSeconds.Text = Globals.readText[i + 14].ToString() + Globals.readText[i + 15].ToString() + Globals.readText[i + 16].ToString() + Globals.readText[i + 17].ToString() + Globals.readText[i + 18].ToString() + Globals.readText[i + 19].ToString() + Globals.readText[i + 20].ToString() + Globals.readText[i + 21].ToString();

                    //save current i position
                    Globals.i = i+21;
                    return;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            arrayID = 1;
            btnNext.Enabled = true;
            drawBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height); // create a bitmap with the same size as the canvas;
 
            int[] number = new int[10];
            string readText = "";
            openFileDialog1.FileName = "cycleTime.txt";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string ext = Path.GetExtension(openFileDialog1.FileName);
                if(!ext.Contains(".txt"))
                {
                    MessageBox.Show("Forkert fil format" + "\n" + "Kun .txt filer kan læses");
                    return;
                }

                readText = File.ReadAllText(openFileDialog1.FileName);
                Globals.readText = readText;

                if (!readText.Contains("Auto Mode"))
                {
                    MessageBox.Show("Ser ikke ud til filen der er valgt indeholder korrekt formateret data");
                    return;
                }
                iterateText();
                k = 0;
            }
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            tbEndDate.Text = "";
            tbStartDate.Text = "";
            tbStartTime.Text = "";
            tbEndTime.Text = "";
            tbTime.Text = "";
            tbArea.Text = "";
            tbSeconds.Text = "";
            tbTotalNotPlanned.Text = "";
            tbTotalPlanned.Text = "";
            tbAutoP.Text = "";

            pictureBox1.Refresh();
            pictureBox1.Image = null;
            pictureBox1.Invalidate();
            Globals.readText = null;
            Globals.i = 0;
            Globals.j = 0;
            btnPrev.Enabled = false;
            btnNext.Enabled = false;
            btnChart.Enabled = false;
            Array.Clear(iArray, 0, iArray.Length);
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            k = 0;
            Globals.VALUE = trackBar1.Value;
            if (Globals.VALUE == 0)
            {
                Globals.VALUE = 1;
            }

            if (Globals.readText == null)
            {
                return;
            }
            else
            {
                pictureBox1.Refresh();                
                iterateText();
            }
            this.Text = "Scale: 1 : " + trackBar1.Value + " (pixel:seconds)";
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            //this.Text = String.Format(e.X.ToString(), e.Y.ToString());            
        }

        private void bNext_Click(object sender, EventArgs e)
        {
            //reset total integers
            totalAutoMode = 0;
            totalNotPlanned = 0;
            totalPlanned = 0;
            
            //reset text boxes
            tbEndDate.Text = "N/A";
            tbStartDate.Text = "N/A";
            tbStartTime.Text = "N/A";
            tbEndTime.Text = "N/A";
            tbTime.Text = "";
            tbArea.Text = "";
            tbSeconds.Text = "N/A";
            tbTotalNotPlanned.Text = "N/A";
            tbTotalPlanned.Text = "N/A";
            tbAutoP.Text = "N/A";

            k = 0;
            iArray[arrayID] = Globals.i;
            if (iArray[arrayID] + 30 > Globals.readText.Length)
            {
                btnNext.Enabled = false;
                btnPrev.Enabled = true;
            }
            else
            {
                btnPrev.Enabled = true;
                Globals.j = iArray[arrayID];
                arrayID++; //increment arrayID for next iteration.
                iterateText();
                if (Globals.i + 30 > Globals.readText.Length)
                {
                    btnNext.Enabled = false;
                }
            }
			if ((string)tbSeconds.Text == "N/A" && pictureBox1.Image.Tag != "jhn")
            {
                btnNext.Enabled = false;
                pictureBox1.Image.Tag = null;
            }
        }

        private void bPrev_Click(object sender, EventArgs e)
        {
            //reset total integers
            totalAutoMode = 0;
            totalNotPlanned = 0;
            totalPlanned = 0;

            k = 0;
            arrayID--; //decrement arrayID for next iteration.
            if (arrayID <= 1)
            {
                btnNext.Enabled = true;
                btnPrev.Enabled = false;
                arrayID = 1;
                iArray[arrayID] = 0;
                Globals.j = 0;
                iterateText();
            }
            else
            {
                btnNext.Enabled = true;
                Globals.j = iArray[arrayID-1];
                iterateText();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            var mouseEventArgs = e as MouseEventArgs;

            for (k = 0; k < xyValues.Length; k++)
            {
                if (mouseEventArgs.X < (xyValues[k].Y / Globals.VALUE) && mouseEventArgs.X > (xyValues[k].X / Globals.VALUE) && mouseEventArgs.Y > 20 && mouseEventArgs.Y < 100)
                {
                    tbArea.Text = (((float)(xyValues[k].Y - xyValues[k].X) / (totalProcessTime / 100)).ToString("n2") + "%");

                    int time = (int)(xyValues[k].Y - xyValues[k].X);

                    int seconds = 0;
                    int minutes = 0;
                    int hours = 0;
                    //timer

                    if (time >= 60)
                    {
                        minutes = time / 60;
                        seconds = time % 60;
                    }
                    else
                    {
                        seconds = time;
                    }

                    if (minutes >= 60)
                    {
                        hours = minutes / 60;
                        minutes = minutes % 60;
                    }
                    else
                    {
                        minutes = minutes;
                    }


                    if (seconds < 10 && minutes < 10 && hours < 10)
                    {
                        tbTime.Text = ("0" + hours.ToString() + ":0" + minutes.ToString()) + ":0" + seconds.ToString();
                        return;
                    }
                    else if (hours < 10 && minutes < 10)
                    {
                        tbTime.Text = ("0" + hours.ToString()) + ":0" + minutes.ToString() + ":" + seconds.ToString();
                        return;
                    }
                    else if (hours < 10 && seconds < 10)
                    {
                        tbTime.Text = ("0" + hours.ToString()) + ":" + minutes.ToString() + ":0" + seconds.ToString();
                        return;
                    }
                    else if (hours < 10)
                    {
                        tbTime.Text = ("0" + hours.ToString()) + ":" + minutes.ToString() + ":" + seconds.ToString();
                        return;
                    }
                    else if (minutes < 10)
                    {
                        tbTime.Text = (hours.ToString() + ":0" + minutes.ToString()) + ":" + seconds.ToString();
                        return;
                    }
                    else if (seconds < 10)
                    {
                        tbTime.Text = (hours.ToString() + ":" + minutes.ToString()) + ":0" + seconds.ToString();
                        return;
                    }
                    else
                    {
                        tbTime.Text = (hours.ToString() + ":" + minutes.ToString()) + ":" + seconds.ToString();
                        return;
                    }
                }
                if (mouseEventArgs.X < (xyValues[k].Y / Globals.VALUE) && mouseEventArgs.X > (xyValues[k].X / Globals.VALUE) && mouseEventArgs.Y > 130 && mouseEventArgs.Y < 210)
                {
                    tbArea.Text = ((totalProcessTime / (totalProcessTime / 100)).ToString() + "%");
                    tbTime.Text = tbSeconds.Text;
                }
            }
            k = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //init 
            arrayID = 1;
            btnNext.Enabled = true;
            btnChart.Enabled = true;
            drawBitmap = new Bitmap(pictureBox1.Width, pictureBox1.Height); // create a bitmap with the same size as the canvas;
            int[] number = new int[10];
            string ipaddress = "192.168.16.170"; //Change this if IP address of server/controlmodule changes

            Ping pingSender = new Ping();

            PingReply reply = pingSender.Send(ipaddress);

            WebClient request = new WebClient();

            string url = "ftp://" + ipaddress + "/cycleTime.txt";
            request.Credentials = new NetworkCredential();

            if (reply.Status == IPStatus.Success)
            {
                try
                {
                    byte[] newFileData = request.DownloadData(url);
                    Globals.readText = System.Text.Encoding.UTF8.GetString(newFileData);
                    iterateText();
                    k = 0;
                }
                catch (WebException ex)
                {
                    MessageBox.Show(ex.Message);
                    btnChart.Enabled = false;
                    btnNext.Enabled = false;
                }
            }
            else
            {
                MessageBox.Show(reply.Status.ToString() + "\n" + "Tjek netværksforbindelsen" );

            }
        }

        private void btnChart_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data 
            xlWorkSheet.Cells[1, 2] = "Overview";
            xlWorkSheet.Cells[1, 3] = "Time";

            xlWorkSheet.Cells[2, 1] = "Planned \n" + toTime(totalPlanned).ToString();
            xlWorkSheet.Cells[2, 2] = totalPlanned;

            xlWorkSheet.Cells[3, 1] = "Not Planned \n" + toTime(totalNotPlanned).ToString();
            xlWorkSheet.Cells[3, 2] = totalNotPlanned;

            xlWorkSheet.Cells[4, 1] = "Auto Mode \n" + toTime(totalAutoMode).ToString();
            xlWorkSheet.Cells[4, 2] = totalAutoMode;
            xlWorkSheet.Cells[4, 3] = toTime(totalAutoMode).ToString();

            //xlWorkBook.SaveAs("csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            Excel.Range chartRange;

            Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlWorkSheet.ChartObjects(Type.Missing);
            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 80, 300, 250);
            Excel.Chart chartPage = myChart.Chart;

            chartRange = xlWorkSheet.get_Range("A1", "c4");
            chartPage.SetSourceData(chartRange, misValue);
            chartPage.ChartType = Excel.XlChartType.xlPie;

            //chartPage.SeriesCollection(2).Format.Fill.ForeColor.RGB = System.Drawing.Color.Red.ToArgb();

            //chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
           // if (!File.Exists(@"C:\excel_chart_export.bmp"))
            //{
                chartPage.Export(@"C:\excel_chart_export.bmp", "BMP", misValue);
           
            //}
            //export chart as picture file

            //xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            this.Hide();

            Image img;
            using (var bmpTemp = new Bitmap(@"C:\excel_chart_export.bmp"))
            {
                img = new Bitmap(bmpTemp);
            }

            using (Form form = new Form())
            {
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Size = new Size(510, 460);

                PictureBox pb = new PictureBox();
                pb.Dock = DockStyle.Fill;
                pb.Image = img;
                                
                form.Controls.Add(pb);
                form.ShowDialog();
            }
            this.Show();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            releaseObject(chartPage);
            
            File.Delete(@"C:\excel_chart_export.bmp");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            File.Delete(@"C:\excel_chart_export.bmp");
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                Uri ftpUri = new Uri("ftp://192.168.16.170/cycleTime.txt");

                DialogResult result = MessageBox.Show("Are you sure you want to delete the file? This cannot be undone!", "Confirm delete", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    FtpWebRequest request1 = (FtpWebRequest)WebRequest.Create(ftpUri);
                    request1.Credentials = new NetworkCredential();
                    request1.Method = WebRequestMethods.Ftp.DeleteFile;
                    FtpWebResponse response = (FtpWebResponse)request1.GetResponse();
                    MessageBox.Show(response.StatusDescription.ToString());
                    response.Close();
                }
                else
                {
                    return;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
			if (Globals.readText == null) {
				MessageBox.Show ("No data available. \nLoad some data first");
				return;
			}
					
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

			string totalAM = "";
			string test = Globals.readText;
			int l = 3; //index counter for action writeout to excel
			int m = 1; //index counter for start/stop writeout to excel
			//split all text into sub-strings. Keyword is "\r" which indicates that a substring is made everytime a carriage return occurs. 
			string[] words = Regex.Split(Globals.readText, "\r");
			//iterate through each new sub-string in the array
			foreach (string auto in words) {
				//start time writeout to excel
				if (auto.Contains ("Start time:")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						if (Regex.IsMatch (newline, "^\\d{4}-\\d{2}-\\d{2}$")) {
							xlWorkSheet.Cells [m+1, 1] = "Start time:";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [m+1, 3] = onelinestring;
						}
						if (Regex.IsMatch (newline, "^\\d{2}:\\d{2}:\\d{2}$")) {
							xlWorkSheet.Cells [m+1, 1] = "Start time:";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [m+1, 2] = onelinestring;
						}
					} 
					//m = l;
				}
				//stop time writeout to excel
				if (auto.Contains("Stop time:")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						if (Regex.IsMatch (newline, "^\\d{4}-\\d{2}-\\d{2}")) {
							xlWorkSheet.Cells [m + 2, 1] = "Stop time:";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [m + 2, 3] = onelinestring;
						}
						if (Regex.IsMatch (newline, "^\\d{2}:\\d{2}:\\d{2}$")) {
							xlWorkSheet.Cells [m+2, 1] = "Stop time:";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [m+2, 2] = onelinestring;
						}
					}

				}
				//sektionstype writeout to excel
				if (auto.Contains ("Sektionstype")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						//xlWorkSheet.Cells [m, 1] = "Sektion:";
						string onelinestring = newline.Replace ("\n", "");
						xlWorkSheet.Cells [m, 3] = onelinestring;

					}
				} 

				//Sektionsnummer writeout to excel
				if (auto.Contains ("Sektionsnummer")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						xlWorkSheet.Cells [m, 1] = "Sektion:";
						string onelinestring = newline.Replace ("\n", "");
						xlWorkSheet.Cells [m, 2] = onelinestring;
					}
					m = l;
					l = l + 2;
				} 

				if (auto.Contains ("Auto Mode:")) {
					string[] time = Regex.Split (auto, " ");
					foreach (string newline in time) {
						if (Regex.IsMatch (newline, "^\\d{2}:\\d{2}:\\d{2}$")) {
							xlWorkSheet.Cells [l, 6] = "Auto Mode";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [l, 5] = onelinestring;
						}
					}
					l++;
				}
				//planned stop writeout to excel
				  if (auto.Contains ("Planned Stop:")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						if (Regex.IsMatch (newline, "^\\d{2}:\\d{2}:\\d{2}$")) {
							xlWorkSheet.Cells [l, 6] = "Planned Stop";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [l, 5] = onelinestring;
						}
					}
					l++;
				}
				//unplanned stop writeout to excel
				if (auto.Contains ("Unplanned Stop:")) {
					string[] timer = Regex.Split (auto, " ");
					foreach (string newline in timer) {
						if (Regex.IsMatch (newline, "^\\d{2}:\\d{2}:\\d{2}$")) {
							xlWorkSheet.Cells [l, 6] = "Unplanned Stop";
							string onelinestring = newline.Replace ("\n", "");
							xlWorkSheet.Cells [l, 5] = onelinestring;
						}
					}
					l++;
				}
			}

				xlApp.Visible = true;
				//add data 
				xlWorkSheet.Cells [1, 5] = "Time";
				xlWorkSheet.Cells [1, 6] = "Action";

            //xlWorkBook.Close(false, misValue, misValue);
            //xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        } // end function
    } // 
} // end class

public class something
{
    public static string fileString;

    public void openftp()
    {
            int[] number = new int[10];

            WebClient request = new WebClient();
            string url = "ftp://192.168.16.170/" + "cycleTime.txt";
            request.Credentials = new NetworkCredential();

            try
            {
                byte[] newFileData = request.DownloadData(url);
                fileString = System.Text.Encoding.UTF8.GetString(newFileData);
                MessageBox.Show(fileString);
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.GetType().FullName);
                MessageBox.Show(ex.GetBaseException().ToString());
                // Do something such as log error, but this is based on OP's original code
                // so for now we do nothing.
            }
            return;
    } // end function
};
