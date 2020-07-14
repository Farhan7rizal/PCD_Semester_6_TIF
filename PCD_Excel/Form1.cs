using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//Tambah Emgu Library
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.UI;
using Emgu.CV.Util;

using OfficeOpenXml;


namespace ProjectPCD1
{
    public partial class Form1 : Form
    {
        //inisialisai tipe gambar
        Image<Bgr, byte> imgInput;
        Image<Gray, byte> imgGray;
        Image<Gray, byte> imgBinarize;
        Image<Gray, byte> imgOut;
        
        string red;
        string green;
        string blue;
        string fitur_area;
        string fitur_perimeter;
        string fitur_titik_tengah;
        string nomor;     
        
        int nomorPlus;
        int nomorData;

        double KomaPerimeter;
        double Blue_Koma;
        double Green_Koma;
        double Red_Koma;
        public Form1()
        {
            
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label10.Text = "";
            label11.Text = "";           
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Fungsi Open image/browse image
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                imgInput = new Image<Bgr, byte>(ofd.FileName);
                pictureBox1.Image = imgInput.Bitmap;                
            }
        }

        private void grayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label10.Text = "";
            label11.Text = "";
            if (imgInput == null)
            {
                MessageBox.Show("Please select an image");
                return;
            }
            //Konversi RGB to Grayscale
            Image<Gray, byte> imgOutput = new Image<Gray, byte>(imgInput.Width, imgInput.Height, new Gray());
            //imgOutput = imgInput.Convert<Gray, byte>();

            CvInvoke.CvtColor(imgInput, imgOutput, Emgu.CV.CvEnum.ColorConversion.Bgr2Gray);

            histogramBox1.ClearHistogram();
            histogramBox1.Refresh();

            DenseHistogram hist = new DenseHistogram(256, new RangeF(0, 255));
            hist.Calculate(new Image<Gray, byte>[] { imgOutput }, false, null);

            Mat m = new Mat();
            hist.CopyTo(m);

            histogramBox1.AddHistogram("Grayscale Histogram", Color.Blue, m, 256, new float[] { 0, 256 });
            histogramBox1.Refresh();

            
            pictureBox2.Update();
            pictureBox2.Image = imgOutput.Bitmap;
        }

        private void binerizationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (imgInput == null)
            {
                MessageBox.Show("Please select an image");
                return;
            }

            histogramBox1.ClearHistogram();
            histogramBox1.Refresh();
            imgGray = imgInput.Convert<Gray, byte>();
            //pictureBox2.Image = imgGray.Bitmap;

            // Binarization Thresholding
            imgBinarize = new Image<Gray, byte>(imgGray.Width, imgGray.Height, new Gray(0));
            double threshold = CvInvoke.Threshold(imgGray, imgBinarize, 80, 255, Emgu.CV.CvEnum.ThresholdType.Binary);
            
            DenseHistogram hist = new DenseHistogram(256, new RangeF(0, 255));
            hist.Calculate(new Image<Gray, byte>[] { imgBinarize }, false, null);

            Mat m = new Mat();
            hist.CopyTo(m);

            histogramBox1.AddHistogram("Binerisasi Histogram", Color.Blue, m, 256, new float[] { 0, 256 });
            histogramBox1.Refresh();

            pictureBox2.Image = imgBinarize.Bitmap;
            label10.Text = "Nilai Threshold : ";
            label11.Text = threshold.ToString();
            // MessageBox.Show(threshold.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (imgInput == null)
            {
                MessageBox.Show("Please select an image");
                return;
            }

            Double Blue_avg = 0.0;
            for (int j = 0; j < imgInput.Cols; j++)
            {
                for (int i = 0; i < imgInput.Rows; i++)
                {
                    Blue_avg += imgInput.Data[i, j, 0];
                }
            }
            Blue_avg = Blue_avg / (imgInput.Cols * imgInput.Rows);

            Double Green_avg = 0.0;
            for (int j = 0; j < imgInput.Cols; j++)
            {
                for (int i = 0; i < imgInput.Rows; i++)
                {
                    Green_avg += imgInput.Data[i, j, 1];
                }
            }
            Green_avg = Green_avg / (imgInput.Cols * imgInput.Rows);

            Double Red_avg = 0.0;
            for (int j = 0; j < imgInput.Cols; j++)
            {
                for (int i = 0; i < imgInput.Rows; i++)
                {
                    Red_avg += imgInput.Data[i, j, 2];
                }
            }
            Red_avg = Red_avg / (imgInput.Cols * imgInput.Rows);

            Blue_Koma = Math.Round(Blue_avg, 3);
            textBox1.Text = Blue_Koma.ToString();
            blue = Blue_Koma.ToString();

            Green_Koma = Math.Round(Green_avg, 3);
            textBox2.Text = Green_Koma.ToString();
            green = Green_Koma.ToString();

            Red_Koma = Math.Round(Red_avg, 3);
            textBox3.Text = Red_Koma.ToString();
            red = Red_Koma.ToString();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure want to close ?", "System Message", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Close();
            }
        }

        private void deteksiTepicannyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label10.Text = "";
            label11.Text = "";
            if (imgInput == null)
            {
                MessageBox.Show("Please select an image");
                return;
            }

            histogramBox1.ClearHistogram();
            histogramBox1.Refresh();

            //Canny method return a byte Variable
            Image<Gray, byte> imgCanny = new Image<Gray, byte>(imgInput.Width, imgInput.Height, new Gray(0));
            //new Gray(0) is zero values
            imgCanny = imgInput.Canny(140, 50);
            pictureBox2.Image = imgCanny.Bitmap;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            try
            {
                if (imgInput == null)
                {
                    MessageBox.Show("Please select an image");
                    return;
                }

                var img = imgInput;
               
                imgOut = new Image<Gray, byte>(imgInput.Width, imgInput.Height, new Gray(0));
               
                var gray = imgInput.Convert<Gray, byte>().ThresholdBinaryInv(new Gray(140), new Gray(255));
                var gray2 = imgInput.Convert<Gray, byte>().ThresholdBinaryInv(new Gray(140), new Gray(255));
                //var gray = imgCanny.Convert<Gray, byte>().ThresholdBinaryInv(new Gray(125), new Gray(255));

                //Contours
                VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
                Mat h = new Mat();

                CvInvoke.FindContours(gray, contours, h, Emgu.CV.CvEnum.RetrType.External,
                    Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);

                VectorOfPoint approx = new VectorOfPoint();

                Dictionary<int, double> shapes = new Dictionary<int, double>();
                Dictionary<int, double> shapes2 = new Dictionary<int, double>();
                Dictionary<int, double> shapes3 = new Dictionary<int, double>();
                Dictionary<int, double> shapes4 = new Dictionary<int, double>();

                for (int i = 0; i < contours.Size; i++)
                {
                    approx.Clear();
                    double perimeter = CvInvoke.ArcLength(contours[i], true);
                    KomaPerimeter = Math.Round(perimeter, 2); // membatasi angka

                    CvInvoke.ApproxPolyDP(contours[i], approx, 0.04 * perimeter, true);

                    double area = CvInvoke.ContourArea(contours[i]);                
                    
                    if (approx.Size > 0 && area < 220000 && area > 60000 && perimeter < 2100)
                    {                        
                        shapes.Add(i, area);
                        shapes2.Add(i, KomaPerimeter);
                        shapes3.Add(i, KomaPerimeter);                           
                    }
                }

                if (shapes.Count > 0 )
                {
                    var sortedShapes = (from item in shapes
                                        orderby item.Value ascending
                                        select item).ToList();

                    for (int i = 0; i < sortedShapes.Count; i++)
                    {
                        CvInvoke.DrawContours(img, contours, sortedShapes[i].Key, new MCvScalar(255, 0, 0), 2);

                        var moments = CvInvoke.Moments(contours[sortedShapes[i].Key]);

                        int x = (int)(moments.M10 / moments.M00);
                        int y = (int)(moments.M01 / moments.M00);                        
                        
                        textBox6.Text = sortedShapes[i].Value.ToString();
                        fitur_area = sortedShapes[i].Value.ToString();

                    }

                    var sortedShapes2 = (from item in shapes2
                                         orderby item.Value ascending
                                         select item).ToList();

                    for (int i = 0; i < sortedShapes2.Count; i++)
                    {
                        CvInvoke.DrawContours(img, contours, sortedShapes2[i].Key, new MCvScalar(255, 255, 0), 1);
                        var moments = CvInvoke.Moments(contours[sortedShapes2[i].Key]);
                        int x = (int)(moments.M10 / moments.M00);
                        int y = (int)(moments.M01 / moments.M00);

                        textBox5.Text = sortedShapes2[i].Value.ToString();
                        fitur_perimeter = sortedShapes2[i].Value.ToString();
                    }
                    
                    var sortedShapes3 = (from item in shapes3
                                         orderby item.Value descending
                                         select item).ToList();

                    for (int i = 0; i < sortedShapes3.Count; i++)
                    {                       
                        var moments = CvInvoke.Moments(contours[sortedShapes3[i].Key]);
                        int x = (int)(moments.M10 / moments.M00);
                        int y = (int)(moments.M01 / moments.M00);
                        
                        img.Draw(new CircleF(new Point(x, y), 3), new Bgr(Color.Yellow), 3);
                        //Draw titik tengah
                        textBox4.Text = x.ToString() + ", " + y.ToString();
                        fitur_titik_tengah = x.ToString() + ", " + y.ToString();                        

                    }                    
                }        
                
                pictureBox2.Image = img.Bitmap;                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            if (textBox7.Text == null)
            {
                MessageBox.Show("Please fill textbox");
                return;
            }

            nomor = textBox7.Text;
            nomorPlus = int.Parse(textBox7.Text) + 2;            
            nomorData = nomorPlus;
            nomorData.ToString();

            FileInfo file = new FileInfo(@"D:\DaunPCD\test.xlsx");
                        
            if (file.Exists)
            {
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    ExcelWorksheet ws = excelPackage.Workbook.Worksheets.First();
                   
                    var headerRow = new List<string[]>()
                    {
                        new string[] { nomor, "Daun Ke-"+nomor,red, green, blue, fitur_area,fitur_perimeter,fitur_titik_tengah }
                    };

                    // Determine the header range (e.g. A1:H1)                    
                    string headerRange = "A" + nomorData + ":" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + nomorData;

                    // Popular header row data                    
                    ws.Cells[headerRange].LoadFromArrays(headerRow);
                    
                    bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                    if (isExcelInstalled)
                    {
                        System.Diagnostics.Process.Start(file.ToString());
                    }
                                        
                    excelPackage.Save();
                }
            }
            
        }

                private void pictureBox3_Click(object sender, EventArgs e)
        {

        }

        private void imageBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
}
