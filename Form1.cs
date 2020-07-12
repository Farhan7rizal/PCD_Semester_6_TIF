using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//Tambah Emgu Library
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.UI;
using Emgu.CV.Util;

namespace ProjectPCD1
{
    public partial class Form1 : Form
    {
        //inisialisai tipe gambar
        Image<Bgr, byte> imgInput;
        Image<Gray, byte> imgGray;
        Image<Gray, byte> imgBinarize;
        Image<Gray, byte> imgOut;
        Image<Gray, byte> imgOut2;
        

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
                //pictureBox2.Image = imgInput.Bitmap;
                //pictureBox3.Image = imgInput.Bitmap;
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
            double threshold = CvInvoke.Threshold(imgGray, imgBinarize, 125, 255, Emgu.CV.CvEnum.ThresholdType.Binary);
            //CvInvoke.AdaptiveThreshold(imgGray, imgBinarize, 255, Emgu.CV.CvEnum.AdaptiveThresholdType.GaussianC,
            //   Emgu.CV.CvEnum.ThresholdType.Binary, 5, 0.0);

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
            //imageBox1.Image = _InputImage;
            //imageBox1.FunctionalMode = Emgu.CV.UI.ImageBox.FunctionalModeOption.Everything;
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

            Green_Koma = Math.Round(Green_avg, 3);
            textBox2.Text = Green_Koma.ToString();

            Red_Koma = Math.Round(Red_avg, 3);
            textBox3.Text = Red_Koma.ToString();
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
            imgCanny = imgInput.Canny(125, 50);
            pictureBox2.Image = imgCanny.Bitmap;
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
                    
                    //Rectangle rect = CvInvoke.BoundingRectangle(contours[i]);
                    
                    if (approx.Size > 2 && area > 60000)
                    {
                        shapes.Add(i, area);
                        shapes2.Add(i, KomaPerimeter);
                        shapes3.Add(i, KomaPerimeter);
                        //shapes4.Add(i, rect);   
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

                        
                        
                        //CvInvoke.PutText(img, "Area: " + sortedShapes[i].Value.ToString(), new Point(x, y + 60), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                         //   new MCvScalar(255, 0, 0));
                        textBox5.Text = sortedShapes[i].Value.ToString();


                    }

                    var sortedShapes2 = (from item in shapes2
                                         orderby item.Value ascending
                                         select item).ToList();

                    for (int i = 0; i < sortedShapes2.Count; i++)
                    {
                        CvInvoke.DrawContours(img, contours, sortedShapes2[i].Key, new MCvScalar(255, 0, 0), 1);
                        var moments = CvInvoke.Moments(contours[sortedShapes2[i].Key]);
                        int x = (int)(moments.M10 / moments.M00);
                        int y = (int)(moments.M01 / moments.M00);

                        //CvInvoke.PutText(img, (i + 1).ToString(), new Point(x, y), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                        //    new MCvScalar(0, 255, 255));


                        //CvInvoke.PutText(img, "Perimeter: " + sortedShapes2[i].Value.ToString(), new Point(x, y + 30), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                        //    new MCvScalar(255, 0, 0));
                        textBox6.Text = sortedShapes2[i].Value.ToString();

                    }
                    
                    var sortedShapes3 = (from item in shapes3
                                         orderby item.Value ascending
                                         select item).ToList();

                    for (int i = 0; i < sortedShapes3.Count; i++)
                    {
                       // CvInvoke.DrawContours(img, contours, sortedShapes3[i].Key, new MCvScalar(0, 255, 255), 2);
                        var moments = CvInvoke.Moments(contours[sortedShapes3[i].Key]);
                        int x = (int)(moments.M10 / moments.M00);
                        int y = (int)(moments.M01 / moments.M00);

                        //CvInvoke.PutText(img, (i + 1).ToString(), new Point(x, y), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                        //    new MCvScalar(0, 255, 255));
                        //CvInvoke.PutText(img, "X: " + x.ToString(), new Point(x, y - 60), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                        //    new MCvScalar(255, 0, 0));
                        //CvInvoke.PutText(img, "Y: " + y.ToString(), new Point(x, y - 30), Emgu.CV.CvEnum.FontFace.HersheyTriplex, 1.0,
                         //   new MCvScalar(255, 0, 0));
                        img.Draw(new CircleF(new Point(x, y), 1), new Bgr(Color.Yellow), 3);
                        //Draw titik tengah
                        textBox4.Text = x.ToString() + ", " + y.ToString();
                        //CvInvoke.Rectangle(img, rect, new MCvScalar(255, 0, 0), 1);

                    }
                    
                }
                
                pictureBox2.Image = img.Bitmap;
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
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
    }
}
