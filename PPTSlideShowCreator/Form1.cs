using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Packaging;

namespace PPTSlideShowCreator
{
    public partial class Form1 : Form
    {
        SlideShowCreator slideShowCreator = new SlideShowCreator();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string newPresentation = "";
            List<string> imageFiles = new List<string>();
            SaveFileDialog queryPPTname = new SaveFileDialog();
            queryPPTname.DefaultExt = "pptx";
            if (queryPPTname.ShowDialog() == DialogResult.OK)
            {
                newPresentation = queryPPTname.FileName;
            }
            else
            {
                throw new NotImplementedException();
            }
             OpenFileDialog queryImages = new OpenFileDialog();
            queryImages.Title = "Escolha as imagens a serem adicionadas.";
            queryImages.Multiselect = true;
            if (queryImages.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < queryImages.FileNames.Length; i++)
                {
                    imageFiles.Add(queryImages.FileNames[i]);
                }
            }
            else
            {
                throw new NotImplementedException();
            }
            

            slideShowCreator.CreateSlideShow(newPresentation, imageFiles);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           Image temp = Utils.OpenImage(listBox1.Items[listBox1.SelectedIndex].ToString());
            Size size = new Size(pictureBox1.Width, (temp.Height * temp.Width) / pictureBox1.Width);
            pictureBox1.Image = Utils.ResizeImage(temp, size);
        }

       

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog queryImages = new OpenFileDialog();
            queryImages.Title = "Escolha as imagens a serem adicionadas.";
            queryImages.Multiselect = true;
            if (queryImages.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < queryImages.FileNames.Length; i++)
                {
                    listBox1.Items.Add(queryImages.FileNames[i]);
                }
            }
        }
        float barProgress;
        float barStep;

        void incBar()
        {
            barProgress += barStep;           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            barProgress = 0;
            barStep = ((float) 100) /((float) listBox1.Items.Count);
            string newPresentation = "";
            List<string> imageFiles = new List<string>();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                imageFiles.Add(listBox1.Items[i].ToString());
            }
            SaveFileDialog queryPPTname = new SaveFileDialog();
            queryPPTname.DefaultExt = "pptx";
            if (queryPPTname.ShowDialog() == DialogResult.OK)
            {
                newPresentation = queryPPTname.FileName;
            }
            else
            {
                throw new NotImplementedException();
            }
            slideShowCreator.CreateSlideShowThread(imageFiles, newPresentation, incBar);
 
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = Convert.ToInt32(barProgress);

            if (progressBar1.Value == 99)
                progressBar1.Value = 100;
        }


    }
}
