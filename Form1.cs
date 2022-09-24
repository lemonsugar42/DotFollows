using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ExcelApp;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        PaintedDot dot;
        Bitmap bmp;
        System.Drawing.Point _point;
        Microsoft.Office.Interop.Excel.Application excelApp;
        public Form1()
        {
            InitializeComponent();
            dot = new PaintedDot(new GraphicsPath(), Pens.Red, Brushes.Red);
            dot.Path.AddEllipse(System.Drawing.Rectangle.FromLTRB(20, 20, 27, 27));
            bmp = new Bitmap(this.ClientSize.Width, this.ClientSize.Height);
            _point.X = 20;
            _point.Y = 20;
            RefreshBitmap();
            this.DoubleBuffered = true;
            //this.MouseLeave += PictureBox1_MouseLeave;
            //this.MouseCaptureChanged += PictureBox1_MouseLeave;
            this.MouseMove += PictureBox1_MouseMove;
            this.Paint += PictureBox1_Paint;
            excelApp = ExcelApp.ExcelApp.NewApp();
        }

        void RefreshBitmap()
        {
            if (bmp != null) bmp.Dispose();
            bmp = new Bitmap(this.ClientSize.Width, this.ClientSize.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.DrawPath(dot.MyPen, dot.Path);
                g.FillPath(dot.MyBrush, dot.Path);
            }
        }
        void PictureBox1_MouseEnter(object sender, EventArgs e)
        {
            int deltaX, deltaY;
            deltaX = MousePosition.X - _point.X;
            deltaY = MousePosition.Y - _point.Y;
            dot.Path.Transform(new Matrix(1, 0, 0, 1, deltaX, deltaY));
            RefreshBitmap();
            _point.X = MousePosition.X;
            _point.Y = MousePosition.Y;
        }

        //void PictureBox1_MouseLeave(object sender, EventArgs e)
        //{
        //    using (Graphics g = Graphics.FromImage(bmp)) g.Clear(Color.Green);
        //}

        private void PictureBox1_Paint(object sender, PaintEventArgs e)
        {
            if (bmp == null) return;
            RefreshBitmap();
            e.Graphics.DrawImage(bmp, 0, 0);
        }

        void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            int deltaX, deltaY;
            deltaX = e.Location.X - _point.X;
            deltaY = e.Location.Y - _point.Y;
            dot.Path.Transform(new Matrix(1, 0, 0, 1, deltaX, deltaY));
            _point = e.Location;
            this.Refresh();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            ExcelApp.ExcelApp.Update(excelApp, TextBox1.Text);
            Button1.Text = "Done";
        }
    }

    class PaintedDot
    {
        private GraphicsPath path;
        public GraphicsPath Path
        {
            get { return path; }
            set { path = value; }
        }
        private Pen pen;
        public Pen MyPen
        {
            get { return pen; }
            set { pen = value; }
        }
        private Brush brush;
        public Brush MyBrush
        {
            get { return brush; }
            set { brush = value; }
        }
        public PaintedDot(GraphicsPath path, Pen pen, Brush brush)
        {
            this.path = path;
            this.pen = pen;
            this.brush = brush;
        }
    }
}