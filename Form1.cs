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
        private PaintedDot dot;
        private Bitmap bmp;
        private System.Drawing.Point _point;
        private Microsoft.Office.Interop.Excel.Application excelApp;
        public Form1()
        {
            InitializeComponent();
            dot = new PaintedDot();
            bmp = new Bitmap(this.ClientSize.Width, this.ClientSize.Height);
            InitDot();
            this.DoubleBuffered = true;
            pictureBox1.MouseEnter += PictureBox1_MouseEnter;
            pictureBox1.MouseMove += PictureBox1_MouseMove;
            pictureBox1.MouseLeave += PictureBox1_MouseLeave;
            pictureBox1.Paint += PictureBox1_Paint;
            excelApp = Excel.ExcelApp();
        }
        private void InitDot()
        {
            dot.MyPen = Pens.Transparent;
            dot.MyBrush = Brushes.Transparent;
            dot.Path = new GraphicsPath();
            dot.Path.AddEllipse(System.Drawing.Rectangle.FromLTRB(0, 0, 7, 7));
            _point.X = 0;
            _point.Y = 0;
            RefreshBitmap();
        }
        private void RefreshBitmap()
        {
            if (bmp != null) bmp.Dispose();
            bmp = new Bitmap(this.ClientSize.Width, this.ClientSize.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.DrawPath(dot.MyPen, dot.Path);
                g.FillPath(dot.MyBrush, dot.Path);
            }
        }
        private void PictureBox1_MouseEnter(object sender, EventArgs e)
        {
            dot.MyPen = Pens.Red;
            dot.MyBrush = Brushes.Red;
        }
        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            int deltaX, deltaY;
            deltaX = e.Location.X - _point.X;
            deltaY = e.Location.Y - _point.Y;
            dot.Path.Transform(new Matrix(1, 0, 0, 1, deltaX, deltaY));
            _point = e.Location;
            this.Refresh();
        }
        private void PictureBox1_MouseLeave(object sender, EventArgs e)
        {
            dot.Path.Reset();
            this.Refresh();
            InitDot();
        }
        private void PictureBox1_Paint(object sender, PaintEventArgs e)
        {
            if (bmp == null) return;
            RefreshBitmap();
            e.Graphics.DrawImage(bmp, 0, 0);
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Excel.Update(TextBox1.Text, ListBox1.Text);
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
        //public PaintedDot(GraphicsPath path, Pen pen, Brush brush)
        //{
        //    this.path = path;
        //    this.pen = pen;
        //    this.brush = brush;
        //}
    }
}