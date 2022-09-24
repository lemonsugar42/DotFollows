
namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.Button1 = new System.Windows.Forms.Button();
            this.TextBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.Window;
            this.pictureBox1.Location = new System.Drawing.Point(9, 9);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(779, 353);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Paint += new System.Windows.Forms.PaintEventHandler(this.PictureBox1_Paint);
            this.pictureBox1.MouseEnter += new System.EventHandler(this.PictureBox1_MouseEnter);
            this.pictureBox1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.PictureBox1_MouseMove);
            // 
            // Button1
            // 
            this.Button1.Location = new System.Drawing.Point(329, 389);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(87, 37);
            this.Button1.TabIndex = 6;
            this.Button1.Text = "Submit";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // TextBox1
            // 
            this.TextBox1.Location = new System.Drawing.Point(46, 396);
            this.TextBox1.Name = "TextBox1";
            this.TextBox1.Size = new System.Drawing.Size(236, 22);
            this.TextBox1.TabIndex = 7;
            this.TextBox1.Text = "Enter your name";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.TextBox1);
            this.Controls.Add(this.Button1);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button Button1;
        private System.Windows.Forms.TextBox TextBox1;
    }
}

