namespace ExcelVTEntegrasyonProjesi
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btn_VTdenOku = new Button();
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            btn_ExceldenOku = new Button();
            SuspendLayout();
            // 
            // btn_VTdenOku
            // 
            btn_VTdenOku.BackColor = SystemColors.ControlDarkDark;
            btn_VTdenOku.ForeColor = SystemColors.Control;
            btn_VTdenOku.Location = new Point(507, 84);
            btn_VTdenOku.Name = "btn_VTdenOku";
            btn_VTdenOku.Size = new Size(241, 41);
            btn_VTdenOku.TabIndex = 0;
            btn_VTdenOku.Text = "VERİ TABANINDAN OKU EXCELE YAZ";
            btn_VTdenOku.UseVisualStyleBackColor = false;
            btn_VTdenOku.Click += btn_VTdenOku_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.BackColor = SystemColors.ButtonFace;
            richTextBox1.Location = new Point(39, 32);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(407, 155);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.BackColor = SystemColors.ButtonFace;
            richTextBox2.Location = new Point(39, 242);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(407, 155);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // btn_ExceldenOku
            // 
            btn_ExceldenOku.BackColor = Color.SpringGreen;
            btn_ExceldenOku.Location = new Point(507, 281);
            btn_ExceldenOku.Name = "btn_ExceldenOku";
            btn_ExceldenOku.Size = new Size(241, 41);
            btn_ExceldenOku.TabIndex = 3;
            btn_ExceldenOku.Text = "EXCELDEN OKU VERİTABANINA YAZ";
            btn_ExceldenOku.UseVisualStyleBackColor = false;
            btn_ExceldenOku.Click += btn_ExceldenOku_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.MenuHighlight;
            ClientSize = new Size(800, 450);
            Controls.Add(btn_ExceldenOku);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Controls.Add(btn_VTdenOku);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion

        private Button btn_VTdenOku;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button btn_ExceldenOku;
    }
}