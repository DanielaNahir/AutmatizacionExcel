namespace procesadorDeExel
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
            btnCargarExcelNuevo = new Button();
            txtExcelNuevo = new TextBox();
            btnProcesar = new Button();
            btnComparar = new Button();
            textBox1 = new TextBox();
            button2 = new Button();
            textBox2 = new TextBox();
            button3 = new Button();
            SuspendLayout();
            // 
            // btnCargarExcelNuevo
            // 
            btnCargarExcelNuevo.Location = new Point(22, 28);
            btnCargarExcelNuevo.Name = "btnCargarExcelNuevo";
            btnCargarExcelNuevo.Size = new Size(109, 23);
            btnCargarExcelNuevo.TabIndex = 0;
            btnCargarExcelNuevo.Text = "Cargar Excel";
            btnCargarExcelNuevo.UseVisualStyleBackColor = true;
            btnCargarExcelNuevo.Click += btnCargarExcelNuevo_Click;
            // 
            // txtExcelNuevo
            // 
            txtExcelNuevo.Location = new Point(155, 29);
            txtExcelNuevo.Name = "txtExcelNuevo";
            txtExcelNuevo.Size = new Size(208, 23);
            txtExcelNuevo.TabIndex = 1;
            // 
            // btnProcesar
            // 
            btnProcesar.Location = new Point(288, 58);
            btnProcesar.Name = "btnProcesar";
            btnProcesar.Size = new Size(75, 23);
            btnProcesar.TabIndex = 2;
            btnProcesar.Text = "Procesar";
            btnProcesar.UseVisualStyleBackColor = true;
            btnProcesar.Click += btnProcesar_Click;
            // 
            // btnComparar
            // 
            btnComparar.Location = new Point(288, 180);
            btnComparar.Name = "btnComparar";
            btnComparar.Size = new Size(75, 23);
            btnComparar.TabIndex = 5;
            btnComparar.Text = "Comparar";
            btnComparar.UseVisualStyleBackColor = true;
            btnComparar.Click += button1_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(155, 109);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(208, 23);
            textBox1.TabIndex = 4;
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // button2
            // 
            button2.Location = new Point(22, 108);
            button2.Name = "button2";
            button2.Size = new Size(109, 23);
            button2.TabIndex = 3;
            button2.Text = "Excel viejo";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(155, 138);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(208, 23);
            textBox2.TabIndex = 7;
            // 
            // button3
            // 
            button3.Location = new Point(22, 137);
            button3.Name = "button3";
            button3.Size = new Size(109, 23);
            button3.TabIndex = 6;
            button3.Text = "Excel nuevo";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(392, 215);
            Controls.Add(textBox2);
            Controls.Add(button3);
            Controls.Add(btnComparar);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(btnProcesar);
            Controls.Add(txtExcelNuevo);
            Controls.Add(btnCargarExcelNuevo);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnCargarExcelNuevo;
        private TextBox txtExcelNuevo;
        private Button btnProcesar;
        private Button btnComparar;
        private TextBox textBox1;
        private Button button2;
        private TextBox textBox2;
        private Button button3;
    }
}