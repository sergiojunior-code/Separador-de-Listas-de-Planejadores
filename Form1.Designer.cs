namespace Separador_de_Listas_de_Planejadores
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnCarregarXLS = new Button();
            lblAgExcel = new Label();
            btnCarregarPDF = new Button();
            lblAgPDF = new Label();
            richtxtPainel = new RichTextBox();
            btnIniciarSep = new Button();
            lblStatus = new Label();
            SuspendLayout();
            // 
            // btnCarregarXLS
            // 
            btnCarregarXLS.Location = new Point(12, 12);
            btnCarregarXLS.Name = "btnCarregarXLS";
            btnCarregarXLS.Size = new Size(101, 31);
            btnCarregarXLS.TabIndex = 0;
            btnCarregarXLS.Text = "Carregar Excel";
            btnCarregarXLS.UseVisualStyleBackColor = true;
            btnCarregarXLS.Click += btnCarregarXLS_Click;
            // 
            // lblAgExcel
            // 
            lblAgExcel.AutoSize = true;
            lblAgExcel.Location = new Point(126, 20);
            lblAgExcel.Name = "lblAgExcel";
            lblAgExcel.Size = new Size(154, 15);
            lblAgExcel.TabIndex = 1;
            lblAgExcel.Text = "Aguardando arquivo excel...";
            lblAgExcel.Click += lblAgExcel_Click;
            // 
            // btnCarregarPDF
            // 
            btnCarregarPDF.Location = new Point(12, 49);
            btnCarregarPDF.Name = "btnCarregarPDF";
            btnCarregarPDF.Size = new Size(101, 31);
            btnCarregarPDF.TabIndex = 2;
            btnCarregarPDF.Text = "Carregar PDF";
            btnCarregarPDF.UseVisualStyleBackColor = true;
            btnCarregarPDF.Click += btnCarregarPDF_Click;
            // 
            // lblAgPDF
            // 
            lblAgPDF.AutoSize = true;
            lblAgPDF.Location = new Point(126, 57);
            lblAgPDF.Name = "lblAgPDF";
            lblAgPDF.Size = new Size(149, 15);
            lblAgPDF.TabIndex = 3;
            lblAgPDF.Text = "Aguardando arquivo PDF...";
            lblAgPDF.Click += lblAgPDF_Click;
            // 
            // richtxtPainel
            // 
            richtxtPainel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            richtxtPainel.BackColor = SystemColors.Info;
            richtxtPainel.Location = new Point(12, 144);
            richtxtPainel.Name = "richtxtPainel";
            richtxtPainel.ReadOnly = true;
            richtxtPainel.Size = new Size(511, 189);
            richtxtPainel.TabIndex = 4;
            richtxtPainel.Text = "";
            richtxtPainel.TextChanged += richtxtPainel_TextChanged;
            // 
            // btnIniciarSep
            // 
            btnIniciarSep.Location = new Point(12, 86);
            btnIniciarSep.Name = "btnIniciarSep";
            btnIniciarSep.Size = new Size(101, 52);
            btnIniciarSep.TabIndex = 5;
            btnIniciarSep.Text = "Iniciar separação";
            btnIniciarSep.UseVisualStyleBackColor = true;
            btnIniciarSep.Click += btnIniciarSep_Click;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(126, 105);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(82, 15);
            lblStatus.TabIndex = 6;
            lblStatus.Text = "Status: Ocioso";
            lblStatus.Click += lblStatus_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(542, 345);
            Controls.Add(lblStatus);
            Controls.Add(btnIniciarSep);
            Controls.Add(richtxtPainel);
            Controls.Add(lblAgPDF);
            Controls.Add(btnCarregarPDF);
            Controls.Add(lblAgExcel);
            Controls.Add(btnCarregarXLS);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Separador de Listas de Planejadores";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnCarregarXLS;
        private Label lblAgExcel;
        private Button btnCarregarPDF;
        private Label lblAgPDF;
        private RichTextBox richtxtPainel;
        private Button btnIniciarSep;
        private Label lblStatus;
    }
}
