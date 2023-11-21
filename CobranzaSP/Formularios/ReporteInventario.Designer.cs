namespace CobranzaSP.Formularios
{
    partial class ReporteInventario
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReporteInventario));
            this.pSuperior = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.PictureBox();
            this.btnGenerarReporte = new System.Windows.Forms.Button();
            this.cboOpcionReporte = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpModelo = new System.Windows.Forms.GroupBox();
            this.radUnModelo = new System.Windows.Forms.RadioButton();
            this.rdTodosLosModelos = new System.Windows.Forms.RadioButton();
            this.cboModelos = new System.Windows.Forms.ComboBox();
            this.rdTodasLasMarcas = new System.Windows.Forms.RadioButton();
            this.rdUnaMarca = new System.Windows.Forms.RadioButton();
            this.cboMarcas = new System.Windows.Forms.ComboBox();
            this.pSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
            this.grpModelo.SuspendLayout();
            this.SuspendLayout();
            // 
            // pSuperior
            // 
            this.pSuperior.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            this.pSuperior.Controls.Add(this.btnCerrar);
            this.pSuperior.Dock = System.Windows.Forms.DockStyle.Top;
            this.pSuperior.Location = new System.Drawing.Point(0, 0);
            this.pSuperior.Name = "pSuperior";
            this.pSuperior.Size = new System.Drawing.Size(501, 36);
            this.pSuperior.TabIndex = 0;
            this.pSuperior.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pSuperior_MouseDown);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(473, 3);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(25, 28);
            this.btnCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnCerrar.TabIndex = 6;
            this.btnCerrar.TabStop = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnGenerarReporte
            // 
            this.btnGenerarReporte.FlatAppearance.BorderColor = System.Drawing.Color.Crimson;
            this.btnGenerarReporte.FlatAppearance.BorderSize = 2;
            this.btnGenerarReporte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerarReporte.ForeColor = System.Drawing.Color.White;
            this.btnGenerarReporte.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarReporte.Image")));
            this.btnGenerarReporte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGenerarReporte.Location = new System.Drawing.Point(372, 251);
            this.btnGenerarReporte.Name = "btnGenerarReporte";
            this.btnGenerarReporte.Size = new System.Drawing.Size(117, 51);
            this.btnGenerarReporte.TabIndex = 28;
            this.btnGenerarReporte.Text = "Generar";
            this.btnGenerarReporte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGenerarReporte.UseVisualStyleBackColor = true;
            this.btnGenerarReporte.Click += new System.EventHandler(this.btnGenerarReporte_Click);
            // 
            // cboOpcionReporte
            // 
            this.cboOpcionReporte.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboOpcionReporte.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboOpcionReporte.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOpcionReporte.FormattingEnabled = true;
            this.cboOpcionReporte.Location = new System.Drawing.Point(163, 47);
            this.cboOpcionReporte.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionReporte.Name = "cboOpcionReporte";
            this.cboOpcionReporte.Size = new System.Drawing.Size(304, 27);
            this.cboOpcionReporte.Sorted = true;
            this.cboOpcionReporte.TabIndex = 21;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(56, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 19);
            this.label1.TabIndex = 20;
            this.label1.Text = "Generar por:";
            // 
            // grpModelo
            // 
            this.grpModelo.Controls.Add(this.radUnModelo);
            this.grpModelo.Controls.Add(this.rdTodosLosModelos);
            this.grpModelo.Controls.Add(this.cboModelos);
            this.grpModelo.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpModelo.Location = new System.Drawing.Point(162, 145);
            this.grpModelo.Name = "grpModelo";
            this.grpModelo.Size = new System.Drawing.Size(327, 100);
            this.grpModelo.TabIndex = 65;
            this.grpModelo.TabStop = false;
            this.grpModelo.Text = "Modelos";
            this.grpModelo.Visible = false;
            // 
            // radUnModelo
            // 
            this.radUnModelo.AutoSize = true;
            this.radUnModelo.Location = new System.Drawing.Point(6, 26);
            this.radUnModelo.Name = "radUnModelo";
            this.radUnModelo.Size = new System.Drawing.Size(56, 23);
            this.radUnModelo.TabIndex = 55;
            this.radUnModelo.TabStop = true;
            this.radUnModelo.Text = "Uno";
            this.radUnModelo.UseVisualStyleBackColor = true;
            this.radUnModelo.CheckedChanged += new System.EventHandler(this.radUnModelo_CheckedChanged);
            // 
            // rdTodosLosModelos
            // 
            this.rdTodosLosModelos.AutoSize = true;
            this.rdTodosLosModelos.Location = new System.Drawing.Point(67, 26);
            this.rdTodosLosModelos.Name = "rdTodosLosModelos";
            this.rdTodosLosModelos.Size = new System.Drawing.Size(71, 23);
            this.rdTodosLosModelos.TabIndex = 56;
            this.rdTodosLosModelos.TabStop = true;
            this.rdTodosLosModelos.Text = "Todos";
            this.rdTodosLosModelos.UseVisualStyleBackColor = true;
            this.rdTodosLosModelos.CheckedChanged += new System.EventHandler(this.rdTodosLosModelos_CheckedChanged);
            // 
            // cboModelos
            // 
            this.cboModelos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelos.FormattingEnabled = true;
            this.cboModelos.Location = new System.Drawing.Point(7, 56);
            this.cboModelos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelos.Name = "cboModelos";
            this.cboModelos.Size = new System.Drawing.Size(297, 27);
            this.cboModelos.Sorted = true;
            this.cboModelos.TabIndex = 54;
            this.cboModelos.Visible = false;
            // 
            // rdTodasLasMarcas
            // 
            this.rdTodasLasMarcas.AutoSize = true;
            this.rdTodasLasMarcas.Location = new System.Drawing.Point(224, 81);
            this.rdTodasLasMarcas.Name = "rdTodasLasMarcas";
            this.rdTodasLasMarcas.Size = new System.Drawing.Size(70, 23);
            this.rdTodasLasMarcas.TabIndex = 64;
            this.rdTodasLasMarcas.TabStop = true;
            this.rdTodasLasMarcas.Text = "Todas";
            this.rdTodasLasMarcas.UseVisualStyleBackColor = true;
            this.rdTodasLasMarcas.CheckedChanged += new System.EventHandler(this.rdTodasLasMarcas_CheckedChanged);
            // 
            // rdUnaMarca
            // 
            this.rdUnaMarca.AutoSize = true;
            this.rdUnaMarca.Location = new System.Drawing.Point(163, 81);
            this.rdUnaMarca.Name = "rdUnaMarca";
            this.rdUnaMarca.Size = new System.Drawing.Size(55, 23);
            this.rdUnaMarca.TabIndex = 63;
            this.rdUnaMarca.TabStop = true;
            this.rdUnaMarca.Text = "Una";
            this.rdUnaMarca.UseVisualStyleBackColor = true;
            this.rdUnaMarca.CheckedChanged += new System.EventHandler(this.rdUnaMarca_CheckedChanged);
            // 
            // cboMarcas
            // 
            this.cboMarcas.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarcas.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarcas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarcas.FormattingEnabled = true;
            this.cboMarcas.Location = new System.Drawing.Point(163, 111);
            this.cboMarcas.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarcas.Name = "cboMarcas";
            this.cboMarcas.Size = new System.Drawing.Size(303, 27);
            this.cboMarcas.Sorted = true;
            this.cboMarcas.TabIndex = 62;
            this.cboMarcas.Visible = false;
            this.cboMarcas.SelectedIndexChanged += new System.EventHandler(this.cboMarcas_SelectedIndexChanged);
            // 
            // ReporteInventario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.ClientSize = new System.Drawing.Size(501, 318);
            this.Controls.Add(this.grpModelo);
            this.Controls.Add(this.rdTodasLasMarcas);
            this.Controls.Add(this.rdUnaMarca);
            this.Controls.Add(this.cboMarcas);
            this.Controls.Add(this.btnGenerarReporte);
            this.Controls.Add(this.cboOpcionReporte);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pSuperior);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ReporteInventario";
            this.Text = "ReporteInventario";
            this.pSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
            this.grpModelo.ResumeLayout(false);
            this.grpModelo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel pSuperior;
        private System.Windows.Forms.PictureBox btnCerrar;
        private System.Windows.Forms.Button btnGenerarReporte;
        private System.Windows.Forms.ComboBox cboOpcionReporte;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpModelo;
        private System.Windows.Forms.RadioButton radUnModelo;
        private System.Windows.Forms.RadioButton rdTodosLosModelos;
        private System.Windows.Forms.ComboBox cboModelos;
        private System.Windows.Forms.RadioButton rdTodasLasMarcas;
        private System.Windows.Forms.RadioButton rdUnaMarca;
        private System.Windows.Forms.ComboBox cboMarcas;
    }
}