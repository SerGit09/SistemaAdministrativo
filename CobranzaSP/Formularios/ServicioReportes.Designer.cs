namespace CobranzaSP.Formularios
{
    partial class ServicioReportes
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ServicioReportes));
            this.pSuperior = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.PictureBox();
            this.chkRango = new System.Windows.Forms.CheckBox();
            this.btnGenerarReporte = new System.Windows.Forms.Button();
            this.dtpFechaFinal = new System.Windows.Forms.DateTimePicker();
            this.dtpFechaInicial = new System.Windows.Forms.DateTimePicker();
            this.lblFechaFinal = new System.Windows.Forms.Label();
            this.lblFechaInicio = new System.Windows.Forms.Label();
            this.cboOpcionElegida = new System.Windows.Forms.ComboBox();
            this.cboOpcionReporte = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpModelo = new System.Windows.Forms.GroupBox();
            this.radUnModelo = new System.Windows.Forms.RadioButton();
            this.rdTodosLosModelos = new System.Windows.Forms.RadioButton();
            this.cboModelo = new System.Windows.Forms.ComboBox();
            this.rdTodasLasMarcas = new System.Windows.Forms.RadioButton();
            this.rdUnaMarca = new System.Windows.Forms.RadioButton();
            this.cboMarca = new System.Windows.Forms.ComboBox();
            this.grpMarca = new System.Windows.Forms.GroupBox();
            this.grpFusor = new System.Windows.Forms.GroupBox();
            this.rdUnFusor = new System.Windows.Forms.RadioButton();
            this.cboFusor = new System.Windows.Forms.ComboBox();
            this.rdTodosLosFusores = new System.Windows.Forms.RadioButton();
            this.erReporte = new System.Windows.Forms.ErrorProvider(this.components);
            this.pSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
            this.grpModelo.SuspendLayout();
            this.grpMarca.SuspendLayout();
            this.grpFusor.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.erReporte)).BeginInit();
            this.SuspendLayout();
            // 
            // pSuperior
            // 
            this.pSuperior.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            this.pSuperior.Controls.Add(this.btnCerrar);
            this.pSuperior.Dock = System.Windows.Forms.DockStyle.Top;
            this.pSuperior.Location = new System.Drawing.Point(0, 0);
            this.pSuperior.Name = "pSuperior";
            this.pSuperior.Size = new System.Drawing.Size(516, 38);
            this.pSuperior.TabIndex = 1;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(488, 3);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(25, 28);
            this.btnCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnCerrar.TabIndex = 5;
            this.btnCerrar.TabStop = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // chkRango
            // 
            this.chkRango.AutoSize = true;
            this.chkRango.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.chkRango.Location = new System.Drawing.Point(196, 302);
            this.chkRango.Name = "chkRango";
            this.chkRango.Size = new System.Drawing.Size(125, 23);
            this.chkRango.TabIndex = 28;
            this.chkRango.Text = "Rango Fechas";
            this.chkRango.UseVisualStyleBackColor = true;
            this.chkRango.Visible = false;
            this.chkRango.CheckedChanged += new System.EventHandler(this.chkRango_CheckedChanged);
            // 
            // btnGenerarReporte
            // 
            this.btnGenerarReporte.FlatAppearance.BorderColor = System.Drawing.Color.Crimson;
            this.btnGenerarReporte.FlatAppearance.BorderSize = 2;
            this.btnGenerarReporte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerarReporte.ForeColor = System.Drawing.Color.White;
            this.btnGenerarReporte.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarReporte.Image")));
            this.btnGenerarReporte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGenerarReporte.Location = new System.Drawing.Point(396, 404);
            this.btnGenerarReporte.Name = "btnGenerarReporte";
            this.btnGenerarReporte.Size = new System.Drawing.Size(117, 51);
            this.btnGenerarReporte.TabIndex = 27;
            this.btnGenerarReporte.Text = "Generar";
            this.btnGenerarReporte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGenerarReporte.UseVisualStyleBackColor = true;
            this.btnGenerarReporte.Click += new System.EventHandler(this.btnGenerarReporte_Click);
            // 
            // dtpFechaFinal
            // 
            this.dtpFechaFinal.Location = new System.Drawing.Point(209, 369);
            this.dtpFechaFinal.Name = "dtpFechaFinal";
            this.dtpFechaFinal.Size = new System.Drawing.Size(301, 27);
            this.dtpFechaFinal.TabIndex = 26;
            this.dtpFechaFinal.Visible = false;
            // 
            // dtpFechaInicial
            // 
            this.dtpFechaInicial.Location = new System.Drawing.Point(207, 332);
            this.dtpFechaInicial.Name = "dtpFechaInicial";
            this.dtpFechaInicial.Size = new System.Drawing.Size(303, 27);
            this.dtpFechaInicial.TabIndex = 25;
            this.dtpFechaInicial.Visible = false;
            // 
            // lblFechaFinal
            // 
            this.lblFechaFinal.AutoSize = true;
            this.lblFechaFinal.ForeColor = System.Drawing.Color.White;
            this.lblFechaFinal.Location = new System.Drawing.Point(108, 373);
            this.lblFechaFinal.Name = "lblFechaFinal";
            this.lblFechaFinal.Size = new System.Drawing.Size(93, 19);
            this.lblFechaFinal.TabIndex = 24;
            this.lblFechaFinal.Text = "Fecha Final:";
            this.lblFechaFinal.Visible = false;
            // 
            // lblFechaInicio
            // 
            this.lblFechaInicio.AutoSize = true;
            this.lblFechaInicio.ForeColor = System.Drawing.Color.White;
            this.lblFechaInicio.Location = new System.Drawing.Point(99, 338);
            this.lblFechaInicio.Name = "lblFechaInicio";
            this.lblFechaInicio.Size = new System.Drawing.Size(102, 19);
            this.lblFechaInicio.TabIndex = 23;
            this.lblFechaInicio.Text = "Fecha Inicial:";
            this.lblFechaInicio.Visible = false;
            // 
            // cboOpcionElegida
            // 
            this.cboOpcionElegida.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboOpcionElegida.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboOpcionElegida.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOpcionElegida.FormattingEnabled = true;
            this.cboOpcionElegida.Location = new System.Drawing.Point(197, 82);
            this.cboOpcionElegida.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionElegida.Name = "cboOpcionElegida";
            this.cboOpcionElegida.Size = new System.Drawing.Size(303, 27);
            this.cboOpcionElegida.Sorted = true;
            this.cboOpcionElegida.TabIndex = 21;
            this.cboOpcionElegida.Visible = false;
            this.cboOpcionElegida.SelectedIndexChanged += new System.EventHandler(this.cboOpcionElegida_SelectedIndexChanged);
            // 
            // cboOpcionReporte
            // 
            this.cboOpcionReporte.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboOpcionReporte.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboOpcionReporte.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOpcionReporte.FormattingEnabled = true;
            this.cboOpcionReporte.Location = new System.Drawing.Point(197, 47);
            this.cboOpcionReporte.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionReporte.Name = "cboOpcionReporte";
            this.cboOpcionReporte.Size = new System.Drawing.Size(304, 27);
            this.cboOpcionReporte.Sorted = true;
            this.cboOpcionReporte.TabIndex = 20;
            this.cboOpcionReporte.SelectedIndexChanged += new System.EventHandler(this.cboOpcionReporte_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(90, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 19);
            this.label1.TabIndex = 19;
            this.label1.Text = "Generar por:";
            // 
            // grpModelo
            // 
            this.grpModelo.Controls.Add(this.radUnModelo);
            this.grpModelo.Controls.Add(this.rdTodosLosModelos);
            this.grpModelo.Controls.Add(this.cboModelo);
            this.grpModelo.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpModelo.Location = new System.Drawing.Point(21, 209);
            this.grpModelo.Name = "grpModelo";
            this.grpModelo.Size = new System.Drawing.Size(492, 87);
            this.grpModelo.TabIndex = 65;
            this.grpModelo.TabStop = false;
            this.grpModelo.Text = "Modelo";
            this.grpModelo.Visible = false;
            // 
            // radUnModelo
            // 
            this.radUnModelo.AutoSize = true;
            this.radUnModelo.Location = new System.Drawing.Point(185, 17);
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
            this.rdTodosLosModelos.Location = new System.Drawing.Point(279, 18);
            this.rdTodosLosModelos.Name = "rdTodosLosModelos";
            this.rdTodosLosModelos.Size = new System.Drawing.Size(71, 23);
            this.rdTodosLosModelos.TabIndex = 56;
            this.rdTodosLosModelos.TabStop = true;
            this.rdTodosLosModelos.Text = "Todos";
            this.rdTodosLosModelos.UseVisualStyleBackColor = true;
            this.rdTodosLosModelos.CheckedChanged += new System.EventHandler(this.rdTodosLosModelos_CheckedChanged);
            // 
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(179, 48);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(298, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 54;
            this.cboModelo.Visible = false;
            // 
            // rdTodasLasMarcas
            // 
            this.rdTodasLasMarcas.AutoSize = true;
            this.rdTodasLasMarcas.ForeColor = System.Drawing.SystemColors.Control;
            this.rdTodasLasMarcas.Location = new System.Drawing.Point(291, 14);
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
            this.rdUnaMarca.ForeColor = System.Drawing.SystemColors.Control;
            this.rdUnaMarca.Location = new System.Drawing.Point(186, 14);
            this.rdUnaMarca.Name = "rdUnaMarca";
            this.rdUnaMarca.Size = new System.Drawing.Size(55, 23);
            this.rdUnaMarca.TabIndex = 63;
            this.rdUnaMarca.TabStop = true;
            this.rdUnaMarca.Text = "Una";
            this.rdUnaMarca.UseVisualStyleBackColor = true;
            this.rdUnaMarca.CheckedChanged += new System.EventHandler(this.rdUnaMarca_CheckedChanged);
            // 
            // cboMarca
            // 
            this.cboMarca.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarca.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarca.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarca.FormattingEnabled = true;
            this.cboMarca.Location = new System.Drawing.Point(186, 44);
            this.cboMarca.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarca.Name = "cboMarca";
            this.cboMarca.Size = new System.Drawing.Size(292, 27);
            this.cboMarca.Sorted = true;
            this.cboMarca.TabIndex = 62;
            this.cboMarca.Visible = false;
            this.cboMarca.SelectedIndexChanged += new System.EventHandler(this.cboMarca_SelectedIndexChanged);
            // 
            // grpMarca
            // 
            this.grpMarca.Controls.Add(this.rdUnaMarca);
            this.grpMarca.Controls.Add(this.cboMarca);
            this.grpMarca.Controls.Add(this.rdTodasLasMarcas);
            this.grpMarca.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpMarca.Location = new System.Drawing.Point(21, 116);
            this.grpMarca.Name = "grpMarca";
            this.grpMarca.Size = new System.Drawing.Size(492, 88);
            this.grpMarca.TabIndex = 66;
            this.grpMarca.TabStop = false;
            this.grpMarca.Text = "Marca";
            this.grpMarca.Visible = false;
            // 
            // grpFusor
            // 
            this.grpFusor.Controls.Add(this.rdUnFusor);
            this.grpFusor.Controls.Add(this.cboFusor);
            this.grpFusor.Controls.Add(this.rdTodosLosFusores);
            this.grpFusor.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpFusor.Location = new System.Drawing.Point(21, 116);
            this.grpFusor.Name = "grpFusor";
            this.grpFusor.Size = new System.Drawing.Size(492, 88);
            this.grpFusor.TabIndex = 67;
            this.grpFusor.TabStop = false;
            this.grpFusor.Text = "Fusor";
            this.grpFusor.Visible = false;
            // 
            // rdUnFusor
            // 
            this.rdUnFusor.AutoSize = true;
            this.rdUnFusor.ForeColor = System.Drawing.SystemColors.Control;
            this.rdUnFusor.Location = new System.Drawing.Point(186, 14);
            this.rdUnFusor.Name = "rdUnFusor";
            this.rdUnFusor.Size = new System.Drawing.Size(56, 23);
            this.rdUnFusor.TabIndex = 63;
            this.rdUnFusor.TabStop = true;
            this.rdUnFusor.Text = "Uno";
            this.rdUnFusor.UseVisualStyleBackColor = true;
            this.rdUnFusor.CheckedChanged += new System.EventHandler(this.rdUnFusor_CheckedChanged);
            // 
            // cboFusor
            // 
            this.cboFusor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboFusor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboFusor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFusor.FormattingEnabled = true;
            this.cboFusor.Location = new System.Drawing.Point(186, 44);
            this.cboFusor.Margin = new System.Windows.Forms.Padding(4);
            this.cboFusor.Name = "cboFusor";
            this.cboFusor.Size = new System.Drawing.Size(292, 27);
            this.cboFusor.Sorted = true;
            this.cboFusor.TabIndex = 62;
            this.cboFusor.Visible = false;
            // 
            // rdTodosLosFusores
            // 
            this.rdTodosLosFusores.AutoSize = true;
            this.rdTodosLosFusores.ForeColor = System.Drawing.SystemColors.Control;
            this.rdTodosLosFusores.Location = new System.Drawing.Point(291, 14);
            this.rdTodosLosFusores.Name = "rdTodosLosFusores";
            this.rdTodosLosFusores.Size = new System.Drawing.Size(71, 23);
            this.rdTodosLosFusores.TabIndex = 64;
            this.rdTodosLosFusores.TabStop = true;
            this.rdTodosLosFusores.Text = "Todos";
            this.rdTodosLosFusores.UseVisualStyleBackColor = true;
            this.rdTodosLosFusores.CheckedChanged += new System.EventHandler(this.rdTodosLosFusores_CheckedChanged);
            // 
            // erReporte
            // 
            this.erReporte.ContainerControl = this;
            // 
            // ServicioReportes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.ClientSize = new System.Drawing.Size(516, 467);
            this.Controls.Add(this.grpFusor);
            this.Controls.Add(this.grpMarca);
            this.Controls.Add(this.grpModelo);
            this.Controls.Add(this.chkRango);
            this.Controls.Add(this.btnGenerarReporte);
            this.Controls.Add(this.dtpFechaFinal);
            this.Controls.Add(this.dtpFechaInicial);
            this.Controls.Add(this.lblFechaFinal);
            this.Controls.Add(this.lblFechaInicio);
            this.Controls.Add(this.cboOpcionElegida);
            this.Controls.Add(this.cboOpcionReporte);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pSuperior);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ServicioReportes";
            this.Text = "ServicioReportes";
            this.pSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
            this.grpModelo.ResumeLayout(false);
            this.grpModelo.PerformLayout();
            this.grpMarca.ResumeLayout(false);
            this.grpMarca.PerformLayout();
            this.grpFusor.ResumeLayout(false);
            this.grpFusor.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.erReporte)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel pSuperior;
        private System.Windows.Forms.PictureBox btnCerrar;
        private System.Windows.Forms.CheckBox chkRango;
        private System.Windows.Forms.Button btnGenerarReporte;
        private System.Windows.Forms.DateTimePicker dtpFechaFinal;
        private System.Windows.Forms.DateTimePicker dtpFechaInicial;
        private System.Windows.Forms.Label lblFechaFinal;
        private System.Windows.Forms.Label lblFechaInicio;
        private System.Windows.Forms.ComboBox cboOpcionElegida;
        private System.Windows.Forms.ComboBox cboOpcionReporte;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpModelo;
        private System.Windows.Forms.RadioButton radUnModelo;
        private System.Windows.Forms.RadioButton rdTodosLosModelos;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.RadioButton rdTodasLasMarcas;
        private System.Windows.Forms.RadioButton rdUnaMarca;
        private System.Windows.Forms.ComboBox cboMarca;
        private System.Windows.Forms.GroupBox grpMarca;
        private System.Windows.Forms.GroupBox grpFusor;
        private System.Windows.Forms.RadioButton rdUnFusor;
        private System.Windows.Forms.ComboBox cboFusor;
        private System.Windows.Forms.RadioButton rdTodosLosFusores;
        private System.Windows.Forms.ErrorProvider erReporte;
    }
}