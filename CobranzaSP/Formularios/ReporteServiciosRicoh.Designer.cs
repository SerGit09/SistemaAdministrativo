namespace CobranzaSP.Formularios
{
    partial class ReporteServiciosRicoh
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReporteServiciosRicoh));
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
            this.cboModulos = new System.Windows.Forms.ComboBox();
            this.lblOpcionReporte = new System.Windows.Forms.Label();
            this.lblModulo = new System.Windows.Forms.Label();
            this.lblClaves = new System.Windows.Forms.Label();
            this.cboClaves = new System.Windows.Forms.ComboBox();
            this.erReporte = new System.Windows.Forms.ErrorProvider(this.components);
            this.txtClaveModulo = new System.Windows.Forms.TextBox();
            this.pSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
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
            this.pSuperior.Size = new System.Drawing.Size(548, 38);
            this.pSuperior.TabIndex = 1;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(520, 3);
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
            this.chkRango.Location = new System.Drawing.Point(235, 318);
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
            this.btnGenerarReporte.Location = new System.Drawing.Point(419, 421);
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
            this.dtpFechaFinal.Location = new System.Drawing.Point(234, 388);
            this.dtpFechaFinal.Name = "dtpFechaFinal";
            this.dtpFechaFinal.Size = new System.Drawing.Size(303, 27);
            this.dtpFechaFinal.TabIndex = 26;
            this.dtpFechaFinal.Visible = false;
            // 
            // dtpFechaInicial
            // 
            this.dtpFechaInicial.Location = new System.Drawing.Point(235, 347);
            this.dtpFechaInicial.Name = "dtpFechaInicial";
            this.dtpFechaInicial.Size = new System.Drawing.Size(303, 27);
            this.dtpFechaInicial.TabIndex = 25;
            this.dtpFechaInicial.Visible = false;
            // 
            // lblFechaFinal
            // 
            this.lblFechaFinal.AutoSize = true;
            this.lblFechaFinal.ForeColor = System.Drawing.Color.White;
            this.lblFechaFinal.Location = new System.Drawing.Point(135, 392);
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
            this.lblFechaInicio.Location = new System.Drawing.Point(126, 353);
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
            this.cboOpcionElegida.Location = new System.Drawing.Point(208, 83);
            this.cboOpcionElegida.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionElegida.Name = "cboOpcionElegida";
            this.cboOpcionElegida.Size = new System.Drawing.Size(324, 27);
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
            this.cboOpcionReporte.Location = new System.Drawing.Point(209, 47);
            this.cboOpcionReporte.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionReporte.Name = "cboOpcionReporte";
            this.cboOpcionReporte.Size = new System.Drawing.Size(324, 27);
            this.cboOpcionReporte.Sorted = true;
            this.cboOpcionReporte.TabIndex = 20;
            this.cboOpcionReporte.SelectedIndexChanged += new System.EventHandler(this.cboOpcionReporte_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(102, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 19);
            this.label1.TabIndex = 19;
            this.label1.Text = "Generar por:";
            // 
            // cboModulos
            // 
            this.cboModulos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModulos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModulos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModulos.FormattingEnabled = true;
            this.cboModulos.Location = new System.Drawing.Point(216, 206);
            this.cboModulos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModulos.Name = "cboModulos";
            this.cboModulos.Size = new System.Drawing.Size(322, 27);
            this.cboModulos.Sorted = true;
            this.cboModulos.TabIndex = 59;
            this.cboModulos.Visible = false;
            this.cboModulos.SelectedIndexChanged += new System.EventHandler(this.cboModulos_SelectedIndexChanged);
            // 
            // lblOpcionReporte
            // 
            this.lblOpcionReporte.AutoSize = true;
            this.lblOpcionReporte.ForeColor = System.Drawing.Color.White;
            this.lblOpcionReporte.Location = new System.Drawing.Point(126, 86);
            this.lblOpcionReporte.Name = "lblOpcionReporte";
            this.lblOpcionReporte.Size = new System.Drawing.Size(66, 19);
            this.lblOpcionReporte.TabIndex = 61;
            this.lblOpcionReporte.Text = "Modelo:";
            this.lblOpcionReporte.Visible = false;
            // 
            // lblModulo
            // 
            this.lblModulo.AutoSize = true;
            this.lblModulo.ForeColor = System.Drawing.Color.White;
            this.lblModulo.Location = new System.Drawing.Point(135, 209);
            this.lblModulo.Name = "lblModulo";
            this.lblModulo.Size = new System.Drawing.Size(67, 19);
            this.lblModulo.TabIndex = 62;
            this.lblModulo.Text = "Modulo:";
            this.lblModulo.Visible = false;
            // 
            // lblClaves
            // 
            this.lblClaves.AutoSize = true;
            this.lblClaves.ForeColor = System.Drawing.Color.White;
            this.lblClaves.Location = new System.Drawing.Point(149, 245);
            this.lblClaves.Name = "lblClaves";
            this.lblClaves.Size = new System.Drawing.Size(53, 19);
            this.lblClaves.TabIndex = 64;
            this.lblClaves.Text = "Clave:";
            this.lblClaves.Visible = false;
            // 
            // cboClaves
            // 
            this.cboClaves.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClaves.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClaves.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClaves.FormattingEnabled = true;
            this.cboClaves.Location = new System.Drawing.Point(216, 242);
            this.cboClaves.Margin = new System.Windows.Forms.Padding(4);
            this.cboClaves.Name = "cboClaves";
            this.cboClaves.Size = new System.Drawing.Size(322, 27);
            this.cboClaves.Sorted = true;
            this.cboClaves.TabIndex = 63;
            this.cboClaves.Visible = false;
            // 
            // erReporte
            // 
            this.erReporte.ContainerControl = this;
            // 
            // txtClaveModulo
            // 
            this.txtClaveModulo.Location = new System.Drawing.Point(208, 83);
            this.txtClaveModulo.Name = "txtClaveModulo";
            this.txtClaveModulo.Size = new System.Drawing.Size(325, 27);
            this.txtClaveModulo.TabIndex = 65;
            this.txtClaveModulo.Visible = false;
            // 
            // ReporteServiciosRicoh
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.ClientSize = new System.Drawing.Size(548, 542);
            this.Controls.Add(this.txtClaveModulo);
            this.Controls.Add(this.lblClaves);
            this.Controls.Add(this.cboClaves);
            this.Controls.Add(this.lblModulo);
            this.Controls.Add(this.lblOpcionReporte);
            this.Controls.Add(this.cboModulos);
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
            this.Name = "ReporteServiciosRicoh";
            this.Text = "ReporteServiciosRicoh";
            this.pSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
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
        private System.Windows.Forms.ComboBox cboModulos;
        private System.Windows.Forms.Label lblOpcionReporte;
        private System.Windows.Forms.Label lblModulo;
        private System.Windows.Forms.Label lblClaves;
        private System.Windows.Forms.ComboBox cboClaves;
        private System.Windows.Forms.ErrorProvider erReporte;
        private System.Windows.Forms.TextBox txtClaveModulo;
    }
}