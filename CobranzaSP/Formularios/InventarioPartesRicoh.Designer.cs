namespace CobranzaSP.Formularios
{
    partial class InventarioPartesRicoh
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InventarioPartesRicoh));
            this.dtgInventario = new System.Windows.Forms.DataGridView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabInventario = new System.Windows.Forms.TabPage();
            this.grpDatosInventario = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.radNuevo = new System.Windows.Forms.RadioButton();
            this.radUsado = new System.Windows.Forms.RadioButton();
            this.rtxtDescripcion = new System.Windows.Forms.RichTextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.txtNumeroParte = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tabRegistros = new System.Windows.Forms.TabPage();
            this.grpDatosRegistro = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkCliente = new System.Windows.Forms.CheckBox();
            this.radEntrada = new System.Windows.Forms.RadioButton();
            this.radSalida = new System.Windows.Forms.RadioButton();
            this.btnGenerarReporte = new System.Windows.Forms.Button();
            this.lblTipoPersona = new System.Windows.Forms.Label();
            this.cboTipoPersona = new System.Windows.Forms.ComboBox();
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cboModelo = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.dtpFechaRegistro = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.Buscar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dtgInventario)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabInventario.SuspendLayout();
            this.grpDatosInventario.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabRegistros.SuspendLayout();
            this.grpDatosRegistro.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtgInventario
            // 
            this.dtgInventario.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgInventario.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgInventario.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgInventario.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgInventario.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgInventario.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgInventario.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgInventario.EnableHeadersVisualStyles = false;
            this.dtgInventario.Location = new System.Drawing.Point(12, 284);
            this.dtgInventario.Name = "dtgInventario";
            this.dtgInventario.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgInventario.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgInventario.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgInventario.RowTemplate.Height = 25;
            this.dtgInventario.Size = new System.Drawing.Size(1326, 681);
            this.dtgInventario.TabIndex = 39;
            this.dtgInventario.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgInventario_CellClick);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabInventario);
            this.tabControl1.Controls.Add(this.tabRegistros);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1141, 209);
            this.tabControl1.TabIndex = 35;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabInventario
            // 
            this.tabInventario.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.tabInventario.Controls.Add(this.grpDatosInventario);
            this.tabInventario.Location = new System.Drawing.Point(4, 28);
            this.tabInventario.Name = "tabInventario";
            this.tabInventario.Padding = new System.Windows.Forms.Padding(3);
            this.tabInventario.Size = new System.Drawing.Size(1133, 177);
            this.tabInventario.TabIndex = 0;
            this.tabInventario.Text = "Catalogo";
            // 
            // grpDatosInventario
            // 
            this.grpDatosInventario.Controls.Add(this.groupBox3);
            this.grpDatosInventario.Controls.Add(this.rtxtDescripcion);
            this.grpDatosInventario.Controls.Add(this.label4);
            this.grpDatosInventario.Controls.Add(this.btnImprimir);
            this.grpDatosInventario.Controls.Add(this.txtNumeroParte);
            this.grpDatosInventario.Controls.Add(this.label1);
            this.grpDatosInventario.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpDatosInventario.Location = new System.Drawing.Point(6, 6);
            this.grpDatosInventario.Name = "grpDatosInventario";
            this.grpDatosInventario.Size = new System.Drawing.Size(1111, 156);
            this.grpDatosInventario.TabIndex = 0;
            this.grpDatosInventario.TabStop = false;
            this.grpDatosInventario.Text = "Datos";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.radNuevo);
            this.groupBox3.Controls.Add(this.radUsado);
            this.groupBox3.ForeColor = System.Drawing.SystemColors.Control;
            this.groupBox3.Location = new System.Drawing.Point(392, 16);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(265, 64);
            this.groupBox3.TabIndex = 50;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Estado:";
            // 
            // radNuevo
            // 
            this.radNuevo.AutoSize = true;
            this.radNuevo.Location = new System.Drawing.Point(6, 26);
            this.radNuevo.Name = "radNuevo";
            this.radNuevo.Size = new System.Drawing.Size(72, 23);
            this.radNuevo.TabIndex = 46;
            this.radNuevo.TabStop = true;
            this.radNuevo.Text = "Nuevo";
            this.radNuevo.UseVisualStyleBackColor = true;
            // 
            // radUsado
            // 
            this.radUsado.AutoSize = true;
            this.radUsado.Location = new System.Drawing.Point(84, 26);
            this.radUsado.Name = "radUsado";
            this.radUsado.Size = new System.Drawing.Size(71, 23);
            this.radUsado.TabIndex = 47;
            this.radUsado.TabStop = true;
            this.radUsado.Text = "Usado";
            this.radUsado.UseVisualStyleBackColor = true;
            // 
            // rtxtDescripcion
            // 
            this.rtxtDescripcion.Location = new System.Drawing.Point(6, 76);
            this.rtxtDescripcion.Name = "rtxtDescripcion";
            this.rtxtDescripcion.Size = new System.Drawing.Size(370, 73);
            this.rtxtDescripcion.TabIndex = 40;
            this.rtxtDescripcion.Text = "";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(6, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 19);
            this.label4.TabIndex = 39;
            this.label4.Text = "Descripcion";
            // 
            // btnImprimir
            // 
            this.btnImprimir.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnImprimir.FlatAppearance.BorderSize = 2;
            this.btnImprimir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnImprimir.ForeColor = System.Drawing.Color.White;
            this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
            this.btnImprimir.Location = new System.Drawing.Point(1047, 101);
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.Size = new System.Drawing.Size(58, 49);
            this.btnImprimir.TabIndex = 36;
            this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnImprimir.UseVisualStyleBackColor = true;
            this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
            // 
            // txtNumeroParte
            // 
            this.txtNumeroParte.Location = new System.Drawing.Point(125, 16);
            this.txtNumeroParte.Name = "txtNumeroParte";
            this.txtNumeroParte.Size = new System.Drawing.Size(237, 27);
            this.txtNumeroParte.TabIndex = 8;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 19);
            this.label1.TabIndex = 4;
            this.label1.Text = "Numero Parte:";
            // 
            // tabRegistros
            // 
            this.tabRegistros.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.tabRegistros.Controls.Add(this.grpDatosRegistro);
            this.tabRegistros.Location = new System.Drawing.Point(4, 28);
            this.tabRegistros.Name = "tabRegistros";
            this.tabRegistros.Padding = new System.Windows.Forms.Padding(3);
            this.tabRegistros.Size = new System.Drawing.Size(1133, 177);
            this.tabRegistros.TabIndex = 1;
            this.tabRegistros.Text = "Movimientos";
            // 
            // grpDatosRegistro
            // 
            this.grpDatosRegistro.Controls.Add(this.groupBox1);
            this.grpDatosRegistro.Controls.Add(this.btnGenerarReporte);
            this.grpDatosRegistro.Controls.Add(this.lblTipoPersona);
            this.grpDatosRegistro.Controls.Add(this.cboTipoPersona);
            this.grpDatosRegistro.Controls.Add(this.txtCantidad);
            this.grpDatosRegistro.Controls.Add(this.label7);
            this.grpDatosRegistro.Controls.Add(this.cboModelo);
            this.grpDatosRegistro.Controls.Add(this.label12);
            this.grpDatosRegistro.Controls.Add(this.dtpFechaRegistro);
            this.grpDatosRegistro.Controls.Add(this.label9);
            this.grpDatosRegistro.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpDatosRegistro.Location = new System.Drawing.Point(11, 5);
            this.grpDatosRegistro.Name = "grpDatosRegistro";
            this.grpDatosRegistro.Size = new System.Drawing.Size(1111, 166);
            this.grpDatosRegistro.TabIndex = 1;
            this.grpDatosRegistro.TabStop = false;
            this.grpDatosRegistro.Text = "Datos";
            this.grpDatosRegistro.Enter += new System.EventHandler(this.grpDatosRegistro_Enter);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkCliente);
            this.groupBox1.Controls.Add(this.radEntrada);
            this.groupBox1.Controls.Add(this.radSalida);
            this.groupBox1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.groupBox1.Location = new System.Drawing.Point(10, 87);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(312, 65);
            this.groupBox1.TabIndex = 48;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Movimiento:";
            // 
            // chkCliente
            // 
            this.chkCliente.AutoSize = true;
            this.chkCliente.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkCliente.ForeColor = System.Drawing.Color.White;
            this.chkCliente.Location = new System.Drawing.Point(230, 27);
            this.chkCliente.Name = "chkCliente";
            this.chkCliente.Size = new System.Drawing.Size(76, 23);
            this.chkCliente.TabIndex = 51;
            this.chkCliente.Text = "Cliente";
            this.chkCliente.UseVisualStyleBackColor = true;
            this.chkCliente.Visible = false;
            this.chkCliente.CheckedChanged += new System.EventHandler(this.chkCliente_CheckedChanged);
            // 
            // radEntrada
            // 
            this.radEntrada.AutoSize = true;
            this.radEntrada.Location = new System.Drawing.Point(91, 26);
            this.radEntrada.Name = "radEntrada";
            this.radEntrada.Size = new System.Drawing.Size(81, 23);
            this.radEntrada.TabIndex = 50;
            this.radEntrada.TabStop = true;
            this.radEntrada.Text = "Entrada";
            this.radEntrada.UseVisualStyleBackColor = true;
            // 
            // radSalida
            // 
            this.radSalida.AutoSize = true;
            this.radSalida.Location = new System.Drawing.Point(12, 26);
            this.radSalida.Name = "radSalida";
            this.radSalida.Size = new System.Drawing.Size(69, 23);
            this.radSalida.TabIndex = 49;
            this.radSalida.TabStop = true;
            this.radSalida.Text = "Salida";
            this.radSalida.UseVisualStyleBackColor = true;
            this.radSalida.CheckedChanged += new System.EventHandler(this.radSalida_CheckedChanged);
            // 
            // btnGenerarReporte
            // 
            this.btnGenerarReporte.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnGenerarReporte.FlatAppearance.BorderSize = 2;
            this.btnGenerarReporte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGenerarReporte.ForeColor = System.Drawing.Color.White;
            this.btnGenerarReporte.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerarReporte.Image")));
            this.btnGenerarReporte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGenerarReporte.Location = new System.Drawing.Point(972, 18);
            this.btnGenerarReporte.Name = "btnGenerarReporte";
            this.btnGenerarReporte.Size = new System.Drawing.Size(133, 56);
            this.btnGenerarReporte.TabIndex = 35;
            this.btnGenerarReporte.Text = "Reporte";
            this.btnGenerarReporte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGenerarReporte.UseVisualStyleBackColor = true;
            this.btnGenerarReporte.Click += new System.EventHandler(this.btnGenerarReporte_Click);
            // 
            // lblTipoPersona
            // 
            this.lblTipoPersona.AutoSize = true;
            this.lblTipoPersona.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTipoPersona.ForeColor = System.Drawing.Color.White;
            this.lblTipoPersona.Location = new System.Drawing.Point(383, 62);
            this.lblTipoPersona.Name = "lblTipoPersona";
            this.lblTipoPersona.Size = new System.Drawing.Size(87, 19);
            this.lblTipoPersona.TabIndex = 44;
            this.lblTipoPersona.Text = "Proveedor:";
            // 
            // cboTipoPersona
            // 
            this.cboTipoPersona.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboTipoPersona.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboTipoPersona.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboTipoPersona.FormattingEnabled = true;
            this.cboTipoPersona.Location = new System.Drawing.Point(477, 57);
            this.cboTipoPersona.Margin = new System.Windows.Forms.Padding(4);
            this.cboTipoPersona.Name = "cboTipoPersona";
            this.cboTipoPersona.Size = new System.Drawing.Size(400, 27);
            this.cboTipoPersona.Sorted = true;
            this.cboTipoPersona.TabIndex = 5;
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(93, 54);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(93, 27);
            this.txtCantidad.TabIndex = 4;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(10, 57);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(77, 19);
            this.label7.TabIndex = 39;
            this.label7.Text = "Cantidad:";
            // 
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(477, 14);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(400, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 3;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.White;
            this.label12.Location = new System.Drawing.Point(404, 20);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(66, 19);
            this.label12.TabIndex = 36;
            this.label12.Text = "Modelo:";
            // 
            // dtpFechaRegistro
            // 
            this.dtpFechaRegistro.Location = new System.Drawing.Point(67, 20);
            this.dtpFechaRegistro.Name = "dtpFechaRegistro";
            this.dtpFechaRegistro.Size = new System.Drawing.Size(305, 27);
            this.dtpFechaRegistro.TabIndex = 1;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(6, 26);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(55, 19);
            this.label9.TabIndex = 6;
            this.label9.Text = "Fecha:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(410, 227);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(134, 19);
            this.label5.TabIndex = 34;
            this.label5.Text = "Modelo Cartucho:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(414, 249);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(237, 27);
            this.textBox1.TabIndex = 34;
            // 
            // Buscar
            // 
            this.Buscar.FlatAppearance.BorderColor = System.Drawing.Color.DarkOrange;
            this.Buscar.FlatAppearance.BorderSize = 2;
            this.Buscar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Buscar.ForeColor = System.Drawing.Color.White;
            this.Buscar.Image = ((System.Drawing.Image)(resources.GetObject("Buscar.Image")));
            this.Buscar.Location = new System.Drawing.Point(657, 249);
            this.Buscar.Name = "Buscar";
            this.Buscar.Size = new System.Drawing.Size(50, 27);
            this.Buscar.TabIndex = 34;
            this.Buscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Buscar.UseVisualStyleBackColor = true;
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(258, 227);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(117, 51);
            this.btnCancelar.TabIndex = 38;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // btnEliminar
            // 
            this.btnEliminar.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.btnEliminar.FlatAppearance.BorderSize = 2;
            this.btnEliminar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEliminar.ForeColor = System.Drawing.Color.White;
            this.btnEliminar.Image = ((System.Drawing.Image)(resources.GetObject("btnEliminar.Image")));
            this.btnEliminar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEliminar.Location = new System.Drawing.Point(135, 227);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(117, 51);
            this.btnEliminar.TabIndex = 37;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.btnGuardar.FlatAppearance.BorderSize = 2;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
            this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGuardar.Location = new System.Drawing.Point(12, 227);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 36;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // InventarioPartesRicoh
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1350, 977);
            this.Controls.Add(this.dtgInventario);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.Buscar);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "InventarioPartesRicoh";
            this.Text = "InventarioPartesRicoh";
            ((System.ComponentModel.ISupportInitialize)(this.dtgInventario)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabInventario.ResumeLayout(false);
            this.grpDatosInventario.ResumeLayout(false);
            this.grpDatosInventario.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabRegistros.ResumeLayout(false);
            this.grpDatosRegistro.ResumeLayout(false);
            this.grpDatosRegistro.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dtgInventario;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabInventario;
        private System.Windows.Forms.GroupBox grpDatosInventario;
        private System.Windows.Forms.RichTextBox rtxtDescripcion;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnImprimir;
        private System.Windows.Forms.TextBox txtNumeroParte;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage tabRegistros;
        private System.Windows.Forms.GroupBox grpDatosRegistro;
        private System.Windows.Forms.Button btnGenerarReporte;
        private System.Windows.Forms.Label lblTipoPersona;
        private System.Windows.Forms.ComboBox cboTipoPersona;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DateTimePicker dtpFechaRegistro;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Buscar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radEntrada;
        private System.Windows.Forms.RadioButton radSalida;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton radNuevo;
        private System.Windows.Forms.RadioButton radUsado;
        private System.Windows.Forms.CheckBox chkCliente;
    }
}