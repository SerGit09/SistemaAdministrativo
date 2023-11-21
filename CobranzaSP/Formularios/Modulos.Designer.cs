namespace CobranzaSP.Formularios
{
    partial class Modulos
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Modulos));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.grpDatosPartes = new System.Windows.Forms.GroupBox();
            this.cboModelo = new System.Windows.Forms.ComboBox();
            this.radUsado = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.radNuevo = new System.Windows.Forms.RadioButton();
            this.label6 = new System.Windows.Forms.Label();
            this.rtxtDescripcion = new System.Windows.Forms.RichTextBox();
            this.btnCancelarParte = new System.Windows.Forms.Button();
            this.btnEliminarParte = new System.Windows.Forms.Button();
            this.btnGuardarParte = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txtFolioSeleccionado = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.grpDatosHoja = new System.Windows.Forms.GroupBox();
            this.cboPersonaRecibido = new System.Windows.Forms.ComboBox();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.txtConcepto = new System.Windows.Forms.TextBox();
            this.dtpFecha = new System.Windows.Forms.DateTimePicker();
            this.txtFolio = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dtgFolioos = new System.Windows.Forms.DataGridView();
            this.dtgPartes = new System.Windows.Forms.DataGridView();
            this.btnReportes = new System.Windows.Forms.Button();
            this.grpDatosPartes.SuspendLayout();
            this.grpDatosHoja.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgFolioos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgPartes)).BeginInit();
            this.SuspendLayout();
            // 
            // grpDatosPartes
            // 
            this.grpDatosPartes.Controls.Add(this.cboModelo);
            this.grpDatosPartes.Controls.Add(this.radUsado);
            this.grpDatosPartes.Controls.Add(this.label8);
            this.grpDatosPartes.Controls.Add(this.txtCantidad);
            this.grpDatosPartes.Controls.Add(this.radNuevo);
            this.grpDatosPartes.Controls.Add(this.label6);
            this.grpDatosPartes.Controls.Add(this.rtxtDescripcion);
            this.grpDatosPartes.Controls.Add(this.btnCancelarParte);
            this.grpDatosPartes.Controls.Add(this.btnEliminarParte);
            this.grpDatosPartes.Controls.Add(this.btnGuardarParte);
            this.grpDatosPartes.Controls.Add(this.label5);
            this.grpDatosPartes.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.grpDatosPartes.Location = new System.Drawing.Point(25, 57);
            this.grpDatosPartes.Name = "grpDatosPartes";
            this.grpDatosPartes.Size = new System.Drawing.Size(456, 322);
            this.grpDatosPartes.TabIndex = 0;
            this.grpDatosPartes.TabStop = false;
            this.grpDatosPartes.Text = "Partes:";
            // 
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(79, 49);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(330, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 51;
            this.cboModelo.SelectedIndexChanged += new System.EventHandler(this.cboModelo_SelectedIndexChanged);
            // 
            // radUsado
            // 
            this.radUsado.AutoSize = true;
            this.radUsado.Location = new System.Drawing.Point(305, 213);
            this.radUsado.Name = "radUsado";
            this.radUsado.Size = new System.Drawing.Size(71, 23);
            this.radUsado.TabIndex = 50;
            this.radUsado.TabStop = true;
            this.radUsado.Text = "Usado";
            this.radUsado.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(7, 215);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 19);
            this.label8.TabIndex = 37;
            this.label8.Text = "Cantidad:";
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(90, 212);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(85, 27);
            this.txtCantidad.TabIndex = 41;
            // 
            // radNuevo
            // 
            this.radNuevo.AutoSize = true;
            this.radNuevo.Location = new System.Drawing.Point(218, 213);
            this.radNuevo.Name = "radNuevo";
            this.radNuevo.Size = new System.Drawing.Size(72, 23);
            this.radNuevo.TabIndex = 49;
            this.radNuevo.TabStop = true;
            this.radNuevo.Text = "Nuevo";
            this.radNuevo.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(7, 81);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(96, 19);
            this.label6.TabIndex = 39;
            this.label6.Text = "Descripcion:";
            // 
            // rtxtDescripcion
            // 
            this.rtxtDescripcion.Location = new System.Drawing.Point(11, 103);
            this.rtxtDescripcion.Name = "rtxtDescripcion";
            this.rtxtDescripcion.Size = new System.Drawing.Size(410, 96);
            this.rtxtDescripcion.TabIndex = 48;
            this.rtxtDescripcion.Text = "";
            // 
            // btnCancelarParte
            // 
            this.btnCancelarParte.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelarParte.FlatAppearance.BorderSize = 2;
            this.btnCancelarParte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelarParte.ForeColor = System.Drawing.Color.White;
            this.btnCancelarParte.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelarParte.Image")));
            this.btnCancelarParte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelarParte.Location = new System.Drawing.Point(271, 262);
            this.btnCancelarParte.Name = "btnCancelarParte";
            this.btnCancelarParte.Size = new System.Drawing.Size(117, 51);
            this.btnCancelarParte.TabIndex = 47;
            this.btnCancelarParte.Text = "Cancelar";
            this.btnCancelarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCancelarParte.UseVisualStyleBackColor = true;
            this.btnCancelarParte.Click += new System.EventHandler(this.btnCancelarParte_Click);
            // 
            // btnEliminarParte
            // 
            this.btnEliminarParte.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.btnEliminarParte.FlatAppearance.BorderSize = 2;
            this.btnEliminarParte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEliminarParte.ForeColor = System.Drawing.Color.White;
            this.btnEliminarParte.Image = ((System.Drawing.Image)(resources.GetObject("btnEliminarParte.Image")));
            this.btnEliminarParte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEliminarParte.Location = new System.Drawing.Point(148, 262);
            this.btnEliminarParte.Name = "btnEliminarParte";
            this.btnEliminarParte.Size = new System.Drawing.Size(117, 51);
            this.btnEliminarParte.TabIndex = 46;
            this.btnEliminarParte.Text = "Eliminar";
            this.btnEliminarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEliminarParte.UseVisualStyleBackColor = true;
            this.btnEliminarParte.Click += new System.EventHandler(this.btnEliminarParte_Click);
            // 
            // btnGuardarParte
            // 
            this.btnGuardarParte.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.btnGuardarParte.FlatAppearance.BorderSize = 2;
            this.btnGuardarParte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardarParte.ForeColor = System.Drawing.Color.White;
            this.btnGuardarParte.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardarParte.Image")));
            this.btnGuardarParte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGuardarParte.Location = new System.Drawing.Point(25, 262);
            this.btnGuardarParte.Name = "btnGuardarParte";
            this.btnGuardarParte.Size = new System.Drawing.Size(117, 51);
            this.btnGuardarParte.TabIndex = 45;
            this.btnGuardarParte.Text = "Guardar";
            this.btnGuardarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardarParte.UseVisualStyleBackColor = true;
            this.btnGuardarParte.Click += new System.EventHandler(this.btnGuardarParte_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(7, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 19);
            this.label5.TabIndex = 40;
            this.label5.Text = "Modelo:";
            // 
            // txtFolioSeleccionado
            // 
            this.txtFolioSeleccionado.Location = new System.Drawing.Point(76, 14);
            this.txtFolioSeleccionado.Name = "txtFolioSeleccionado";
            this.txtFolioSeleccionado.Size = new System.Drawing.Size(85, 27);
            this.txtFolioSeleccionado.TabIndex = 53;
            this.txtFolioSeleccionado.TextChanged += new System.EventHandler(this.txtFolioSeleccionado_TextChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(21, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(49, 19);
            this.label7.TabIndex = 52;
            this.label7.Text = "Folio:";
            // 
            // grpDatosHoja
            // 
            this.grpDatosHoja.Controls.Add(this.cboPersonaRecibido);
            this.grpDatosHoja.Controls.Add(this.btnCancelar);
            this.grpDatosHoja.Controls.Add(this.btnEliminar);
            this.grpDatosHoja.Controls.Add(this.btnGuardar);
            this.grpDatosHoja.Controls.Add(this.txtConcepto);
            this.grpDatosHoja.Controls.Add(this.dtpFecha);
            this.grpDatosHoja.Controls.Add(this.txtFolio);
            this.grpDatosHoja.Controls.Add(this.label4);
            this.grpDatosHoja.Controls.Add(this.label2);
            this.grpDatosHoja.Controls.Add(this.label1);
            this.grpDatosHoja.Controls.Add(this.label3);
            this.grpDatosHoja.ForeColor = System.Drawing.SystemColors.ControlLight;
            this.grpDatosHoja.Location = new System.Drawing.Point(702, 12);
            this.grpDatosHoja.Name = "grpDatosHoja";
            this.grpDatosHoja.Size = new System.Drawing.Size(421, 256);
            this.grpDatosHoja.TabIndex = 1;
            this.grpDatosHoja.TabStop = false;
            this.grpDatosHoja.Text = "Datos Folio:";
            // 
            // cboPersonaRecibido
            // 
            this.cboPersonaRecibido.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboPersonaRecibido.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboPersonaRecibido.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPersonaRecibido.FormattingEnabled = true;
            this.cboPersonaRecibido.Location = new System.Drawing.Point(79, 103);
            this.cboPersonaRecibido.Margin = new System.Windows.Forms.Padding(4);
            this.cboPersonaRecibido.Name = "cboPersonaRecibido";
            this.cboPersonaRecibido.Size = new System.Drawing.Size(336, 27);
            this.cboPersonaRecibido.Sorted = true;
            this.cboPersonaRecibido.TabIndex = 3;
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(276, 183);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(117, 51);
            this.btnCancelar.TabIndex = 36;
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
            this.btnEliminar.Location = new System.Drawing.Point(153, 183);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(117, 51);
            this.btnEliminar.TabIndex = 35;
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
            this.btnGuardar.Location = new System.Drawing.Point(30, 183);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 34;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // txtConcepto
            // 
            this.txtConcepto.Location = new System.Drawing.Point(93, 139);
            this.txtConcepto.Name = "txtConcepto";
            this.txtConcepto.Size = new System.Drawing.Size(322, 27);
            this.txtConcepto.TabIndex = 13;
            // 
            // dtpFecha
            // 
            this.dtpFecha.Location = new System.Drawing.Point(67, 63);
            this.dtpFecha.Name = "dtpFecha";
            this.dtpFecha.Size = new System.Drawing.Size(348, 27);
            this.dtpFecha.TabIndex = 11;
            // 
            // txtFolio
            // 
            this.txtFolio.Location = new System.Drawing.Point(61, 26);
            this.txtFolio.Name = "txtFolio";
            this.txtFolio.Size = new System.Drawing.Size(354, 27);
            this.txtFolio.TabIndex = 10;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(6, 142);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(81, 19);
            this.label4.TabIndex = 9;
            this.label4.Text = "Concepto:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(6, 106);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 19);
            this.label2.TabIndex = 8;
            this.label2.Text = "Recibio:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 19);
            this.label1.TabIndex = 7;
            this.label1.Text = "Fecha:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 32);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(49, 19);
            this.label3.TabIndex = 6;
            this.label3.Text = "Folio:";
            // 
            // dtgFolioos
            // 
            this.dtgFolioos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgFolioos.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgFolioos.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgFolioos.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgFolioos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgFolioos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgFolioos.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgFolioos.EnableHeadersVisualStyles = false;
            this.dtgFolioos.Location = new System.Drawing.Point(702, 274);
            this.dtgFolioos.Name = "dtgFolioos";
            this.dtgFolioos.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgFolioos.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgFolioos.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgFolioos.RowTemplate.Height = 25;
            this.dtgFolioos.Size = new System.Drawing.Size(645, 682);
            this.dtgFolioos.TabIndex = 33;
            this.dtgFolioos.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgFolioos_CellClick);
            // 
            // dtgPartes
            // 
            this.dtgPartes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgPartes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgPartes.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgPartes.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dtgPartes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgPartes.DefaultCellStyle = dataGridViewCellStyle6;
            this.dtgPartes.EnableHeadersVisualStyles = false;
            this.dtgPartes.Location = new System.Drawing.Point(12, 385);
            this.dtgPartes.Name = "dtgPartes";
            this.dtgPartes.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.RowHeadersDefaultCellStyle = dataGridViewCellStyle7;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.RowsDefaultCellStyle = dataGridViewCellStyle8;
            this.dtgPartes.RowTemplate.Height = 25;
            this.dtgPartes.Size = new System.Drawing.Size(670, 571);
            this.dtgPartes.TabIndex = 34;
            this.dtgPartes.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgPartes_CellClick);
            // 
            // btnReportes
            // 
            this.btnReportes.FlatAppearance.BorderColor = System.Drawing.Color.LightSlateGray;
            this.btnReportes.FlatAppearance.BorderSize = 2;
            this.btnReportes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReportes.ForeColor = System.Drawing.Color.White;
            this.btnReportes.Image = ((System.Drawing.Image)(resources.GetObject("btnReportes.Image")));
            this.btnReportes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnReportes.Location = new System.Drawing.Point(1199, 207);
            this.btnReportes.Name = "btnReportes";
            this.btnReportes.Size = new System.Drawing.Size(139, 61);
            this.btnReportes.TabIndex = 37;
            this.btnReportes.Text = "Reportes";
            this.btnReportes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnReportes.UseVisualStyleBackColor = true;
            this.btnReportes.Click += new System.EventHandler(this.btnReportes_Click);
            // 
            // Modulos
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1350, 977);
            this.Controls.Add(this.btnReportes);
            this.Controls.Add(this.txtFolioSeleccionado);
            this.Controls.Add(this.dtgPartes);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.dtgFolioos);
            this.Controls.Add(this.grpDatosHoja);
            this.Controls.Add(this.grpDatosPartes);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Modulos";
            this.Text = "Modulos";
            this.grpDatosPartes.ResumeLayout(false);
            this.grpDatosPartes.PerformLayout();
            this.grpDatosHoja.ResumeLayout(false);
            this.grpDatosHoja.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgFolioos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgPartes)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grpDatosPartes;
        private System.Windows.Forms.GroupBox grpDatosHoja;
        private System.Windows.Forms.TextBox txtConcepto;
        private System.Windows.Forms.DateTimePicker dtpFecha;
        private System.Windows.Forms.TextBox txtFolio;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.RadioButton radUsado;
        private System.Windows.Forms.RadioButton radNuevo;
        private System.Windows.Forms.RichTextBox rtxtDescripcion;
        private System.Windows.Forms.Button btnCancelarParte;
        private System.Windows.Forms.Button btnEliminarParte;
        private System.Windows.Forms.Button btnGuardarParte;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.ComboBox cboPersonaRecibido;
        private System.Windows.Forms.TextBox txtFolioSeleccionado;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridView dtgFolioos;
        private System.Windows.Forms.DataGridView dtgPartes;
        private System.Windows.Forms.Button btnReportes;
    }
}