namespace CobranzaSP.Formularios
{
    partial class Servicios
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Servicios));
            this.label1 = new System.Windows.Forms.Label();
            this.grpDatos = new System.Windows.Forms.GroupBox();
            this.txtModelo = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.chkFusor = new System.Windows.Forms.CheckBox();
            this.cboFusorRetirado = new System.Windows.Forms.ComboBox();
            this.cboFusor = new System.Windows.Forms.ComboBox();
            this.rtxtFallas = new System.Windows.Forms.RichTextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.rtxtServicio = new System.Windows.Forms.RichTextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.lblFusorRetirado = new System.Windows.Forms.Label();
            this.lblFusorInstalado = new System.Windows.Forms.Label();
            this.cboTecnico = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtNumeroSerie = new System.Windows.Forms.TextBox();
            this.chkSerie = new System.Windows.Forms.CheckBox();
            this.dtpFecha = new System.Windows.Forms.DateTimePicker();
            this.label7 = new System.Windows.Forms.Label();
            this.txtContador = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.cboNumeroSerie = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboModelos = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboMarca = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboClientes = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNumeroFolio = new System.Windows.Forms.TextBox();
            this.cboMostrar = new System.Windows.Forms.ComboBox();
            this.dtgServicios = new System.Windows.Forms.DataGridView();
            this.erServicios = new System.Windows.Forms.ErrorProvider(this.components);
            this.txtBusqueda = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.btnBusqueda = new System.Windows.Forms.Button();
            this.btnreportes2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.grpDatos.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgServicios)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.erServicios)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Numero Folio:";
            // 
            // grpDatos
            // 
            this.grpDatos.Controls.Add(this.btnreportes2);
            this.grpDatos.Controls.Add(this.txtModelo);
            this.grpDatos.Controls.Add(this.checkBox1);
            this.grpDatos.Controls.Add(this.button1);
            this.grpDatos.Controls.Add(this.chkFusor);
            this.grpDatos.Controls.Add(this.btnCancelar);
            this.grpDatos.Controls.Add(this.btnEliminar);
            this.grpDatos.Controls.Add(this.btnGuardar);
            this.grpDatos.Controls.Add(this.cboFusorRetirado);
            this.grpDatos.Controls.Add(this.cboFusor);
            this.grpDatos.Controls.Add(this.rtxtFallas);
            this.grpDatos.Controls.Add(this.label12);
            this.grpDatos.Controls.Add(this.rtxtServicio);
            this.grpDatos.Controls.Add(this.label11);
            this.grpDatos.Controls.Add(this.lblFusorRetirado);
            this.grpDatos.Controls.Add(this.lblFusorInstalado);
            this.grpDatos.Controls.Add(this.cboTecnico);
            this.grpDatos.Controls.Add(this.label8);
            this.grpDatos.Controls.Add(this.txtNumeroSerie);
            this.grpDatos.Controls.Add(this.chkSerie);
            this.grpDatos.Controls.Add(this.dtpFecha);
            this.grpDatos.Controls.Add(this.label7);
            this.grpDatos.Controls.Add(this.txtContador);
            this.grpDatos.Controls.Add(this.label6);
            this.grpDatos.Controls.Add(this.cboNumeroSerie);
            this.grpDatos.Controls.Add(this.label5);
            this.grpDatos.Controls.Add(this.cboModelos);
            this.grpDatos.Controls.Add(this.label4);
            this.grpDatos.Controls.Add(this.cboMarca);
            this.grpDatos.Controls.Add(this.label3);
            this.grpDatos.Controls.Add(this.cboClientes);
            this.grpDatos.Controls.Add(this.label2);
            this.grpDatos.Controls.Add(this.txtNumeroFolio);
            this.grpDatos.Controls.Add(this.label1);
            this.grpDatos.Location = new System.Drawing.Point(12, 12);
            this.grpDatos.Name = "grpDatos";
            this.grpDatos.Size = new System.Drawing.Size(1318, 331);
            this.grpDatos.TabIndex = 1;
            this.grpDatos.TabStop = false;
            this.grpDatos.Text = "Datos";
            // 
            // txtModelo
            // 
            this.txtModelo.Location = new System.Drawing.Point(267, 80);
            this.txtModelo.Name = "txtModelo";
            this.txtModelo.Size = new System.Drawing.Size(166, 26);
            this.txtModelo.TabIndex = 33;
            this.txtModelo.Visible = false;
            this.txtModelo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtModelo_KeyPress);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox1.ForeColor = System.Drawing.Color.White;
            this.checkBox1.Location = new System.Drawing.Point(440, 81);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(73, 23);
            this.checkBox1.TabIndex = 32;
            this.checkBox1.Text = "Nuevo";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // chkFusor
            // 
            this.chkFusor.AutoSize = true;
            this.chkFusor.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkFusor.ForeColor = System.Drawing.Color.White;
            this.chkFusor.Location = new System.Drawing.Point(331, 225);
            this.chkFusor.Name = "chkFusor";
            this.chkFusor.Size = new System.Drawing.Size(134, 23);
            this.chkFusor.TabIndex = 11;
            this.chkFusor.Text = "¿Uso de fusor?";
            this.chkFusor.UseVisualStyleBackColor = true;
            this.chkFusor.CheckedChanged += new System.EventHandler(this.chkFusor_CheckedChanged);
            // 
            // cboFusorRetirado
            // 
            this.cboFusorRetirado.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboFusorRetirado.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboFusorRetirado.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFusorRetirado.FormattingEnabled = true;
            this.cboFusorRetirado.Location = new System.Drawing.Point(651, 46);
            this.cboFusorRetirado.Margin = new System.Windows.Forms.Padding(4);
            this.cboFusorRetirado.Name = "cboFusorRetirado";
            this.cboFusorRetirado.Size = new System.Drawing.Size(412, 26);
            this.cboFusorRetirado.Sorted = true;
            this.cboFusorRetirado.TabIndex = 13;
            this.cboFusorRetirado.Visible = false;
            // 
            // cboFusor
            // 
            this.cboFusor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboFusor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboFusor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFusor.FormattingEnabled = true;
            this.cboFusor.Location = new System.Drawing.Point(651, 15);
            this.cboFusor.Margin = new System.Windows.Forms.Padding(4);
            this.cboFusor.Name = "cboFusor";
            this.cboFusor.Size = new System.Drawing.Size(412, 26);
            this.cboFusor.Sorted = true;
            this.cboFusor.TabIndex = 12;
            this.cboFusor.Visible = false;
            // 
            // rtxtFallas
            // 
            this.rtxtFallas.Location = new System.Drawing.Point(524, 102);
            this.rtxtFallas.Name = "rtxtFallas";
            this.rtxtFallas.Size = new System.Drawing.Size(539, 96);
            this.rtxtFallas.TabIndex = 14;
            this.rtxtFallas.Text = "";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.White;
            this.label12.Location = new System.Drawing.Point(519, 80);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(124, 19);
            this.label12.TabIndex = 23;
            this.label12.Text = "Falla Reportada:";
            // 
            // rtxtServicio
            // 
            this.rtxtServicio.Location = new System.Drawing.Point(523, 223);
            this.rtxtServicio.Name = "rtxtServicio";
            this.rtxtServicio.Size = new System.Drawing.Size(539, 96);
            this.rtxtServicio.TabIndex = 15;
            this.rtxtServicio.Text = "";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.White;
            this.label11.Location = new System.Drawing.Point(520, 201);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(142, 19);
            this.label11.TabIndex = 21;
            this.label11.Text = "Servicio Realizado:";
            // 
            // lblFusorRetirado
            // 
            this.lblFusorRetirado.AutoSize = true;
            this.lblFusorRetirado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFusorRetirado.ForeColor = System.Drawing.Color.White;
            this.lblFusorRetirado.Location = new System.Drawing.Point(520, 48);
            this.lblFusorRetirado.Name = "lblFusorRetirado";
            this.lblFusorRetirado.Size = new System.Drawing.Size(118, 19);
            this.lblFusorRetirado.TabIndex = 20;
            this.lblFusorRetirado.Text = "Fusor Retirado:";
            this.lblFusorRetirado.Visible = false;
            // 
            // lblFusorInstalado
            // 
            this.lblFusorInstalado.AutoSize = true;
            this.lblFusorInstalado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFusorInstalado.ForeColor = System.Drawing.Color.White;
            this.lblFusorInstalado.Location = new System.Drawing.Point(520, 17);
            this.lblFusorInstalado.Name = "lblFusorInstalado";
            this.lblFusorInstalado.Size = new System.Drawing.Size(124, 19);
            this.lblFusorInstalado.TabIndex = 19;
            this.lblFusorInstalado.Text = "Fusor Instalado:";
            this.lblFusorInstalado.Visible = false;
            // 
            // cboTecnico
            // 
            this.cboTecnico.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboTecnico.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboTecnico.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboTecnico.FormattingEnabled = true;
            this.cboTecnico.Location = new System.Drawing.Point(82, 225);
            this.cboTecnico.Margin = new System.Windows.Forms.Padding(4);
            this.cboTecnico.Name = "cboTecnico";
            this.cboTecnico.Size = new System.Drawing.Size(242, 26);
            this.cboTecnico.Sorted = true;
            this.cboTecnico.TabIndex = 10;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(6, 232);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(69, 19);
            this.label8.TabIndex = 17;
            this.label8.Text = "Técnico:";
            // 
            // txtNumeroSerie
            // 
            this.txtNumeroSerie.Location = new System.Drawing.Point(69, 114);
            this.txtNumeroSerie.Name = "txtNumeroSerie";
            this.txtNumeroSerie.Size = new System.Drawing.Size(255, 26);
            this.txtNumeroSerie.TabIndex = 6;
            this.txtNumeroSerie.Visible = false;
            this.txtNumeroSerie.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumeroSerie_KeyPress);
            // 
            // chkSerie
            // 
            this.chkSerie.AutoSize = true;
            this.chkSerie.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSerie.ForeColor = System.Drawing.Color.White;
            this.chkSerie.Location = new System.Drawing.Point(331, 117);
            this.chkSerie.Name = "chkSerie";
            this.chkSerie.Size = new System.Drawing.Size(132, 23);
            this.chkSerie.TabIndex = 5;
            this.chkSerie.Text = "Serie Diferente";
            this.chkSerie.UseVisualStyleBackColor = true;
            this.chkSerie.CheckedChanged += new System.EventHandler(this.chkSerie_CheckedChanged);
            // 
            // dtpFecha
            // 
            this.dtpFecha.Location = new System.Drawing.Point(69, 188);
            this.dtpFecha.Name = "dtpFecha";
            this.dtpFecha.Size = new System.Drawing.Size(412, 26);
            this.dtpFecha.TabIndex = 9;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(6, 195);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 19);
            this.label7.TabIndex = 13;
            this.label7.Text = "Fecha:";
            // 
            // txtContador
            // 
            this.txtContador.Location = new System.Drawing.Point(92, 149);
            this.txtContador.Name = "txtContador";
            this.txtContador.Size = new System.Drawing.Size(232, 26);
            this.txtContador.TabIndex = 8;
            this.txtContador.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtContador_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(6, 157);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 19);
            this.label6.TabIndex = 10;
            this.label6.Text = "Contador:";
            // 
            // cboNumeroSerie
            // 
            this.cboNumeroSerie.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboNumeroSerie.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboNumeroSerie.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboNumeroSerie.FormattingEnabled = true;
            this.cboNumeroSerie.Location = new System.Drawing.Point(69, 116);
            this.cboNumeroSerie.Margin = new System.Windows.Forms.Padding(4);
            this.cboNumeroSerie.Name = "cboNumeroSerie";
            this.cboNumeroSerie.Size = new System.Drawing.Size(255, 26);
            this.cboNumeroSerie.Sorted = true;
            this.cboNumeroSerie.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(6, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 19);
            this.label5.TabIndex = 8;
            this.label5.Text = "Serie:";
            // 
            // cboModelos
            // 
            this.cboModelos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelos.FormattingEnabled = true;
            this.cboModelos.Location = new System.Drawing.Point(267, 80);
            this.cboModelos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelos.Name = "cboModelos";
            this.cboModelos.Size = new System.Drawing.Size(166, 26);
            this.cboModelos.Sorted = true;
            this.cboModelos.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(194, 84);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 19);
            this.label4.TabIndex = 6;
            this.label4.Text = "Modelo:";
            // 
            // cboMarca
            // 
            this.cboMarca.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarca.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarca.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarca.FormattingEnabled = true;
            this.cboMarca.Location = new System.Drawing.Point(69, 82);
            this.cboMarca.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarca.Name = "cboMarca";
            this.cboMarca.Size = new System.Drawing.Size(118, 26);
            this.cboMarca.Sorted = true;
            this.cboMarca.TabIndex = 3;
            this.cboMarca.SelectedIndexChanged += new System.EventHandler(this.cboMarca_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 87);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 19);
            this.label3.TabIndex = 4;
            this.label3.Text = "Marca:";
            // 
            // cboClientes
            // 
            this.cboClientes.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClientes.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClientes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClientes.FormattingEnabled = true;
            this.cboClientes.Location = new System.Drawing.Point(69, 48);
            this.cboClientes.Margin = new System.Windows.Forms.Padding(4);
            this.cboClientes.Name = "cboClientes";
            this.cboClientes.Size = new System.Drawing.Size(412, 26);
            this.cboClientes.Sorted = true;
            this.cboClientes.TabIndex = 2;
            this.cboClientes.SelectedIndexChanged += new System.EventHandler(this.cboClientes_SelectedIndexChanged);
            this.cboClientes.SelectionChangeCommitted += new System.EventHandler(this.cboClientes_SelectionChangeCommitted);
            this.cboClientes.DropDownClosed += new System.EventHandler(this.cboClientes_DropDownClosed);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(6, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 19);
            this.label2.TabIndex = 2;
            this.label2.Text = "Cliente:";
            // 
            // txtNumeroFolio
            // 
            this.txtNumeroFolio.Location = new System.Drawing.Point(123, 15);
            this.txtNumeroFolio.Name = "txtNumeroFolio";
            this.txtNumeroFolio.Size = new System.Drawing.Size(358, 26);
            this.txtNumeroFolio.TabIndex = 1;
            this.txtNumeroFolio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumeroFolio_KeyPress);
            // 
            // cboMostrar
            // 
            this.cboMostrar.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMostrar.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMostrar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMostrar.FormattingEnabled = true;
            this.cboMostrar.Location = new System.Drawing.Point(577, 349);
            this.cboMostrar.Margin = new System.Windows.Forms.Padding(4);
            this.cboMostrar.Name = "cboMostrar";
            this.cboMostrar.Size = new System.Drawing.Size(412, 26);
            this.cboMostrar.Sorted = true;
            this.cboMostrar.TabIndex = 31;
            this.cboMostrar.SelectedIndexChanged += new System.EventHandler(this.cboMostrar_SelectedIndexChanged);
            // 
            // dtgServicios
            // 
            this.dtgServicios.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgServicios.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgServicios.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgServicios.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgServicios.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgServicios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgServicios.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgServicios.EnableHeadersVisualStyles = false;
            this.dtgServicios.Location = new System.Drawing.Point(12, 383);
            this.dtgServicios.Name = "dtgServicios";
            this.dtgServicios.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgServicios.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgServicios.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgServicios.RowTemplate.Height = 25;
            this.dtgServicios.Size = new System.Drawing.Size(1326, 517);
            this.dtgServicios.TabIndex = 32;
            this.dtgServicios.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgServicios_CellClick);
            // 
            // erServicios
            // 
            this.erServicios.ContainerControl = this;
            // 
            // txtBusqueda
            // 
            this.txtBusqueda.Location = new System.Drawing.Point(129, 349);
            this.txtBusqueda.Name = "txtBusqueda";
            this.txtBusqueda.Size = new System.Drawing.Size(232, 26);
            this.txtBusqueda.TabIndex = 34;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(12, 353);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(111, 19);
            this.label9.TabIndex = 34;
            this.label9.Text = "Numero Folio:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.White;
            this.label10.Location = new System.Drawing.Point(473, 351);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(97, 19);
            this.label10.TabIndex = 36;
            this.label10.Text = "Mostrar por:";
            // 
            // btnBusqueda
            // 
            this.btnBusqueda.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.btnBusqueda.Image = ((System.Drawing.Image)(resources.GetObject("btnBusqueda.Image")));
            this.btnBusqueda.Location = new System.Drawing.Point(367, 344);
            this.btnBusqueda.Name = "btnBusqueda";
            this.btnBusqueda.Size = new System.Drawing.Size(66, 35);
            this.btnBusqueda.TabIndex = 35;
            this.btnBusqueda.UseVisualStyleBackColor = false;
            this.btnBusqueda.Click += new System.EventHandler(this.btnBusqueda_Click);
            // 
            // btnreportes2
            // 
            this.btnreportes2.FlatAppearance.BorderColor = System.Drawing.Color.Coral;
            this.btnreportes2.FlatAppearance.BorderSize = 2;
            this.btnreportes2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnreportes2.ForeColor = System.Drawing.Color.White;
            this.btnreportes2.Image = ((System.Drawing.Image)(resources.GetObject("btnreportes2.Image")));
            this.btnreportes2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnreportes2.Location = new System.Drawing.Point(1177, 210);
            this.btnreportes2.Name = "btnreportes2";
            this.btnreportes2.Size = new System.Drawing.Size(135, 53);
            this.btnreportes2.TabIndex = 34;
            this.btnreportes2.Text = "Reportes 2";
            this.btnreportes2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnreportes2.UseVisualStyleBackColor = true;
            this.btnreportes2.Click += new System.EventHandler(this.btnreportes2_Click);
            // 
            // button1
            // 
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.OrangeRed;
            this.button1.FlatAppearance.BorderSize = 2;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(1177, 270);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(135, 51);
            this.button1.TabIndex = 31;
            this.button1.Text = "Reportes";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(256, 270);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(117, 51);
            this.btnCancelar.TabIndex = 30;
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
            this.btnEliminar.Location = new System.Drawing.Point(133, 270);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(117, 51);
            this.btnEliminar.TabIndex = 29;
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
            this.btnGuardar.Location = new System.Drawing.Point(10, 270);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 16;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // Servicios
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1350, 977);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.btnBusqueda);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtBusqueda);
            this.Controls.Add(this.dtgServicios);
            this.Controls.Add(this.cboMostrar);
            this.Controls.Add(this.grpDatos);
            this.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Servicios";
            this.grpDatos.ResumeLayout(false);
            this.grpDatos.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgServicios)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.erServicios)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpDatos;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNumeroFolio;
        private System.Windows.Forms.ComboBox cboModelos;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboMarca;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboClientes;
        private System.Windows.Forms.ComboBox cboNumeroSerie;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.CheckBox chkSerie;
        private System.Windows.Forms.DateTimePicker dtpFecha;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtContador;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cboTecnico;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.RichTextBox rtxtFallas;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.RichTextBox rtxtServicio;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lblFusorRetirado;
        private System.Windows.Forms.Label lblFusorInstalado;
        private System.Windows.Forms.ComboBox cboFusorRetirado;
        private System.Windows.Forms.ComboBox cboFusor;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.CheckBox chkFusor;
        private System.Windows.Forms.ComboBox cboMostrar;
        private System.Windows.Forms.DataGridView dtgServicios;
        private System.Windows.Forms.ErrorProvider erServicios;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TextBox txtBusqueda;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnBusqueda;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button btnreportes2;
        private System.Windows.Forms.TextBox txtNumeroSerie;
        private System.Windows.Forms.TextBox txtModelo;
    }
}