namespace CobranzaSP.Formularios
{
    partial class Inventario
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Inventario));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabInventario = new System.Windows.Forms.TabPage();
            this.grpDatosInventario = new System.Windows.Forms.GroupBox();
            this.btnAbrirReporteExistencias = new System.Windows.Forms.Button();
            this.txtPrecio = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.cboMarcaSeleccionada = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dtpFecha = new System.Windows.Forms.DateTimePicker();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.txtOficina = new System.Windows.Forms.TextBox();
            this.Buscar = new System.Windows.Forms.Button();
            this.txtModelo = new System.Windows.Forms.TextBox();
            this.cboMarca = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.grpDatosRegistro = new System.Windows.Forms.GroupBox();
            this.label15 = new System.Windows.Forms.Label();
            this.radGarantia = new System.Windows.Forms.RadioButton();
            this.radSalida = new System.Windows.Forms.RadioButton();
            this.radEntrada = new System.Windows.Forms.RadioButton();
            this.btnAbrirReportes = new System.Windows.Forms.Button();
            this.cboFusor = new System.Windows.Forms.ComboBox();
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.cboClientes = new System.Windows.Forms.ComboBox();
            this.lblSerieP = new System.Windows.Forms.Label();
            this.cboModelos = new System.Windows.Forms.ComboBox();
            this.label20 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.cboMarcas = new System.Windows.Forms.ComboBox();
            this.dtpFechaRegistro = new System.Windows.Forms.DateTimePicker();
            this.label22 = new System.Windows.Forms.Label();
            this.dtgCartuchos = new System.Windows.Forms.DataGridView();
            this.erInventario = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnGarantias = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabInventario.SuspendLayout();
            this.grpDatosInventario.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.grpDatosRegistro.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgCartuchos)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.erInventario)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabInventario);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1141, 199);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabInventario
            // 
            this.tabInventario.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.tabInventario.Controls.Add(this.grpDatosInventario);
            this.tabInventario.Location = new System.Drawing.Point(4, 28);
            this.tabInventario.Name = "tabInventario";
            this.tabInventario.Padding = new System.Windows.Forms.Padding(3);
            this.tabInventario.Size = new System.Drawing.Size(1133, 167);
            this.tabInventario.TabIndex = 0;
            this.tabInventario.Text = "Inventario";
            // 
            // grpDatosInventario
            // 
            this.grpDatosInventario.Controls.Add(this.btnAbrirReporteExistencias);
            this.grpDatosInventario.Controls.Add(this.txtPrecio);
            this.grpDatosInventario.Controls.Add(this.label8);
            this.grpDatosInventario.Controls.Add(this.btnImprimir);
            this.grpDatosInventario.Controls.Add(this.cboMarcaSeleccionada);
            this.grpDatosInventario.Controls.Add(this.label5);
            this.grpDatosInventario.Controls.Add(this.label6);
            this.grpDatosInventario.Controls.Add(this.dtpFecha);
            this.grpDatosInventario.Controls.Add(this.textBox1);
            this.grpDatosInventario.Controls.Add(this.txtOficina);
            this.grpDatosInventario.Controls.Add(this.Buscar);
            this.grpDatosInventario.Controls.Add(this.txtModelo);
            this.grpDatosInventario.Controls.Add(this.cboMarca);
            this.grpDatosInventario.Controls.Add(this.label4);
            this.grpDatosInventario.Controls.Add(this.label3);
            this.grpDatosInventario.Controls.Add(this.label1);
            this.grpDatosInventario.Controls.Add(this.label2);
            this.grpDatosInventario.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpDatosInventario.Location = new System.Drawing.Point(6, 6);
            this.grpDatosInventario.Name = "grpDatosInventario";
            this.grpDatosInventario.Size = new System.Drawing.Size(1111, 156);
            this.grpDatosInventario.TabIndex = 0;
            this.grpDatosInventario.TabStop = false;
            this.grpDatosInventario.Text = "Datos";
            // 
            // btnAbrirReporteExistencias
            // 
            this.btnAbrirReporteExistencias.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.btnAbrirReporteExistencias.FlatAppearance.BorderColor = System.Drawing.Color.Crimson;
            this.btnAbrirReporteExistencias.FlatAppearance.BorderSize = 2;
            this.btnAbrirReporteExistencias.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAbrirReporteExistencias.ForeColor = System.Drawing.Color.White;
            this.btnAbrirReporteExistencias.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrirReporteExistencias.Image")));
            this.btnAbrirReporteExistencias.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAbrirReporteExistencias.Location = new System.Drawing.Point(958, 26);
            this.btnAbrirReporteExistencias.Name = "btnAbrirReporteExistencias";
            this.btnAbrirReporteExistencias.Size = new System.Drawing.Size(147, 54);
            this.btnAbrirReporteExistencias.TabIndex = 35;
            this.btnAbrirReporteExistencias.Text = "Reporte Existencias";
            this.btnAbrirReporteExistencias.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAbrirReporteExistencias.UseVisualStyleBackColor = false;
            this.btnAbrirReporteExistencias.Click += new System.EventHandler(this.btnAbrirReporteExistencias_Click);
            // 
            // txtPrecio
            // 
            this.txtPrecio.Location = new System.Drawing.Point(383, 19);
            this.txtPrecio.Name = "txtPrecio";
            this.txtPrecio.Size = new System.Drawing.Size(92, 27);
            this.txtPrecio.TabIndex = 38;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(313, 23);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(58, 19);
            this.label8.TabIndex = 37;
            this.label8.Text = "Precio:";
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
            // cboMarcaSeleccionada
            // 
            this.cboMarcaSeleccionada.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarcaSeleccionada.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarcaSeleccionada.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarcaSeleccionada.FormattingEnabled = true;
            this.cboMarcaSeleccionada.Location = new System.Drawing.Point(706, 115);
            this.cboMarcaSeleccionada.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarcaSeleccionada.Name = "cboMarcaSeleccionada";
            this.cboMarcaSeleccionada.Size = new System.Drawing.Size(237, 27);
            this.cboMarcaSeleccionada.Sorted = true;
            this.cboMarcaSeleccionada.TabIndex = 34;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(328, 99);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(134, 19);
            this.label5.TabIndex = 34;
            this.label5.Text = "Modelo Cartucho:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(702, 92);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(114, 19);
            this.label6.TabIndex = 35;
            this.label6.Text = "Mostrar Marca:";
            // 
            // dtpFecha
            // 
            this.dtpFecha.Location = new System.Drawing.Point(373, 60);
            this.dtpFecha.Name = "dtpFecha";
            this.dtpFecha.Size = new System.Drawing.Size(305, 27);
            this.dtpFecha.TabIndex = 10;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(332, 121);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(237, 27);
            this.textBox1.TabIndex = 34;
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            // 
            // txtOficina
            // 
            this.txtOficina.Location = new System.Drawing.Point(69, 104);
            this.txtOficina.Name = "txtOficina";
            this.txtOficina.Size = new System.Drawing.Size(157, 27);
            this.txtOficina.TabIndex = 9;
            this.txtOficina.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtOficina_KeyPress);
            // 
            // Buscar
            // 
            this.Buscar.FlatAppearance.BorderColor = System.Drawing.Color.DarkOrange;
            this.Buscar.FlatAppearance.BorderSize = 2;
            this.Buscar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Buscar.ForeColor = System.Drawing.Color.White;
            this.Buscar.Image = ((System.Drawing.Image)(resources.GetObject("Buscar.Image")));
            this.Buscar.Location = new System.Drawing.Point(575, 121);
            this.Buscar.Name = "Buscar";
            this.Buscar.Size = new System.Drawing.Size(50, 27);
            this.Buscar.TabIndex = 34;
            this.Buscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.Buscar.UseVisualStyleBackColor = true;
            // 
            // txtModelo
            // 
            this.txtModelo.Location = new System.Drawing.Point(69, 66);
            this.txtModelo.Name = "txtModelo";
            this.txtModelo.Size = new System.Drawing.Size(237, 27);
            this.txtModelo.TabIndex = 8;
            this.txtModelo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtModelo_KeyPress);
            // 
            // cboMarca
            // 
            this.cboMarca.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarca.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarca.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarca.FormattingEnabled = true;
            this.cboMarca.Location = new System.Drawing.Point(69, 27);
            this.cboMarca.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarca.Name = "cboMarca";
            this.cboMarca.Size = new System.Drawing.Size(237, 27);
            this.cboMarca.Sorted = true;
            this.cboMarca.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(312, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 19);
            this.label4.TabIndex = 6;
            this.label4.Text = "Fecha:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 107);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 19);
            this.label3.TabIndex = 5;
            this.label3.Text = "Oficina:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 69);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 19);
            this.label1.TabIndex = 4;
            this.label1.Text = "Modelo:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(8, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 19);
            this.label2.TabIndex = 3;
            this.label2.Text = "Marca:";
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.tabPage1.Controls.Add(this.grpDatosRegistro);
            this.tabPage1.Location = new System.Drawing.Point(4, 28);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(1133, 167);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "Registros";
            // 
            // grpDatosRegistro
            // 
            this.grpDatosRegistro.Controls.Add(this.btnGarantias);
            this.grpDatosRegistro.Controls.Add(this.label15);
            this.grpDatosRegistro.Controls.Add(this.radGarantia);
            this.grpDatosRegistro.Controls.Add(this.radSalida);
            this.grpDatosRegistro.Controls.Add(this.radEntrada);
            this.grpDatosRegistro.Controls.Add(this.btnAbrirReportes);
            this.grpDatosRegistro.Controls.Add(this.cboFusor);
            this.grpDatosRegistro.Controls.Add(this.txtCantidad);
            this.grpDatosRegistro.Controls.Add(this.label16);
            this.grpDatosRegistro.Controls.Add(this.label17);
            this.grpDatosRegistro.Controls.Add(this.cboClientes);
            this.grpDatosRegistro.Controls.Add(this.lblSerieP);
            this.grpDatosRegistro.Controls.Add(this.cboModelos);
            this.grpDatosRegistro.Controls.Add(this.label20);
            this.grpDatosRegistro.Controls.Add(this.label21);
            this.grpDatosRegistro.Controls.Add(this.cboMarcas);
            this.grpDatosRegistro.Controls.Add(this.dtpFechaRegistro);
            this.grpDatosRegistro.Controls.Add(this.label22);
            this.grpDatosRegistro.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpDatosRegistro.Location = new System.Drawing.Point(11, 5);
            this.grpDatosRegistro.Name = "grpDatosRegistro";
            this.grpDatosRegistro.Size = new System.Drawing.Size(1111, 156);
            this.grpDatosRegistro.TabIndex = 2;
            this.grpDatosRegistro.TabStop = false;
            this.grpDatosRegistro.Text = "Datos";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.White;
            this.label15.Location = new System.Drawing.Point(396, 20);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(91, 19);
            this.label15.TabIndex = 51;
            this.label15.Text = "Movimiento";
            // 
            // radGarantia
            // 
            this.radGarantia.AutoSize = true;
            this.radGarantia.Location = new System.Drawing.Point(562, 42);
            this.radGarantia.Name = "radGarantia";
            this.radGarantia.Size = new System.Drawing.Size(86, 23);
            this.radGarantia.TabIndex = 50;
            this.radGarantia.TabStop = true;
            this.radGarantia.Text = "Garantia";
            this.radGarantia.UseVisualStyleBackColor = true;
            // 
            // radSalida
            // 
            this.radSalida.AutoSize = true;
            this.radSalida.Location = new System.Drawing.Point(487, 42);
            this.radSalida.Name = "radSalida";
            this.radSalida.Size = new System.Drawing.Size(69, 23);
            this.radSalida.TabIndex = 49;
            this.radSalida.TabStop = true;
            this.radSalida.Text = "Salida";
            this.radSalida.UseVisualStyleBackColor = true;
            // 
            // radEntrada
            // 
            this.radEntrada.AutoSize = true;
            this.radEntrada.Location = new System.Drawing.Point(400, 42);
            this.radEntrada.Name = "radEntrada";
            this.radEntrada.Size = new System.Drawing.Size(81, 23);
            this.radEntrada.TabIndex = 48;
            this.radEntrada.TabStop = true;
            this.radEntrada.Text = "Entrada";
            this.radEntrada.UseVisualStyleBackColor = true;
            // 
            // btnAbrirReportes
            // 
            this.btnAbrirReportes.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnAbrirReportes.FlatAppearance.BorderSize = 2;
            this.btnAbrirReportes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAbrirReportes.ForeColor = System.Drawing.Color.White;
            this.btnAbrirReportes.Image = ((System.Drawing.Image)(resources.GetObject("btnAbrirReportes.Image")));
            this.btnAbrirReportes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAbrirReportes.Location = new System.Drawing.Point(972, 91);
            this.btnAbrirReportes.Name = "btnAbrirReportes";
            this.btnAbrirReportes.Size = new System.Drawing.Size(133, 56);
            this.btnAbrirReportes.TabIndex = 35;
            this.btnAbrirReportes.Text = "Reporte";
            this.btnAbrirReportes.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAbrirReportes.UseVisualStyleBackColor = true;
            this.btnAbrirReportes.Click += new System.EventHandler(this.btnAbrirReportes_Click);
            // 
            // cboFusor
            // 
            this.cboFusor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboFusor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboFusor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFusor.FormattingEnabled = true;
            this.cboFusor.Location = new System.Drawing.Point(465, 105);
            this.cboFusor.Margin = new System.Windows.Forms.Padding(4);
            this.cboFusor.Name = "cboFusor";
            this.cboFusor.Size = new System.Drawing.Size(303, 27);
            this.cboFusor.Sorted = true;
            this.cboFusor.TabIndex = 46;
            this.cboFusor.Visible = false;
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(479, 73);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(99, 27);
            this.txtCantidad.TabIndex = 6;
            this.txtCantidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantidad_KeyPress);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.ForeColor = System.Drawing.Color.White;
            this.label16.Location = new System.Drawing.Point(396, 79);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(77, 19);
            this.label16.TabIndex = 45;
            this.label16.Text = "Cantidad:";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.ForeColor = System.Drawing.Color.White;
            this.label17.Location = new System.Drawing.Point(7, 122);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(63, 19);
            this.label17.TabIndex = 44;
            this.label17.Text = "Cliente:";
            // 
            // cboClientes
            // 
            this.cboClientes.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClientes.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClientes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClientes.FormattingEnabled = true;
            this.cboClientes.Location = new System.Drawing.Point(80, 119);
            this.cboClientes.Margin = new System.Windows.Forms.Padding(4);
            this.cboClientes.Name = "cboClientes";
            this.cboClientes.Size = new System.Drawing.Size(293, 27);
            this.cboClientes.Sorted = true;
            this.cboClientes.TabIndex = 5;
            // 
            // lblSerieP
            // 
            this.lblSerieP.AutoSize = true;
            this.lblSerieP.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSerieP.ForeColor = System.Drawing.Color.White;
            this.lblSerieP.Location = new System.Drawing.Point(396, 110);
            this.lblSerieP.Name = "lblSerieP";
            this.lblSerieP.Size = new System.Drawing.Size(62, 19);
            this.lblSerieP.TabIndex = 41;
            this.lblSerieP.Text = "#Serie:";
            // 
            // cboModelos
            // 
            this.cboModelos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelos.FormattingEnabled = true;
            this.cboModelos.Location = new System.Drawing.Point(80, 87);
            this.cboModelos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelos.Name = "cboModelos";
            this.cboModelos.Size = new System.Drawing.Size(293, 27);
            this.cboModelos.Sorted = true;
            this.cboModelos.TabIndex = 3;
            this.cboModelos.SelectedIndexChanged += new System.EventHandler(this.cboModelosP_SelectedIndexChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label20.ForeColor = System.Drawing.Color.White;
            this.label20.Location = new System.Drawing.Point(14, 57);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(56, 19);
            this.label20.TabIndex = 37;
            this.label20.Text = "Marca:";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label21.ForeColor = System.Drawing.Color.White;
            this.label21.Location = new System.Drawing.Point(7, 91);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(66, 19);
            this.label21.TabIndex = 36;
            this.label21.Text = "Modelo:";
            // 
            // cboMarcas
            // 
            this.cboMarcas.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarcas.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarcas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarcas.FormattingEnabled = true;
            this.cboMarcas.Location = new System.Drawing.Point(80, 54);
            this.cboMarcas.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarcas.Name = "cboMarcas";
            this.cboMarcas.Size = new System.Drawing.Size(292, 27);
            this.cboMarcas.Sorted = true;
            this.cboMarcas.TabIndex = 2;
            this.cboMarcas.SelectedIndexChanged += new System.EventHandler(this.cboMarcaP_SelectedIndexChanged);
            // 
            // dtpFechaRegistro
            // 
            this.dtpFechaRegistro.Location = new System.Drawing.Point(80, 20);
            this.dtpFechaRegistro.Name = "dtpFechaRegistro";
            this.dtpFechaRegistro.Size = new System.Drawing.Size(292, 27);
            this.dtpFechaRegistro.TabIndex = 1;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label22.ForeColor = System.Drawing.Color.White;
            this.label22.Location = new System.Drawing.Point(15, 26);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(55, 19);
            this.label22.TabIndex = 6;
            this.label22.Text = "Fecha:";
            // 
            // dtgCartuchos
            // 
            this.dtgCartuchos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgCartuchos.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgCartuchos.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgCartuchos.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgCartuchos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgCartuchos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgCartuchos.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgCartuchos.EnableHeadersVisualStyles = false;
            this.dtgCartuchos.Location = new System.Drawing.Point(12, 274);
            this.dtgCartuchos.Name = "dtgCartuchos";
            this.dtgCartuchos.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgCartuchos.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgCartuchos.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgCartuchos.RowTemplate.Height = 25;
            this.dtgCartuchos.Size = new System.Drawing.Size(1326, 691);
            this.dtgCartuchos.TabIndex = 34;
            this.dtgCartuchos.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgCartuchos_CellClick);
            // 
            // erInventario
            // 
            this.erInventario.ContainerControl = this;
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(258, 217);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(117, 51);
            this.btnCancelar.TabIndex = 33;
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
            this.btnEliminar.Location = new System.Drawing.Point(135, 217);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(117, 51);
            this.btnEliminar.TabIndex = 32;
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
            this.btnGuardar.Location = new System.Drawing.Point(12, 217);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 31;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnGarantias
            // 
            this.btnGarantias.FlatAppearance.BorderColor = System.Drawing.Color.Black;
            this.btnGarantias.FlatAppearance.BorderSize = 2;
            this.btnGarantias.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGarantias.ForeColor = System.Drawing.Color.White;
            this.btnGarantias.Image = ((System.Drawing.Image)(resources.GetObject("btnGarantias.Image")));
            this.btnGarantias.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGarantias.Location = new System.Drawing.Point(833, 91);
            this.btnGarantias.Name = "btnGarantias";
            this.btnGarantias.Size = new System.Drawing.Size(133, 56);
            this.btnGarantias.TabIndex = 52;
            this.btnGarantias.Text = "Garantia";
            this.btnGarantias.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGarantias.UseVisualStyleBackColor = true;
            this.btnGarantias.Click += new System.EventHandler(this.btnGarantias_Click);
            // 
            // Inventario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1350, 977);
            this.Controls.Add(this.dtgCartuchos);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.btnGuardar);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Inventario";
            this.Text = "Inventario";
            this.tabControl1.ResumeLayout(false);
            this.tabInventario.ResumeLayout(false);
            this.grpDatosInventario.ResumeLayout(false);
            this.grpDatosInventario.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.grpDatosRegistro.ResumeLayout(false);
            this.grpDatosRegistro.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgCartuchos)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.erInventario)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabInventario;
        private System.Windows.Forms.GroupBox grpDatosInventario;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboMarca;
        private System.Windows.Forms.DateTimePicker dtpFecha;
        private System.Windows.Forms.TextBox txtOficina;
        private System.Windows.Forms.TextBox txtModelo;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button Buscar;
        private System.Windows.Forms.ComboBox cboMarcaSeleccionada;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DataGridView dtgCartuchos;
        private System.Windows.Forms.Button btnImprimir;
        private System.Windows.Forms.ErrorProvider erInventario;
        private System.Windows.Forms.TextBox txtPrecio;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnAbrirReporteExistencias;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox grpDatosRegistro;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.RadioButton radGarantia;
        private System.Windows.Forms.RadioButton radSalida;
        private System.Windows.Forms.RadioButton radEntrada;
        private System.Windows.Forms.Button btnAbrirReportes;
        private System.Windows.Forms.ComboBox cboFusor;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.ComboBox cboClientes;
        private System.Windows.Forms.Label lblSerieP;
        private System.Windows.Forms.ComboBox cboModelos;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.ComboBox cboMarcas;
        private System.Windows.Forms.DateTimePicker dtpFechaRegistro;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Button btnGarantias;
    }
}