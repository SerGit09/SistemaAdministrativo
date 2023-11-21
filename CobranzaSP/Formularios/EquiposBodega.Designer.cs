namespace CobranzaSP.Formularios
{
    partial class EquiposBodega
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EquiposBodega));
            this.grpDatos = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.rdUsada = new System.Windows.Forms.RadioButton();
            this.rdNueva = new System.Windows.Forms.RadioButton();
            this.rtxtNotas = new System.Windows.Forms.RichTextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtPrecio = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSerie = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboModelos = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboMarcas = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dtgEquipos = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpModelo = new System.Windows.Forms.GroupBox();
            this.radUnModelo = new System.Windows.Forms.RadioButton();
            this.rdTodosLosModelos = new System.Windows.Forms.RadioButton();
            this.cboModelo = new System.Windows.Forms.ComboBox();
            this.rdTodasLasMarcas = new System.Windows.Forms.RadioButton();
            this.rdUnaMarca = new System.Windows.Forms.RadioButton();
            this.cboBusqueda = new System.Windows.Forms.ComboBox();
            this.cboOpcionMostrar = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.erEquipos = new System.Windows.Forms.ErrorProvider(this.components);
            this.btnMostrar = new System.Windows.Forms.Button();
            this.btnRentada = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnVentaEquipo = new System.Windows.Forms.Button();
            this.grpDatos.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgEquipos)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.grpModelo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.erEquipos)).BeginInit();
            this.SuspendLayout();
            // 
            // grpDatos
            // 
            this.grpDatos.Controls.Add(this.btnVentaEquipo);
            this.grpDatos.Controls.Add(this.btnRentada);
            this.grpDatos.Controls.Add(this.label8);
            this.grpDatos.Controls.Add(this.rdUsada);
            this.grpDatos.Controls.Add(this.rdNueva);
            this.grpDatos.Controls.Add(this.rtxtNotas);
            this.grpDatos.Controls.Add(this.label7);
            this.grpDatos.Controls.Add(this.txtPrecio);
            this.grpDatos.Controls.Add(this.label6);
            this.grpDatos.Controls.Add(this.txtSerie);
            this.grpDatos.Controls.Add(this.label1);
            this.grpDatos.Controls.Add(this.cboModelos);
            this.grpDatos.Controls.Add(this.label4);
            this.grpDatos.Controls.Add(this.cboMarcas);
            this.grpDatos.Controls.Add(this.label3);
            this.grpDatos.Controls.Add(this.btnCancelar);
            this.grpDatos.Controls.Add(this.btnEliminar);
            this.grpDatos.Controls.Add(this.btnGuardar);
            this.grpDatos.Location = new System.Drawing.Point(12, 12);
            this.grpDatos.Name = "grpDatos";
            this.grpDatos.Size = new System.Drawing.Size(917, 336);
            this.grpDatos.TabIndex = 3;
            this.grpDatos.TabStop = false;
            this.grpDatos.Text = "Datos";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(268, 113);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(62, 19);
            this.label8.TabIndex = 45;
            this.label8.Text = "Estado:";
            // 
            // rdUsada
            // 
            this.rdUsada.AutoSize = true;
            this.rdUsada.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.rdUsada.Location = new System.Drawing.Point(379, 142);
            this.rdUsada.Name = "rdUsada";
            this.rdUsada.Size = new System.Drawing.Size(70, 23);
            this.rdUsada.TabIndex = 44;
            this.rdUsada.TabStop = true;
            this.rdUsada.Text = "Usada";
            this.rdUsada.UseVisualStyleBackColor = true;
            // 
            // rdNueva
            // 
            this.rdNueva.AutoSize = true;
            this.rdNueva.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.rdNueva.Location = new System.Drawing.Point(272, 142);
            this.rdNueva.Name = "rdNueva";
            this.rdNueva.Size = new System.Drawing.Size(71, 23);
            this.rdNueva.TabIndex = 43;
            this.rdNueva.TabStop = true;
            this.rdNueva.Text = "Nueva";
            this.rdNueva.UseVisualStyleBackColor = true;
            // 
            // rtxtNotas
            // 
            this.rtxtNotas.Location = new System.Drawing.Point(548, 113);
            this.rtxtNotas.Name = "rtxtNotas";
            this.rtxtNotas.Size = new System.Drawing.Size(356, 122);
            this.rtxtNotas.TabIndex = 42;
            this.rtxtNotas.Text = "";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(544, 88);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 19);
            this.label7.TabIndex = 41;
            this.label7.Text = "Notas:";
            // 
            // txtPrecio
            // 
            this.txtPrecio.Location = new System.Drawing.Point(87, 110);
            this.txtPrecio.Name = "txtPrecio";
            this.txtPrecio.Size = new System.Drawing.Size(158, 27);
            this.txtPrecio.TabIndex = 7;
            this.txtPrecio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPrecio_KeyPress);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(23, 116);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 19);
            this.label6.TabIndex = 40;
            this.label6.Text = "Precio:";
            // 
            // txtSerie
            // 
            this.txtSerie.Location = new System.Drawing.Point(87, 69);
            this.txtSerie.Name = "txtSerie";
            this.txtSerie.Size = new System.Drawing.Size(443, 27);
            this.txtSerie.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(23, 77);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(50, 19);
            this.label1.TabIndex = 36;
            this.label1.Text = "Serie:";
            // 
            // cboModelos
            // 
            this.cboModelos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelos.FormattingEnabled = true;
            this.cboModelos.Location = new System.Drawing.Point(305, 27);
            this.cboModelos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelos.Name = "cboModelos";
            this.cboModelos.Size = new System.Drawing.Size(225, 27);
            this.cboModelos.Sorted = true;
            this.cboModelos.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(232, 30);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 19);
            this.label4.TabIndex = 34;
            this.label4.Text = "Modelo:";
            // 
            // cboMarcas
            // 
            this.cboMarcas.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboMarcas.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboMarcas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboMarcas.FormattingEnabled = true;
            this.cboMarcas.Location = new System.Drawing.Point(88, 27);
            this.cboMarcas.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarcas.Name = "cboMarcas";
            this.cboMarcas.Size = new System.Drawing.Size(137, 27);
            this.cboMarcas.Sorted = true;
            this.cboMarcas.TabIndex = 3;
            this.cboMarcas.SelectedIndexChanged += new System.EventHandler(this.cboMarcas_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(25, 32);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 19);
            this.label3.TabIndex = 33;
            this.label3.Text = "Marca:";
            // 
            // dtgEquipos
            // 
            this.dtgEquipos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgEquipos.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgEquipos.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgEquipos.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgEquipos.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgEquipos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgEquipos.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgEquipos.EnableHeadersVisualStyles = false;
            this.dtgEquipos.Location = new System.Drawing.Point(12, 354);
            this.dtgEquipos.Name = "dtgEquipos";
            this.dtgEquipos.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgEquipos.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgEquipos.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgEquipos.RowTemplate.Height = 25;
            this.dtgEquipos.Size = new System.Drawing.Size(1326, 611);
            this.dtgEquipos.TabIndex = 33;
            this.dtgEquipos.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgEquipos_CellClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpModelo);
            this.groupBox1.Controls.Add(this.rdTodasLasMarcas);
            this.groupBox1.Controls.Add(this.rdUnaMarca);
            this.groupBox1.Controls.Add(this.btnMostrar);
            this.groupBox1.Controls.Add(this.cboBusqueda);
            this.groupBox1.Controls.Add(this.cboOpcionMostrar);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.ForeColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Location = new System.Drawing.Point(949, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(389, 336);
            this.groupBox1.TabIndex = 34;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Reportes";
            // 
            // grpModelo
            // 
            this.grpModelo.Controls.Add(this.radUnModelo);
            this.grpModelo.Controls.Add(this.rdTodosLosModelos);
            this.grpModelo.Controls.Add(this.cboModelo);
            this.grpModelo.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpModelo.Location = new System.Drawing.Point(19, 146);
            this.grpModelo.Name = "grpModelo";
            this.grpModelo.Size = new System.Drawing.Size(356, 100);
            this.grpModelo.TabIndex = 57;
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
            // 
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(7, 56);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(337, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 54;
            this.cboModelo.Visible = false;
            // 
            // rdTodasLasMarcas
            // 
            this.rdTodasLasMarcas.AutoSize = true;
            this.rdTodasLasMarcas.Location = new System.Drawing.Point(80, 82);
            this.rdTodasLasMarcas.Name = "rdTodasLasMarcas";
            this.rdTodasLasMarcas.Size = new System.Drawing.Size(70, 23);
            this.rdTodasLasMarcas.TabIndex = 53;
            this.rdTodasLasMarcas.TabStop = true;
            this.rdTodasLasMarcas.Text = "Todas";
            this.rdTodasLasMarcas.UseVisualStyleBackColor = true;
            this.rdTodasLasMarcas.Visible = false;
            // 
            // rdUnaMarca
            // 
            this.rdUnaMarca.AutoSize = true;
            this.rdUnaMarca.Location = new System.Drawing.Point(19, 82);
            this.rdUnaMarca.Name = "rdUnaMarca";
            this.rdUnaMarca.Size = new System.Drawing.Size(55, 23);
            this.rdUnaMarca.TabIndex = 52;
            this.rdUnaMarca.TabStop = true;
            this.rdUnaMarca.Text = "Una";
            this.rdUnaMarca.UseVisualStyleBackColor = true;
            this.rdUnaMarca.Visible = false;
            this.rdUnaMarca.CheckedChanged += new System.EventHandler(this.rdUnaMarca_CheckedChanged);
            // 
            // cboBusqueda
            // 
            this.cboBusqueda.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboBusqueda.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboBusqueda.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboBusqueda.FormattingEnabled = true;
            this.cboBusqueda.Location = new System.Drawing.Point(19, 112);
            this.cboBusqueda.Margin = new System.Windows.Forms.Padding(4);
            this.cboBusqueda.Name = "cboBusqueda";
            this.cboBusqueda.Size = new System.Drawing.Size(356, 27);
            this.cboBusqueda.Sorted = true;
            this.cboBusqueda.TabIndex = 50;
            this.cboBusqueda.Visible = false;
            this.cboBusqueda.SelectedIndexChanged += new System.EventHandler(this.cboBusqueda_SelectedIndexChanged);
            // 
            // cboOpcionMostrar
            // 
            this.cboOpcionMostrar.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboOpcionMostrar.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboOpcionMostrar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOpcionMostrar.FormattingEnabled = true;
            this.cboOpcionMostrar.Location = new System.Drawing.Point(19, 48);
            this.cboOpcionMostrar.Margin = new System.Windows.Forms.Padding(4);
            this.cboOpcionMostrar.Name = "cboOpcionMostrar";
            this.cboOpcionMostrar.Size = new System.Drawing.Size(356, 27);
            this.cboOpcionMostrar.Sorted = true;
            this.cboOpcionMostrar.TabIndex = 49;
            this.cboOpcionMostrar.SelectedIndexChanged += new System.EventHandler(this.cboOpcionMostrar_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(15, 25);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(100, 19);
            this.label9.TabIndex = 48;
            this.label9.Text = "Generar por:";
            // 
            // erEquipos
            // 
            this.erEquipos.ContainerControl = this;
            // 
            // btnMostrar
            // 
            this.btnMostrar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnMostrar.FlatAppearance.BorderSize = 2;
            this.btnMostrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMostrar.ForeColor = System.Drawing.Color.White;
            this.btnMostrar.Image = ((System.Drawing.Image)(resources.GetObject("btnMostrar.Image")));
            this.btnMostrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnMostrar.Location = new System.Drawing.Point(266, 267);
            this.btnMostrar.Name = "btnMostrar";
            this.btnMostrar.Size = new System.Drawing.Size(117, 51);
            this.btnMostrar.TabIndex = 51;
            this.btnMostrar.Text = "Generar";
            this.btnMostrar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnMostrar.UseVisualStyleBackColor = true;
            this.btnMostrar.Click += new System.EventHandler(this.btnMostrar_Click);
            // 
            // btnRentada
            // 
            this.btnRentada.FlatAppearance.BorderColor = System.Drawing.Color.Crimson;
            this.btnRentada.FlatAppearance.BorderSize = 2;
            this.btnRentada.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRentada.ForeColor = System.Drawing.Color.White;
            this.btnRentada.Image = ((System.Drawing.Image)(resources.GetObject("btnRentada.Image")));
            this.btnRentada.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRentada.Location = new System.Drawing.Point(777, 261);
            this.btnRentada.Name = "btnRentada";
            this.btnRentada.Size = new System.Drawing.Size(134, 63);
            this.btnRentada.TabIndex = 46;
            this.btnRentada.Text = "Rentada";
            this.btnRentada.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnRentada.UseVisualStyleBackColor = true;
            this.btnRentada.Click += new System.EventHandler(this.btnRentada_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(262, 207);
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
            this.btnEliminar.Location = new System.Drawing.Point(139, 207);
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
            this.btnGuardar.Location = new System.Drawing.Point(16, 207);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 9;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // btnVentaEquipo
            // 
            this.btnVentaEquipo.FlatAppearance.BorderColor = System.Drawing.Color.DarkBlue;
            this.btnVentaEquipo.FlatAppearance.BorderSize = 2;
            this.btnVentaEquipo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnVentaEquipo.ForeColor = System.Drawing.Color.White;
            this.btnVentaEquipo.Image = ((System.Drawing.Image)(resources.GetObject("btnVentaEquipo.Image")));
            this.btnVentaEquipo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnVentaEquipo.Location = new System.Drawing.Point(637, 261);
            this.btnVentaEquipo.Name = "btnVentaEquipo";
            this.btnVentaEquipo.Size = new System.Drawing.Size(134, 63);
            this.btnVentaEquipo.TabIndex = 47;
            this.btnVentaEquipo.Text = "Vendido";
            this.btnVentaEquipo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnVentaEquipo.UseVisualStyleBackColor = true;
            this.btnVentaEquipo.Click += new System.EventHandler(this.btnVentaEquipo_Click);
            // 
            // EquiposBodega
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1350, 977);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dtgEquipos);
            this.Controls.Add(this.grpDatos);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "EquiposBodega";
            this.Text = "EquiposBodega";
            this.grpDatos.ResumeLayout(false);
            this.grpDatos.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgEquipos)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.grpModelo.ResumeLayout(false);
            this.grpModelo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.erEquipos)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpDatos;
        private System.Windows.Forms.TextBox txtPrecio;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtSerie;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboModelos;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboMarcas;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.RichTextBox rtxtNotas;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridView dtgEquipos;
        private System.Windows.Forms.RadioButton rdUsada;
        private System.Windows.Forms.RadioButton rdNueva;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnMostrar;
        private System.Windows.Forms.ComboBox cboOpcionMostrar;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ErrorProvider erEquipos;
        private System.Windows.Forms.RadioButton rdTodosLosModelos;
        private System.Windows.Forms.RadioButton radUnModelo;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.RadioButton rdTodasLasMarcas;
        private System.Windows.Forms.RadioButton rdUnaMarca;
        private System.Windows.Forms.ComboBox cboBusqueda;
        private System.Windows.Forms.GroupBox grpModelo;
        private System.Windows.Forms.Button btnRentada;
        private System.Windows.Forms.Button btnVentaEquipo;
    }
}