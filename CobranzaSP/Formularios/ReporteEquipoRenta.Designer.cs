namespace CobranzaSP.Formularios
{
    partial class ReporteEquipoRenta
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ReporteEquipoRenta));
            this.pSuperior = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.PictureBox();
            this.cboOpcionMostrar = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.grpClientes = new System.Windows.Forms.GroupBox();
            this.radUnCliente = new System.Windows.Forms.RadioButton();
            this.radTodosLosClientes = new System.Windows.Forms.RadioButton();
            this.cboClienteElegido = new System.Windows.Forms.ComboBox();
            this.grpMarcas = new System.Windows.Forms.GroupBox();
            this.rdUnaMarca = new System.Windows.Forms.RadioButton();
            this.cboMarca = new System.Windows.Forms.ComboBox();
            this.rdTodasLasMarcas = new System.Windows.Forms.RadioButton();
            this.grpModelo = new System.Windows.Forms.GroupBox();
            this.radUnModelo = new System.Windows.Forms.RadioButton();
            this.rdTodosLosModelos = new System.Windows.Forms.RadioButton();
            this.cboModelo = new System.Windows.Forms.ComboBox();
            this.btnMostrar = new System.Windows.Forms.Button();
            this.pSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
            this.grpClientes.SuspendLayout();
            this.grpMarcas.SuspendLayout();
            this.grpModelo.SuspendLayout();
            this.SuspendLayout();
            // 
            // pSuperior
            // 
            this.pSuperior.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            this.pSuperior.Controls.Add(this.btnCerrar);
            this.pSuperior.Dock = System.Windows.Forms.DockStyle.Top;
            this.pSuperior.Location = new System.Drawing.Point(0, 0);
            this.pSuperior.Margin = new System.Windows.Forms.Padding(4);
            this.pSuperior.Name = "pSuperior";
            this.pSuperior.Size = new System.Drawing.Size(566, 38);
            this.pSuperior.TabIndex = 1;
            this.pSuperior.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pSuperior_MouseDown);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(528, 6);
            this.btnCerrar.Margin = new System.Windows.Forms.Padding(4);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(25, 28);
            this.btnCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnCerrar.TabIndex = 5;
            this.btnCerrar.TabStop = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // cboOpcionMostrar
            // 
            this.cboOpcionMostrar.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboOpcionMostrar.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboOpcionMostrar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboOpcionMostrar.FormattingEnabled = true;
            this.cboOpcionMostrar.Location = new System.Drawing.Point(73, 105);
            this.cboOpcionMostrar.Margin = new System.Windows.Forms.Padding(6);
            this.cboOpcionMostrar.Name = "cboOpcionMostrar";
            this.cboOpcionMostrar.Size = new System.Drawing.Size(358, 27);
            this.cboOpcionMostrar.Sorted = true;
            this.cboOpcionMostrar.TabIndex = 47;
            this.cboOpcionMostrar.SelectedIndexChanged += new System.EventHandler(this.cboOpcionMostrar_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(70, 80);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(100, 19);
            this.label9.TabIndex = 46;
            this.label9.Text = "Generar por:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(190, 42);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(241, 19);
            this.label1.TabIndex = 48;
            this.label1.Text = "REPORTES EQUIPOS EN RENTA";
            // 
            // grpClientes
            // 
            this.grpClientes.Controls.Add(this.radUnCliente);
            this.grpClientes.Controls.Add(this.radTodosLosClientes);
            this.grpClientes.Controls.Add(this.cboClienteElegido);
            this.grpClientes.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpClientes.Location = new System.Drawing.Point(75, 141);
            this.grpClientes.Name = "grpClientes";
            this.grpClientes.Size = new System.Drawing.Size(465, 100);
            this.grpClientes.TabIndex = 63;
            this.grpClientes.TabStop = false;
            this.grpClientes.Text = "Cliente";
            this.grpClientes.Visible = false;
            // 
            // radUnCliente
            // 
            this.radUnCliente.AutoSize = true;
            this.radUnCliente.Location = new System.Drawing.Point(6, 26);
            this.radUnCliente.Name = "radUnCliente";
            this.radUnCliente.Size = new System.Drawing.Size(56, 23);
            this.radUnCliente.TabIndex = 55;
            this.radUnCliente.TabStop = true;
            this.radUnCliente.Text = "Uno";
            this.radUnCliente.UseVisualStyleBackColor = true;
            this.radUnCliente.CheckedChanged += new System.EventHandler(this.radUnCliente_CheckedChanged);
            // 
            // radTodosLosClientes
            // 
            this.radTodosLosClientes.AutoSize = true;
            this.radTodosLosClientes.Location = new System.Drawing.Point(67, 26);
            this.radTodosLosClientes.Name = "radTodosLosClientes";
            this.radTodosLosClientes.Size = new System.Drawing.Size(71, 23);
            this.radTodosLosClientes.TabIndex = 56;
            this.radTodosLosClientes.TabStop = true;
            this.radTodosLosClientes.Text = "Todos";
            this.radTodosLosClientes.UseVisualStyleBackColor = true;
            this.radTodosLosClientes.CheckedChanged += new System.EventHandler(this.radTodosLosClientes_CheckedChanged);
            // 
            // cboClienteElegido
            // 
            this.cboClienteElegido.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClienteElegido.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClienteElegido.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClienteElegido.FormattingEnabled = true;
            this.cboClienteElegido.Location = new System.Drawing.Point(7, 56);
            this.cboClienteElegido.Margin = new System.Windows.Forms.Padding(4);
            this.cboClienteElegido.Name = "cboClienteElegido";
            this.cboClienteElegido.Size = new System.Drawing.Size(349, 27);
            this.cboClienteElegido.Sorted = true;
            this.cboClienteElegido.TabIndex = 54;
            this.cboClienteElegido.Visible = false;
            // 
            // grpMarcas
            // 
            this.grpMarcas.Controls.Add(this.rdUnaMarca);
            this.grpMarcas.Controls.Add(this.cboMarca);
            this.grpMarcas.Controls.Add(this.rdTodasLasMarcas);
            this.grpMarcas.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpMarcas.Location = new System.Drawing.Point(75, 247);
            this.grpMarcas.Name = "grpMarcas";
            this.grpMarcas.Size = new System.Drawing.Size(465, 100);
            this.grpMarcas.TabIndex = 64;
            this.grpMarcas.TabStop = false;
            this.grpMarcas.Text = "Marca";
            this.grpMarcas.Visible = false;
            // 
            // rdUnaMarca
            // 
            this.rdUnaMarca.AutoSize = true;
            this.rdUnaMarca.Location = new System.Drawing.Point(24, 26);
            this.rdUnaMarca.Name = "rdUnaMarca";
            this.rdUnaMarca.Size = new System.Drawing.Size(55, 23);
            this.rdUnaMarca.TabIndex = 59;
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
            this.cboMarca.Location = new System.Drawing.Point(24, 56);
            this.cboMarca.Margin = new System.Windows.Forms.Padding(4);
            this.cboMarca.Name = "cboMarca";
            this.cboMarca.Size = new System.Drawing.Size(332, 27);
            this.cboMarca.Sorted = true;
            this.cboMarca.TabIndex = 58;
            this.cboMarca.Visible = false;
            this.cboMarca.SelectedIndexChanged += new System.EventHandler(this.cboMarca_SelectedIndexChanged);
            // 
            // rdTodasLasMarcas
            // 
            this.rdTodasLasMarcas.AutoSize = true;
            this.rdTodasLasMarcas.Location = new System.Drawing.Point(85, 26);
            this.rdTodasLasMarcas.Name = "rdTodasLasMarcas";
            this.rdTodasLasMarcas.Size = new System.Drawing.Size(70, 23);
            this.rdTodasLasMarcas.TabIndex = 60;
            this.rdTodasLasMarcas.TabStop = true;
            this.rdTodasLasMarcas.Text = "Todas";
            this.rdTodasLasMarcas.UseVisualStyleBackColor = true;
            this.rdTodasLasMarcas.CheckedChanged += new System.EventHandler(this.rdTodasLasMarcas_CheckedChanged);
            // 
            // grpModelo
            // 
            this.grpModelo.Controls.Add(this.radUnModelo);
            this.grpModelo.Controls.Add(this.rdTodosLosModelos);
            this.grpModelo.Controls.Add(this.cboModelo);
            this.grpModelo.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.grpModelo.Location = new System.Drawing.Point(75, 353);
            this.grpModelo.Name = "grpModelo";
            this.grpModelo.Size = new System.Drawing.Size(465, 100);
            this.grpModelo.TabIndex = 65;
            this.grpModelo.TabStop = false;
            this.grpModelo.Text = "Modelo";
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
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(7, 56);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(349, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 54;
            this.cboModelo.Visible = false;
            // 
            // btnMostrar
            // 
            this.btnMostrar.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.btnMostrar.FlatAppearance.BorderSize = 2;
            this.btnMostrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMostrar.ForeColor = System.Drawing.Color.White;
            this.btnMostrar.Image = ((System.Drawing.Image)(resources.GetObject("btnMostrar.Image")));
            this.btnMostrar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnMostrar.Location = new System.Drawing.Point(437, 517);
            this.btnMostrar.Name = "btnMostrar";
            this.btnMostrar.Size = new System.Drawing.Size(117, 51);
            this.btnMostrar.TabIndex = 66;
            this.btnMostrar.Text = "Generar";
            this.btnMostrar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnMostrar.UseVisualStyleBackColor = true;
            this.btnMostrar.Click += new System.EventHandler(this.btnMostrar_Click);
            // 
            // ReporteEquipoRenta
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(39)))), ((int)(((byte)(57)))), ((int)(((byte)(80)))));
            this.ClientSize = new System.Drawing.Size(566, 580);
            this.Controls.Add(this.btnMostrar);
            this.Controls.Add(this.grpModelo);
            this.Controls.Add(this.grpMarcas);
            this.Controls.Add(this.grpClientes);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboOpcionMostrar);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.pSuperior);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ReporteEquipoRenta";
            this.Text = "ReporteEquipoRenta";
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ReporteEquipoRenta_MouseDown);
            this.pSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
            this.grpClientes.ResumeLayout(false);
            this.grpClientes.PerformLayout();
            this.grpMarcas.ResumeLayout(false);
            this.grpMarcas.PerformLayout();
            this.grpModelo.ResumeLayout(false);
            this.grpModelo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel pSuperior;
        private System.Windows.Forms.PictureBox btnCerrar;
        private System.Windows.Forms.ComboBox cboOpcionMostrar;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpClientes;
        private System.Windows.Forms.RadioButton radUnCliente;
        private System.Windows.Forms.RadioButton radTodosLosClientes;
        private System.Windows.Forms.ComboBox cboClienteElegido;
        private System.Windows.Forms.GroupBox grpMarcas;
        private System.Windows.Forms.RadioButton rdUnaMarca;
        private System.Windows.Forms.ComboBox cboMarca;
        private System.Windows.Forms.RadioButton rdTodasLasMarcas;
        private System.Windows.Forms.GroupBox grpModelo;
        private System.Windows.Forms.RadioButton radUnModelo;
        private System.Windows.Forms.RadioButton rdTodosLosModelos;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.Button btnMostrar;
    }
}