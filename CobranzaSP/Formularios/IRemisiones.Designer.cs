namespace CobranzaSP.Formularios
{
    partial class IRemisiones
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IRemisiones));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.grpDatos = new System.Windows.Forms.GroupBox();
            this.cboClientes = new System.Windows.Forms.ComboBox();
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.dtpFechaRemision = new System.Windows.Forms.DateTimePicker();
            this.txtDiasCredito = new System.Windows.Forms.TextBox();
            this.txtFolio = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblDiasCredito = new System.Windows.Forms.Label();
            this.radMostrador = new System.Windows.Forms.RadioButton();
            this.radRemisiones = new System.Windows.Forms.RadioButton();
            this.lblRemisiones = new System.Windows.Forms.Label();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnEliminar = new System.Windows.Forms.Button();
            this.btnCobrado = new System.Windows.Forms.Button();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.dtgRemisiones = new System.Windows.Forms.DataGridView();
            this.btnEnviarCorreo = new System.Windows.Forms.Button();
            this.erRemisiones = new System.Windows.Forms.ErrorProvider(this.components);
            this.grpDatos.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgRemisiones)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.erRemisiones)).BeginInit();
            this.SuspendLayout();
            // 
            // grpDatos
            // 
            this.grpDatos.Controls.Add(this.cboClientes);
            this.grpDatos.Controls.Add(this.txtCantidad);
            this.grpDatos.Controls.Add(this.dtpFechaRemision);
            this.grpDatos.Controls.Add(this.txtDiasCredito);
            this.grpDatos.Controls.Add(this.txtFolio);
            this.grpDatos.Controls.Add(this.label5);
            this.grpDatos.Controls.Add(this.label4);
            this.grpDatos.Controls.Add(this.label3);
            this.grpDatos.Controls.Add(this.lblDiasCredito);
            this.grpDatos.Controls.Add(this.radMostrador);
            this.grpDatos.Controls.Add(this.radRemisiones);
            this.grpDatos.Controls.Add(this.lblRemisiones);
            this.grpDatos.Location = new System.Drawing.Point(12, 12);
            this.grpDatos.Name = "grpDatos";
            this.grpDatos.Size = new System.Drawing.Size(573, 234);
            this.grpDatos.TabIndex = 0;
            this.grpDatos.TabStop = false;
            this.grpDatos.Text = "Datos:";
            // 
            // cboClientes
            // 
            this.cboClientes.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClientes.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClientes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClientes.FormattingEnabled = true;
            this.cboClientes.Location = new System.Drawing.Point(151, 191);
            this.cboClientes.Margin = new System.Windows.Forms.Padding(4);
            this.cboClientes.Name = "cboClientes";
            this.cboClientes.Size = new System.Drawing.Size(311, 27);
            this.cboClientes.Sorted = true;
            this.cboClientes.TabIndex = 12;
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(150, 153);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(100, 27);
            this.txtCantidad.TabIndex = 11;
            this.txtCantidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantidad_KeyPress);
            // 
            // dtpFechaRemision
            // 
            this.dtpFechaRemision.Location = new System.Drawing.Point(150, 120);
            this.dtpFechaRemision.Name = "dtpFechaRemision";
            this.dtpFechaRemision.Size = new System.Drawing.Size(312, 27);
            this.dtpFechaRemision.TabIndex = 10;
            // 
            // txtDiasCredito
            // 
            this.txtDiasCredito.Location = new System.Drawing.Point(150, 88);
            this.txtDiasCredito.Name = "txtDiasCredito";
            this.txtDiasCredito.Size = new System.Drawing.Size(100, 27);
            this.txtDiasCredito.TabIndex = 9;
            this.txtDiasCredito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasCredito_KeyPress);
            // 
            // txtFolio
            // 
            this.txtFolio.Location = new System.Drawing.Point(150, 55);
            this.txtFolio.Name = "txtFolio";
            this.txtFolio.Size = new System.Drawing.Size(312, 27);
            this.txtFolio.TabIndex = 8;
            this.txtFolio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFolio_KeyPress);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(81, 194);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 19);
            this.label5.TabIndex = 7;
            this.label5.Text = "Cliente:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(67, 161);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 19);
            this.label4.TabIndex = 6;
            this.label4.Text = "Cantidad:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(49, 126);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 19);
            this.label3.TabIndex = 5;
            this.label3.Text = "Fecha pago:";
            // 
            // lblDiasCredito
            // 
            this.lblDiasCredito.AutoSize = true;
            this.lblDiasCredito.ForeColor = System.Drawing.Color.White;
            this.lblDiasCredito.Location = new System.Drawing.Point(46, 93);
            this.lblDiasCredito.Name = "lblDiasCredito";
            this.lblDiasCredito.Size = new System.Drawing.Size(98, 19);
            this.lblDiasCredito.TabIndex = 4;
            this.lblDiasCredito.Text = "Días crédito:";
            // 
            // radMostrador
            // 
            this.radMostrador.AutoSize = true;
            this.radMostrador.ForeColor = System.Drawing.Color.White;
            this.radMostrador.Location = new System.Drawing.Point(136, 26);
            this.radMostrador.Name = "radMostrador";
            this.radMostrador.Size = new System.Drawing.Size(98, 23);
            this.radMostrador.TabIndex = 3;
            this.radMostrador.TabStop = true;
            this.radMostrador.Text = "Mostrador";
            this.radMostrador.UseVisualStyleBackColor = true;
            this.radMostrador.CheckedChanged += new System.EventHandler(this.radMostrador_CheckedChanged);
            // 
            // radRemisiones
            // 
            this.radRemisiones.AutoSize = true;
            this.radRemisiones.ForeColor = System.Drawing.Color.White;
            this.radRemisiones.Location = new System.Drawing.Point(23, 26);
            this.radRemisiones.Name = "radRemisiones";
            this.radRemisiones.Size = new System.Drawing.Size(107, 23);
            this.radRemisiones.TabIndex = 2;
            this.radRemisiones.TabStop = true;
            this.radRemisiones.Text = "Remisiones";
            this.radRemisiones.UseVisualStyleBackColor = true;
            this.radRemisiones.CheckedChanged += new System.EventHandler(this.radRemisiones_CheckedChanged);
            // 
            // lblRemisiones
            // 
            this.lblRemisiones.AutoSize = true;
            this.lblRemisiones.ForeColor = System.Drawing.Color.White;
            this.lblRemisiones.Location = new System.Drawing.Point(6, 63);
            this.lblRemisiones.Name = "lblRemisiones";
            this.lblRemisiones.Size = new System.Drawing.Size(138, 19);
            this.lblRemisiones.TabIndex = 1;
            this.lblRemisiones.Text = "Numero remision:";
            // 
            // btnCancelar
            // 
            this.btnCancelar.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.btnCancelar.FlatAppearance.BorderSize = 2;
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.ForeColor = System.Drawing.Color.White;
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCancelar.Location = new System.Drawing.Point(591, 195);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(117, 51);
            this.btnCancelar.TabIndex = 8;
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
            this.btnEliminar.Location = new System.Drawing.Point(591, 138);
            this.btnEliminar.Name = "btnEliminar";
            this.btnEliminar.Size = new System.Drawing.Size(117, 51);
            this.btnEliminar.TabIndex = 7;
            this.btnEliminar.Text = "Eliminar";
            this.btnEliminar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEliminar.UseVisualStyleBackColor = true;
            this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
            // 
            // btnCobrado
            // 
            this.btnCobrado.FlatAppearance.BorderColor = System.Drawing.Color.MediumSlateBlue;
            this.btnCobrado.FlatAppearance.BorderSize = 2;
            this.btnCobrado.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCobrado.ForeColor = System.Drawing.Color.White;
            this.btnCobrado.Image = ((System.Drawing.Image)(resources.GetObject("btnCobrado.Image")));
            this.btnCobrado.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCobrado.Location = new System.Drawing.Point(591, 81);
            this.btnCobrado.Name = "btnCobrado";
            this.btnCobrado.Size = new System.Drawing.Size(117, 51);
            this.btnCobrado.TabIndex = 6;
            this.btnCobrado.Text = "Cobrado";
            this.btnCobrado.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCobrado.UseVisualStyleBackColor = true;
            this.btnCobrado.Click += new System.EventHandler(this.btnCobrado_Click);
            // 
            // btnGuardar
            // 
            this.btnGuardar.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.btnGuardar.FlatAppearance.BorderSize = 2;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
            this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGuardar.Location = new System.Drawing.Point(591, 24);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(117, 51);
            this.btnGuardar.TabIndex = 5;
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // dtgRemisiones
            // 
            this.dtgRemisiones.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgRemisiones.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgRemisiones.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgRemisiones.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgRemisiones.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgRemisiones.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgRemisiones.EnableHeadersVisualStyles = false;
            this.dtgRemisiones.Location = new System.Drawing.Point(12, 252);
            this.dtgRemisiones.Name = "dtgRemisiones";
            this.dtgRemisiones.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgRemisiones.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgRemisiones.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgRemisiones.RowTemplate.Height = 25;
            this.dtgRemisiones.Size = new System.Drawing.Size(1008, 296);
            this.dtgRemisiones.TabIndex = 18;
            this.dtgRemisiones.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgRemisiones_CellClick);
            // 
            // btnEnviarCorreo
            // 
            this.btnEnviarCorreo.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.btnEnviarCorreo.FlatAppearance.BorderSize = 2;
            this.btnEnviarCorreo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEnviarCorreo.Image = ((System.Drawing.Image)(resources.GetObject("btnEnviarCorreo.Image")));
            this.btnEnviarCorreo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEnviarCorreo.Location = new System.Drawing.Point(744, 195);
            this.btnEnviarCorreo.Name = "btnEnviarCorreo";
            this.btnEnviarCorreo.Size = new System.Drawing.Size(157, 51);
            this.btnEnviarCorreo.TabIndex = 19;
            this.btnEnviarCorreo.Text = "Enviar Correo";
            this.btnEnviarCorreo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEnviarCorreo.UseVisualStyleBackColor = true;
            this.btnEnviarCorreo.Click += new System.EventHandler(this.btnEnviarCorreo_Click);
            // 
            // erRemisiones
            // 
            this.erRemisiones.ContainerControl = this;
            // 
            // IRemisiones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1387, 862);
            this.Controls.Add(this.btnEnviarCorreo);
            this.Controls.Add(this.dtgRemisiones);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnEliminar);
            this.Controls.Add(this.btnCobrado);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.grpDatos);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "IRemisiones";
            this.Text = "IRemisiones";
            this.grpDatos.ResumeLayout(false);
            this.grpDatos.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtgRemisiones)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.erRemisiones)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpDatos;
        private System.Windows.Forms.RadioButton radMostrador;
        private System.Windows.Forms.RadioButton radRemisiones;
        private System.Windows.Forms.Label lblRemisiones;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.DateTimePicker dtpFechaRemision;
        private System.Windows.Forms.TextBox txtDiasCredito;
        private System.Windows.Forms.TextBox txtFolio;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblDiasCredito;
        private System.Windows.Forms.ComboBox cboClientes;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnEliminar;
        private System.Windows.Forms.Button btnCobrado;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.DataGridView dtgRemisiones;
        private System.Windows.Forms.Button btnEnviarCorreo;
        private System.Windows.Forms.ErrorProvider erRemisiones;
    }
}