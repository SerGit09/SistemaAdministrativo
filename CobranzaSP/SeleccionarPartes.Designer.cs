namespace CobranzaSP
{
    partial class SeleccionarPartes
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SeleccionarPartes));
            this.dtgPartes = new System.Windows.Forms.DataGridView();
            this.grpDatosPartes = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtReporteServicio = new System.Windows.Forms.TextBox();
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
            ((System.ComponentModel.ISupportInitialize)(this.dtgPartes)).BeginInit();
            this.grpDatosPartes.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtgPartes
            // 
            this.dtgPartes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dtgPartes.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dtgPartes.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            this.dtgPartes.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dtgPartes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InfoText;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dtgPartes.DefaultCellStyle = dataGridViewCellStyle2;
            this.dtgPartes.EnableHeadersVisualStyles = false;
            this.dtgPartes.Location = new System.Drawing.Point(12, 278);
            this.dtgPartes.Name = "dtgPartes";
            this.dtgPartes.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(45)))), ((int)(((byte)(66)))), ((int)(((byte)(90)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.White;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dtgPartes.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dtgPartes.RowTemplate.Height = 25;
            this.dtgPartes.Size = new System.Drawing.Size(682, 505);
            this.dtgPartes.TabIndex = 36;
            // 
            // grpDatosPartes
            // 
            this.grpDatosPartes.Controls.Add(this.label1);
            this.grpDatosPartes.Controls.Add(this.txtReporteServicio);
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
            this.grpDatosPartes.Location = new System.Drawing.Point(12, 12);
            this.grpDatosPartes.Name = "grpDatosPartes";
            this.grpDatosPartes.Size = new System.Drawing.Size(682, 260);
            this.grpDatosPartes.TabIndex = 35;
            this.grpDatosPartes.TabStop = false;
            this.grpDatosPartes.Text = "Partes:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(132, 19);
            this.label1.TabIndex = 52;
            this.label1.Text = "Numero Reporte:";
            // 
            // txtReporteServicio
            // 
            this.txtReporteServicio.Location = new System.Drawing.Point(144, 20);
            this.txtReporteServicio.Name = "txtReporteServicio";
            this.txtReporteServicio.Size = new System.Drawing.Size(132, 27);
            this.txtReporteServicio.TabIndex = 53;
            // 
            // cboModelo
            // 
            this.cboModelo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModelo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModelo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModelo.FormattingEnabled = true;
            this.cboModelo.Location = new System.Drawing.Point(79, 62);
            this.cboModelo.Margin = new System.Windows.Forms.Padding(4);
            this.cboModelo.Name = "cboModelo";
            this.cboModelo.Size = new System.Drawing.Size(342, 27);
            this.cboModelo.Sorted = true;
            this.cboModelo.TabIndex = 51;
            // 
            // radUsado
            // 
            this.radUsado.AutoSize = true;
            this.radUsado.Location = new System.Drawing.Point(305, 221);
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
            this.label8.Location = new System.Drawing.Point(7, 223);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 19);
            this.label8.TabIndex = 37;
            this.label8.Text = "Cantidad:";
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(90, 220);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(85, 27);
            this.txtCantidad.TabIndex = 41;
            // 
            // radNuevo
            // 
            this.radNuevo.AutoSize = true;
            this.radNuevo.Location = new System.Drawing.Point(218, 221);
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
            this.label6.Location = new System.Drawing.Point(7, 94);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(96, 19);
            this.label6.TabIndex = 39;
            this.label6.Text = "Descripcion:";
            // 
            // rtxtDescripcion
            // 
            this.rtxtDescripcion.Location = new System.Drawing.Point(11, 116);
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
            this.btnCancelarParte.Location = new System.Drawing.Point(559, 135);
            this.btnCancelarParte.Name = "btnCancelarParte";
            this.btnCancelarParte.Size = new System.Drawing.Size(117, 51);
            this.btnCancelarParte.TabIndex = 47;
            this.btnCancelarParte.Text = "Cancelar";
            this.btnCancelarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCancelarParte.UseVisualStyleBackColor = true;
            // 
            // btnEliminarParte
            // 
            this.btnEliminarParte.FlatAppearance.BorderColor = System.Drawing.Color.Red;
            this.btnEliminarParte.FlatAppearance.BorderSize = 2;
            this.btnEliminarParte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEliminarParte.ForeColor = System.Drawing.Color.White;
            this.btnEliminarParte.Image = ((System.Drawing.Image)(resources.GetObject("btnEliminarParte.Image")));
            this.btnEliminarParte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEliminarParte.Location = new System.Drawing.Point(559, 78);
            this.btnEliminarParte.Name = "btnEliminarParte";
            this.btnEliminarParte.Size = new System.Drawing.Size(117, 51);
            this.btnEliminarParte.TabIndex = 46;
            this.btnEliminarParte.Text = "Eliminar";
            this.btnEliminarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEliminarParte.UseVisualStyleBackColor = true;
            // 
            // btnGuardarParte
            // 
            this.btnGuardarParte.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.btnGuardarParte.FlatAppearance.BorderSize = 2;
            this.btnGuardarParte.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardarParte.ForeColor = System.Drawing.Color.White;
            this.btnGuardarParte.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardarParte.Image")));
            this.btnGuardarParte.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGuardarParte.Location = new System.Drawing.Point(559, 20);
            this.btnGuardarParte.Name = "btnGuardarParte";
            this.btnGuardarParte.Size = new System.Drawing.Size(117, 51);
            this.btnGuardarParte.TabIndex = 45;
            this.btnGuardarParte.Text = "Guardar";
            this.btnGuardarParte.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnGuardarParte.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(7, 65);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 19);
            this.label5.TabIndex = 40;
            this.label5.Text = "Modelo:";
            // 
            // SeleccionarPartes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.ClientSize = new System.Drawing.Size(1510, 795);
            this.Controls.Add(this.dtgPartes);
            this.Controls.Add(this.grpDatosPartes);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SeleccionarPartes";
            this.Text = "SeleccionarPartes";
            ((System.ComponentModel.ISupportInitialize)(this.dtgPartes)).EndInit();
            this.grpDatosPartes.ResumeLayout(false);
            this.grpDatosPartes.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dtgPartes;
        private System.Windows.Forms.GroupBox grpDatosPartes;
        private System.Windows.Forms.ComboBox cboModelo;
        private System.Windows.Forms.RadioButton radUsado;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.RadioButton radNuevo;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.RichTextBox rtxtDescripcion;
        private System.Windows.Forms.Button btnCancelarParte;
        private System.Windows.Forms.Button btnEliminarParte;
        private System.Windows.Forms.Button btnGuardarParte;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtReporteServicio;
    }
}