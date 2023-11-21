namespace CobranzaSP.Formularios
{
    partial class ModuloNuevo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ModuloNuevo));
            this.panelSuperior = new System.Windows.Forms.Panel();
            this.btnCerrar = new System.Windows.Forms.PictureBox();
            this.btnAgregarModulo = new System.Windows.Forms.Button();
            this.cboModulos = new System.Windows.Forms.ComboBox();
            this.txtClave = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.cboClaves = new System.Windows.Forms.ComboBox();
            this.lblAnteriorClave = new System.Windows.Forms.Label();
            this.txtClaveRetirada = new System.Windows.Forms.TextBox();
            this.erModulo = new System.Windows.Forms.ErrorProvider(this.components);
            this.grpModulo = new System.Windows.Forms.GroupBox();
            this.rtxtObsrevacion = new System.Windows.Forms.RichTextBox();
            this.lblObservacion = new System.Windows.Forms.Label();
            this.panelSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.erModulo)).BeginInit();
            this.grpModulo.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelSuperior
            // 
            this.panelSuperior.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            this.panelSuperior.Controls.Add(this.btnCerrar);
            this.panelSuperior.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSuperior.Location = new System.Drawing.Point(0, 0);
            this.panelSuperior.Name = "panelSuperior";
            this.panelSuperior.Size = new System.Drawing.Size(682, 35);
            this.panelSuperior.TabIndex = 1;
            this.panelSuperior.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelSuperior_MouseDown_1);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(645, 4);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(25, 28);
            this.btnCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnCerrar.TabIndex = 2;
            this.btnCerrar.TabStop = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnAgregarModulo
            // 
            this.btnAgregarModulo.FlatAppearance.BorderColor = System.Drawing.Color.Lime;
            this.btnAgregarModulo.FlatAppearance.BorderSize = 2;
            this.btnAgregarModulo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAgregarModulo.ForeColor = System.Drawing.Color.White;
            this.btnAgregarModulo.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregarModulo.Image")));
            this.btnAgregarModulo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAgregarModulo.Location = new System.Drawing.Point(553, 329);
            this.btnAgregarModulo.Name = "btnAgregarModulo";
            this.btnAgregarModulo.Size = new System.Drawing.Size(117, 51);
            this.btnAgregarModulo.TabIndex = 52;
            this.btnAgregarModulo.Text = "Guardar";
            this.btnAgregarModulo.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAgregarModulo.UseVisualStyleBackColor = true;
            this.btnAgregarModulo.Click += new System.EventHandler(this.btnAgregarModulo_Click);
            // 
            // cboModulos
            // 
            this.cboModulos.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboModulos.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboModulos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboModulos.FormattingEnabled = true;
            this.cboModulos.Location = new System.Drawing.Point(95, 28);
            this.cboModulos.Margin = new System.Windows.Forms.Padding(4);
            this.cboModulos.Name = "cboModulos";
            this.cboModulos.Size = new System.Drawing.Size(289, 27);
            this.cboModulos.Sorted = true;
            this.cboModulos.TabIndex = 62;
            this.cboModulos.SelectedIndexChanged += new System.EventHandler(this.cboModulos_SelectedIndexChanged);
            // 
            // txtClave
            // 
            this.txtClave.Location = new System.Drawing.Point(94, 63);
            this.txtClave.Name = "txtClave";
            this.txtClave.Size = new System.Drawing.Size(289, 27);
            this.txtClave.TabIndex = 59;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.White;
            this.label14.Location = new System.Drawing.Point(33, 66);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(53, 19);
            this.label14.TabIndex = 58;
            this.label14.Text = "Clave:";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.White;
            this.label15.Location = new System.Drawing.Point(21, 34);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(67, 19);
            this.label15.TabIndex = 57;
            this.label15.Text = "Modulo:";
            // 
            // cboClaves
            // 
            this.cboClaves.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cboClaves.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboClaves.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboClaves.FormattingEnabled = true;
            this.cboClaves.Location = new System.Drawing.Point(93, 63);
            this.cboClaves.Margin = new System.Windows.Forms.Padding(4);
            this.cboClaves.Name = "cboClaves";
            this.cboClaves.Size = new System.Drawing.Size(289, 27);
            this.cboClaves.Sorted = true;
            this.cboClaves.TabIndex = 64;
            this.cboClaves.Visible = false;
            // 
            // lblAnteriorClave
            // 
            this.lblAnteriorClave.AutoSize = true;
            this.lblAnteriorClave.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnteriorClave.ForeColor = System.Drawing.Color.White;
            this.lblAnteriorClave.Location = new System.Drawing.Point(21, 107);
            this.lblAnteriorClave.Name = "lblAnteriorClave";
            this.lblAnteriorClave.Size = new System.Drawing.Size(116, 19);
            this.lblAnteriorClave.TabIndex = 65;
            this.lblAnteriorClave.Text = "Clave Retirada:";
            this.lblAnteriorClave.Visible = false;
            // 
            // txtClaveRetirada
            // 
            this.txtClaveRetirada.Location = new System.Drawing.Point(143, 104);
            this.txtClaveRetirada.Name = "txtClaveRetirada";
            this.txtClaveRetirada.Size = new System.Drawing.Size(156, 27);
            this.txtClaveRetirada.TabIndex = 66;
            this.txtClaveRetirada.Visible = false;
            // 
            // erModulo
            // 
            this.erModulo.ContainerControl = this;
            // 
            // grpModulo
            // 
            this.grpModulo.Controls.Add(this.rtxtObsrevacion);
            this.grpModulo.Controls.Add(this.lblObservacion);
            this.grpModulo.Controls.Add(this.txtClaveRetirada);
            this.grpModulo.Controls.Add(this.lblAnteriorClave);
            this.grpModulo.Controls.Add(this.label15);
            this.grpModulo.Controls.Add(this.cboClaves);
            this.grpModulo.Controls.Add(this.cboModulos);
            this.grpModulo.Controls.Add(this.label14);
            this.grpModulo.Controls.Add(this.txtClave);
            this.grpModulo.ForeColor = System.Drawing.SystemColors.Control;
            this.grpModulo.Location = new System.Drawing.Point(12, 41);
            this.grpModulo.Name = "grpModulo";
            this.grpModulo.Size = new System.Drawing.Size(556, 282);
            this.grpModulo.TabIndex = 67;
            this.grpModulo.TabStop = false;
            this.grpModulo.Text = "Datos Modulo:";
            // 
            // rtxtObsrevacion
            // 
            this.rtxtObsrevacion.Location = new System.Drawing.Point(10, 170);
            this.rtxtObsrevacion.Name = "rtxtObsrevacion";
            this.rtxtObsrevacion.Size = new System.Drawing.Size(527, 96);
            this.rtxtObsrevacion.TabIndex = 68;
            this.rtxtObsrevacion.Text = "";
            // 
            // lblObservacion
            // 
            this.lblObservacion.AutoSize = true;
            this.lblObservacion.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblObservacion.ForeColor = System.Drawing.Color.White;
            this.lblObservacion.Location = new System.Drawing.Point(6, 148);
            this.lblObservacion.Name = "lblObservacion";
            this.lblObservacion.Size = new System.Drawing.Size(102, 19);
            this.lblObservacion.TabIndex = 67;
            this.lblObservacion.Text = "Observacion:";
            // 
            // ModuloNuevo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(28)))), ((int)(((byte)(48)))));
            this.ClientSize = new System.Drawing.Size(682, 392);
            this.Controls.Add(this.grpModulo);
            this.Controls.Add(this.btnAgregarModulo);
            this.Controls.Add(this.panelSuperior);
            this.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ModuloNuevo";
            this.Text = "ModuloNuevo";
            this.Load += new System.EventHandler(this.ModuloNuevo_Load);
            this.panelSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.erModulo)).EndInit();
            this.grpModulo.ResumeLayout(false);
            this.grpModulo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelSuperior;
        private System.Windows.Forms.PictureBox btnCerrar;
        private System.Windows.Forms.Button btnAgregarModulo;
        private System.Windows.Forms.ComboBox cboModulos;
        private System.Windows.Forms.TextBox txtClave;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox cboClaves;
        private System.Windows.Forms.Label lblAnteriorClave;
        private System.Windows.Forms.TextBox txtClaveRetirada;
        private System.Windows.Forms.ErrorProvider erModulo;
        private System.Windows.Forms.GroupBox grpModulo;
        private System.Windows.Forms.RichTextBox rtxtObsrevacion;
        private System.Windows.Forms.Label lblObservacion;
    }
}