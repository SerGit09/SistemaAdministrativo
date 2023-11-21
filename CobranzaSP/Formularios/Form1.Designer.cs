using System.Windows.Media.TextFormatting;

namespace CobranzaSP
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.panelSuperior = new System.Windows.Forms.Panel();
            this.btnRestaurar = new System.Windows.Forms.PictureBox();
            this.btnMinimizar = new System.Windows.Forms.PictureBox();
            this.btnMaximizar = new System.Windows.Forms.PictureBox();
            this.btnCerrar = new System.Windows.Forms.PictureBox();
            this.pLateral = new System.Windows.Forms.Panel();
            this.SubMenuModulos = new System.Windows.Forms.Panel();
            this.btnModulosRenta = new System.Windows.Forms.Button();
            this.btnModuloBodega = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.SubMenuEquipos = new System.Windows.Forms.Panel();
            this.btnEquiposVendidos = new System.Windows.Forms.Button();
            this.btnRentaEquipo = new System.Windows.Forms.Button();
            this.btnBodega = new System.Windows.Forms.Button();
            this.btnEquipos = new System.Windows.Forms.Button();
            this.btnFusores = new System.Windows.Forms.Button();
            this.btnProveedores = new System.Windows.Forms.Button();
            this.SubMenuInventario = new System.Windows.Forms.Panel();
            this.btnModulos = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.btnInventarioToners = new System.Windows.Forms.Button();
            this.btnInventario = new System.Windows.Forms.Button();
            this.btnClientes = new System.Windows.Forms.Button();
            this.SubMenuServicios = new System.Windows.Forms.Panel();
            this.btnServiciosEquipos = new System.Windows.Forms.Button();
            this.btnServiciosRicoh = new System.Windows.Forms.Button();
            this.btnServicios = new System.Windows.Forms.Button();
            this.panelSubMenuCobranza = new System.Windows.Forms.Panel();
            this.btnCompras = new System.Windows.Forms.Button();
            this.btnCobranza = new System.Windows.Forms.Button();
            this.btnCobrado = new System.Windows.Forms.Button();
            this.btnGraficas = new System.Windows.Forms.Button();
            this.btnRemisiones = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panelCentral = new System.Windows.Forms.Panel();
            this.panelSuperior.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnRestaurar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnMinimizar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnMaximizar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).BeginInit();
            this.pLateral.SuspendLayout();
            this.SubMenuModulos.SuspendLayout();
            this.SubMenuEquipos.SuspendLayout();
            this.SubMenuInventario.SuspendLayout();
            this.SubMenuServicios.SuspendLayout();
            this.panelSubMenuCobranza.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelSuperior
            // 
            this.panelSuperior.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(117)))), ((int)(((byte)(20)))));
            this.panelSuperior.Controls.Add(this.btnRestaurar);
            this.panelSuperior.Controls.Add(this.btnMinimizar);
            this.panelSuperior.Controls.Add(this.btnMaximizar);
            this.panelSuperior.Controls.Add(this.btnCerrar);
            this.panelSuperior.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSuperior.Location = new System.Drawing.Point(0, 0);
            this.panelSuperior.Name = "panelSuperior";
            this.panelSuperior.Size = new System.Drawing.Size(1600, 43);
            this.panelSuperior.TabIndex = 0;
            this.panelSuperior.Paint += new System.Windows.Forms.PaintEventHandler(this.panelSuperior_Paint);
            this.panelSuperior.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panelSuperior_MouseDown);
            // 
            // btnRestaurar
            // 
            this.btnRestaurar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRestaurar.ErrorImage = null;
            this.btnRestaurar.Image = ((System.Drawing.Image)(resources.GetObject("btnRestaurar.Image")));
            this.btnRestaurar.InitialImage = null;
            this.btnRestaurar.Location = new System.Drawing.Point(1541, 11);
            this.btnRestaurar.Name = "btnRestaurar";
            this.btnRestaurar.Size = new System.Drawing.Size(25, 28);
            this.btnRestaurar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnRestaurar.TabIndex = 4;
            this.btnRestaurar.TabStop = false;
            this.btnRestaurar.Visible = false;
            this.btnRestaurar.Click += new System.EventHandler(this.btnRestaurar_Click);
            // 
            // btnMinimizar
            // 
            this.btnMinimizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMinimizar.ErrorImage = null;
            this.btnMinimizar.Image = ((System.Drawing.Image)(resources.GetObject("btnMinimizar.Image")));
            this.btnMinimizar.InitialImage = null;
            this.btnMinimizar.Location = new System.Drawing.Point(1510, 11);
            this.btnMinimizar.Name = "btnMinimizar";
            this.btnMinimizar.Size = new System.Drawing.Size(25, 28);
            this.btnMinimizar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnMinimizar.TabIndex = 3;
            this.btnMinimizar.TabStop = false;
            this.btnMinimizar.Click += new System.EventHandler(this.btnMinimizar_Click);
            // 
            // btnMaximizar
            // 
            this.btnMaximizar.ErrorImage = null;
            this.btnMaximizar.Image = ((System.Drawing.Image)(resources.GetObject("btnMaximizar.Image")));
            this.btnMaximizar.InitialImage = null;
            this.btnMaximizar.Location = new System.Drawing.Point(1541, 11);
            this.btnMaximizar.Name = "btnMaximizar";
            this.btnMaximizar.Size = new System.Drawing.Size(25, 28);
            this.btnMaximizar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnMaximizar.TabIndex = 2;
            this.btnMaximizar.TabStop = false;
            this.btnMaximizar.Click += new System.EventHandler(this.btnMaximizar_Click);
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.ErrorImage = null;
            this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
            this.btnCerrar.InitialImage = null;
            this.btnCerrar.Location = new System.Drawing.Point(1572, 11);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(25, 28);
            this.btnCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.btnCerrar.TabIndex = 1;
            this.btnCerrar.TabStop = false;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // pLateral
            // 
            this.pLateral.AutoScroll = true;
            this.pLateral.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(26)))), ((int)(((byte)(32)))), ((int)(((byte)(40)))));
            this.pLateral.Controls.Add(this.SubMenuModulos);
            this.pLateral.Controls.Add(this.button4);
            this.pLateral.Controls.Add(this.SubMenuEquipos);
            this.pLateral.Controls.Add(this.btnEquipos);
            this.pLateral.Controls.Add(this.btnFusores);
            this.pLateral.Controls.Add(this.btnProveedores);
            this.pLateral.Controls.Add(this.SubMenuInventario);
            this.pLateral.Controls.Add(this.btnInventario);
            this.pLateral.Controls.Add(this.btnClientes);
            this.pLateral.Controls.Add(this.SubMenuServicios);
            this.pLateral.Controls.Add(this.btnServicios);
            this.pLateral.Controls.Add(this.panelSubMenuCobranza);
            this.pLateral.Controls.Add(this.button1);
            this.pLateral.Controls.Add(this.panel1);
            this.pLateral.Dock = System.Windows.Forms.DockStyle.Left;
            this.pLateral.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.pLateral.Location = new System.Drawing.Point(0, 43);
            this.pLateral.Name = "pLateral";
            this.pLateral.Size = new System.Drawing.Size(250, 977);
            this.pLateral.TabIndex = 1;
            // 
            // SubMenuModulos
            // 
            this.SubMenuModulos.Controls.Add(this.btnModulosRenta);
            this.SubMenuModulos.Controls.Add(this.btnModuloBodega);
            this.SubMenuModulos.Dock = System.Windows.Forms.DockStyle.Top;
            this.SubMenuModulos.Location = new System.Drawing.Point(0, 1546);
            this.SubMenuModulos.Name = "SubMenuModulos";
            this.SubMenuModulos.Size = new System.Drawing.Size(233, 147);
            this.SubMenuModulos.TabIndex = 44;
            // 
            // btnModulosRenta
            // 
            this.btnModulosRenta.FlatAppearance.BorderSize = 0;
            this.btnModulosRenta.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnModulosRenta.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnModulosRenta.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnModulosRenta.ForeColor = System.Drawing.Color.White;
            this.btnModulosRenta.Image = ((System.Drawing.Image)(resources.GetObject("btnModulosRenta.Image")));
            this.btnModulosRenta.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnModulosRenta.Location = new System.Drawing.Point(22, 71);
            this.btnModulosRenta.Name = "btnModulosRenta";
            this.btnModulosRenta.Size = new System.Drawing.Size(250, 65);
            this.btnModulosRenta.TabIndex = 16;
            this.btnModulosRenta.Text = "Renta";
            this.btnModulosRenta.UseCompatibleTextRendering = true;
            this.btnModulosRenta.UseVisualStyleBackColor = true;
            // 
            // btnModuloBodega
            // 
            this.btnModuloBodega.FlatAppearance.BorderSize = 0;
            this.btnModuloBodega.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnModuloBodega.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnModuloBodega.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnModuloBodega.ForeColor = System.Drawing.Color.White;
            this.btnModuloBodega.Image = ((System.Drawing.Image)(resources.GetObject("btnModuloBodega.Image")));
            this.btnModuloBodega.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnModuloBodega.Location = new System.Drawing.Point(22, 6);
            this.btnModuloBodega.Name = "btnModuloBodega";
            this.btnModuloBodega.Size = new System.Drawing.Size(250, 65);
            this.btnModuloBodega.TabIndex = 17;
            this.btnModuloBodega.Text = "Bodega";
            this.btnModuloBodega.UseCompatibleTextRendering = true;
            this.btnModuloBodega.UseVisualStyleBackColor = true;
            this.btnModuloBodega.Click += new System.EventHandler(this.btnModuloBodega_Click);
            // 
            // button4
            // 
            this.button4.Dock = System.Windows.Forms.DockStyle.Top;
            this.button4.FlatAppearance.BorderSize = 0;
            this.button4.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.ForeColor = System.Drawing.Color.White;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.Location = new System.Drawing.Point(0, 1481);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(233, 65);
            this.button4.TabIndex = 43;
            this.button4.Text = "Modulos";
            this.button4.UseCompatibleTextRendering = true;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // SubMenuEquipos
            // 
            this.SubMenuEquipos.Controls.Add(this.btnEquiposVendidos);
            this.SubMenuEquipos.Controls.Add(this.btnRentaEquipo);
            this.SubMenuEquipos.Controls.Add(this.btnBodega);
            this.SubMenuEquipos.Dock = System.Windows.Forms.DockStyle.Top;
            this.SubMenuEquipos.Location = new System.Drawing.Point(0, 1269);
            this.SubMenuEquipos.Name = "SubMenuEquipos";
            this.SubMenuEquipos.Size = new System.Drawing.Size(233, 212);
            this.SubMenuEquipos.TabIndex = 42;
            // 
            // btnEquiposVendidos
            // 
            this.btnEquiposVendidos.FlatAppearance.BorderSize = 0;
            this.btnEquiposVendidos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnEquiposVendidos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEquiposVendidos.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEquiposVendidos.ForeColor = System.Drawing.Color.White;
            this.btnEquiposVendidos.Image = ((System.Drawing.Image)(resources.GetObject("btnEquiposVendidos.Image")));
            this.btnEquiposVendidos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEquiposVendidos.Location = new System.Drawing.Point(22, 132);
            this.btnEquiposVendidos.Name = "btnEquiposVendidos";
            this.btnEquiposVendidos.Size = new System.Drawing.Size(250, 65);
            this.btnEquiposVendidos.TabIndex = 18;
            this.btnEquiposVendidos.Text = "Vendidos";
            this.btnEquiposVendidos.UseCompatibleTextRendering = true;
            this.btnEquiposVendidos.UseVisualStyleBackColor = true;
            this.btnEquiposVendidos.Click += new System.EventHandler(this.btnEquiposVendidos_Click);
            // 
            // btnRentaEquipo
            // 
            this.btnRentaEquipo.FlatAppearance.BorderSize = 0;
            this.btnRentaEquipo.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnRentaEquipo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRentaEquipo.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRentaEquipo.ForeColor = System.Drawing.Color.White;
            this.btnRentaEquipo.Image = ((System.Drawing.Image)(resources.GetObject("btnRentaEquipo.Image")));
            this.btnRentaEquipo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRentaEquipo.Location = new System.Drawing.Point(22, 71);
            this.btnRentaEquipo.Name = "btnRentaEquipo";
            this.btnRentaEquipo.Size = new System.Drawing.Size(250, 65);
            this.btnRentaEquipo.TabIndex = 16;
            this.btnRentaEquipo.Text = "Renta";
            this.btnRentaEquipo.UseCompatibleTextRendering = true;
            this.btnRentaEquipo.UseVisualStyleBackColor = true;
            this.btnRentaEquipo.Click += new System.EventHandler(this.btnRentaEquipo_Click);
            // 
            // btnBodega
            // 
            this.btnBodega.FlatAppearance.BorderSize = 0;
            this.btnBodega.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnBodega.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBodega.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBodega.ForeColor = System.Drawing.Color.White;
            this.btnBodega.Image = ((System.Drawing.Image)(resources.GetObject("btnBodega.Image")));
            this.btnBodega.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnBodega.Location = new System.Drawing.Point(22, 6);
            this.btnBodega.Name = "btnBodega";
            this.btnBodega.Size = new System.Drawing.Size(250, 65);
            this.btnBodega.TabIndex = 17;
            this.btnBodega.Text = "Bodega";
            this.btnBodega.UseCompatibleTextRendering = true;
            this.btnBodega.UseVisualStyleBackColor = true;
            this.btnBodega.Click += new System.EventHandler(this.btnBodega_Click);
            // 
            // btnEquipos
            // 
            this.btnEquipos.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnEquipos.FlatAppearance.BorderSize = 0;
            this.btnEquipos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnEquipos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEquipos.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEquipos.ForeColor = System.Drawing.Color.White;
            this.btnEquipos.Image = ((System.Drawing.Image)(resources.GetObject("btnEquipos.Image")));
            this.btnEquipos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEquipos.Location = new System.Drawing.Point(0, 1204);
            this.btnEquipos.Name = "btnEquipos";
            this.btnEquipos.Size = new System.Drawing.Size(233, 65);
            this.btnEquipos.TabIndex = 41;
            this.btnEquipos.Text = "Equipos";
            this.btnEquipos.UseCompatibleTextRendering = true;
            this.btnEquipos.UseVisualStyleBackColor = true;
            this.btnEquipos.Click += new System.EventHandler(this.btnEquipos_Click);
            // 
            // btnFusores
            // 
            this.btnFusores.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnFusores.FlatAppearance.BorderSize = 0;
            this.btnFusores.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnFusores.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFusores.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFusores.ForeColor = System.Drawing.Color.White;
            this.btnFusores.Image = ((System.Drawing.Image)(resources.GetObject("btnFusores.Image")));
            this.btnFusores.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFusores.Location = new System.Drawing.Point(0, 1139);
            this.btnFusores.Name = "btnFusores";
            this.btnFusores.Size = new System.Drawing.Size(233, 65);
            this.btnFusores.TabIndex = 40;
            this.btnFusores.Text = "Fusores";
            this.btnFusores.UseCompatibleTextRendering = true;
            this.btnFusores.UseVisualStyleBackColor = true;
            this.btnFusores.Click += new System.EventHandler(this.btnFusores_Click);
            // 
            // btnProveedores
            // 
            this.btnProveedores.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnProveedores.FlatAppearance.BorderSize = 0;
            this.btnProveedores.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnProveedores.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProveedores.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProveedores.ForeColor = System.Drawing.Color.White;
            this.btnProveedores.Image = ((System.Drawing.Image)(resources.GetObject("btnProveedores.Image")));
            this.btnProveedores.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnProveedores.Location = new System.Drawing.Point(0, 1074);
            this.btnProveedores.Name = "btnProveedores";
            this.btnProveedores.Size = new System.Drawing.Size(233, 65);
            this.btnProveedores.TabIndex = 39;
            this.btnProveedores.Text = "Proveedores";
            this.btnProveedores.UseVisualStyleBackColor = true;
            this.btnProveedores.Click += new System.EventHandler(this.btnProveedores_Click);
            // 
            // SubMenuInventario
            // 
            this.SubMenuInventario.Controls.Add(this.btnModulos);
            this.SubMenuInventario.Controls.Add(this.button2);
            this.SubMenuInventario.Controls.Add(this.btnInventarioToners);
            this.SubMenuInventario.Dock = System.Windows.Forms.DockStyle.Top;
            this.SubMenuInventario.Location = new System.Drawing.Point(0, 877);
            this.SubMenuInventario.Name = "SubMenuInventario";
            this.SubMenuInventario.Size = new System.Drawing.Size(233, 197);
            this.SubMenuInventario.TabIndex = 38;
            // 
            // btnModulos
            // 
            this.btnModulos.FlatAppearance.BorderSize = 0;
            this.btnModulos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnModulos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnModulos.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnModulos.ForeColor = System.Drawing.Color.White;
            this.btnModulos.Image = ((System.Drawing.Image)(resources.GetObject("btnModulos.Image")));
            this.btnModulos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnModulos.Location = new System.Drawing.Point(24, 126);
            this.btnModulos.Name = "btnModulos";
            this.btnModulos.Size = new System.Drawing.Size(229, 65);
            this.btnModulos.TabIndex = 4;
            this.btnModulos.Text = "Modulos";
            this.btnModulos.UseVisualStyleBackColor = true;
            this.btnModulos.Click += new System.EventHandler(this.btnModulos_Click);
            // 
            // button2
            // 
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(21, 64);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(229, 65);
            this.button2.TabIndex = 2;
            this.button2.Text = "Partes Ricoh";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnInventarioToners
            // 
            this.btnInventarioToners.FlatAppearance.BorderSize = 0;
            this.btnInventarioToners.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnInventarioToners.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInventarioToners.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInventarioToners.ForeColor = System.Drawing.Color.White;
            this.btnInventarioToners.Image = ((System.Drawing.Image)(resources.GetObject("btnInventarioToners.Image")));
            this.btnInventarioToners.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInventarioToners.Location = new System.Drawing.Point(21, 3);
            this.btnInventarioToners.Name = "btnInventarioToners";
            this.btnInventarioToners.Size = new System.Drawing.Size(229, 65);
            this.btnInventarioToners.TabIndex = 1;
            this.btnInventarioToners.Text = "Toners";
            this.btnInventarioToners.UseVisualStyleBackColor = true;
            this.btnInventarioToners.Click += new System.EventHandler(this.btnInventarioToners_Click);
            // 
            // btnInventario
            // 
            this.btnInventario.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnInventario.FlatAppearance.BorderSize = 0;
            this.btnInventario.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnInventario.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInventario.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInventario.ForeColor = System.Drawing.Color.White;
            this.btnInventario.Image = ((System.Drawing.Image)(resources.GetObject("btnInventario.Image")));
            this.btnInventario.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInventario.Location = new System.Drawing.Point(0, 812);
            this.btnInventario.Name = "btnInventario";
            this.btnInventario.Size = new System.Drawing.Size(233, 65);
            this.btnInventario.TabIndex = 37;
            this.btnInventario.Text = "Inventario";
            this.btnInventario.UseCompatibleTextRendering = true;
            this.btnInventario.UseVisualStyleBackColor = true;
            this.btnInventario.Click += new System.EventHandler(this.btnInventario_Click);
            // 
            // btnClientes
            // 
            this.btnClientes.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnClientes.FlatAppearance.BorderSize = 0;
            this.btnClientes.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnClientes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClientes.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClientes.ForeColor = System.Drawing.Color.White;
            this.btnClientes.Image = ((System.Drawing.Image)(resources.GetObject("btnClientes.Image")));
            this.btnClientes.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnClientes.Location = new System.Drawing.Point(0, 747);
            this.btnClientes.Name = "btnClientes";
            this.btnClientes.Size = new System.Drawing.Size(233, 65);
            this.btnClientes.TabIndex = 30;
            this.btnClientes.Text = "Clientes";
            this.btnClientes.UseVisualStyleBackColor = true;
            this.btnClientes.Click += new System.EventHandler(this.btnClientes_Click);
            // 
            // SubMenuServicios
            // 
            this.SubMenuServicios.Controls.Add(this.btnServiciosEquipos);
            this.SubMenuServicios.Controls.Add(this.btnServiciosRicoh);
            this.SubMenuServicios.Dock = System.Windows.Forms.DockStyle.Top;
            this.SubMenuServicios.Location = new System.Drawing.Point(0, 616);
            this.SubMenuServicios.Name = "SubMenuServicios";
            this.SubMenuServicios.Size = new System.Drawing.Size(233, 131);
            this.SubMenuServicios.TabIndex = 21;
            // 
            // btnServiciosEquipos
            // 
            this.btnServiciosEquipos.FlatAppearance.BorderSize = 0;
            this.btnServiciosEquipos.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnServiciosEquipos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnServiciosEquipos.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnServiciosEquipos.ForeColor = System.Drawing.Color.White;
            this.btnServiciosEquipos.Image = ((System.Drawing.Image)(resources.GetObject("btnServiciosEquipos.Image")));
            this.btnServiciosEquipos.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnServiciosEquipos.Location = new System.Drawing.Point(21, 64);
            this.btnServiciosEquipos.Name = "btnServiciosEquipos";
            this.btnServiciosEquipos.Size = new System.Drawing.Size(229, 65);
            this.btnServiciosEquipos.TabIndex = 2;
            this.btnServiciosEquipos.Text = "Otros";
            this.btnServiciosEquipos.UseVisualStyleBackColor = true;
            this.btnServiciosEquipos.Click += new System.EventHandler(this.btnServiciosEquipos_Click);
            // 
            // btnServiciosRicoh
            // 
            this.btnServiciosRicoh.FlatAppearance.BorderSize = 0;
            this.btnServiciosRicoh.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnServiciosRicoh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnServiciosRicoh.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnServiciosRicoh.ForeColor = System.Drawing.Color.White;
            this.btnServiciosRicoh.Image = ((System.Drawing.Image)(resources.GetObject("btnServiciosRicoh.Image")));
            this.btnServiciosRicoh.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnServiciosRicoh.Location = new System.Drawing.Point(21, 3);
            this.btnServiciosRicoh.Name = "btnServiciosRicoh";
            this.btnServiciosRicoh.Size = new System.Drawing.Size(229, 65);
            this.btnServiciosRicoh.TabIndex = 1;
            this.btnServiciosRicoh.Text = "Ricoh";
            this.btnServiciosRicoh.UseVisualStyleBackColor = true;
            this.btnServiciosRicoh.Click += new System.EventHandler(this.btnServiciosRicoh_Click);
            // 
            // btnServicios
            // 
            this.btnServicios.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnServicios.FlatAppearance.BorderSize = 0;
            this.btnServicios.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnServicios.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnServicios.Font = new System.Drawing.Font("Century Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnServicios.ForeColor = System.Drawing.Color.White;
            this.btnServicios.Image = ((System.Drawing.Image)(resources.GetObject("btnServicios.Image")));
            this.btnServicios.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnServicios.Location = new System.Drawing.Point(0, 551);
            this.btnServicios.Name = "btnServicios";
            this.btnServicios.Size = new System.Drawing.Size(233, 65);
            this.btnServicios.TabIndex = 20;
            this.btnServicios.Text = "Servicios";
            this.btnServicios.UseCompatibleTextRendering = true;
            this.btnServicios.UseVisualStyleBackColor = true;
            this.btnServicios.Click += new System.EventHandler(this.btnServicios_Click);
            // 
            // panelSubMenuCobranza
            // 
            this.panelSubMenuCobranza.Controls.Add(this.btnCompras);
            this.panelSubMenuCobranza.Controls.Add(this.btnCobranza);
            this.panelSubMenuCobranza.Controls.Add(this.btnCobrado);
            this.panelSubMenuCobranza.Controls.Add(this.btnGraficas);
            this.panelSubMenuCobranza.Controls.Add(this.btnRemisiones);
            this.panelSubMenuCobranza.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSubMenuCobranza.Location = new System.Drawing.Point(0, 223);
            this.panelSubMenuCobranza.Name = "panelSubMenuCobranza";
            this.panelSubMenuCobranza.Size = new System.Drawing.Size(233, 328);
            this.panelSubMenuCobranza.TabIndex = 19;
            // 
            // btnCompras
            // 
            this.btnCompras.FlatAppearance.BorderSize = 0;
            this.btnCompras.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnCompras.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCompras.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCompras.ForeColor = System.Drawing.Color.White;
            this.btnCompras.Image = ((System.Drawing.Image)(resources.GetObject("btnCompras.Image")));
            this.btnCompras.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCompras.Location = new System.Drawing.Point(12, 255);
            this.btnCompras.Name = "btnCompras";
            this.btnCompras.Size = new System.Drawing.Size(250, 65);
            this.btnCompras.TabIndex = 5;
            this.btnCompras.Text = "Compras";
            this.btnCompras.UseVisualStyleBackColor = true;
            this.btnCompras.Click += new System.EventHandler(this.btnCompras_Click);
            // 
            // btnCobranza
            // 
            this.btnCobranza.FlatAppearance.BorderSize = 0;
            this.btnCobranza.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnCobranza.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCobranza.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCobranza.ForeColor = System.Drawing.Color.White;
            this.btnCobranza.Image = ((System.Drawing.Image)(resources.GetObject("btnCobranza.Image")));
            this.btnCobranza.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCobranza.Location = new System.Drawing.Point(12, 3);
            this.btnCobranza.Name = "btnCobranza";
            this.btnCobranza.Size = new System.Drawing.Size(250, 65);
            this.btnCobranza.TabIndex = 1;
            this.btnCobranza.Text = "Facturacion";
            this.btnCobranza.UseVisualStyleBackColor = true;
            this.btnCobranza.Click += new System.EventHandler(this.btnCobranza_Click);
            // 
            // btnCobrado
            // 
            this.btnCobrado.FlatAppearance.BorderSize = 0;
            this.btnCobrado.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnCobrado.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCobrado.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCobrado.ForeColor = System.Drawing.Color.White;
            this.btnCobrado.Image = ((System.Drawing.Image)(resources.GetObject("btnCobrado.Image")));
            this.btnCobrado.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCobrado.Location = new System.Drawing.Point(12, 64);
            this.btnCobrado.Name = "btnCobrado";
            this.btnCobrado.Size = new System.Drawing.Size(250, 65);
            this.btnCobrado.TabIndex = 2;
            this.btnCobrado.Text = "Cobrado";
            this.btnCobrado.UseVisualStyleBackColor = true;
            this.btnCobrado.Click += new System.EventHandler(this.btnCobrado_Click);
            // 
            // btnGraficas
            // 
            this.btnGraficas.FlatAppearance.BorderSize = 0;
            this.btnGraficas.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnGraficas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGraficas.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGraficas.ForeColor = System.Drawing.Color.White;
            this.btnGraficas.Image = ((System.Drawing.Image)(resources.GetObject("btnGraficas.Image")));
            this.btnGraficas.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnGraficas.Location = new System.Drawing.Point(12, 184);
            this.btnGraficas.Name = "btnGraficas";
            this.btnGraficas.Size = new System.Drawing.Size(250, 65);
            this.btnGraficas.TabIndex = 4;
            this.btnGraficas.Text = "Gráficas";
            this.btnGraficas.UseVisualStyleBackColor = true;
            this.btnGraficas.Click += new System.EventHandler(this.btnGraficas_Click);
            // 
            // btnRemisiones
            // 
            this.btnRemisiones.FlatAppearance.BorderSize = 0;
            this.btnRemisiones.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.btnRemisiones.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRemisiones.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRemisiones.ForeColor = System.Drawing.Color.White;
            this.btnRemisiones.Image = ((System.Drawing.Image)(resources.GetObject("btnRemisiones.Image")));
            this.btnRemisiones.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRemisiones.Location = new System.Drawing.Point(21, 126);
            this.btnRemisiones.Name = "btnRemisiones";
            this.btnRemisiones.Size = new System.Drawing.Size(232, 65);
            this.btnRemisiones.TabIndex = 3;
            this.btnRemisiones.Text = "Remisiones y mostrador";
            this.btnRemisiones.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnRemisiones.UseVisualStyleBackColor = true;
            this.btnRemisiones.Click += new System.EventHandler(this.btnRemisiones_Click);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Top;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(80)))), ((int)(((byte)(200)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Century Gothic", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(0, 158);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(233, 65);
            this.button1.TabIndex = 18;
            this.button1.Text = "Cobranza";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // panel1
            // 
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Controls.Add(this.panel3);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(233, 158);
            this.panel1.TabIndex = 17;
            // 
            // panel3
            // 
            this.panel3.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("panel3.BackgroundImage")));
            this.panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel3.Location = new System.Drawing.Point(3, 6);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(241, 138);
            this.panel3.TabIndex = 0;
            // 
            // panelCentral
            // 
            this.panelCentral.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(49)))), ((int)(((byte)(69)))), ((int)(((byte)(82)))));
            this.panelCentral.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentral.Location = new System.Drawing.Point(250, 43);
            this.panelCentral.Name = "panelCentral";
            this.panelCentral.Size = new System.Drawing.Size(1350, 977);
            this.panelCentral.TabIndex = 2;
            this.panelCentral.Paint += new System.Windows.Forms.PaintEventHandler(this.panelCentral_Paint);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1600, 1020);
            this.Controls.Add(this.panelCentral);
            this.Controls.Add(this.pLateral);
            this.Controls.Add(this.panelSuperior);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.Text = "Form1";
            this.panelSuperior.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.btnRestaurar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnMinimizar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnMaximizar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCerrar)).EndInit();
            this.pLateral.ResumeLayout(false);
            this.SubMenuModulos.ResumeLayout(false);
            this.SubMenuEquipos.ResumeLayout(false);
            this.SubMenuInventario.ResumeLayout(false);
            this.SubMenuServicios.ResumeLayout(false);
            this.panelSubMenuCobranza.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelSuperior;
        private System.Windows.Forms.PictureBox btnRestaurar;
        private System.Windows.Forms.PictureBox btnMinimizar;
        private System.Windows.Forms.PictureBox btnMaximizar;
        private System.Windows.Forms.PictureBox btnCerrar;
        private System.Windows.Forms.Panel pLateral;
        private System.Windows.Forms.Panel panelCentral;
        private System.Windows.Forms.Button btnCobranza;
        private System.Windows.Forms.Button btnRemisiones;
        private System.Windows.Forms.Button btnCobrado;
        private System.Windows.Forms.Button btnGraficas;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panelSubMenuCobranza;
        private System.Windows.Forms.Button btnServicios;
        private System.Windows.Forms.Button btnCompras;
        private System.Windows.Forms.Panel SubMenuServicios;
        private System.Windows.Forms.Button btnServiciosRicoh;
        private System.Windows.Forms.Button btnServiciosEquipos;
        private System.Windows.Forms.Button btnClientes;
        private System.Windows.Forms.Panel SubMenuInventario;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnInventarioToners;
        private System.Windows.Forms.Button btnInventario;
        private System.Windows.Forms.Button btnProveedores;
        private System.Windows.Forms.Button btnFusores;
        private System.Windows.Forms.Button btnEquipos;
        private System.Windows.Forms.Button btnModulos;
        private System.Windows.Forms.Panel SubMenuEquipos;
        private System.Windows.Forms.Button btnEquiposVendidos;
        private System.Windows.Forms.Button btnRentaEquipo;
        private System.Windows.Forms.Button btnBodega;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Panel SubMenuModulos;
        private System.Windows.Forms.Button btnModulosRenta;
        private System.Windows.Forms.Button btnModuloBodega;
    }
}

