namespace ExcelZugriff
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miLadenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miSpeichernUnterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.miBeendenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miErzeugenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miHilfeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.miÜberToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel2 = new System.Windows.Forms.Panel();
            this.tcWorksheets = new System.Windows.Forms.TabControl();
            this.panel3 = new System.Windows.Forms.Panel();
            this.cbWorksheets = new System.Windows.Forms.CheckedListBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.menuStrip1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(886, 62);
            this.panel1.TabIndex = 0;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 27);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(781, 20);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "z:\\2009.xlsx";
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(799, 24);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Analyse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuToolStripMenuItem,
            this.miErzeugenToolStripMenuItem,
            this.miHilfeToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(886, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.miLadenToolStripMenuItem,
            this.miSpeichernUnterToolStripMenuItem,
            this.toolStripMenuItem1,
            this.miBeendenToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            this.menuToolStripMenuItem.Text = "Datei";
            // 
            // miLadenToolStripMenuItem
            // 
            this.miLadenToolStripMenuItem.Name = "miLadenToolStripMenuItem";
            this.miLadenToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.miLadenToolStripMenuItem.Text = "Laden";
            this.miLadenToolStripMenuItem.Click += new System.EventHandler(this.miLadenToolStripMenuItem_Click);
            // 
            // miSpeichernUnterToolStripMenuItem
            // 
            this.miSpeichernUnterToolStripMenuItem.Name = "miSpeichernUnterToolStripMenuItem";
            this.miSpeichernUnterToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.miSpeichernUnterToolStripMenuItem.Text = "Speichern unter...";
            this.miSpeichernUnterToolStripMenuItem.Click += new System.EventHandler(this.miSpeichernUnterToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(163, 6);
            // 
            // miBeendenToolStripMenuItem
            // 
            this.miBeendenToolStripMenuItem.Name = "miBeendenToolStripMenuItem";
            this.miBeendenToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.miBeendenToolStripMenuItem.Text = "Beenden";
            // 
            // miErzeugenToolStripMenuItem
            // 
            this.miErzeugenToolStripMenuItem.Name = "miErzeugenToolStripMenuItem";
            this.miErzeugenToolStripMenuItem.Size = new System.Drawing.Size(67, 20);
            this.miErzeugenToolStripMenuItem.Text = "Erzeugen";
            this.miErzeugenToolStripMenuItem.Click += new System.EventHandler(this.miErzeugenToolStripMenuItem_Click);
            // 
            // miHilfeToolStripMenuItem
            // 
            this.miHilfeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.miÜberToolStripMenuItem});
            this.miHilfeToolStripMenuItem.Name = "miHilfeToolStripMenuItem";
            this.miHilfeToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.miHilfeToolStripMenuItem.Text = "Hilfe";
            this.miHilfeToolStripMenuItem.Click += new System.EventHandler(this.miHilfeToolStripMenuItem_Click);
            // 
            // miÜberToolStripMenuItem
            // 
            this.miÜberToolStripMenuItem.Name = "miÜberToolStripMenuItem";
            this.miÜberToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            this.miÜberToolStripMenuItem.Text = "Über";
            this.miÜberToolStripMenuItem.Click += new System.EventHandler(this.miÜberToolStripMenuItem_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tcWorksheets);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.textBox2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 62);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(886, 493);
            this.panel2.TabIndex = 2;
            // 
            // tcWorksheets
            // 
            this.tcWorksheets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tcWorksheets.Location = new System.Drawing.Point(221, 0);
            this.tcWorksheets.Name = "tcWorksheets";
            this.tcWorksheets.SelectedIndex = 0;
            this.tcWorksheets.Size = new System.Drawing.Size(665, 296);
            this.tcWorksheets.TabIndex = 2;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.cbWorksheets);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(221, 296);
            this.panel3.TabIndex = 1;
            // 
            // cbWorksheets
            // 
            this.cbWorksheets.CheckOnClick = true;
            this.cbWorksheets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cbWorksheets.FormattingEnabled = true;
            this.cbWorksheets.Location = new System.Drawing.Point(0, 0);
            this.cbWorksheets.Name = "cbWorksheets";
            this.cbWorksheets.Size = new System.Drawing.Size(221, 296);
            this.cbWorksheets.TabIndex = 0;
            this.cbWorksheets.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cbWorksheets_KeyDown);
            // 
            // textBox2
            // 
            this.textBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBox2.Location = new System.Drawing.Point(0, 296);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(886, 197);
            this.textBox2.TabIndex = 0;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(886, 555);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Leave += new System.EventHandler(this.Form1_Leave);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TabControl tcWorksheets;
        private System.Windows.Forms.CheckedListBox cbWorksheets;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem miLadenToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem miSpeichernUnterToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem miBeendenToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem miErzeugenToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem miHilfeToolStripMenuItem;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem miÜberToolStripMenuItem;
    }
}

