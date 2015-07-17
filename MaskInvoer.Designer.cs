namespace PowerpointGenerater2
{
    partial class MaskInvoer
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MaskInvoer));
            this.TxtBoxRealName = new System.Windows.Forms.TextBox();
            this.TxtBoxVirtualName = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.Masks = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.type = new System.Windows.Forms.GroupBox();
            this.tekst = new System.Windows.Forms.RadioButton();
            this.lezen = new System.Windows.Forms.RadioButton();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.Masks.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.type.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtBoxRealName
            // 
            this.TxtBoxRealName.Location = new System.Drawing.Point(7, 30);
            this.TxtBoxRealName.Name = "TxtBoxRealName";
            this.TxtBoxRealName.Size = new System.Drawing.Size(98, 20);
            this.TxtBoxRealName.TabIndex = 0;
            // 
            // TxtBoxVirtualName
            // 
            this.TxtBoxVirtualName.Location = new System.Drawing.Point(7, 81);
            this.TxtBoxVirtualName.Name = "TxtBoxVirtualName";
            this.TxtBoxVirtualName.Size = new System.Drawing.Size(98, 20);
            this.TxtBoxVirtualName.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(111, 14);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(64, 26);
            this.button1.TabIndex = 2;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Echte naam";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Virtuele naam";
            // 
            // listBox1
            // 
            this.listBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(181, 6);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(195, 95);
            this.listBox1.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(272, 159);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(62, 30);
            this.button2.TabIndex = 6;
            this.button2.Text = "Annuleer";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button3.Location = new System.Drawing.Point(340, 159);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(62, 30);
            this.button3.TabIndex = 7;
            this.button3.Text = "Save";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(111, 43);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(64, 26);
            this.button4.TabIndex = 8;
            this.button4.Text = "Wijzig";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(111, 75);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(64, 26);
            this.button5.TabIndex = 9;
            this.button5.Text = "Verwijder";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Masks
            // 
            this.Masks.Controls.Add(this.tabPage1);
            this.Masks.Controls.Add(this.tabPage2);
            this.Masks.Location = new System.Drawing.Point(12, 12);
            this.Masks.Name = "Masks";
            this.Masks.SelectedIndex = 0;
            this.Masks.Size = new System.Drawing.Size(390, 141);
            this.Masks.TabIndex = 10;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.listBox1);
            this.tabPage1.Controls.Add(this.button5);
            this.tabPage1.Controls.Add(this.TxtBoxRealName);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Controls.Add(this.TxtBoxVirtualName);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(382, 115);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Masks";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.textBox1);
            this.tabPage2.Controls.Add(this.type);
            this.tabPage2.Controls.Add(this.listBox2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(382, 115);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Lezen";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 3);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(38, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Naam:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(108, 20);
            this.textBox1.TabIndex = 8;
            // 
            // type
            // 
            this.type.Controls.Add(this.tekst);
            this.type.Controls.Add(this.lezen);
            this.type.Location = new System.Drawing.Point(6, 42);
            this.type.Name = "type";
            this.type.Size = new System.Drawing.Size(108, 61);
            this.type.TabIndex = 7;
            this.type.TabStop = false;
            this.type.Text = "type";
            // 
            // tekst
            // 
            this.tekst.AutoSize = true;
            this.tekst.Location = new System.Drawing.Point(7, 38);
            this.tekst.Name = "tekst";
            this.tekst.Size = new System.Drawing.Size(52, 17);
            this.tekst.TabIndex = 1;
            this.tekst.TabStop = true;
            this.tekst.Text = "Tekst";
            this.tekst.UseVisualStyleBackColor = true;
            // 
            // lezen
            // 
            this.lezen.AutoSize = true;
            this.lezen.Location = new System.Drawing.Point(7, 20);
            this.lezen.Name = "lezen";
            this.lezen.Size = new System.Drawing.Size(54, 17);
            this.lezen.TabIndex = 0;
            this.lezen.TabStop = true;
            this.lezen.Text = "Lezen";
            this.lezen.UseVisualStyleBackColor = true;
            // 
            // listBox2
            // 
            this.listBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(186, 8);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(190, 95);
            this.listBox2.TabIndex = 6;
            // 
            // MaskInvoer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 199);
            this.Controls.Add(this.Masks);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MaskInvoer";
            this.Text = "MaskInvoer";
            this.Masks.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.type.ResumeLayout(false);
            this.type.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox TxtBoxRealName;
        private System.Windows.Forms.TextBox TxtBoxVirtualName;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        public System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.TabControl Masks;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.GroupBox type;
        private System.Windows.Forms.RadioButton tekst;
        private System.Windows.Forms.RadioButton lezen;
        public System.Windows.Forms.ListBox listBox2;
    }
}