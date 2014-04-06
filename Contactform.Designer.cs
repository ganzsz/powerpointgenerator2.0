namespace PowerpointGenerater2
{
    partial class Contactform
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
            this.lblBuild = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblBuild
            // 
            this.lblBuild.Location = new System.Drawing.Point(49, 332);
            this.lblBuild.Name = "lblBuild";
            this.lblBuild.Size = new System.Drawing.Size(532, 23);
            this.lblBuild.TabIndex = 0;
            this.lblBuild.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Contactform
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(671, 384);
            this.Controls.Add(this.lblBuild);
            this.MaximumSize = new System.Drawing.Size(687, 422);
            this.MinimumSize = new System.Drawing.Size(687, 422);
            this.Name = "Contactform";
            this.ShowIcon = false;
            this.Text = "Visitekaartje";
            this.Load += new System.EventHandler(this.Contactform_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblBuild;

    }
}