namespace Perde_Evim
{
    partial class Parol
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
            this.txtParolGiris = new System.Windows.Forms.Button();
            this.txtParol = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtParolGiris
            // 
            this.txtParolGiris.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtParolGiris.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtParolGiris.Location = new System.Drawing.Point(39, 63);
            this.txtParolGiris.Name = "txtParolGiris";
            this.txtParolGiris.Size = new System.Drawing.Size(97, 24);
            this.txtParolGiris.TabIndex = 73;
            this.txtParolGiris.Text = "GİRİŞ";
            this.txtParolGiris.UseVisualStyleBackColor = true;
            this.txtParolGiris.Click += new System.EventHandler(this.TxtParolGiris_Click);
            // 
            // txtParol
            // 
            this.txtParol.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtParol.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtParol.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtParol.ForeColor = System.Drawing.Color.Red;
            this.txtParol.Location = new System.Drawing.Point(12, 22);
            this.txtParol.Name = "txtParol";
            this.txtParol.PasswordChar = '•';
            this.txtParol.Size = new System.Drawing.Size(149, 24);
            this.txtParol.TabIndex = 72;
            this.txtParol.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtParol_KeyDown);
            // 
            // Parol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(179, 112);
            this.Controls.Add(this.txtParolGiris);
            this.Controls.Add(this.txtParol);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Parol";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Təhlükəsizlik";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button txtParolGiris;
        private System.Windows.Forms.TextBox txtParol;
    }
}