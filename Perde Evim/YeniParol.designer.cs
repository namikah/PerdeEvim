namespace Perde_Evim
{
    partial class YeniParol
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
            this.btEnter = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtYeniParol = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtUserName = new System.Windows.Forms.TextBox();
            this.label43 = new System.Windows.Forms.Label();
            this.txtHazirkiParol = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btEnter
            // 
            this.btEnter.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btEnter.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btEnter.Location = new System.Drawing.Point(126, 107);
            this.btEnter.Name = "btEnter";
            this.btEnter.Size = new System.Drawing.Size(97, 24);
            this.btEnter.TabIndex = 86;
            this.btEnter.Text = "Yadda Saxla";
            this.btEnter.UseVisualStyleBackColor = true;
            this.btEnter.Click += new System.EventHandler(this.BtEnter_Click);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(38, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 85;
            this.label2.Text = "Yeni Parol";
            // 
            // txtYeniParol
            // 
            this.txtYeniParol.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtYeniParol.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtYeniParol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtYeniParol.Location = new System.Drawing.Point(99, 70);
            this.txtYeniParol.Name = "txtYeniParol";
            this.txtYeniParol.PasswordChar = '•';
            this.txtYeniParol.Size = new System.Drawing.Size(157, 20);
            this.txtYeniParol.TabIndex = 84;
            this.txtYeniParol.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtYeniParol_KeyDown);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(26, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 83;
            this.label1.Text = "Istifadəçi Adı";
            // 
            // txtUserName
            // 
            this.txtUserName.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtUserName.Enabled = false;
            this.txtUserName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtUserName.Location = new System.Drawing.Point(99, 44);
            this.txtUserName.Name = "txtUserName";
            this.txtUserName.Size = new System.Drawing.Size(157, 20);
            this.txtUserName.TabIndex = 82;
            // 
            // label43
            // 
            this.label43.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label43.AutoSize = true;
            this.label43.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label43.Location = new System.Drawing.Point(27, 20);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(66, 13);
            this.label43.TabIndex = 81;
            this.label43.Text = "Hazırki Parol";
            // 
            // txtHazirkiParol
            // 
            this.txtHazirkiParol.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtHazirkiParol.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtHazirkiParol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtHazirkiParol.Location = new System.Drawing.Point(99, 18);
            this.txtHazirkiParol.Name = "txtHazirkiParol";
            this.txtHazirkiParol.PasswordChar = '•';
            this.txtHazirkiParol.Size = new System.Drawing.Size(157, 20);
            this.txtHazirkiParol.TabIndex = 80;
            // 
            // YeniParol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(313, 149);
            this.Controls.Add(this.btEnter);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtYeniParol);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtUserName);
            this.Controls.Add(this.label43);
            this.Controls.Add(this.txtHazirkiParol);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "YeniParol";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Parol";
            this.Load += new System.EventHandler(this.YeniParol_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btEnter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtYeniParol;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtUserName;
        private System.Windows.Forms.Label label43;
        private System.Windows.Forms.TextBox txtHazirkiParol;
    }
}