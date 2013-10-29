namespace Direct_Intervention
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.Process = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtConsole = new System.Windows.Forms.TextBox();
            this.Quit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Process
            // 
            this.Process.Location = new System.Drawing.Point(12, 212);
            this.Process.Name = "Process";
            this.Process.Size = new System.Drawing.Size(169, 45);
            this.Process.TabIndex = 0;
            this.Process.Text = "Process";
            this.Process.UseVisualStyleBackColor = true;
            this.Process.Click += new System.EventHandler(this.Process_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(56, -4);
            this.label1.MinimumSize = new System.Drawing.Size(50, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 50);
            this.label1.TabIndex = 1;
            this.label1.Text = "Drag And Drop Excel File To Begin";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtConsole
            // 
            this.txtConsole.Location = new System.Drawing.Point(11, 43);
            this.txtConsole.Multiline = true;
            this.txtConsole.Name = "txtConsole";
            this.txtConsole.Size = new System.Drawing.Size(261, 163);
            this.txtConsole.TabIndex = 2;
            // 
            // Quit
            // 
            this.Quit.Location = new System.Drawing.Point(187, 212);
            this.Quit.Name = "Quit";
            this.Quit.Size = new System.Drawing.Size(86, 45);
            this.Quit.TabIndex = 3;
            this.Quit.Text = "Quit";
            this.Quit.UseVisualStyleBackColor = true;
            this.Quit.Click += new System.EventHandler(this.Quit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.ControlBox = false;
            this.Controls.Add(this.Quit);
            this.Controls.Add(this.txtConsole);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Process);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Direct Intervention";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Process;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtConsole;
        private System.Windows.Forms.Button Quit;
    }
}

