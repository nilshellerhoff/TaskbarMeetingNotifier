
namespace TaskbarMeetingNotifierUI
{
    partial class TaskbarMeetingNotifierUI
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.meetingLabel = new System.Windows.Forms.Label();
            this.showButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // meetingLabel
            // 
            this.meetingLabel.AutoEllipsis = true;
            this.meetingLabel.AutoSize = true;
            this.meetingLabel.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.meetingLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.meetingLabel.Location = new System.Drawing.Point(3, 9);
            this.meetingLabel.MaximumSize = new System.Drawing.Size(290, 0);
            this.meetingLabel.Name = "meetingLabel";
            this.meetingLabel.Size = new System.Drawing.Size(239, 25);
            this.meetingLabel.TabIndex = 0;
            this.meetingLabel.Text = "(not connected to Outlook)";
            // 
            // showButton
            // 
            this.showButton.Location = new System.Drawing.Point(304, 9);
            this.showButton.Name = "showButton";
            this.showButton.Size = new System.Drawing.Size(43, 25);
            this.showButton.TabIndex = 1;
            this.showButton.Text = "Show";
            this.showButton.UseVisualStyleBackColor = true;
            this.showButton.Click += new System.EventHandler(this.showButton_click);
            // 
            // TaskbarMeetingNotifierUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.Controls.Add(this.showButton);
            this.Controls.Add(this.meetingLabel);
            this.Name = "TaskbarMeetingNotifierUI";
            this.Size = new System.Drawing.Size(350, 42);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label meetingLabel;
        private System.Windows.Forms.Button showButton;
    }
}
