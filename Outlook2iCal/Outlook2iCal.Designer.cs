namespace Outlook2iCal
{
    partial class Outlook2iCal
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Outlook2iCal));
            this.eventLabel = new System.Windows.Forms.Label();
            this.exceptLabel = new System.Windows.Forms.Label();
            this.startButton = new System.Windows.Forms.Button();
            this.eventText = new System.Windows.Forms.Label();
            this.exceptText = new System.Windows.Forms.Label();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.currentLabel = new System.Windows.Forms.Label();
            this.currentBox = new ReadOnlyTextBox();
            this.exceptBar = new EventProgressBar();
            this.eventBar = new EventProgressBar();
            this.SuspendLayout();
            // 
            // eventLabel
            // 
            this.eventLabel.AutoSize = true;
            this.eventLabel.Location = new System.Drawing.Point(12, 9);
            this.eventLabel.Name = "eventLabel";
            this.eventLabel.Size = new System.Drawing.Size(43, 13);
            this.eventLabel.TabIndex = 2;
            this.eventLabel.Text = "Events:";
            // 
            // exceptLabel
            // 
            this.exceptLabel.AutoSize = true;
            this.exceptLabel.Location = new System.Drawing.Point(12, 51);
            this.exceptLabel.Name = "exceptLabel";
            this.exceptLabel.Size = new System.Drawing.Size(62, 13);
            this.exceptLabel.TabIndex = 3;
            this.exceptLabel.Text = "Exceptions:";
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(179, 135);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(75, 23);
            this.startButton.TabIndex = 6;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // eventText
            // 
            this.eventText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.eventText.AutoSize = true;
            this.eventText.Location = new System.Drawing.Point(318, 9);
            this.eventText.MinimumSize = new System.Drawing.Size(100, 13);
            this.eventText.Name = "eventText";
            this.eventText.Size = new System.Drawing.Size(100, 13);
            this.eventText.TabIndex = 7;
            this.eventText.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // exceptText
            // 
            this.exceptText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.exceptText.AutoSize = true;
            this.exceptText.Location = new System.Drawing.Point(318, 51);
            this.exceptText.MinimumSize = new System.Drawing.Size(100, 13);
            this.exceptText.Name = "exceptText";
            this.exceptText.Size = new System.Drawing.Size(100, 13);
            this.exceptText.TabIndex = 7;
            this.exceptText.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // currentLabel
            // 
            this.currentLabel.AutoSize = true;
            this.currentLabel.Location = new System.Drawing.Point(12, 93);
            this.currentLabel.Name = "currentLabel";
            this.currentLabel.Size = new System.Drawing.Size(44, 13);
            this.currentLabel.TabIndex = 9;
            this.currentLabel.Text = "Current:";
            // 
            // currentBox
            // 
            this.currentBox.Enabled = false;
            this.currentBox.Location = new System.Drawing.Point(12, 109);
            this.currentBox.Name = "currentBox";
            this.currentBox.ReadOnly = true;
            this.currentBox.Size = new System.Drawing.Size(406, 20);
            this.currentBox.TabIndex = 10;
            // 
            // exceptBar
            // 
            this.exceptBar.Location = new System.Drawing.Point(12, 67);
            this.exceptBar.Name = "exceptBar";
            this.exceptBar.Size = new System.Drawing.Size(406, 23);
            this.exceptBar.TabIndex = 1;
            this.exceptBar.ValueChanged += new EventProgressBar.OnValueChanged(this.exceptBar_ValueChanged);
            // 
            // eventBar
            // 
            this.eventBar.Location = new System.Drawing.Point(12, 25);
            this.eventBar.Name = "eventBar";
            this.eventBar.Size = new System.Drawing.Size(406, 23);
            this.eventBar.TabIndex = 0;
            this.eventBar.ValueChanged += new EventProgressBar.OnValueChanged(this.eventBar_ValueChanged);
            // 
            // Outlook2iCal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(430, 170);
            this.Controls.Add(this.currentBox);
            this.Controls.Add(this.currentLabel);
            this.Controls.Add(this.exceptText);
            this.Controls.Add(this.eventText);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.exceptLabel);
            this.Controls.Add(this.eventLabel);
            this.Controls.Add(this.exceptBar);
            this.Controls.Add(this.eventBar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Outlook2iCal";
            this.Text = "Outlook2iCal";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private EventProgressBar eventBar;
        private EventProgressBar exceptBar;
        private System.Windows.Forms.Label eventLabel;
        private System.Windows.Forms.Label exceptLabel;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Label eventText;
        private System.Windows.Forms.Label exceptText;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.Label currentLabel;
        private ReadOnlyTextBox currentBox;
    }
}

