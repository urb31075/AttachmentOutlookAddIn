namespace AttachmentOutlookAddIn
{
    partial class AttachmentUserControl
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
            this.FolderComboBox = new System.Windows.Forms.ComboBox();
            this.ExtractButton = new System.Windows.Forms.Button();
            this.InfoListBox = new System.Windows.Forms.ListBox();
            this.MainStatusStrip = new System.Windows.Forms.StatusStrip();
            this.InfoToolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar = new System.Windows.Forms.ToolStripProgressBar();
            this.ClearButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.GetReceivedButton = new System.Windows.Forms.Button();
            this.SaveLogButton = new System.Windows.Forms.Button();
            this.MainStatusStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // FolderComboBox
            // 
            this.FolderComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FolderComboBox.DisplayMember = "FullFolderPath";
            this.FolderComboBox.FormattingEnabled = true;
            this.FolderComboBox.Location = new System.Drawing.Point(3, 26);
            this.FolderComboBox.Name = "FolderComboBox";
            this.FolderComboBox.Size = new System.Drawing.Size(364, 21);
            this.FolderComboBox.TabIndex = 0;
            this.FolderComboBox.SelectedIndexChanged += new System.EventHandler(this.FolderComboBoxSelectedIndexChanged);
            // 
            // ExtractButton
            // 
            this.ExtractButton.BackColor = System.Drawing.Color.LightGreen;
            this.ExtractButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ExtractButton.Location = new System.Drawing.Point(3, 49);
            this.ExtractButton.Name = "ExtractButton";
            this.ExtractButton.Size = new System.Drawing.Size(120, 23);
            this.ExtractButton.TabIndex = 1;
            this.ExtractButton.Text = "Извлечь";
            this.ExtractButton.UseVisualStyleBackColor = false;
            this.ExtractButton.Click += new System.EventHandler(this.ExtractButtonClick);
            // 
            // InfoListBox
            // 
            this.InfoListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.InfoListBox.FormattingEnabled = true;
            this.InfoListBox.Location = new System.Drawing.Point(4, 75);
            this.InfoListBox.Name = "InfoListBox";
            this.InfoListBox.Size = new System.Drawing.Size(363, 407);
            this.InfoListBox.TabIndex = 2;
            // 
            // MainStatusStrip
            // 
            this.MainStatusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.InfoToolStripStatusLabel,
            this.progressBar});
            this.MainStatusStrip.Location = new System.Drawing.Point(0, 511);
            this.MainStatusStrip.Name = "MainStatusStrip";
            this.MainStatusStrip.Size = new System.Drawing.Size(370, 22);
            this.MainStatusStrip.TabIndex = 23;
            this.MainStatusStrip.Text = "statusStrip1";
            // 
            // InfoToolStripStatusLabel
            // 
            this.InfoToolStripStatusLabel.AutoSize = false;
            this.InfoToolStripStatusLabel.Name = "InfoToolStripStatusLabel";
            this.InfoToolStripStatusLabel.Size = new System.Drawing.Size(200, 17);
            this.InfoToolStripStatusLabel.Text = "toolStripStatusLabel1";
            this.InfoToolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // progressBar
            // 
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(100, 16);
            // 
            // ClearButton
            // 
            this.ClearButton.BackColor = System.Drawing.Color.Crimson;
            this.ClearButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ClearButton.Location = new System.Drawing.Point(125, 49);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(120, 23);
            this.ClearButton.TabIndex = 24;
            this.ClearButton.Text = "Отодрать";
            this.ClearButton.UseVisualStyleBackColor = false;
            this.ClearButton.Click += new System.EventHandler(this.ClearButtonClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(179, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Экстрактор вложений. версия 1.1";
            // 
            // GetReceivedButton
            // 
            this.GetReceivedButton.BackColor = System.Drawing.Color.DodgerBlue;
            this.GetReceivedButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.GetReceivedButton.Location = new System.Drawing.Point(247, 49);
            this.GetReceivedButton.Name = "GetReceivedButton";
            this.GetReceivedButton.Size = new System.Drawing.Size(120, 23);
            this.GetReceivedButton.TabIndex = 26;
            this.GetReceivedButton.Text = "Взять Received";
            this.GetReceivedButton.UseVisualStyleBackColor = false;
            this.GetReceivedButton.Click += new System.EventHandler(this.GetReceivedButtonClick);
            // 
            // SaveLogButton
            // 
            this.SaveLogButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.SaveLogButton.Location = new System.Drawing.Point(4, 485);
            this.SaveLogButton.Name = "SaveLogButton";
            this.SaveLogButton.Size = new System.Drawing.Size(75, 23);
            this.SaveLogButton.TabIndex = 23;
            this.SaveLogButton.Text = "Сохранить";
            this.SaveLogButton.UseVisualStyleBackColor = true;
            this.SaveLogButton.Click += new System.EventHandler(this.SaveLogButtonClick);
            // 
            // AttachmentUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.SaveLogButton);
            this.Controls.Add(this.GetReceivedButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ClearButton);
            this.Controls.Add(this.MainStatusStrip);
            this.Controls.Add(this.InfoListBox);
            this.Controls.Add(this.ExtractButton);
            this.Controls.Add(this.FolderComboBox);
            this.Name = "AttachmentUserControl";
            this.Size = new System.Drawing.Size(370, 533);
            this.Load += new System.EventHandler(this.AttachmentUserControlLoad);
            this.MainStatusStrip.ResumeLayout(false);
            this.MainStatusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox FolderComboBox;
        private System.Windows.Forms.Button ExtractButton;
        private System.Windows.Forms.ListBox InfoListBox;
        private System.Windows.Forms.StatusStrip MainStatusStrip;
        private System.Windows.Forms.ToolStripStatusLabel InfoToolStripStatusLabel;
        private System.Windows.Forms.Button ClearButton;
        private System.Windows.Forms.ToolStripProgressBar progressBar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button GetReceivedButton;
        private System.Windows.Forms.Button SaveLogButton;
    }
}
