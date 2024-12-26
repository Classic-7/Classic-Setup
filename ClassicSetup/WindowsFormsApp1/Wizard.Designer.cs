namespace ClassicSetup
{
    partial class Wizard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Wizard));
            this.wizardPageContainer1 = new AeroWizard.WizardPageContainer();
            this.wizardPage1 = new AeroWizard.WizardPage();
            this.welcomeWizard = new AeroWizard.WizardControl();
            this.wizardPage2 = new AeroWizard.WizardPage();
            this.label1 = new System.Windows.Forms.Label();
            this.wizardPage3 = new AeroWizard.WizardPage();
            this.cmdlinkpanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.wizardPage6 = new AeroWizard.WizardPage();
            this.bwsrlinkpanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.wizardPage4 = new AeroWizard.WizardPage();
            this.openwithPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.wizardPage5 = new AeroWizard.WizardPage();
            this.rebootpanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).BeginInit();
            this.wizardPageContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.welcomeWizard)).BeginInit();
            this.wizardPage2.SuspendLayout();
            this.wizardPage3.SuspendLayout();
            this.wizardPage6.SuspendLayout();
            this.wizardPage4.SuspendLayout();
            this.wizardPage5.SuspendLayout();
            this.SuspendLayout();
            // 
            // wizardPageContainer1
            // 
            this.wizardPageContainer1.BackButton = null;
            this.wizardPageContainer1.CancelButton = null;
            this.wizardPageContainer1.Controls.Add(this.wizardPage1);
            this.wizardPageContainer1.Location = new System.Drawing.Point(193, 270);
            this.wizardPageContainer1.Name = "wizardPageContainer1";
            this.wizardPageContainer1.NextButton = null;
            this.wizardPageContainer1.Pages.Add(this.wizardPage1);
            this.wizardPageContainer1.Size = new System.Drawing.Size(75, 23);
            this.wizardPageContainer1.TabIndex = 0;
            // 
            // wizardPage1
            // 
            this.wizardPage1.Name = "wizardPage1";
            this.wizardPage1.Size = new System.Drawing.Size(75, 23);
            this.wizardPage1.TabIndex = 0;
            this.wizardPage1.Text = "Page Title";
            // 
            // welcomeWizard
            // 
            this.welcomeWizard.Location = new System.Drawing.Point(0, 0);
            this.welcomeWizard.Name = "welcomeWizard";
            this.welcomeWizard.Pages.Add(this.wizardPage2);
            this.welcomeWizard.Pages.Add(this.wizardPage3);
            this.welcomeWizard.Pages.Add(this.wizardPage6);
            this.welcomeWizard.Pages.Add(this.wizardPage4);
            this.welcomeWizard.Pages.Add(this.wizardPage5);
            this.welcomeWizard.Size = new System.Drawing.Size(584, 462);
            this.welcomeWizard.TabIndex = 1;
            this.welcomeWizard.Title = "Welcome to Classic 7";
            this.welcomeWizard.TitleIcon = ((System.Drawing.Icon)(resources.GetObject("welcomeWizard.TitleIcon")));
            // 
            // wizardPage2
            // 
            this.wizardPage2.Controls.Add(this.label1);
            this.wizardPage2.Name = "wizardPage2";
            this.wizardPage2.Size = new System.Drawing.Size(537, 308);
            this.wizardPage2.TabIndex = 0;
            this.wizardPage2.Text = "Welcome!";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(500, 105);
            this.label1.TabIndex = 0;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // wizardPage3
            // 
            this.wizardPage3.Controls.Add(this.cmdlinkpanel);
            this.wizardPage3.Controls.Add(this.label2);
            this.wizardPage3.Name = "wizardPage3";
            this.wizardPage3.ShowNext = false;
            this.wizardPage3.Size = new System.Drawing.Size(537, 308);
            this.wizardPage3.TabIndex = 1;
            this.wizardPage3.Text = "Choose your branding";
            // 
            // cmdlinkpanel
            // 
            this.cmdlinkpanel.Location = new System.Drawing.Point(4, 27);
            this.cmdlinkpanel.Name = "cmdlinkpanel";
            this.cmdlinkpanel.Size = new System.Drawing.Size(500, 240);
            this.cmdlinkpanel.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(1, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(239, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "Choose your branding type for your edition.\r\n";
            // 
            // wizardPage6
            // 
            this.wizardPage6.Controls.Add(this.bwsrlinkpanel);
            this.wizardPage6.Controls.Add(this.label5);
            this.wizardPage6.Name = "wizardPage6";
            this.wizardPage6.ShowNext = false;
            this.wizardPage6.Size = new System.Drawing.Size(537, 308);
            this.wizardPage6.TabIndex = 4;
            this.wizardPage6.Text = "Choose your browser style";
            // 
            // bwsrlinkpanel
            // 
            this.bwsrlinkpanel.Location = new System.Drawing.Point(4, 21);
            this.bwsrlinkpanel.Name = "bwsrlinkpanel";
            this.bwsrlinkpanel.Size = new System.Drawing.Size(500, 190);
            this.bwsrlinkpanel.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(1, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(228, 15);
            this.label5.TabIndex = 0;
            this.label5.Text = "Choose your preferred style of Firefox skin\r\n";
            // 
            // wizardPage4
            // 
            this.wizardPage4.Controls.Add(this.openwithPanel);
            this.wizardPage4.Controls.Add(this.label3);
            this.wizardPage4.Name = "wizardPage4";
            this.wizardPage4.Size = new System.Drawing.Size(537, 308);
            this.wizardPage4.TabIndex = 5;
            this.wizardPage4.Text = "Choose your OpenWith version";
            // 
            // openwithPanel
            // 
            this.openwithPanel.Location = new System.Drawing.Point(3, 37);
            this.openwithPanel.Name = "openwithPanel";
            this.openwithPanel.Size = new System.Drawing.Size(500, 190);
            this.openwithPanel.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(248, 30);
            this.label3.TabIndex = 2;
            this.label3.Text = "Choose your preferred OpenWith version\r\n(Just a note! Windows 7\'s style can be bu" +
    "ggy!)";
            // 
            // wizardPage5
            // 
            this.wizardPage5.AllowCancel = false;
            this.wizardPage5.Controls.Add(this.rebootpanel);
            this.wizardPage5.Controls.Add(this.label4);
            this.wizardPage5.IsFinishPage = true;
            this.wizardPage5.Name = "wizardPage5";
            this.wizardPage5.ShowNext = false;
            this.wizardPage5.Size = new System.Drawing.Size(537, 308);
            this.wizardPage5.TabIndex = 3;
            this.wizardPage5.Text = "Thank you!";
            // 
            // rebootpanel
            // 
            this.rebootpanel.Location = new System.Drawing.Point(6, 52);
            this.rebootpanel.Name = "rebootpanel";
            this.rebootpanel.Size = new System.Drawing.Size(500, 50);
            this.rebootpanel.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(327, 45);
            this.label4.TabIndex = 0;
            this.label4.Text = "That\'s all the information we need right now.\r\n\r\nClassic 7 will reboot for you to" +
    " apply some finishing touches.";
            // 
            // Wizard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 462);
            this.Controls.Add(this.welcomeWizard);
            this.Controls.Add(this.wizardPageContainer1);
            this.Name = "Wizard";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "hai!!!";
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).EndInit();
            this.wizardPageContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.welcomeWizard)).EndInit();
            this.wizardPage2.ResumeLayout(false);
            this.wizardPage2.PerformLayout();
            this.wizardPage3.ResumeLayout(false);
            this.wizardPage3.PerformLayout();
            this.wizardPage6.ResumeLayout(false);
            this.wizardPage6.PerformLayout();
            this.wizardPage4.ResumeLayout(false);
            this.wizardPage4.PerformLayout();
            this.wizardPage5.ResumeLayout(false);
            this.wizardPage5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private AeroWizard.WizardPageContainer wizardPageContainer1;
        private AeroWizard.WizardPage wizardPage1;
        private AeroWizard.WizardControl welcomeWizard;
        private AeroWizard.WizardPage wizardPage2;
        private System.Windows.Forms.Label label1;
        private AeroWizard.WizardPage wizardPage3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FlowLayoutPanel cmdlinkpanel;
        private AeroWizard.WizardPage wizardPage5;
        private System.Windows.Forms.Label label4;
        private AeroWizard.WizardPage wizardPage6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.FlowLayoutPanel bwsrlinkpanel;
        private System.Windows.Forms.FlowLayoutPanel rebootpanel;
        private AeroWizard.WizardPage wizardPage4;
        private System.Windows.Forms.FlowLayoutPanel openwithPanel;
        private System.Windows.Forms.Label label3;
    }
}

