﻿using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Microsoft.Win32;

namespace ClassicSetup
{
    public partial class Wizard : Form
    {
        private readonly string logFilePath = @"C:\Classic Files\firsttime.log";

        public Wizard()
        {
            InitializeComponent();
            AddCommandLinkButtons();
        }

        private void AddCommandLinkButtons()
        {
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Ultimate branding", BrandingButton_Click);
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Professional branding", BrandingButton_Click);
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Home Premium branding", BrandingButton_Click);
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Home Basic branding", BrandingButton_Click);
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Starter branding", BrandingButton_Click);

            AddCommandLinkButton(bwsrlinkpanel, "Internet Explorer 11 style (BeautyFox)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Firefox 14 - 28 style (Echelon)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Chrome 23 style (Silverfox)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Firefox 115 (Unmodified)", BrowserButton_Click);

            AddCommandLinkButton(rebootpanel, "Reboot now and finish the post-install stage", RebootButton_Click);
            Log("Added buttons");
        }

        private void AddCommandLinkButton(Panel panel, string text, EventHandler clickHandler)
        {
            CommandLinkButton cmdLinkButton = new CommandLinkButton
            {
                Text = text,
                Size = new System.Drawing.Size(400, 42)
            };
            cmdLinkButton.Click += clickHandler;
            panel.Controls.Add(cmdLinkButton);
        }

        private void BrandingButton_Click(object sender, EventArgs e)
        {
            var button = (CommandLinkButton)sender;
            switch (button.Text)
            {
                case "Windows 7 Ultimate branding":
                    ApplyBranding("Ultimate");
                    break;
                case "Windows 7 Professional branding":
                    ApplyBranding("Professional");
                    break;
                case "Windows 7 Home Premium branding":
                    ApplyBranding("Home Premium");
                    break;
                case "Windows 7 Home Basic branding":
                    ApplyBranding("Home Basic");
                    break;
                case "Windows 7 Starter branding":
                    ApplyBranding("Starter");
                    break;
            }
            welcomeWizard.NextPage();
        }

        private void ApplyBranding(string edition)
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string brandingSourcePath = Path.Combine(appDirectory, "Branding", edition);
            string brandingDestinationPath = @"C:\Windows\Branding";

            try
            {
                CopyDirectory(brandingSourcePath, brandingDestinationPath);
                Log($"Applied {edition} branding to {brandingDestinationPath}");

                SimulateWinR();
                SendKeys.SendWait($"\"C:\\Classic Files\\Classic Setup\\branding.exe\" -branding \"{edition}\"{Environment.NewLine}");

                Log($"Executed branding.exe for {edition}");

                if (edition == "Starter")
                {
                    ApplyStarterWallpaper();
                }
            }
            catch (UnauthorizedAccessException uex)
            {
                MessageBox.Show($"Permission denied while applying {edition} branding: {uex.Message}");
                Log($"Permission denied: {uex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to apply {edition} branding: {ex.Message}");
                Log($"Error applying {edition} branding: {ex.Message}");
            }
        }

        private void SimulateWinR()
        {
            keybd_event((byte)Keys.LWin, 0, 0, 0);
            keybd_event((byte)Keys.R, 0, 0, 0);
            keybd_event((byte)Keys.R, 0, 0x0002, 0);
            keybd_event((byte)Keys.LWin, 0, 0x0002, 0);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

        private void RunAsAdmin(string fileName, string arguments)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = $"/c \"{fileName} {arguments}\"";
                process.StartInfo.Verb = "runas";
                process.StartInfo.UseShellExecute = true;
                process.StartInfo.CreateNoWindow = false;

                process.Start();
                process.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to run {fileName} as Administrator: {ex.Message}");
                Log($"Error running {fileName} with arguments {arguments}: {ex.Message}");
            }
        }

        private void CopyDirectory(string sourceDir, string destinationDir)
        {
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
            }

            foreach (string filePath in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
            {
                string relativePath = filePath.Substring(sourceDir.Length + 1);
                string destinationFilePath = Path.Combine(destinationDir, relativePath);

                string destinationFileDirectory = Path.GetDirectoryName(destinationFilePath);
                if (!Directory.Exists(destinationFileDirectory))
                {
                    Directory.CreateDirectory(destinationFileDirectory);
                }
                File.Copy(filePath, destinationFilePath, true);
            }
        }

        private void ApplyStarterWallpaper()
        {
            string wallpaperPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Branding", "Starter", "wallpaper.jpg");

            if (File.Exists(wallpaperPath))
            {
                SetWallpaper(wallpaperPath);
                Log("Applied Starter Wallpaper");
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        const int SPI_SETDESKWALLPAPER = 0x0014;
        const int SPIF_UPDATEINIFILE = 0x01;
        const int SPIF_SENDCHANGE = 0x02;

        private void SetWallpaper(string wallpaperPath)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wallpaperPath, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }

        private void BrowserButton_Click(object sender, EventArgs e)
        {
            var button = (CommandLinkButton)sender;
            switch (button.Text)
            {
                case "Internet Explorer 11 style (BeautyFox)":
                    ApplyIE11Style();
                    break;
                case "Firefox 14 - 28 style (Echelon)":
                    ApplyFirefox10To13Style();
                    break;
                case "Chrome 23 style (Silverfox)":
                    ApplyChrome2012Style();
                    break;
                case "Firefox 115 (Unmodified)":
                    ApplyFirefox115Style();
                    break;
            }
            welcomeWizard.NextPage();
        }

        private void ApplyIE11Style()
        {
            ApplyBrowserStyle("BeautyFox");
            Log("Applied Internet Explorer 11 Style");
        }

        private void ApplyFirefox10To13Style()
        {
            ApplyBrowserStyle("Echelon");
            Log("Applied Firefox 14 - 28 Style");
        }

        private void ApplyChrome2012Style()
        {
            ApplyBrowserStyle("Silverfox");
            Log("Applied Chrome 23 Style");
        }

        private void ApplyFirefox115Style()
        {
            ApplyBrowserStyle("FF115");
            Log("Applied Default Firefox 115 Style");
        }

        private void ApplyBrowserStyle(string folderName)
        {
            string sourceFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Browser", folderName);
            string programFilesFolder = @"C:\Program Files\Mozilla Firefox";
            string appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string startMenuProgramsFolder = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs";

            if (Directory.Exists(sourceFolder))
            {
                CopyStyleFiles(sourceFolder, programFilesFolder, appDataFolder);
                CopyShortcutFiles(sourceFolder, startMenuProgramsFolder);
            }
            else
            {
                MessageBox.Show($"Source folder does not exist: {sourceFolder}");
                Log($"Error: Source folder does not exist: {sourceFolder}");
            }
        }

        private void CopyShortcutFiles(string sourceFolder, string destinationFolder)
        {
            foreach (var file in Directory.GetFiles(sourceFolder, "*.lnk"))
            {
                string fileName = Path.GetFileName(file);
                string destinationPath = Path.Combine(destinationFolder, fileName);
                File.Copy(file, destinationPath, true);
            }
        }

        private void CopyStyleFiles(string sourceFolder, string programFilesFolder, string appDataFolder)
        {
            string sourceFFFolder = Path.Combine(sourceFolder, "FFFolder");
            string sourceADFolder = Path.Combine(sourceFolder, "ADFolder");

            if (Directory.Exists(sourceFFFolder))
            {
                CopyFilesRecursively(new DirectoryInfo(sourceFFFolder), new DirectoryInfo(programFilesFolder));
            }
            else
            {
                MessageBox.Show($"FFFolder does not exist: {sourceFFFolder}");
            }

            if (Directory.Exists(sourceADFolder))
            {
                CopyFilesRecursively(new DirectoryInfo(sourceADFolder), new DirectoryInfo(appDataFolder));
            }
            else
            {
                MessageBox.Show($"ADFolder does not exist: {sourceADFolder}");
            }
        }

        private void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (var dir in source.GetDirectories())
            {
                CopyFilesRecursively(dir, target.CreateSubdirectory(dir.Name));
            }

            foreach (var file in source.GetFiles())
            {
                string path = Path.Combine(target.FullName, file.Name);
                file.CopyTo(path, true);
            }
        }

        private void RebootButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to reboot?", "Reboot", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                RebootSystem();
            }
        }

        private void RebootSystem()
        {
            try
            {
                Process.Start(new ProcessStartInfo("shutdown", "/r /t 0")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to reboot system: {ex.Message}");
                Log($"Error rebooting system: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to write log: {ex.Message}");
            }
        }
    }
}
