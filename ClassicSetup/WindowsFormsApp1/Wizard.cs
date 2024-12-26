using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Net.Http;
using System.Linq;

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
            AddCommandLinkButton(cmdlinkpanel, "Windows 7 Enterprise branding", BrandingButton_Click);

            AddCommandLinkButton(bwsrlinkpanel, "Internet Explorer 11 style (BeautyFox)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Firefox 14 - 28 style (Echelon)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Chrome 1 - 58 style (Geckium)", BrowserButton_Click);
            AddCommandLinkButton(bwsrlinkpanel, "Firefox 115 (Unmodified)", BrowserButton_Click);

            AddCommandLinkButton(openwithPanel, "Windows 10 OpenWith style", BrowserButton_Click);
            AddCommandLinkButton(openwithPanel, "Windows 7 OpenWith style", BrowserButton_Click);

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
                    ApplyBranding("Premium");
                    break;
                case "Windows 7 Enterprise branding":
                    ApplyBranding("Enterprise");
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

                RunCLHBranding($"\"{edition}\"");

                Log($"Executed branding.exe for {edition}");
            }
            catch (UnauthorizedAccessException uex)
            {
                Log($"Permission denied: {uex.Message}");
            }
            catch (Exception ex)
            {
                Log($"Error applying {edition} branding: {ex.Message}");
            }
        }

        private async void RunCLHBranding(string edition)
        {
            try
            {
                string executablePath = @"C:\Classic Files\Classic Setup\branding.exe";
                string arguments = $"-branding \"{edition}\"";

                var process = new System.Diagnostics.Process();
                process.StartInfo.FileName = executablePath;
                process.StartInfo.Arguments = arguments;      
                process.StartInfo.RedirectStandardOutput = true;  
                process.StartInfo.RedirectStandardError = true;   
                process.StartInfo.UseShellExecute = false;        
                process.StartInfo.CreateNoWindow = true;        

                process.Start();

                string output = await process.StandardOutput.ReadToEndAsync();
                string error = await process.StandardError.ReadToEndAsync();

                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                {
                    Log($"Output: {output}");
                }
                if (!string.IsNullOrEmpty(error))
                {
                    Log($"Error: {error}");
                }
            }
            catch (Exception ex)
            {
                Log($"Exception: {ex.Message}");
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
                case "Chrome 1 - 58 style (Geckium)":
                    ApplyChrome2012Style();
                    MessageBox.Show("Don't reboot until you see a message confirming Geckium is done downloading!");
                    break;
                case "Firefox 115 (Unmodified)":
                    ApplyFirefox115Style();
                    break;
            }

            RunIe4uinit();
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
            ApplyBrowserStyle("Geckium");
            Log("Applied Geckium");
        }

        private void ApplyFirefox115Style()
        {
            ApplyBrowserStyle("FF115");
            Log("Applied Default Firefox 115 Style");
        }

        private async void ApplyBrowserStyle(string folderName)
        {
            if (folderName == "Geckium")
            {
                await DownloadAndApplyGeckium();
                Log("Downloaded and applied Geckium style.");
                return;
            }

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

        private async Task DownloadAndApplyGeckium()
        {
            string tagsUrl = "https://api.github.com/repos/angelbruni/Geckium/tags";
            string downloadPath = @"C:\Classic Files\Classic Setup\Browser\Geckium";
            string firefoxInstallPath = @"C:\Program Files\Mozilla Firefox";
            string ff115ProfileBasePath = @"C:\Classic Files\Classic Setup\Browser\FF115";
            string userAgent = "ClassicSetupApp";
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string ff115AppDataProfilePath = Path.Combine(ff115ProfileBasePath, "ADFolder");
            string geckiumProfileFolder = @"Profile Folder";

            try
            {
                Debug.WriteLine("Initializing HTTP client.");
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd(userAgent);

                    Debug.WriteLine($"Fetching tags from: {tagsUrl}");
                    HttpResponseMessage tagsResponse = await client.GetAsync(tagsUrl);

                    Debug.WriteLine($"Response status code: {tagsResponse.StatusCode}");
                    if (!tagsResponse.IsSuccessStatusCode)
                    {
                        Log($"Failed to fetch tags: {tagsResponse.StatusCode}");
                        MessageBox.Show($"Failed to fetch tags: {tagsResponse.ReasonPhrase} (Code {tagsResponse.StatusCode})");
                        return;
                    }

                    string tagsJson = await tagsResponse.Content.ReadAsStringAsync();
                    Debug.WriteLine($"Tags JSON received: {tagsJson}");
                    dynamic tags = Newtonsoft.Json.JsonConvert.DeserializeObject(tagsJson);

                    if (tags.Count == 0)
                    {
                        Debug.WriteLine("No tags found in the repository.");
                        MessageBox.Show("No tags found in the repository.");
                        Log("No tags found in the repository.");
                        return;
                    }

                    Debug.WriteLine("Parsing tag names.");
                    List<string> tagNames = new List<string>();
                    foreach (var tag in tags)
                    {
                        string tagName = (string)tag.name;
                        tagNames.Add(tagName);
                        Debug.WriteLine($"Tag found: {tagName}");
                    }

                    Debug.WriteLine("Sorting tags to find the latest.");
                    string latestTag = GetLatestTag(tagNames);
                    Debug.WriteLine($"Latest tag determined: {latestTag}");

                    string assetUrl = $"https://github.com/angelbruni/Geckium/archive/refs/tags/{latestTag}.zip";
                    string fileName = $"{latestTag}.zip";
                    Debug.WriteLine($"Asset URL: {assetUrl}");
                    Debug.WriteLine($"File will be saved as: {fileName}");

                    Directory.CreateDirectory(downloadPath);
                    Debug.WriteLine($"Download directory created/verified: {downloadPath}");

                    string zipPath = Path.Combine(downloadPath, fileName);

                    Debug.WriteLine($"Starting download of: {assetUrl}");
                    using (HttpResponseMessage downloadResponse = await client.GetAsync(assetUrl, HttpCompletionOption.ResponseHeadersRead))
                    {
                        downloadResponse.EnsureSuccessStatusCode();

                        using (var contentStream = await downloadResponse.Content.ReadAsStreamAsync())
                        using (var fileStream = new FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            await contentStream.CopyToAsync(fileStream);
                        }
                    }

                    Debug.WriteLine($"Download completed: {zipPath}");
                    Log($"Downloaded Geckium latest tag: {fileName}");

                    Debug.WriteLine($"Extracting zip file: {zipPath}");
                    string extractPath = Path.Combine(downloadPath, "Extracted");
                    System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractPath);

                    string extractedFolder = Path.Combine(extractPath, $"Geckium-{latestTag}");
                    if (!Directory.Exists(extractedFolder))
                    {
                        throw new DirectoryNotFoundException($"Expected folder not found: {extractedFolder}");
                    }

                    Debug.WriteLine($"Extracted folder found: {extractedFolder}");

                    string firefoxFolder = Path.Combine(extractedFolder, "Firefox Folder");
                    string profileFolder = Path.Combine(extractedFolder, geckiumProfileFolder);

                    Debug.WriteLine($"Copying Firefox Folder contents from {firefoxFolder} to: {firefoxInstallPath}");
                    if (Directory.Exists(firefoxFolder))
                    {
                        CopyFilesRecursively(new DirectoryInfo(firefoxFolder), new DirectoryInfo(firefoxInstallPath));
                    }
                    else
                    {
                        Debug.WriteLine("Firefox Folder does not exist in the extracted content.");
                    }

                    Debug.WriteLine($"Applying FF115 profile folder from {ff115AppDataProfilePath} to {appDataPath}");
                    string appDataFF115Path = Path.Combine(appDataPath);
                    if (Directory.Exists(ff115AppDataProfilePath))
                    {
                        Debug.WriteLine($"Contents of ADFolder: {string.Join(", ", Directory.GetFiles(ff115AppDataProfilePath))}");
                        CopyFilesRecursively(new DirectoryInfo(ff115AppDataProfilePath), new DirectoryInfo(appDataFF115Path));
                    }
                    else
                    {
                        Debug.WriteLine("FF115 profile folder (ADFolder) does not exist.");
                    }

                    string ff115ProfileFolderPath = Path.Combine(appDataFF115Path,"Mozilla", "Firefox", "Profiles", "foky7k51.ff115");
                    Debug.WriteLine($"Merging Geckium Profile Folder into {ff115ProfileFolderPath}");
                    if (Directory.Exists(ff115ProfileFolderPath) && Directory.Exists(profileFolder))
                    {
                        Debug.WriteLine($"Contents of Geckium Profile Folder: {string.Join(", ", Directory.GetFiles(profileFolder))}");
                        CopyFilesRecursively(new DirectoryInfo(profileFolder), new DirectoryInfo(ff115ProfileFolderPath));
                    }
                    else
                    {
                        Debug.WriteLine("Either the FF115 profile or Geckium profile folder is missing.");
                    }

                    Debug.WriteLine("Geckium installation and configuration completed.");
                    Log("Geckium installation and configuration completed successfully.");
                    MessageBox.Show("Geckium is installed! You will need to perform the first-time setup after you launch Firefox.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception occurred: {ex.Message}");
                MessageBox.Show($"Failed to download or apply Geckium: {ex.Message}");
                Log($"Error during Geckium installation: {ex.Message}");
            }
        }

        private string GetLatestTag(List<string> tagNames)
        {
            Debug.WriteLine("Starting semantic sort of tags.");
            string latestTag = tagNames
                .OrderByDescending(tag =>
                {
                    Debug.WriteLine($"Parsing version: {tag}");
                    var parts = tag.TrimStart('b').Split('.');
                    return new Version(
                        int.Parse(parts[0]),
                        int.Parse(parts[1]),
                        parts.Length > 2 ? int.Parse(parts[2]) : 0
                    );
                })
                .First();
            Debug.WriteLine($"Latest tag after sorting: {latestTag}");
            return latestTag;
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

        private void RunIe4uinit()
        {
            try
            {
                string ie4uinitPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ie4uinit.exe");
                if (File.Exists(ie4uinitPath))
                {
                    System.Diagnostics.Process.Start(ie4uinitPath, "-show");
                }
                else
                {
                    MessageBox.Show("ie4uinit.exe not found in the application directory.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to run ie4uinit: {ex.Message}");
            }
        }

        private void CopyAdFolderToAppData()
        {
            string adFolderPath = @"C:\Classic Files\Classic Setup\Browser\FF115\ADFolder";
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            if (!Directory.Exists(appDataPath))
            {
                Directory.CreateDirectory(appDataPath);
            }

            try
            {
                Log($"Starting to copy files from {adFolderPath} to {appDataPath}");
                CopyFilesDirectlyToAppData(new DirectoryInfo(adFolderPath), appDataPath);
                Log("Successfully copied all files from ADFolder to AppData.");
            }
            catch (Exception ex)
            {
                Log($"Error while copying ADFolder to AppData: {ex.Message}");
            }
        }

        private void CopyFilesDirectlyToAppData(DirectoryInfo source, string targetPath)
        {
            foreach (var dir in source.GetDirectories())
            {
                string targetSubDir = Path.Combine(targetPath, dir.Name);

                if (!Directory.Exists(targetSubDir))
                {
                    Directory.CreateDirectory(targetSubDir);
                }

                CopyFilesDirectlyToAppData(dir, targetSubDir);
            }

            foreach (var file in source.GetFiles())
            {
                string targetFilePath = Path.Combine(targetPath, file.Name);

                file.CopyTo(targetFilePath, true);
            }
        }

        private void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (var dir in source.GetDirectories())
            {
                DirectoryInfo targetSubDir = target.CreateSubdirectory(dir.Name);
                CopyFilesRecursively(dir, targetSubDir);
            }

            foreach (var file in source.GetFiles())
            {
                string targetFilePath = Path.Combine(target.FullName, file.Name);
                file.CopyTo(targetFilePath, true);
            }
        }

        private void RebootButton_Click(object sender, EventArgs e)
        {
            RebootSystem();
        }

        private void RebootSystem()
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", true))
                {
                    if (key != null)
                    {
                        key.SetValue("EnableLUA", 1, RegistryValueKind.DWord);
                    }
                }
                Process.Start("shutdown", "/r /t 0");
            }
            catch
            {
                // nothing
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
