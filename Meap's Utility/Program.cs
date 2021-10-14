using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Meaps_Utility
{
    internal class Program
    {
        internal static bool SilentMode = false;
        internal static bool FoundUltimatePerfPlan = false;
        internal static IntPtr handle = GetConsoleWindow();
        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        internal const int SW_HIDE = 0;
        [DllImport("ntdll.dll", EntryPoint = "NtSetTimerResolution")]
        internal static extern void NtSetTimerResolution(uint DesiredResolution, bool SetResolution, ref uint CurrentResolution);

        static void Main()
        {
            Console.Title = "Meap's Performance Utility   V.1.3";
            Console.WindowHeight = 12;
            Console.WindowWidth = 60;

            // Check if the current process has been started with arguments.
            if (Environment.CommandLine.Contains("-silent"))
            {
                ShowWindow(handle, SW_HIDE);
                SilentMode = true;
            }

            // Have the application start up automatically when the OS boots.
            string ExePath = Assembly.GetEntryAssembly().Location;
            if (ExePath.Contains("'"))
            {
                MessageBox.Show("Please move the .exe file of this utility to a different path, the path to the .exe file can not contain the character: ' (single quotation mark)\n\nThe Utility will now close, please move it to a different location and run it again.", "Meap's Performance Utility - Error");
                Environment.Exit(0);
            }
            ExecuteCMDCommand($"schtasks.exe /delete /f /tn \"Meap's Performance Utility\" && schtasks.exe /create /f /rl highest /sc onlogon /tn \"Meap's Performance Utility\" /tr \"'{ExePath}' -silent\"");

            // Fetch Windows power plans.
            ExecuteCMDCommand("powercfg /list");

            // Set the Windows timer resolution as low as possible.
            Console.Write($"Setting timer resolution from 1ms to 0.5ms... ");
            uint DesiredResolution = 5000; // 0.5ms
            bool SetResolution = true;
            uint CurrentResolution = 0;
            NtSetTimerResolution(DesiredResolution, SetResolution, ref CurrentResolution);
            Console.Write("OK.\n");

            // Cleanup temp files.
            string TempPath = Path.GetTempPath();
            if (Directory.Exists(TempPath))
            {
                Console.Write("Cleaning temporary files... ");
                Console.Write(DeleteAllInFolder(TempPath));
            }
            else
                Console.WriteLine("Could not find Prefetch, ignoring.");

            // Cleanup Prefetch files.
            string Prefetchpath = @"C:\Windows\Prefetch";
            if (Directory.Exists(Prefetchpath))
            {
                Console.Write("Cleaning Prefetch... ");
                Console.Write(DeleteAllInFolder(Prefetchpath));
            }
            else
                Console.WriteLine("Could not find Prefetch, ignoring.");

            // Cleanup SoftwareDistribution files.
            string SoftwareDistributionPath = @"C:\Windows\SoftwareDistribution\Download";
            if (Directory.Exists(SoftwareDistributionPath))
            {
                Console.Write("Cleaning SoftwareDistribution... ");
                Console.Write(DeleteAllInFolder(SoftwareDistributionPath));
            }
            else
                Console.WriteLine("Could not find SoftwareDistribution, ignoring.");

            // Cleanup VRChat cache.
            string VRChatCachePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"Low\VRChat";
            if (Directory.Exists(VRChatCachePath))
            {
                Console.Write("Cleaning VRChat Cache... ");
                Console.Write(DeleteAllInFolder(VRChatCachePath));
            }
            else
                Console.WriteLine("Could not find VRChat Cache, ignoring.");

            // Run cleanmanager.
            Console.Write("Running cleanmgr... ");
            Console.Write(ExecuteCMDCommand("cleanmgr.exe /AUTOCLEAN"));

            // Turn off Windows hibernation file to gain more disk space on the C:\ drive.
            Console.Write("Turning off hibernation file... ");
            Console.Write(ExecuteCMDCommand(@"powercfg /h off"));

            // Run gaming performance boost commandline options.
            Console.Write("Enabling gaming performance boosts... ");
            Console.Write(ExecuteCMDCommand("bcdedit /set useplatformtick yes && bcdedit /set disabledynamictick yes && bcdedit /deletevalue useplatformclock"));

            if (!FoundUltimatePerfPlan)
            {
                ExecuteCMDCommand("powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 && powercfg /list");
                Console.WriteLine("Creating Ultimate Performance power plan... OK.");
            }

            // Hide the window.
            if (!SilentMode)
            {
                Thread.Sleep(1500);
                ShowWindow(handle, SW_HIDE);
                if (File.Exists("Rkeys.reg"))
                    Process.Start("RKeys.reg");
                else
                {
                    MessageBox.Show("The file: 'RKeys.reg' could not be found. Please place it in the same directory and run the utility again.", "Meap's Performance Utility - Missing a file");
                    Environment.Exit(0);
                }

                MessageBox.Show("If this is the first time you are running this utility your machine requires a reboot for these changes to take effect.", "Meap's Performance Utility - Reboot required");
            }
            Console.Clear();
            Thread.Sleep(-1); // Make sure the application does not close. If it does the timer resolution will reset back to default.
        }

        internal static string ExecuteCMDCommand(string command)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/c " + command)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            string result = "OK.\n";
            using (Process process = new Process())
            {
                process.ErrorDataReceived += (sender, args) => {
                    if (args.Data != null) { result = "FAIL!\n"; }
                };

                process.OutputDataReceived += (sender, args) => {
                    if (!FoundUltimatePerfPlan && args.Data != null && args.Data.Contains("(Meap's Performance Utility (Ultimate Performance))"))
                    {
                        FoundUltimatePerfPlan = true;
                        if (!args.Data.Contains("*")) // Ultimate Performance plan exists but is not set as default.
                        {
                            string UltimatePerfSchemeID = args.Data.Replace(" ", string.Empty).Replace("PowerSchemeGUID:", string.Empty).Replace("(Meap'sPerformanceUtility(UltimatePerformance))", string.Empty).Replace("*", string.Empty);
                            ExecuteCMDCommand("powercfg /setactive " + UltimatePerfSchemeID + " && powercfg /change standby-timeout-dc 20 && powercfg /change monitor-timeout-dc 0 && powercfg /change standby-timeout-ac 0 && powercfg /change monitor-timeout-ac 0"); // Set the Ultimate Performance plan as default.
                            Console.WriteLine("Setting Ultimate Performance as default power plan... OK.");
                        }
                    }
                    else
                    {
                        if (!FoundUltimatePerfPlan && args.Data != null && args.Data.Contains("(Ultimate Performance)"))
                        {
                            FoundUltimatePerfPlan = true;
                            string UltimatePerfSchemeID = args.Data.Replace(" ", string.Empty).Replace("PowerSchemeGUID:", string.Empty).Replace("(UltimatePerformance)", string.Empty).Replace("*", string.Empty);
                            ExecuteCMDCommand($"powercfg /changename {UltimatePerfSchemeID} \"Meap's Performance Utility (Ultimate Performance)\" \"Provides ultimate performance.\"");
                            if (!args.Data.Contains("*")) // Ultimate Performance plan exists but is not set as default.
                            {
                                ExecuteCMDCommand("powercfg /setactive " + UltimatePerfSchemeID + " && powercfg /change standby-timeout-dc 20 && powercfg /change monitor-timeout-dc 0 && powercfg /change standby-timeout-ac 0 && powercfg /change monitor-timeout-ac 0"); // Set the Ultimate Performance plan as default.
                                Console.WriteLine("Setting Ultimate Performance as default power plan... OK.");
                            }
                        }
                    }
                };

                process.StartInfo = startInfo;
                process.Start();

                process.BeginErrorReadLine();
                process.BeginOutputReadLine();
            }
            return result;
        }

        internal static string DeleteAllInFolder(string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);

            foreach (FileInfo file in di.EnumerateFiles())
            {
                try
                {
                    file.Delete();
                }
                catch { }
            }
            foreach (DirectoryInfo dir in di.EnumerateDirectories())
            {
                try
                {
                    dir.Delete(true);
                }
                catch { }
            }
            return "OK.\n";
        }
    }
}