using System;
using System.Diagnostics;
using System.IO;
using System.Speech.Synthesis;
using System.Windows.Forms;

namespace Luna
{
    public class Command
    {
        public static void ProcessCommand(string result, SpeechSynthesizer speech)
        {
            switch (result)
            {
                case "Hai":
                    speech.SpeakAsync("Hello! My name is Luna. How can I help you?");
                    break;
                case "Who is you Creator":
                    speech.SpeakAsync("\r\nI was brought into existence by a talented programmer named Jahs Morate. Initially introduced as Luna PCX-001 in May 2021, I have since evolved and undergone updates, emerging as the enhanced version known as Luna-032");
                    break;
                case "Open Word":
                    OpenApplication(@"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", "Microsoft Word", speech);
                    break;
                case "Exit Word":
                    CloseApplication("WINWORD", "Microsoft Word", speech);
                    break;
                case "Open Powerpoint":
                    OpenApplication(@"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE", "Microsoft Power Point", speech);
                    break;
                case "Exit Powerpoint":
                    CloseApplication("POWERPNT", "Microsoft Power Point", speech);
                    break;
                // Add more cases for other commands as needed...
                default:
                    // Handle unrecognized commands
                    break;
            }
        }

        private static void OpenApplication(string appPath, string appName, SpeechSynthesizer speech)
        {
            try
            {
                if (File.Exists(appPath))
                {
                    Process.Start(appPath);
                    speech.SpeakAsync($"{appName} is now open.");
                }
                else
                {
                    MessageBox.Show($"{appName} not found at the specified location.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening {appName}: {ex.Message}");
            }
        }

        private static void CloseApplication(string processName, string appName, SpeechSynthesizer speech)
        {
            try
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    process.Kill();
                }
                speech.SpeakAsync($"{appName} is now closed.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error closing {appName}: {ex.Message}");
            }
        }
    }
}
