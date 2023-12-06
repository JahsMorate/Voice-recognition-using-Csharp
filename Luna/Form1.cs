using System;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.IO;
using System.Diagnostics;

namespace Luna
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recognitionEngine = new SpeechRecognitionEngine();

        SpeechSynthesizer speech = new SpeechSynthesizer();

        Choices choices = new Choices();

        public Form1()
        {
            InitializeComponent();

            try
            {
                string filePath = Path.Combine(Environment.CurrentDirectory, "grammar.txt");

                if (File.Exists(filePath))
                {
                    string[] text = File.ReadAllLines(filePath);
                    choices.Add(text);
                    Grammar grammar = new Grammar(new GrammarBuilder(choices));
                    recognitionEngine.LoadGrammar(grammar);
                    recognitionEngine.SetInputToDefaultAudioDevice();
                    recognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
                    recognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognitionEngine_SpeechRecognized);
                }
                else
                {
                    MessageBox.Show("File 'grammar.txt' not found.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }


            speech.SelectVoiceByHints(VoiceGender.Female);
        }

        private void recognitionEngine_SpeechRecognized(object? sender, SpeechRecognizedEventArgs e)
        {
            string result = e.Result.Text;

            if (result == "Hai")
            {
                speech.SpeakAsync("Hello! My name is Luna. How can I help you?");
                return;
            }

            if (result == "Who is your Creator")
            {
                speech.SpeakAsync("\r\nI was created by a developer named Jahs Morate is currently a fourth-year college student. Initially introduced as Luna PCX-001 in May. 2021, I've evolved through numerous updates, transforming into the iteration known as Luna-032. This journey signifies the remarkable progression of AI technology. I'm designed to assist your application through voice commands. ");
                return;
            }

            if (result == "Who is the Adviser of your Creator")
            {
                speech.SpeakAsync("\r\nMark Anthony Hontanosas Erizo is the advisor associated with 'Intelligent System - Elective 3 is a subject delving into the history of AI, exploring its evolution and impact. It delves into the roots of AI, its historical significance, and its transformative role in shaping the future. This course sheds light on AI's journey from inception to becoming a driving force in technological advancement.'");
                return;
            }

            //open word
            if (result == "Open Word")
            {
                try
                {
                    string wordPath = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"; // Replace XX with your Office version
                    if (File.Exists(wordPath))
                    {
                        Process.Start(wordPath);
                        speech.SpeakAsync("Microsoft Word is now open.");
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Microsoft Word not found at the specified location.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening Microsoft Word: " + ex.Message);
                }
            }

            //exit word
            if (result == "Exit Word")
            {
                try
                {
                    // Find the running Word process and close it
                    foreach (Process process in Process.GetProcessesByName("WINWORD"))
                    {
                        process.Kill();
                    }
                    speech.SpeakAsync("Microsoft Word is now closed.");
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error closing Microsoft Word: " + ex.Message);
                }
            }

            //ppt word
            if (result == "Open PPT")
            {
                try
                {
                    string wordPath = @"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"; // Replace XX with your Office version
                    if (File.Exists(wordPath))
                    {
                        Process.Start(wordPath);
                        speech.SpeakAsync("Microsoft Power Point is now open.");
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Microsoft Word not found at the specified location.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening Microsoft Power Point: " + ex.Message);
                }
            }

            //ppt word
            if (result == "Exit PPT")
            {
                try
                {
                    // Find the running Word process and close it
                    foreach (Process process in Process.GetProcessesByName("POWERPNT"))
                    {
                        process.Kill();
                    }
                    speech.SpeakAsync("Microsoft Power Point is now closed.");
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error closing Microsoft Power Point: " + ex.Message);
                }
            }

            //Canva open
            if (result == "Open Canva")
            {
                try
                {
                    string wordPath = @"C:\Users\Jahs M\AppData\Local\Programs\Canva\Canva.exe"; // Replace XX with your Office version
                    if (File.Exists(wordPath))
                    {
                        Process.Start(wordPath);
                        speech.SpeakAsync("Engine Start.");
                        return;
                    }
                    else
                    {
                        MessageBox.Show("Microsoft Word not found at the specified location.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening Microsoft Power Point: " + ex.Message);
                }
            }

            //Canva exit
            if (result == "Exit Canva")
            {
                try
                {
                    // Find the running Word process and close it
                    foreach (Process process in Process.GetProcessesByName("Canva"))
                    {
                        process.Kill();
                    }
                    speech.SpeakAsync("Canva is now closed.");
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error closing Canva: " + ex.Message);
                }
            }

            else if (result == "Exit")
            {
                // Speak a closing message
                speech.SpeakAsync("Closing application.. Goodbye!!");

                // Close the application
                Application.Exit();
            }

            else
            {
                speech.SpeakAsync("Can't understand what you are saying");
                return;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}