using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Synthesis;
using System.Threading;
using System.Speech.Recognition;
using System.Diagnostics;
namespace VoiceRecognizition3
{
    public partial class Form1 : Form
    {
        Random rnd = new Random();
        string[] musics = { "https://www.youtube.com/watch?v=qHm9MG9xw1o", "https://www.youtube.com/watch?v=mTQOzfcdZpQ", "https://www.youtube.com/watch?v=hT_nvWreIhg", "https://www.youtube.com/watch?v=Sv6dMFF_yts", "https://www.youtube.com/watch?v=8UVNT4wvIGY", "https://www.youtube.com/watch?v=RBumgq5yVrA", "https://www.youtube.com/watch?v=PIh2xe4jnpk", "https://www.youtube.com/watch?v=09R8_2nJtjg", "https://www.youtube.com/watch?v=34jC1fmeFD0", "https://www.youtube.com/watch?v=qpgTC9MDx1o", "https://www.youtube.com/watch?v=JGwWNGJdvx8", "https://www.youtube.com/watch?v=09R8_2nJtjg" };
        SpeechSynthesizer ss = new SpeechSynthesizer();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        PromptBuilder pb = new PromptBuilder();
        Choices clist = new Choices();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         
        }

    private void button1_Click(object sender, EventArgs e)
        {
            sre.RecognizeAsyncStop();
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            clist.Add(new string[] {  "how are you", "what is a current time", "open chrome", "close", "open youtube","open edge","some music","file explorer","task manager","word","powerpoint","excel","acces", "parameters" , "Gitter please"});
            Grammar gr = new Grammar(new GrammarBuilder(clist));
            try
            {
                sre.RequestRecognizerUpdate();
                sre.LoadGrammar(gr);
                sre.SpeechRecognized += sre_SpeechRecognized;
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "error");
            }

        }
        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            switch (e.Result.Text.ToString())
            {

                case "hello":
                    ss.SpeakAsync("Hi Djeff");
                    break;
                case "how are you":
                    ss.SpeakAsync("I'm great. What about you?");
                    break;
                case "what is a current time":
                    ss.SpeakAsync("Current time is " + DateTime.Now.ToLongTimeString());
                    break;
                case "open chrome":
                    Process.Start("chrome", "https://msdn.microsoft.com/en-us");
                    break;
                case "open youtube":
                    Process.Start("chrome", "https://youtube.com");
                    break;
                case "open edge":
                    Process.Start("microsoft-edge:http://www.google.com");
                    break;
                case "some music":
                    Process.Start("chrome", musics[rnd.Next(0,musics.Length)]);
                    break;
                case "file explorer":
                    System.Diagnostics.Process.Start("explorer.exe", @"c:\");
                    break;
                case "word":
                    Process.Start("WINWORD.EXE");
                    break;
                case "powerpoint":
                    Process.Start("POWERPNT.EXE");
                    break;
                case "excel":
                    Process.Start("EXCEL.EXE");
                    break;
                case "access":
                    Process.Start("ACCESS.EXE");
                    break;
                case "parameters":
                    Process.Start(@"C:\Windows\ImmersiveControlPanel\SystemSettings.exe");
                    break;
                
                case "close":
                    Application.Exit();
                    break;

                default:
                    ss.SpeakAsync("i do not understand you.");
                    break;
            }
            textBox1.Text += e.Result.Text.ToString() + Environment.NewLine;
        }
    }
    }

