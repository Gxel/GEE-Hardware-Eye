using System;
using System.Text;
using System.Collections.Generic;
using System.Collections;
using System.Collections.ObjectModel;
using System.IO;
using System.Media;
//using System.Speech.AudioFormat;
//using System.Speech.Synthesis;
//using System.Windows.Forms;
using System.Xml;
using System.Windows;
using Microsoft.Win32;
using System.Speech.AudioFormat;
using System.Speech.Synthesis;

namespace BBW_Hardware_Eye
{
    /// <summary> 
    /// Klasse zum bearbeiten und benutzen der Systemstimme<para/> 
    /// .<para/> 
    ///   VERSION     |   DATUM    | AKTION  |  KOMMENTAR  <para/> 
    /// -------------------------------------------------------------------------<para/> 
    ///   01.00.00.00 | 20.08.2016 | feature | Klassen Versionierung eingebaut.  <para/> 
    ///   01.00.00.01 | 21.08.2016 | feature | Fehlerhandling verbessert.        <para/> 
    ///   01.00.00.02 | 22.08.2016 | feature | mDebugIntensity eingebaut.        <para/> 
    /// </summary> 
    internal class aaClassSpeech
    {
        //Fields / Felder 
        private string s = String.Empty;
        private int mDebugIntensity;
        public List<string> ClassLog = new List<string>();


        /// <summary> 
        /// Gibt die aktuelle Klassen Version zurück 
        /// </summary> 
        /// <returns>Gibt die aktuelle Version als String zurück. Beispiel: 01.00.01.02</returns> 
        public String ClassVersion()
        {
            return "01.00.00.02";
        }



        /// <summary> 
        /// Gibt die aktuelle Klassen Version zurück 
        /// </summary> 
        /// <returns>Gibt die aktuelle Version als Integer zurück. Beispiel: 1000102</returns> 
        public int ClassVersionAsInt()
        {
            string s = string.Empty;
            s = ClassVersion().Replace(".", "");
            return Convert.ToInt16(s);
        }



        /// <summary> 
        /// Speichert eine StringListe in eine Textdatei 
        /// </summary> 
        /// <param name="SL">Die zu speichernde Textdatei.</param> 
        /// <param name="Pfad">Das Verzeichnis wo es gespeichert werden soll.</param> 
        /// <param name="aAppend">In die Textdatei hinzufügen (TRUE) oder Textdatei immer neu Erstellen (FALSE).</param> 
        private void StringlistToTextFile(List<string> SL, string Pfad, bool aAppend)
        {
            if (aAppend)
            {
                StreamReader sr = new StreamReader(@Pfad);
                while (sr.Peek() != -1)
                {
                    SL.Add(sr.ReadLine());
                }
                sr.Close();
            }
            // --- 
            StreamWriter sw = new StreamWriter(Pfad, false, System.Text.Encoding.Default);
            for (int i = 0; i < SL.Count; i++)
            {
                sw.WriteLine(SL[i]);
            }
            sw.Close();
        }


        /// <summary> 
        /// Private Methode für die Exceptions 
        /// </summary> 
        /// <param name="sMessage">Exception Nachricht die gespeichert oder ausgegeben wird</param> 
        private void ExceptionHandle(string sMessage)
        {
            sMessage = "Klassen Version " + ClassVersion() + ":" + sMessage;
            ClassLog.Add(sMessage);
            if (mDebugIntensity >= 1)
            {
                StringlistToTextFile(ClassLog, ".\\" + DateTime.Now.ToString("yyyyMMdd") + "_" + this.GetType().Name, false);
            }
            else if (mDebugIntensity >= 2)
            {
                StringlistToTextFile(ClassLog, ".\\" + DateTime.Now.ToString("yyyyMMdd") + "_" + this.GetType().Name, true);
            }
            // --- 
            if (mDebugIntensity >= 3)
            {
                MessageBox.Show(sMessage);
            }
        }



        /// <summary> 
        /// Constructor 
        /// </summary> 
        /// <param name="aaDebugIntensity"> 
        /// Debug intensität:<para/> 
        /// Bei 0 wird nur in der Klasse geloggt<para/> 
        /// Bei 1 wird nur in der Klasse geloggt, die Logs werden in eine Textdatei im Assembly Verzeichnis geschrieben.<para/> 
        /// Bei 2 wird nur in der Klasse geloggt, die Logs werden in eine Textdatei im Assembly Verzeichnis hinzugefügt.<para/> 
        /// Bei 3 wird nur in der Klasse geloggt, die Logs werden in eine Textdatei im Assembly Verzeichnis hinzugefügt und eine Message ausgegeben.<para/> 
        /// </param> 
        public aaClassSpeech(int aaDebugIntensity)
        {
            mDebugIntensity = aaDebugIntensity;
        }


        /// <summary> 
        /// Constructor 
        /// </summary> 
        public aaClassSpeech()
        {
            mDebugIntensity = 0;
        }


        /// <summary> 
        /// Speichert die Stimme mit START und STOP ton. BETA 
        /// </summary> 
        public void SaveSay(String sText, String VoiceName, String StartSound, String StopSound, int iVolume, int iRate)
        {
            var syn = new SpeechSynthesizer();
            var myPlayer = new SoundPlayer();
            var aSaveFileDialog1 = new SaveFileDialog();

            myPlayer.SoundLocation = @StartSound;
            syn.Volume = iVolume;
            syn.Rate = iRate;
            //syn.SetOutputToAudioStream  
            //syn.SelectVoice(VoiceName);  
            myPlayer.SoundLocation = @StopSound;

            //aWave.Stream = new System.IO.MemoryStream();  
            //syn.SetOutputToWaveFile(aWave.Stream);  

            aSaveFileDialog1.Filter = "wave files (*.wav)|*.wav";
            aSaveFileDialog1.DefaultExt = "*.wav";
            aSaveFileDialog1.Title = "Stimme als Wave speichern";
            aSaveFileDialog1.FileName = "";
            aSaveFileDialog1.ShowDialog();
        }


        /// <summary> 
        /// Speichert die Stimme mit START und STOP als WAVE ton. BETA 
        /// </summary> 
        public void SaveSay1(String sText, String VoiceName, String StartSound, String StopSound, int iVolume, int iRate)
        {
            var synthFormat = new SpeechAudioFormatInfo(8000, AudioBitsPerSample.Sixteen, AudioChannel.Mono);
            var synthesizer = new SpeechSynthesizer();
            var waveStream = new MemoryStream();
            var waveFileStream = new FileStream(@".\\mywave.wav", FileMode.OpenOrCreate);
            var pbuilder = new PromptBuilder();
            var pStyle = new PromptStyle();
            var aSaveFileDialog1 = new SaveFileDialog();
            //---  
            pStyle.Emphasis = PromptEmphasis.None;
            pStyle.Rate = PromptRate.Fast;
            pStyle.Volume = PromptVolume.ExtraLoud;
            pbuilder.StartStyle(pStyle);
            pbuilder.StartParagraph();
            pbuilder.StartVoice(VoiceGender.Male, VoiceAge.Teen, 2);
            pbuilder.StartSentence();
            pbuilder.AppendText("This is some text.");
            pbuilder.EndSentence();
            pbuilder.EndVoice();
            pbuilder.EndParagraph();
            pbuilder.EndStyle();
            synthesizer.SetOutputToAudioStream(waveStream, synthFormat);
            synthesizer.Speak(pbuilder);
            synthesizer.SetOutputToNull();
            waveStream.WriteTo(waveFileStream);
            waveFileStream.Close();
            /*  
            aSaveFileDialog1.Filter = "wave files (*.wav)|*.wav";  
            aSaveFileDialog1.DefaultExt = "*.wav";  
            aSaveFileDialog1.Title = "Stimme als Wave speichern";  
            aSaveFileDialog1.FileName = "";  
            aSaveFileDialog1.ShowDialog();  
            */
        }


        /// <summary> 
        /// Speichert die Stimmt mit START und STOP ton. BETA 
        /// </summary> 
        public void SaveSay2(String sText, String VoiceName, String StartSound, String StopSound, int iVolume, int iRate)
        {
            /*  
            System.Speech.Synthesis.SpeechSynthesizer synthesizer = new System.Speech.Synthesis.SpeechSynthesizer();  
            System.Windows.Forms.SaveFileDialog aSaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();  
            WaveFormat waveFormat = new WaveFormat(audioSampleRate, audioBitsPerSample, audioChannels);  
            System.WaveStream waveStream = CreateWaveStreamfromSamples(waveFormat, samples);   
              
            //System.IO.Stream waveFileStream = Microsoft.Xna. TitleContainer.OpenStream(@"Content\48K16BSLoop.wav");  

            //---  
            synthesizer.SetOutputToWaveFile(".\\test1.wav");  
            //synthesizer.SetOutputToWaveStream(aWaveStream);  
            synthesizer.Volume = iVolume;  
            synthesizer.Rate = iRate;  
            synthesizer.SelectVoice(VoiceName);  
            synthesizer.Speak(sText);  
            //synthesizer.SetOutputToDefaultAudioDevice();  
            aSaveFileDialog1.Filter = "wave files (*.wav)|*.wav";  
            aSaveFileDialog1.DefaultExt = "*.wav";  
            aSaveFileDialog1.Title = "Stimme als Wave speichern";  
            aSaveFileDialog1.FileName = "";  
            aSaveFileDialog1.ShowDialog();  
            */
        }


        /// <summary> 
        /// Liest eine entsprechende Stelle aus dem XML und führt diese aus 
        /// </summary> 
        /// <param name="SelectElement">ELEMENT die er lesen soll, bei mehreren in der XML wird über zufall einer ausgelesen</param> 
        /// <param name="SubSelectElement0">ELEMENT das zwischen den Text kommen soll, Text muss ein {0} string haben!</param> 
        /// <param name="SubSelectElement1">ELEMENT das zwischen den Text kommen soll, Text muss ein {1} string haben!</param> 
        /// <param name="SubSelectElement2">ELEMENT das zwischen den Text kommen soll, Text muss ein {2} string haben!</param> 
        /// <param name="XmlFilePath">Systemstandort des XML's</param> 
        public void XMLRandomElementValue(String SelectElement, String SubSelectElement0, String SubSelectElement1, String SubSelectElement2, String XmlFilePath)
        {
            try
            {
                if (XmlFilePath.Trim() == "")
                {
                    XmlFilePath = ".//VoicesFiles.xml";
                }
                XmlReader aXmlReader;
                var aListValues = new ArrayList();
                var aListVoiceNames = new ArrayList();
                var aListRates = new ArrayList();
                var aListVolumes = new ArrayList();
                var aListStartSound = new ArrayList();
                var aListStopSound = new ArrayList();
                var Rnd = new Random();
                int iRand = int.MinValue;
                string sP1 = String.Empty;
                string sP2 = String.Empty;
                string sP3 = String.Empty;
                string sP4 = String.Empty;
                int iP5 = int.MinValue;
                int iP6 = int.MinValue;
                aXmlReader = XmlReader.Create(@XmlFilePath);
                aListValues.Clear();
                aListVoiceNames.Clear();
                aListRates.Clear();
                aListVolumes.Clear();
                aListStartSound.Clear();
                aListStopSound.Clear();
                while (aXmlReader.Read())
                {
                    if (SelectElement == aXmlReader.Name)
                    {
                        aListVoiceNames.Add(aXmlReader.GetAttribute("VoiceName"));
                        aListRates.Add(aXmlReader.GetAttribute("Rate"));
                        aListVolumes.Add(aXmlReader.GetAttribute("Volume"));
                        aListStartSound.Add(aXmlReader.GetAttribute("StartSound"));
                        aListStopSound.Add(aXmlReader.GetAttribute("StopSound"));
                        aListValues.Add(aXmlReader.ReadElementString());
                    }
                }
                iRand = Rnd.Next(aListValues.Count);
                sP1 = aListValues[iRand] as string;
                if (SubSelectElement0.Trim() != String.Empty )
                {
                    sP1 = String.Format(sP1, SubSelectElement0, SubSelectElement1, SubSelectElement2);
                }
                sP2 = aListVoiceNames[iRand] as string;
                sP3 = aListStartSound[iRand] as string;
                sP4 = aListStopSound[iRand] as string;
                iP5 = Convert.ToInt16(aListVolumes[iRand] as string);
                iP6 = Convert.ToInt16(aListRates[iRand] as string);
                Say(sP1, sP2, sP3, sP4, iP5, iP6);
                aXmlReader.Close();
            }
            catch (Exception e)
            {
                ExceptionHandle("Exception in ClassSpeech:XMLRandomElementValue\n\r" + e);
            }
        }



        /// <summary> 
        /// Erstellt ein default XML file  
        /// </summary> 
        /// <param name="sPath">Dateisystem standort</param> 
        public void CreateDefaultXML(string sPath)
        {
            try
            {
                if (sPath.Trim() == String.Empty)
                {
                    sPath = ".//";
                }
                XmlWriter aXmlWriter;
                var settings = new XmlWriterSettings();
                var aDocument = new XmlDocument();
                XmlElement aElementRoot;
                XmlElement aElemente;
                settings.Indent = true;
                settings.NewLineOnAttributes = true;
                aXmlWriter = XmlWriter.Create(@sPath + "VoicesFiles.xml", settings);
                aElementRoot = aDocument.CreateElement("Voices");
                aDocument.AppendChild(aElementRoot);
                //-- 
                aElemente = aDocument.CreateElement("VoiceStart");
                aElemente.SetAttribute("VoiceName", "ScanSoft Steffi_Dri40_16kHz");
                aElemente.SetAttribute("Rate", "0");
                aElemente.SetAttribute("Volume", "100");
                aElemente.SetAttribute("StartSound", ".\\SpeakSounds\\chimes.wav");
                aElemente.SetAttribute("StopSound", ".\\SpeakSounds\\SpeechOff.wav");
                aElemente.InnerText = "Windows Startet";
                aElementRoot.AppendChild(aElemente);
                //-- 
                aElemente = aDocument.CreateElement("VoiceStart");
                aElemente.SetAttribute("VoiceName", "ScanSoft Steffi_Dri40_16kHz");
                aElemente.SetAttribute("Rate", "0");
                aElemente.SetAttribute("Volume", "100");
                aElemente.SetAttribute("StartSound", ".\\SpeakSounds\\chimes.wav");
                aElemente.SetAttribute("StopSound", ".\\SpeakSounds\\SpeechOff.wav");
                aElemente.InnerText = "Windows Startet Neu";
                aElementRoot.AppendChild(aElemente);
                //-- 
                aElemente = aDocument.CreateElement("VoiceSayHallo");
                aElemente.SetAttribute("VoiceName", "ScanSoft Steffi_Dri40_16kHz");
                aElemente.SetAttribute("Rate", "0");
                aElemente.SetAttribute("Volume", "100");
                aElemente.SetAttribute("StartSound", ".\\SpeakSounds\\chimes.wav");
                aElemente.SetAttribute("StopSound", ".\\SpeakSounds\\SpeechOff.wav");
                aElemente.InnerText = "Hallo";
                aElementRoot.AppendChild(aElemente);
                //-- 
                aDocument.WriteTo(aXmlWriter);
                aXmlWriter.Close();
            }
            catch (Exception e)
            {
                ExceptionHandle("Exception in ClassSpeech:CreateDefaultXML\n\r" + e);
            }
        }

        /// <summary> 
        /// Prüft ob das im ersten Parameter definierte Stimmen Name (sVoiceName) im System vorhanden ist. 
        /// </summary> 
        /// <param name="sVoiceName">Stimmen Name die im System gesucht werden soll.</param> 
        /// <returns>Ob Stimme im System vorhanden ist oder nicht.</returns> 
        private bool CheckVoiceExcist(String sVoiceName)
        {
            bool bResult = false;
            var syn = new SpeechSynthesizer();
            ReadOnlyCollection<InstalledVoice> voices = syn.GetInstalledVoices();
            foreach (InstalledVoice voice in voices)
            {
                if (voice.VoiceInfo.Name == sVoiceName)
                {
                    bResult = true;
                    break;
                }
            }
            return bResult;
        }


        /// <summary> 
        /// Die Hauptmethode die den entsprechenden text aufsagt 
        /// </summary> 
        /// <param name="sText">Der Text die gesprochen werden soll.</param> 
        /// <param name="VoiceName">Die System Stimme die benutzt werden soll.</param> 
        /// <param name="StartSound">Der Sound die vor dem Text kommt.</param> 
        /// <param name="StopSound">Der Sound die nach dem Text kommt.</param> 
        /// <param name="iVolume">Die Lautstärke die benutzt werden soll.</param> 
        /// <param name="iRate">Die Schnelligkeit die benutzt werden soll.</param> 
        public void Say(String sText, String VoiceName, String StartSound, String StopSound, int iVolume, int iRate)
        {
            try
            {
                var syn = new SpeechSynthesizer();
                var myPlayer = new SoundPlayer();
                //---  
                if (StartSound.Trim() != "")
                {
                    myPlayer.SoundLocation = @StartSound;
                    myPlayer.PlaySync();
                }
                //--  
                if (sText.Trim() != "")
                {
                    syn.Volume = iVolume;
                    syn.Rate = iRate;
                    if (CheckVoiceExcist(VoiceName))
                    {
                        syn.SelectVoice(VoiceName);
                    }
                    else
                    {
                        syn.SelectVoice("Microsoft Anna");
                    }
                    syn.Speak(sText);
                }
                //--  
                if (StopSound.Trim() != "")
                {
                    myPlayer.SoundLocation = @StopSound;
                    myPlayer.PlaySync();
                }
            }
            catch (Exception e)
            {
                ExceptionHandle("Exception in ClassSpeech:Say \n\r" + e);
            }
        }
    } //END CLASS "ClassSpeech" 


}
