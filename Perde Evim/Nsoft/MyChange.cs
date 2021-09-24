using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Speech.Synthesis;
using System.Windows.Forms;

namespace Nsoft
{
    class MyChange
    {
        public static string HtmlOxumaq(String url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());
            string s = sr.ReadToEnd();
            sr.Close();
            return s;
        }

        public static string UsdMezenne()
        {
            string[] arr = { "http://www.cbar.az/", "", "" };

            try
            {
                arr[1] = MyChange.HtmlOxumaq(arr[0]);

                for (int k = 0; k < arr[1].Length; k++)
                {
                    if (arr[1].Substring(k, 5) == "1 USD")
                    {
                        arr[2] = arr[1].Substring(k + 8, 6);
                        break;
                    }
                }
            }
            catch { };

            return arr[2];
        }

        public static string HavaBaku()
        {
            string url = "http://www.euronews.com/weather/europe/azerbaijan/baku";
            string s = "";
            string result = "";
            
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream());
                s = sr.ReadToEnd();
                sr.Close();

                for (int k = 0; k < s.Length; k++)
                {
                    if (s.Substring(k, 5) == "C ltr")
                    {
                        result = s.Substring(k + 17, 3); // hava temperaturunun kordinasiyasi
                        k = s.Length;
                    }

                }

            }
            catch { };

            return result;
        }

        public static void seslendirme(String metn)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Adult, 1, DilDeyisme); // to change VoiceGender and VoiceAge check out those links below
            synthesizer.Volume = 100;  // (0 - 100)
            synthesizer.Rate = -3;     // (-10 - 10)
            synthesizer.Speak(metn);
        }

        public static void SetKeyboardLayout(InputLanguage layout) //Keyboard Language Change
        {
            InputLanguage.CurrentInputLanguage = layout;
        }

        public static InputLanguage GetInputLanguageByName(string inputName) //Help to Change Keyboard Language
        {
            foreach (InputLanguage lang in InputLanguage.InstalledInputLanguages)
            {
                if (lang.Culture.EnglishName.ToLower().StartsWith(inputName.ToLower()))
                    return lang;
            }
            return null;

        }

        public static string TarixSozle(DateTime date)
        {
            string result = "";

            switch (date.Month)
            {
                case 1: result = "Yanvar"; break;
                case 2: result = "Fevral"; break;
                case 3: result = "Mart"; break;
                case 4: result = "Aprel"; break;
                case 5: result = "May"; break;
                case 6: result = "İyun"; break;
                case 7: result = "İyul"; break;
                case 8: result = "Avqust"; break;
                case 9: result = "Sentyabr"; break;
                case 10: result = "Oktyabr"; break;
                case 11: result = "Noyabr"; break;
                case 12: result = "Dekabr"; break;
                default: break;

            }

            return result;
        }

        public static CultureInfo DilDeyisme = new CultureInfo("AZ");

        public static void FindAndReplace(Word.Application word, object findText, object replaceText)
        {
            word.Selection.Find.ClearFormatting();
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = true;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 1;
            object wrap = 2;

            word.Selection.Find.Execute(ref findText, ref matchCase,
            ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
            ref matchAllWordForms, ref forward, ref wrap, ref format,
            ref replaceText, ref replace, ref matchKashida,
            ref matchDiacritics,
            ref matchAlefHamza, ref matchControl);
        }

        public static string ReqemToMetn(double Reqem) // 12.34 - on iki manat 34 qepik
        {
            string n = Reqem.ToString();
            string Result = "";
            int s2 = 0;

            for (int s = 0; s < n.Length; s++) if (n.Substring(s, 1) == ".") { n = n.Substring(0, s); s2 = s + 1; break; }

            if (n.Length > 9) { return "max: '999999999'";}
            if (n.Length > 8) { if (n.Substring(n.Length - 9, 1) == "1") Result = "bir yüz"; else if (n.Substring(n.Length - 9, 1) == "2") Result = "iki yüz"; else if(n.Substring(n.Length - 9, 1) == "3") Result = "üç yüz"; else if(n.Substring(n.Length - 9, 1) == "4") Result = "dörd yüz"; if (n.Substring(n.Length - 9, 1) == "5") Result = "beş yüz"; else if (n.Substring(n.Length - 9, 1) == "6") Result = "altı yüz"; else if (n.Substring(n.Length - 9, 1) == "7") Result = "yeddi yüz"; else if (n.Substring(n.Length - 9, 1) == "8") Result = "səkkiz yüz"; else if (n.Substring(n.Length - 9, 1) == "9") Result = "doqquz yüz"; }
            if (n.Length > 7) { if (n.Substring(n.Length - 8, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 8, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 8, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 8, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 8, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 8, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 8, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 8, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 8, 1) == "9") Result += " doxsan"; }
            if (n.Length > 6) { if (n.Substring(n.Length - 7, 1) == "0") Result += " milyon"; else if (n.Substring(n.Length - 7, 1) == "1") Result += " bir milyon"; else if (n.Substring(n.Length - 7, 1) == "2") Result += " iki milyon"; else if (n.Substring(n.Length - 7, 1) == "3") Result += " üç milyon"; else if (n.Substring(n.Length - 7, 1) == "4") Result += " dörd milyon"; else if (n.Substring(n.Length - 7, 1) == "5") Result += " beş milyon"; else if (n.Substring(n.Length - 7, 1) == "6") Result += " altı milyon"; else if (n.Substring(n.Length - 7, 1) == "7") Result += " yeddi milyon"; else if (n.Substring(n.Length - 7, 1) == "8") Result += " səkkiz milyon"; else if (n.Substring(n.Length - 7, 1) == "9") Result += " doqquz milyon"; }
            if (n.Length > 5) { if (n.Substring(n.Length - 6, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 6, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 6, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 6, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 6, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 6, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 6, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 6, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 6, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 4) { if (n.Substring(n.Length - 5, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 5, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 5, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 5, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 5, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 5, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 5, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 5, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 5, 1) == "9") Result += " doxsan"; }
            if (n.Length > 3) { if (n.Substring(n.Length - 4, 1) == "0" && n.Length > 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length < 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length == 6) Result += " min"; else if (n.Substring(n.Length - 4, 1) == "1") Result += " bir min"; else if (n.Substring(n.Length - 4, 1) == "2") Result += " iki min"; else if (n.Substring(n.Length - 4, 1) == "3") Result += " üç min"; else if (n.Substring(n.Length - 4, 1) == "4") Result += " dörd min"; else if (n.Substring(n.Length - 4, 1) == "5") Result += " beş min"; else if (n.Substring(n.Length - 4, 1) == "6") Result += " altı min"; else if (n.Substring(n.Length - 4, 1) == "7") Result += " yeddi min"; else if (n.Substring(n.Length - 4, 1) == "8") Result += " səkkiz min"; else if (n.Substring(n.Length - 4, 1) == "9") Result += " doqquz min"; }
            if (n.Length > 2) { if (n.Substring(n.Length - 3, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 3, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 3, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 3, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 3, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 3, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 3, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 3, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 3, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 1) { if (n.Substring(n.Length - 2, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 2, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 2, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 2, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 2, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 2, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 2, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 2, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 2, 1) == "9") Result += " doxsan"; }
            if (n.Length > 0) { if (n.Substring(n.Length - 1, 1) == "1") Result += " bir"; else if (n.Substring(n.Length - 1, 1) == "2") Result += " iki"; else if (n.Substring(n.Length - 1, 1) == "3") Result += " üç"; else if (n.Substring(n.Length - 1, 1) == "4") Result += " dörd"; else if (n.Substring(n.Length - 1, 1) == "5") Result += " beş"; else if (n.Substring(n.Length - 1, 1) == "6") Result += " altı"; else if (n.Substring(n.Length - 1, 1) == "7") Result += " yeddi"; else if (n.Substring(n.Length - 1, 1) == "8") Result += " səkkiz"; else if (n.Substring(n.Length - 1, 1) == "9") Result += " doqquz"; else if (n == "0") Result = "Sıfır"; }

            if (Result.Substring(0, 1) == " ") Result = Result.Substring(1, Result.Length - 1); Result = Result.Substring(0, 1).ToUpper(DilDeyisme) + Result.Substring(1, Result.Length - 1).ToLower(DilDeyisme);
            if (s2 != 0) Result += ", " + Reqem.ToString().Substring(s2, Reqem.ToString().Length - s2);
            else Result += ", 0";
            if (Result.Substring(Result.Length - 3, 1) == ",") Result += "0";

            Result = Result.Substring(0, Result.Length - 4) + " manat " + Result.Substring(Result.Length - 2, 2) + " qəpik";
            return Result;
        }

        public static string ReqemToMetnValyuta(double Reqem, string Valyuta, string Qepik) // 12.34 - on iki manat 34 qepik
        {
            string n = Reqem.ToString();
            string Result = "";
            int s2 = 0;

            for (int s = 0; s < n.Length; s++) if (n.Substring(s, 1) == ".") { n = n.Substring(0, s); s2 = s + 1; break; }

            if (n.Length > 9) { return "max: '999999999'"; }
            if (n.Length > 8) { if (n.Substring(n.Length - 9, 1) == "1") Result = "bir yüz"; else if (n.Substring(n.Length - 9, 1) == "2") Result = "iki yüz"; else if (n.Substring(n.Length - 9, 1) == "3") Result = "üç yüz"; else if (n.Substring(n.Length - 9, 1) == "4") Result = "dörd yüz"; if (n.Substring(n.Length - 9, 1) == "5") Result = "beş yüz"; else if (n.Substring(n.Length - 9, 1) == "6") Result = "altı yüz"; else if (n.Substring(n.Length - 9, 1) == "7") Result = "yeddi yüz"; else if (n.Substring(n.Length - 9, 1) == "8") Result = "səkkiz yüz"; else if (n.Substring(n.Length - 9, 1) == "9") Result = "doqquz yüz"; }
            if (n.Length > 7) { if (n.Substring(n.Length - 8, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 8, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 8, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 8, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 8, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 8, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 8, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 8, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 8, 1) == "9") Result += " doxsan"; }
            if (n.Length > 6) { if (n.Substring(n.Length - 7, 1) == "0") Result += " milyon"; else if (n.Substring(n.Length - 7, 1) == "1") Result += " bir milyon"; else if (n.Substring(n.Length - 7, 1) == "2") Result += " iki milyon"; else if (n.Substring(n.Length - 7, 1) == "3") Result += " üç milyon"; else if (n.Substring(n.Length - 7, 1) == "4") Result += " dörd milyon"; else if (n.Substring(n.Length - 7, 1) == "5") Result += " beş milyon"; else if (n.Substring(n.Length - 7, 1) == "6") Result += " altı milyon"; else if (n.Substring(n.Length - 7, 1) == "7") Result += " yeddi milyon"; else if (n.Substring(n.Length - 7, 1) == "8") Result += " səkkiz milyon"; else if (n.Substring(n.Length - 7, 1) == "9") Result += " doqquz milyon"; }
            if (n.Length > 5) { if (n.Substring(n.Length - 6, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 6, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 6, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 6, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 6, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 6, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 6, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 6, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 6, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 4) { if (n.Substring(n.Length - 5, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 5, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 5, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 5, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 5, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 5, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 5, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 5, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 5, 1) == "9") Result += " doxsan"; }
            if (n.Length > 3) { if (n.Substring(n.Length - 4, 1) == "0" && n.Length > 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length < 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length == 6) Result += " min"; else if (n.Substring(n.Length - 4, 1) == "1") Result += " bir min"; else if (n.Substring(n.Length - 4, 1) == "2") Result += " iki min"; else if (n.Substring(n.Length - 4, 1) == "3") Result += " üç min"; else if (n.Substring(n.Length - 4, 1) == "4") Result += " dörd min"; else if (n.Substring(n.Length - 4, 1) == "5") Result += " beş min"; else if (n.Substring(n.Length - 4, 1) == "6") Result += " altı min"; else if (n.Substring(n.Length - 4, 1) == "7") Result += " yeddi min"; else if (n.Substring(n.Length - 4, 1) == "8") Result += " səkkiz min"; else if (n.Substring(n.Length - 4, 1) == "9") Result += " doqquz min"; }
            if (n.Length > 2) { if (n.Substring(n.Length - 3, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 3, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 3, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 3, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 3, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 3, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 3, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 3, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 3, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 1) { if (n.Substring(n.Length - 2, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 2, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 2, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 2, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 2, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 2, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 2, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 2, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 2, 1) == "9") Result += " doxsan"; }
            if (n.Length > 0) { if (n.Substring(n.Length - 1, 1) == "1") Result += " bir"; else if (n.Substring(n.Length - 1, 1) == "2") Result += " iki"; else if (n.Substring(n.Length - 1, 1) == "3") Result += " üç"; else if (n.Substring(n.Length - 1, 1) == "4") Result += " dörd"; else if (n.Substring(n.Length - 1, 1) == "5") Result += " beş"; else if (n.Substring(n.Length - 1, 1) == "6") Result += " altı"; else if (n.Substring(n.Length - 1, 1) == "7") Result += " yeddi"; else if (n.Substring(n.Length - 1, 1) == "8") Result += " səkkiz"; else if (n.Substring(n.Length - 1, 1) == "9") Result += " doqquz"; else if (n == "0") Result = "Sıfır"; }

            if (Result.Substring(0, 1) == " ") Result = Result.Substring(1, Result.Length - 1); Result = Result.Substring(0, 1).ToUpper(DilDeyisme) + Result.Substring(1, Result.Length - 1).ToLower(DilDeyisme);
            if (s2 != 0) Result += ", " + Reqem.ToString().Substring(s2, Reqem.ToString().Length - s2);
            else Result += ", 0";
            if (Result.Substring(Result.Length - 3, 1) == ",") Result += "0";

            Result = Result.Substring(0, Result.Length - 4) + " " + Valyuta + " " + Result.Substring(Result.Length - 2, 2) + " " + Qepik;
            return Result;
        }

        public static string ReqemToMetn(int Reqem) // 123 - bir yuz iyirmi uc
        {
            string n = Reqem.ToString();
            string Result = "";

            if (n.Length > 9) { return "max: '999999999'"; }
            if (n.Length > 8) { if (n.Substring(n.Length - 9, 1) == "1") Result = "bir yüz"; else if (n.Substring(n.Length - 9, 1) == "2") Result = "iki yüz"; else if (n.Substring(n.Length - 9, 1) == "3") Result = "üç yüz"; else if (n.Substring(n.Length - 9, 1) == "4") Result = "dörd yüz"; if (n.Substring(n.Length - 9, 1) == "5") Result = "beş yüz"; else if (n.Substring(n.Length - 9, 1) == "6") Result = "altı yüz"; else if (n.Substring(n.Length - 9, 1) == "7") Result = "yeddi yüz"; else if (n.Substring(n.Length - 9, 1) == "8") Result = "səkkiz yüz"; else if (n.Substring(n.Length - 9, 1) == "9") Result = "doqquz yüz"; }
            if (n.Length > 7) { if (n.Substring(n.Length - 8, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 8, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 8, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 8, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 8, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 8, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 8, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 8, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 8, 1) == "9") Result += " doxsan"; }
            if (n.Length > 6) { if (n.Substring(n.Length - 7, 1) == "0") Result += " milyon"; else if (n.Substring(n.Length - 7, 1) == "1") Result += " bir milyon"; else if (n.Substring(n.Length - 7, 1) == "2") Result += " iki milyon"; else if (n.Substring(n.Length - 7, 1) == "3") Result += " üç milyon"; else if (n.Substring(n.Length - 7, 1) == "4") Result += " dörd milyon"; else if (n.Substring(n.Length - 7, 1) == "5") Result += " beş milyon"; else if (n.Substring(n.Length - 7, 1) == "6") Result += " altı milyon"; else if (n.Substring(n.Length - 7, 1) == "7") Result += " yeddi milyon"; else if (n.Substring(n.Length - 7, 1) == "8") Result += " səkkiz milyon"; else if (n.Substring(n.Length - 7, 1) == "9") Result += " doqquz milyon"; }
            if (n.Length > 5) { if (n.Substring(n.Length - 6, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 6, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 6, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 6, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 6, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 6, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 6, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 6, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 6, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 4) { if (n.Substring(n.Length - 5, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 5, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 5, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 5, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 5, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 5, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 5, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 5, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 5, 1) == "9") Result += " doxsan"; }
            if (n.Length > 3) { if (n.Substring(n.Length - 4, 1) == "0" && n.Length > 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length < 6 && n.Substring(n.Length - 5, 1) != "0") Result += " min"; else if (n.Substring(n.Length - 4, 1) == "0" && n.Length == 6) Result += " min"; else if (n.Substring(n.Length - 4, 1) == "1") Result += " bir min"; else if (n.Substring(n.Length - 4, 1) == "2") Result += " iki min"; else if (n.Substring(n.Length - 4, 1) == "3") Result += " üç min"; else if (n.Substring(n.Length - 4, 1) == "4") Result += " dörd min"; else if (n.Substring(n.Length - 4, 1) == "5") Result += " beş min"; else if (n.Substring(n.Length - 4, 1) == "6") Result += " altı min"; else if (n.Substring(n.Length - 4, 1) == "7") Result += " yeddi min"; else if (n.Substring(n.Length - 4, 1) == "8") Result += " səkkiz min"; else if (n.Substring(n.Length - 4, 1) == "9") Result += " doqquz min"; }
            if (n.Length > 2) { if (n.Substring(n.Length - 3, 1) == "1") Result += " bir yüz"; else if (n.Substring(n.Length - 3, 1) == "2") Result += " iki yüz"; else if (n.Substring(n.Length - 3, 1) == "3") Result += " üç yüz"; else if (n.Substring(n.Length - 3, 1) == "4") Result += " dörd yüz"; else if (n.Substring(n.Length - 3, 1) == "5") Result += " beş yüz"; else if (n.Substring(n.Length - 3, 1) == "6") Result += " altı yüz"; else if (n.Substring(n.Length - 3, 1) == "7") Result += " yeddi yüz"; else if (n.Substring(n.Length - 3, 1) == "8") Result += " səkkiz yüz"; else if (n.Substring(n.Length - 3, 1) == "9") Result += " doqquz yüz"; }
            if (n.Length > 1) { if (n.Substring(n.Length - 2, 1) == "1") Result += " on"; else if (n.Substring(n.Length - 2, 1) == "2") Result += " iyirmi"; else if (n.Substring(n.Length - 2, 1) == "3") Result += " otuz"; else if (n.Substring(n.Length - 2, 1) == "4") Result += " qırx"; else if (n.Substring(n.Length - 2, 1) == "5") Result += " əlli"; else if (n.Substring(n.Length - 2, 1) == "6") Result += " altmış"; else if (n.Substring(n.Length - 2, 1) == "7") Result += " yetmiş"; else if (n.Substring(n.Length - 2, 1) == "8") Result += " səksən"; else if (n.Substring(n.Length - 2, 1) == "9") Result += " doxsan"; }
            if (n.Length > 0) { if (n.Substring(n.Length - 1, 1) == "1") Result += " bir"; else if (n.Substring(n.Length - 1, 1) == "2") Result += " iki"; else if (n.Substring(n.Length - 1, 1) == "3") Result += " üç"; else if (n.Substring(n.Length - 1, 1) == "4") Result += " dörd"; else if (n.Substring(n.Length - 1, 1) == "5") Result += " beş"; else if (n.Substring(n.Length - 1, 1) == "6") Result += " altı"; else if (n.Substring(n.Length - 1, 1) == "7") Result += " yeddi"; else if (n.Substring(n.Length - 1, 1) == "8") Result += " səkkiz"; else if (n.Substring(n.Length - 1, 1) == "9") Result += " doqquz"; else if (n == "0") Result = "Sıfır"; }

            if (Result.Substring(0, 1) == " ") Result = Result.Substring(1, Result.Length - 1); Result = Result.Substring(0, 1).ToUpper(DilDeyisme) + Result.Substring(1, Result.Length - 1).ToLower(DilDeyisme);
         
            return Result;
        }
    }
}
