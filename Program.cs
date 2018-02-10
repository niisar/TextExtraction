using Microsoft.ProjectOxford.Vision;
using Microsoft.ProjectOxford.Vision.Contract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Text.RegularExpressions;
using System.Globalization;

namespace TextExtraction
{
    class Program
    {
        const string API_key = "eb796219d53646d28a6d7392e2849fb5";
        const string API_location = "https://eastasia.api.cognitive.microsoft.com/vision/v1.0";

        static void Main(string[] args)
        {
            string imgToAnalyze = @"C:\Users\noah\Downloads\Compressed\microsoft-azure-cognitive-services-computer-vision-api\04\demos\Demo - Specific Field Logic\TextExtraction\receipt.jpg";
            TextExtraction(imgToAnalyze, false, true);

            Console.ReadLine();
        }

        public static void PrintResults(string[] res)
        {
            foreach (string r in res)
                Console.WriteLine(r);
        }

        public static void TextExtraction(string fname, bool wrds, bool mlines)
        {
            Task.Run(async () =>
            {
                string[] res = await TextExtractionCore(fname, wrds, mlines);

                if (mlines && !wrds)
                    res = MergeLines(res);

                PrintResults(res);

                Console.WriteLine("\nDate: " + GetDate(res));
                Console.WriteLine("Highest amount: " + HighestAmount(res));

            }).Wait();
        }

        public static async Task<string[]> TextExtractionCore(string fname, bool wrds, bool mergelines)
        {
            VisionServiceClient client = new VisionServiceClient(API_key, API_location);
            string[] textres = null;

            if (File.Exists(fname))
                using (Stream stream = File.OpenRead(fname))
                {
                    OcrResults res = await client.RecognizeTextAsync(stream, "unk", false);
                    textres = GetExtracted(res, wrds, mergelines);
                }

            return textres;
        }

        public static string[] MergeLines(string[] lines)
        {
            SortedDictionary<int, string> dict = MergeLinesCore(lines);
            return dict.Values.ToArray();
        }

        public static SortedDictionary<int, string> MergeLinesCore(string[] lines)
        {
            SortedDictionary<int, string> dict = new SortedDictionary<int, string>();

            foreach (string l in lines)
            {
                string[] parts = l.Split('|');

                if (parts.Length == 3)
                {
                    int top = Convert.ToInt32(parts[0]);
                    string str = parts[1];
                    int region = Convert.ToInt32(parts[2]);

                    if (dict.Count > 0 && region != 1)
                    {
                        KeyValuePair<int, string> item = FindClosest(dict, top);

                        if (item.Key != -1)
                            dict[item.Key] = item.Value + " " + str;
                        else
                            dict.Add(top, str);
                    }
                    else
                        dict.Add(top, str);
                }
            }

            return dict;
        }

        public static KeyValuePair<int, string> FindClosest(SortedDictionary<int, string> dict, int top)
        {
            KeyValuePair<int, string> item = new KeyValuePair<int, string>(-1, string.Empty);

            foreach (KeyValuePair<int, string> i in dict)
            {
                int diff = i.Key - top;
                if (Math.Abs(diff) <= 15)
                {
                    item = i;
                    break;
                }
            }

            return item;
        }

        public static string[] GetExtracted(OcrResults res, bool wrds, bool mergelines)
        {
            List<string> items = new List<string>();
            int reg = 1;

            foreach (Region r in res.Regions)
            {
                foreach (Line l in r.Lines)
                    if (wrds)
                        items.AddRange(GetWords(l));
                    else
                        items.Add(GetLineAsString(l, mergelines, reg));

                reg++;
            }

            return items.ToArray();
        }

        public static List<string> GetWords(Line line)
        {
            List<string> words = new List<string>();

            foreach (Word w in line.Words)
                words.Add(w.Text);

            return words;
        }

        public static string GetLineAsString(Line line, bool mergelines, int reg)
        {
            List<string> words = GetWords(line);
            string txt = string.Join(" ", words);

            if (mergelines)
                txt = line.Rectangle.Top.ToString() + "|" + txt + "|" + reg.ToString();

            return words.Count > 0 ? txt : string.Empty;
        }

        public static string ParseDate(string str)
        {
            string result = string.Empty;
            string[] formats = new string[]
                { "dd MMM yy h:mm", "dd MMM yy hh:mm" };

            foreach (string fmt in formats)
            {
                try
                {
                    str = str.Replace("'", "");

                    if (DateTime.TryParseExact(str, fmt, CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out DateTime dateTime))
                    {
                        result = str;
                        break;
                    }
                }
                catch { }
            }

            return result;
        }

        public static string GetDate(string[] res)
        {
            string result = string.Empty;

            foreach (string l in res)
            {
                result = ParseDate(l);
                if (result != string.Empty) break;
            }

            return result;
        }

        public static string HighestAmount(string[] res)
        {
            string result = string.Empty;
            float highest = 0;

            Regex r = new Regex(@"[0-9]+\.[0-9]+");

            foreach (string l in res)
            {
                Match m = r.Match(l);

                if (m != null && m.Value != string.Empty &&
                    Convert.ToDouble(m.Value) > highest)
                    result = m.Value;
            }

            return result;
        }
    }
}
