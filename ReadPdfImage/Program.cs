using PdfUtils; // Used to convert the scanned PDF to JPEG Image
using System;
using System.IO;
using MODI; // Used to Read the Image with OCR

namespace OCRReadImage
{
    static class Program
    {
        // Author: Mohammed Ilyas
        // Contact: mail2mdilyas@gmail.com
        // Querys can be answered from blog: https://ilyasdotnetdeveloper.blogspot.com
        static void Main()
        {
            var args = Environment.GetCommandLineArgs();

            // Considering the Scanned document is uploaded in teh Folder called ScannedDocs inside root directory 
            var searchPath = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(args[0]))), "ScannedDocs");
            
            // Getting scanned pdf's and converting to image. However, this step can be skipped if the file is already an image formt
            foreach (var filename in Directory.GetFiles(searchPath, "*.pdf", SearchOption.TopDirectoryOnly))
            {  
                    Console.WriteLine("Converting to Image {0}", Path.GetFileName(filename)); 
                    var images = PdfImageExtractor.ExtractImages(filename);
                    var directory = Path.GetDirectoryName(filename);

                    foreach (var name in images.Keys)
                    {
                        try
                        {
                       
                            Console.WriteLine("---------------------------------------------------------------");
                            Console.WriteLine("" + Path.GetFileName(filename) + " OCR User Choice Result:");

                            images[name].Save(Path.Combine(directory, name));
                            var path = Path.Combine(directory, name);
                            string extractText = ExtractTextFromImage(path);

                            /*Start Custom Logic to read between start and end token*/
                            string start_token = "";
                            string end_token = "";
                            string result = "";

                            if (!string.IsNullOrEmpty(start_token))
                            {
                            String line;
                            String text = extractText;
                            StringReader reader = new StringReader(text);
                            while (!(line = reader.ReadLine()).Equals(start_token))
                            {
                            //ignore
                            }
                            while (!(line = reader.ReadLine()).StartsWith(end_token))
                            {
                            result += line;
                            }
                            }
                            else
                            {
                            result = extractText;
                            }
                        

                            string subconSign = Between(result, "SUBCONTRACTOR", "CONTRACTOR");
                            string subconApprove = Before(subconSign, "APPROVE", 1);
                            Console.WriteLine("Subcontractor : Approve:" 
                            + CheckForTickMark(subconApprove));

                            string subconReject = Before(subconSign, "REJECT", 2);
                            Console.WriteLine("Subcontractor : Reject:" 
                            + CheckForTickMark(subconReject));


                            string conSign = Between(result, "CONTRACTOR", "CLIENT");
                            string conApprove = Before(conSign, "APPROVE", 1);
                            Console.WriteLine("Contractor : Approve:" + CheckForTickMark(conApprove));

                            string conReject = Before(conSign, "REJECT", 2);
                            Console.WriteLine("Contractor : Reject:" + CheckForTickMark(conReject));

                            string clSign = After(result, "CLIENT");
                            string clApprove = Before(clSign, "APPROVE", 1);
                            Console.WriteLine("Client : Approve:" + CheckForTickMark(clApprove));

                            string clReject = Before(clSign, "REJECT", 2);
                            Console.WriteLine("Client : Reject:" + CheckForTickMark(clReject));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error ["+ex.Message.ToString()+"] in Reading file: "+ name);
                        }  
                } 
            }
            Console.WriteLine("Done.");
            Console.ReadLine();
        }



        private static string ExtractTextFromImage(string filePath)
        {
            Document modiDocument = new Document();
            modiDocument.Create(filePath);
            modiDocument.OCR(MiLANGUAGES.miLANG_ENGLISH);
            MODI.Image modiImage = (modiDocument.Images[0] as MODI.Image);
            string extractedText = modiImage.Layout.Text;
            modiDocument.Close();
            return extractedText;
        }

        public static string Between(this string value, string a, string b)
        {
            if (value.Contains(a) && value.Contains(b))
            {
                int posA = value.IndexOf("\n"+a);
                int posB = value.LastIndexOf(b);
                if (posA == -1)
                {
                    return "";
                }
                if (posB == -1)
                {
                    return "";
                }
                int adjustedPosA = posA + a.Length;
                if (adjustedPosA >= posB)
                {
                    return "";
                }
                return value.Substring(adjustedPosA+1, posB - adjustedPosA);
            }
            else return "not";
        }

        /// <summary>
        /// Get string value after [first] a.
        /// </summary>
        public static string Before(this string value, string a, int lineNumber)
        {
            if (value.Contains(a))
            {
                int lineNo = 0;
                string mylineTxt = string.Empty;
                foreach (var myString in value.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                {
                    lineNo = lineNo + 1;
                    if (lineNumber == lineNo)
                    {
                        mylineTxt = myString;
                    }
                }
                if (lineNo == 2)
                {
                    value = mylineTxt;
                    int posA = value.IndexOf(a);
                    if (posA == -1)
                    {
                        return "";
                    }
                    try
                    {
                        return value.Substring(0, posA);
                    }
                    catch { return ""; }
                }
                else {

                    int posA = value.IndexOf(a);
                    if (posA == -1)
                    {
                        return "";
                    }
                    try
                    {
                        if (value.Substring(posA - 2, 2).TrimEnd().TrimStart().Length == 2)
                            return value.Substring(posA - 2, 2);
                        else
                            return value.Substring(posA - 3, 2);
                    }
                    catch { return ""; }
                }
            }
            else { return "na"; }
        }

        
        /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string After(this string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }

        /// <summary>
        /// Internal Logic to Detect User Choice
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
public static string CheckForTickMark(string a)
{
if (string.IsNullOrEmpty("\r\n"))
    return "Ticked";
if (string.IsNullOrEmpty(a))
    return "Ticked";

if (a.Equals("na"))
    return "Ticked";

a = a.Replace("\r","").Replace("\n", "").Replace(" ", "");
if (string.IsNullOrEmpty(a))
    return "Ticked";
else if (a.Equals("D") || a.ToUpper().Equals("E") || 
        a.ToUpper().Equals("L") || a.ToUpper().Equals("EI") 
        || a.ToUpper().Equals("EL") || a.ToUpper().Equals("FL") 
        || a.ToUpper().Equals("FI") || a.ToUpper().Equals("LI") 
        || a.ToUpper().Equals("D!") || a.Equals("LI")  || 
        a.ToUpper().Equals("DI") || a.ToUpper().Equals("D|") 
        || a.ToUpper().Equals("E|")
        || a.ToUpper().Equals("L|") || a.ToUpper().Equals("VE") 
        || a.ToUpper().Equals("''") || a.Equals("'") 
        || a.Equals("0") || a.ToUpper().Equals("O"))
{ return "NotTicked"; } 
                   
return "Ticked";
}


    }
}
