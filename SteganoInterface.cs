using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HiddenIntervals
{
    static class ObjectFactory
    {
        public static IStegano CreateObject(string pathFile)
        {
            IStegano obj = null;
            if (File.Exists(pathFile))
            {
                string extension = Path.GetExtension(pathFile);
                switch (extension)
                {
                    case ".docx":
                        obj = new SteganoDOCX(pathFile);
                        break;
                    case ".rtf":
                        obj = new SteganoRTF(pathFile);
                        break;
                    case ".txt":
                        obj = new SteganoTXT(pathFile);
                        break;
                    default:
                        throw new ArgumentException();
                }
            }
            return obj;
        }
    }
    class SteganoTXT : IStegano
    {
        public string pathFile;
        public string data;
        public SteganoTXT(string pathFile)
        {
            this.pathFile = pathFile;
        }

        public string getRawFileData()
        {
            using (var sr = new StreamReader(pathFile))
            {
                data = sr.ReadToEnd();
                return data;
            }
        }
        
        public void putDataToFile(string data)
        {
            if (!File.Exists(Directory.GetCurrentDirectory() + "\\output.txt"))
            {
                StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() +  "\\output.txt", true, Encoding.UTF8);
                foreach (var ch in data)
                {
                    sw.Write(ch);
                }
                sw.Close();
            }
            else
            {
                throw new ApplicationException("File already exist!");
            }
            
        }
    }
    class SteganoRTF : IStegano
    {
        public string pathFile;
        public string data;
        public SteganoRTF(string pathFile)
        {
            this.pathFile = pathFile;
        }
        public string getRawFileData()
        {
                using (var rtf = new RichTextBox())
                {
                    rtf.Rtf = File.ReadAllText(pathFile);
                    data = rtf.Text;
                    return data;
                }
        }

        public void putDataToFile(string data)
        {
            if (!File.Exists(Directory.GetCurrentDirectory() + "\\output.rtf"))
            {
                using (var rtf = new RichTextBox())
                {
                    rtf.AppendText(data);
                    rtf.SaveFile(Directory.GetCurrentDirectory() +  "\\output.rtf", RichTextBoxStreamType.RichText);
                }
            }
            else
            {
                throw new ApplicationException();
            }
        }
    }
    class SteganoDOCX : IStegano
    {
        public List<string> textList = new List<string>();
        public string pathFile;
        public string data;
        public SteganoDOCX(string pathFile)
        {
            this.pathFile = pathFile;
        }
        public void getFileData()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(this.pathFile, false))
            {
                //get full text
                //var text = doc.MainDocumentPart.Document.Body.InnerText;
                //Console.WriteLine(rawText);

                //string mainPart = doc.MainDocumentPart.Document.Body.InnerText;
                using (WordprocessingDocument docNew = WordprocessingDocument.Create(@"C:\Users\User\source\repos\HiddenIntervals\newdoc.docx", WordprocessingDocumentType.Document))
                {
                    // copy parts from source document to new document
                    foreach (var part in doc.Parts)
                        docNew.AddPart(part.OpenXmlPart, part.RelationshipId);

                    var body = docNew.MainDocumentPart.Document.Body;

                    foreach (var para in body.Elements<Paragraph>())
                    {
                        foreach (var run in para.Elements<Run>())
                        {
                            foreach (var text in run.Elements<Text>())
                            {
                                textList.Add(text.Text);
                            }
                        }
                    }
                }
                Console.WriteLine(textList);
                // counter 
                int lenght_probel = 0;
                StringBuilder sb = new StringBuilder();
                foreach (var str in textList)
                {
                    lenght_probel += str.Where(x => x == ' ').Count();
                    foreach (byte b in System.Text.Encoding.Default.GetBytes(str))
                    {
                        sb.Append(Convert.ToString(b, 2).PadLeft(8, '0'));
                    }
                }
            }
        }

        public string getRawFileData()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(this.pathFile, false))
            {
                string buf = doc.MainDocumentPart.Document.Body.InnerText;
                data = buf;
                return data;
            }
        }

        public void putDataToFile(string data)
        {
            throw new NotImplementedException();
        }
    }
    public interface IStegano
    {
        // need preprocessor for trim multiple ' '
        string getRawFileData();
        void putDataToFile(string data);
        //void getFileData();
    }
}
