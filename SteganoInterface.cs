using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using System.Xml;
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
        public string Data;
        public SteganoTXT(string pathFile)
        {
            this.pathFile = pathFile;
        }

        public string GetData()
        {
            return Data;
        }
        
        public List<string> GetData(bool flag)
        {
            throw new Exception("Анта бака?!\nИспользуй не перегруженную версию метода\n");
        }

        public void GetFileData()
        {
            using (var sr = new StreamReader(pathFile))
            {
                Data = sr.ReadToEnd();
            }
        }
        
        public void putDataToFile(string data)
        {
            string date = DateTime.Now.ToString("HH_mm_ss_");
            string path = Directory.GetCurrentDirectory() + "\\" + date + "output.txt";
            if (!File.Exists(path))
            {
                StreamWriter sw = new StreamWriter(path, true, Encoding.UTF8);
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
        
        public bool putDataToFile(List<string> data)
        {
            throw new Exception("Анта бака?!\nИспользуй не перегруженную версию метода\n");
        }
    }
    class SteganoRTF : IStegano
    {
        public string pathFile;
        public string Data;
        public SteganoRTF(string pathFile)
        {
            this.pathFile = pathFile;
        }
        
        public string GetData()
        {
            return Data;
        }
        
        public List<string> GetData(bool flag)
        {
            throw new Exception("Анта бака?!\nИспользуй не перегруженную версию метода\n");
        }
        
        public void GetFileData()
        {
                using (var rtf = new RichTextBox())
                {
                    rtf.Rtf = File.ReadAllText(pathFile);
                    Data = rtf.Text;
                }
        }

        public void putDataToFile(string data)
        {
            string date = DateTime.Now.ToString("HH_mm_ss_");
            string path = Directory.GetCurrentDirectory() + "\\" + date + "output.rtf";
            if (!File.Exists(path))
            {
                using (var rtf = new RichTextBox())
                {
                    rtf.AppendText(data);
                    rtf.SaveFile(path, RichTextBoxStreamType.RichText);
                }
            }
            else
            {
                throw new ApplicationException();
            }
        }
        
        public bool putDataToFile(List<string> data)
        {
            throw new Exception("Анта бака?!\nИспользуй не перегруженную версию метода\n");
        }
    }
    class SteganoDOCX : IStegano
    {
        public List<string> Data = new List<string>();
        public string pathFile;
        //public string data;
        public SteganoDOCX(string pathFile)
        {
            this.pathFile = pathFile;
        }
        
        public string GetData()
        {
            throw new Exception("Анта бака?!\nИспользуй перегруженную версию метода\n");
        }

        public List<string> GetData(bool flag)
        {
            return Data;
        }

        public void GetFileData()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathFile, false))
            {
                var body = doc.MainDocumentPart.Document.Body;
                foreach (var para in body.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            Data.Add(text.Text);
                        }
                    }
                }
                doc.Close();
            }
        }
        

        public void putDataToFile(string data)
        {
            string date = DateTime.Now.ToString("HH_mm_ss_");
            string path = Directory.GetCurrentDirectory() + "\\" + date + "output.docx";
            using ( WordprocessingDocument package = WordprocessingDocument.Create(path,
                WordprocessingDocumentType.Document))
            {
                // Add a new main document part.
                package.AddMainDocumentPart();
                
                var validXmlChars =  data.Where(ch => XmlConvert.IsXmlChar(ch)).ToArray();
                
                string validData = new string(validXmlChars);
                    
                // Create the Document DOM.
                package.MainDocumentPart.Document =
                    new Document(
                        new Body(
                            new Paragraph(
                                new Run(
                                    new Text(validData)))));
                // Save changes to the main document part.
                package.MainDocumentPart.Document.Save();
            }

        }

        public bool putDataToFile(List<string> data)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathFile, false))
            {
                string date = DateTime.Now.ToString("HH_mm_ss_");
                string path = Directory.GetCurrentDirectory() + "\\" + date + "output.docx";
                var docNew = (WordprocessingDocument) doc.Clone(path, true);
                var body = docNew.MainDocumentPart.Document.Body;
                
                var count = 0;
                
                foreach (var para in body.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            text.Text = data[count];
                            count++;
                        }
                    }
                }
                docNew.Close();
                doc.Close();
            }
            return true;
        }
    }
    public interface IStegano
    {
        void GetFileData();
        string GetData();
        List<string> GetData(bool flag);
        void putDataToFile(string data);
        bool putDataToFile(List<string> data);
    }
}
