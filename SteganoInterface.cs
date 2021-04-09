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
                //get full text
                //var text = doc.MainDocumentPart.Document.Body.InnerText;
                //Console.WriteLine(rawText);
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
                //string mainPart = doc.MainDocumentPart.Document.Body.InnerText;
                // using (WordprocessingDocument docNew = WordprocessingDocument.Create(@"C:\Users\User\source\repos\HiddenIntervals\newdoc.docx", WordprocessingDocumentType.Document))
                // {
                //     // copy parts from source document to new document
                //     foreach (var part in doc.Parts)
                //         docNew.AddPart(part.OpenXmlPart, part.RelationshipId);
                //
                //     var body = docNew.MainDocumentPart.Document.Body;
                //
                //     foreach (var para in body.Elements<Paragraph>())
                //     {
                //         foreach (var run in para.Elements<Run>())
                //         {
                //             foreach (var text in run.Elements<Text>())
                //             {
                //                 textList.Add(text.Text);
                //             }
                //         }
                //     }
                // }
                // Console.WriteLine(textList);
                // // counter 
                // int lenght_probel = 0;
                // StringBuilder sb = new StringBuilder();
                // foreach (var str in textList)
                // {
                //     lenght_probel += str.Where(x => x == ' ').Count();
                //     foreach (byte b in System.Text.Encoding.Default.GetBytes(str))
                //     {
                //         sb.Append(Convert.ToString(b, 2).PadLeft(8, '0'));
                //     }
                // }
                doc.Close();
            }
        }

        // public string GetRawFileData()
        // {
        //     using (WordprocessingDocument doc = WordprocessingDocument.Open(this.pathFile, false))
        //     {
        //         string buf = doc.MainDocumentPart.Document.Body.InnerText;
        //         data = buf;
        //         return data;
        //     }
        // }

        public void putDataToFile(string data)
        {
            // Create a document by supplying the filepath. 
            // using (WordprocessingDocument wordDocument =
            //     WordprocessingDocument.Create(Directory.GetCurrentDirectory() +  "\\output.docx", WordprocessingDocumentType.Document))
            // {
            //     // Add a main document part. 
            //     MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            //
            //     // Create the document structure and add some text.
            //     mainPart.Document = new Document();
            //     Body body = mainPart.Document.AppendChild(new Body());
            //     Paragraph para = body.AppendChild(new Paragraph());
            //     Run run = para.AppendChild(new Run());
            //     run.AppendChild(new Text(data));
            // }
            using ( WordprocessingDocument package = WordprocessingDocument.Create(Directory.GetCurrentDirectory() +  "\\output.docx",
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
                var docNew = (WordprocessingDocument) doc.Clone(Directory.GetCurrentDirectory() +  "\\output.docx", true);
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
                // using (WordprocessingDocument docNew = WordprocessingDocument.Create(Directory.GetCurrentDirectory() +  "\\output.docx", WordprocessingDocumentType.Document))
                // {
                //     // copy parts from source document to new document
                //     foreach (var part in doc.Parts)
                //         docNew.AddPart(part.OpenXmlPart, part.RelationshipId);
                //
                //     var body = docNew.MainDocumentPart.Document.Body;
                //
                //     var count = 0;
                //
                //     foreach (var para in body.Elements<Paragraph>())
                //     {
                //         foreach (var run in para.Elements<Run>())
                //         {
                //             foreach (var text in run.Elements<Text>())
                //             {
                //                 text.Text = data[count];
                //                 count++;
                //             }
                //         }
                //     }
                //     docNew.Close();
                // }
                doc.Close();
            }
            return true;
        }
    }
    public interface IStegano
    {
        // need preprocessor for trim multiple ' '
        void GetFileData();
        string GetData();
        List<string> GetData(bool flag);
        void putDataToFile(string data);
        bool putDataToFile(List<string> data);
        //void getFileData();
    }
}
