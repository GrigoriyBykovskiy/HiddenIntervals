using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using HiddenIntervals;

namespace HiddenIntervals
{
    public partial class Form1 : Form
    {
        public IStegano SteganoContainerEncrypt;
        public IStegano SteganoContainerDecrypt;
        public Form1()
        {
            InitializeComponent();
        }

        // private void Form1_Load()
        // {
        //     richTextBox1.AllowDrop = true;
        //     richTextBox1.DragDrop += richTextBox1_DragDrop;
        // }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;
                    SteganoContainerEncrypt = ObjectFactory.CreateObject(fileName);
                    SteganoContainerEncrypt.GetFileData();
                }
            }
            catch
            {
                MessageBox.Show("Анта бака!\n Ошибка при считывании файла-контейнера!\n", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fileName = openFileDialog.FileName;
                    SteganoContainerDecrypt = ObjectFactory.CreateObject(fileName);
                    SteganoContainerDecrypt.GetFileData();
                }
            }
            catch
            {
                MessageBox.Show("Анта бака!\n Ошибка при считывании файла-контейнера!\n", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // private void richTextBox1_DragDrop(object sender, DragEventArgs e)
        // {
        //     try
        //     {
        //         var data = e.Data.GetData(DataFormats.FileDrop);
        //         if (data != null)
        //         {
        //             var fileNames = data as string[];
        //             if (fileNames.Length > 0)
        //                 richTextBox1.LoadFile(fileNames[0]);
        //         }
        //     }
        //     catch
        //     {
        //         MessageBox.Show("Анта бака!\n Что-то пошло не так!\n", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //     }
        // }  
        // private void richTextBox1_DragEnter(object sender, System.Windows.Forms.DragEventArgs e)
        // {
        //     if (e.Data.GetDataPresent(DataFormats.FileDrop))
        //         e.Effect = DragDropEffects.Copy;
        //     else
        //         e.Effect = DragDropEffects.None;
        // }
        //
        // private void richTextBox1_DragDrop(object sender, System.Windows.Forms.DragEventArgs e)
        // {
        //     string FilePath = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
        //     string[] ss = FilePath.Split('.');
        //     if ((ss[ss.Length-1]=="txt") || (ss[ss.Length-1]=="TXT"))
        //     {
        //         richTextBox1.LoadFile( FilePath, RichTextBoxStreamType.PlainText);
        //     }
        //     if ((ss[ss.Length-1] == "rtf") || (ss[ss.Length-1] == "RTF"))
        //     {
        //         richTextBox1.LoadFile(FilePath, RichTextBoxStreamType.RichText);
        //     }        
        // }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                SteganoModel obj = new SteganoModel();
                if (SteganoContainerDecrypt is SteganoDOCX)
                {
                    var result = SteganoModel.Decrypt(SteganoContainerDecrypt.GetData(true));
                    SteganoContainerDecrypt.putDataToFile(result);
                    richTextBox2.Text = result;
                }
                else
                {
                    var result = SteganoModel.Decrypt(SteganoContainerDecrypt.GetData());
                    SteganoContainerDecrypt.putDataToFile(result);
                    richTextBox2.Text = result;
                }

                MessageBox.Show("Успешно!\nФайл находится в директории:\n" + Directory.GetCurrentDirectory(), "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception excptn)
            {
                MessageBox.Show("Анта бака!\nОшибка при дешифровании!\n" + excptn.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                SteganoModel obj = new SteganoModel();
                if (SteganoContainerEncrypt is SteganoDOCX)
                {
                    var result = SteganoModel.Encrypt(SteganoContainerEncrypt.GetData(true), richTextBox1.Text);
                    //richTextBox2.Text = string.Join("", result.ToArray());
                    //richTextBox2.Text = result.Aggregate((x, y) => x + y);
                    SteganoContainerEncrypt.putDataToFile(result);
                }
                else
                {
                    var result = SteganoModel.Encrypt(SteganoContainerEncrypt.GetData(), richTextBox1.Text);
                    SteganoContainerEncrypt.putDataToFile(result);
                }
                MessageBox.Show("Успешно!\nФайл находится в директории:\n" + Directory.GetCurrentDirectory(), "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception excptn)
            {
                MessageBox.Show("Анта бака!\nОшибка при шифровании!\n" + excptn.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}