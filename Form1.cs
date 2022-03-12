using System;
using System.IO;
using System.Windows.Forms;

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