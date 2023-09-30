using DataEncryptionAndDecryption.Helper;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DataEncryptionAndDecryption
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Document wordDoc = wordApp.Documents.Open(filePath);

                // Read the content of the Word document
                string documentText = wordDoc.Content.Text;

                // Close the Word document and application
                wordDoc.Close();
                wordApp.Quit();

                // Set the content to the RichTextBox
                richTextBox1.Text = documentText;
            }
        }

        private void btnEncrypt_Click(object sender, EventArgs e)
        {
            // Generate a new key and IV
            byte[] key = AesHelper.GenerateKey();
            byte[] iv = AesHelper.GenerateIV();

            // Your plaintext data
            string plainText = richTextBox1.Text;

            // Encrypt the data
            byte[] cipherText = AesHelper.Encrypt(plainText, key, iv);

            string stringCipherText = cipherText.ToString(); 
            textBox1.ReadOnly = true;
            textBox1.Multiline = true; 
            textBox1.Text = Convert.ToBase64String(cipherText);
            ;


        }

        private void btnDecrypt_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] key = AesHelper.GenerateKey(); // Use the correct decryption key
                byte[] iv = AesHelper.GenerateIV();   // Use the correct decryption IV

                string base64CipherText = textBox1.Text;
                byte[] cipherTextBytes = Convert.FromBase64String(base64CipherText);

                string decryptedText = AesHelper.Decrypt(cipherTextBytes, key, iv);

                richTextBox1.Text = decryptedText;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error decrypting: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static string Decrypt(byte[] cipherText, byte[] key, byte[] iv)
        {
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = key;
                aesAlg.IV = iv;
                aesAlg.Padding = PaddingMode.PKCS7; // Specify PKCS7 padding

                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);

                using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        {
                            return srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
        }


    }
}
