using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace OfficeAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
          
        }

        private void Ribbon1_AnonymizeData(object sender, RibbonUIEventArgs e)
        {
            var fileDialog = new Microsoft.Office.Tools.Ribbon.FileDialog();
            fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            fileDialog.Filter = "Office Files|*.docx;*.xlsx;*.pptx";
            fileDialog.ShowDialog();

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                var file = fileDialog.FileName;
                var fileStream = new FileStream(file, FileMode.Open);
                var fileContent = new byte[fileStream.Length];
                fileStream.Read(fileContent, 0, fileContent.Length);

                var encryptedData = DataAnonymizer.EncryptData(fileContent, "your_secret_password");

                var encryptedFileStream = new FileStream("Anonimized_" + file, FileMode.Create);
                encryptedFileStream.Write(encryptedData, 0, encryptedData.Length);
                encryptedFileStream.Close();

                MessageBox.Show("Data anonymized successfully!");
            }
        }

        private void Ribbon1_DeAnonymizeData(object sender, RibbonUIEventArgs e)
        {
          
            var fileDialog = new Microsoft.Office.Tools.Ribbon.FileDialog();
            fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            fileDialog.Filter = "Anonymized files|Anonimized_*.docx;Anonimized_*.xlsx;Anonimized_*.pptx";
            fileDialog.ShowDialog();

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                var file = fileDialog.FileName;
                var fileStream = new FileStream(file, FileMode.Open);
                var fileContent = new byte[fileStream.Length];
                fileStream.Read(fileContent, 0, fileContent.Length);

                var decryptedData = DataAnonymizer.DecryptData(fileContent, "your_secret_password");

                var decryptedFileStream = new FileStream(file.Substring(11), FileMode.Create);
                decryptedFileStream.Write(decryptedData, 0, decryptedData.Length);
                decryptedFileStream.Close();

                MessageBox.Show("Successfully deanonymized data!");
            }
        }
    }
}
