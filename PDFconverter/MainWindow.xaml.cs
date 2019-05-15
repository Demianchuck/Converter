using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SautinSoft;
using SautinSoft.Document;
using System.IO;
using Microsoft.Win32;
using MahApps.Metro.Controls;
using System.Diagnostics;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using GemBox.Presentation;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.Common;
using System.Collections.Specialized;

namespace PDFconverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["UsersDatabaseConnection"].ConnectionString;
        private string dataProvider = ConfigurationManager.ConnectionStrings["UsersDatabaseConnection"].ProviderName;
        const int KEY = 10;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click1(object sender, RoutedEventArgs e) // docx
        {
            //SautinSoft.Document library: 
            //Create: DOCX, PDF, RTF, HTML
            //Load: DOCX, PDF, RTF, HTML
            //Save as: DOCX, PDF, RTF, HTML, Text
            //string inpFile = @"C:\Users\user\Desktop\kkkkkkkk\d\d\dokum\test.docx..docx";
            //string outFile = @"C:\Users\user\Desktop\kkkkkkkk\d\d\dokum\111.pdf";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Doc files (*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string inpFile = openFileDialog.FileName;   // Open document
                string outFile = @"C:\Users\Mariia\Desktop\Convertor\PDFfromDOC.pdf";
                DocumentCore dc = DocumentCore.Load(inpFile);
                dc.Save(outFile);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
                //File.OpenRead(filename);
            }
            //DocumentCore dc = DocumentCore.Load(inpFile);
            //dc.Save(outFile);
            // Open the result for demonstation purposes.
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        private void Button_Click2(object sender, RoutedEventArgs e) //.xlsx 
        {
            //SautinSoft.ExcelToPdf convert.xls and .xlsx to PDF, RTF, DOCX
            //SautinSoft.ExcelToPdf x = new SautinSoft.ExcelToPdf();
            ////x.ConvertFile(@"C:\Users\user\Desktop\kkkkkkkk\d\d\dokum\Business-Budget.xlsx", @"d:C:\Users\user\Desktop\kkkkkkkk\d\d\dokum\Table.pdf");
            //Convert Excel to PDF in memory
            ExcelToPdf x = new ExcelToPdf();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = true;
            //openFileDialog.Filter = "Doc files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog.Filter = "Doc files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string excelFile = openFileDialog.FileName;   // Open document
                string pdfFile = @"C:\Users\Mariia\Desktop\Convertor\PDFfromXLS.pdf";
                // Set PDF as output format.
                x.OutputFormat = SautinSoft.ExcelToPdf.eOutputFormat.Pdf;
                byte[] pdfBytes = null;
                try
                {
                    // Let us say, we have a memory stream with Excel data.
                    using (MemoryStream ms = new MemoryStream(File.ReadAllBytes(excelFile)))
                    {
                        pdfBytes = x.ConvertBytes(ms.ToArray());
                    }
                    // Save pdfBytes to a file for demonstration purposes.
                    File.WriteAllBytes(pdfFile, pdfBytes);
                    System.Diagnostics.Process.Start(pdfFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
            }
            //string excelFile = @"D:\Desktop\1ДЗ\Курсач\d3 (1)\d3\d\d\dokum\Business-Budget.xlsx";
            //string pdfFile = System.IO.Path.ChangeExtension(excelFile, ".pdf");
        }

        private void Button_Click3(object sender, RoutedEventArgs e) //.txt
        {
            //Convert TXT to PDF(PdfSharp, PdfSharp.Charting-wpf)
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Doc files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                try
                {
                    string line = null;
                    int yPoint = 0;
                    string inpFile = openFileDialog.FileName;
                    string outFile = @"C:\Users\Mariia\Desktop\Convertor\PDFfromTXT.pdf";
                    System.IO.TextReader readFile = new StreamReader(inpFile);
                    PdfDocument pdf = new PdfDocument(); // Создаем новый PDF документ
                    pdf.Info.Title = "TXT to PDF";
                    PdfPage pdfPage = pdf.AddPage(); // Создаем пустую страницу
                    XGraphics graph = XGraphics.FromPdfPage(pdfPage); // Получаем объект XGraphics для "рисования" элементов на странице
                                                                      //  XTextFormatter graph = XTextFormatter.FromPdfPage(pdfPage);
                    PdfSharp.Drawing.Layout.XTextFormatter measure = new PdfSharp.Drawing.Layout.XTextFormatter(graph);
                    // Специальная опция для шрифта. Это чтобы русский текст нормально отображался
                    XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);// в фоксіт не змінилося
                    XFont font = new XFont("Verdana", 12, XFontStyle.Regular, options); // Создаем шрифт
                    while (true)
                    {
                        line = readFile.ReadLine();
                        if (line == null)
                        {
                            break; // TODO: might not be correct. Was : Exit While
                        }
                        else
                        {
                            //XRect class to create rectangles for elements to sit within.
                            measure.DrawString(line, font, XBrushes.Black,
                                new XRect(30, yPoint + 30, pdfPage.Width.Point - 50, pdfPage.Height.Point - 30), XStringFormats.TopLeft);//  XRect(зліва, зверху, справа, знизу)
                                //yPoint = yPoint + 30;
                        }
                    }
                    //string pdfFilename = @"D:\Desktop\1ДЗ\Курсач\d3 (1)\d3\d\d\dokum\test.pdf"; // Сохраняем файл 
                    pdf.Save(outFile);
                    readFile.Close();
                    readFile = null;
                    Process.Start(outFile); // запускам сразу в программе просмотра pdf файлов
                }
                //yhhyhy
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            ////// картинка
            ////gfx.DrawImage(XImage.FromFile("путь до картинки\\1.jpg"), 110, 10);
        }
        private void Button_Click4(object sender, RoutedEventArgs e) //.pptx
        {
            //Convert PowerPoint to PDF()
            //If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY"); // для використання  безкоштовної ліцензії
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = true;
            //openFileDialog.Filter = "Doc files (*.ppt)|*.ppt|All files (*.*)|*.*";
            openFileDialog.Filter = "Doc files (*.pptx)|*.pptx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                string inpFile = openFileDialog.FileName;   // Open document
                var presentation = PresentationDocument.Load(inpFile);
                string outFile = @"C:\Users\Mariia\Desktop\Convertor\PDFfromPPT.pdf";
                presentation.Save(outFile);
                //DocumentCore dc = DocumentCore.Load(inpFile);
                //dc.Save(outFile);
                //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
            }
            // In order to achieve the conversion of a loaded PowerPoint file to PDF,
            // we just need to save a PresentationDocument object to desired 
            // output file format.
        }
        private void Authorization_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = @"Data Source=DESKTOP-IRFJCUS;" +
                                   "Initial Catalog=Users;" +
                                   "Integrated Security=True";

            SqlConnection connection;
            using (connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;

                command.CommandText = $"USE {connection.Database}; " +
                      $"SELECT * FROM Logi " +
                      $"WHERE Login='{UserLogin.Text}' AND Password='{CodePass(new StringBuilder(password.Password))}'";

                SqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                {
                        if (UserLogin.Text.ToString() == dr.GetValue(1).ToString())
                        {
                            MessageBox.Show("Авторизація пройшла успішно");
                            //this.Close();
                            //break;
                        }
                        else
                        {
                            MessageBox.Show("Невірний логін чи пароль");
                            //this.Close();
                            //break;
                        }
                }
            }
        }
        private void Registration_Click(object sender, RoutedEventArgs e)
        {
            SqlConnection connection;
            using (connection = new SqlConnection(connectionString))
            {

                connection.Open();
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                //CreateDatabase(command);
                // CreateTable(command)
                string Login = UserLogin.Text;
                //string Password = password.Password;
                StringBuilder pass = new StringBuilder(password.Password);
                string Password = CodePass(pass);

                command.CommandText = "USE Users; Insert Into Logi (Login, Password)" +  $"Values( '{Login}', '{Password}');";
                //command.CommandText = "Insert Into Users (Login, Password) Values(@login, @password";

                if (UserLogin.Text.Equals(""))
                {
                    MessageBox.Show("Введіть логін!");
                }
                else if (password.Password.Equals(""))
                {
                    MessageBox.Show("Введіть пароль!");
                }
                else
                {
                    command.Parameters.Add(new SqlParameter("@Login", UserLogin.Text.ToString()));
                    command.Parameters.Add(new SqlParameter("@Password", password.Password.ToString()));
                    command.ExecuteNonQuery();
                    MessageBox.Show("Авторизація пройшла успішно!");
                   // this.Close();
                }
            }
        }
        static void CreateDatabase(DbCommand command)
        {
            command.CommandText = "CREATE DATABASE Users;";
            command.ExecuteNonQuery();
        }
        static void CreateTable(DbCommand command)
        {
            command.CommandText = "USE Users; CREATE TABLE Logi (Id INT IDENTITY NOT NULL, Login NVARCHAR(10) NOT NULL, Password NVARCHAR(10) NOT NULL, UNIQUE (Id), PRIMARY KEY(Id));";
            command.ExecuteNonQuery();
        }
        static string CodePass(StringBuilder password)
        {
            for (int i = 0; i < password.Length; ++i)
            {
                password[i] = (char)((int)password[i] + KEY);
            }
            return password.ToString();
        }
        private bool CheckTextBox()
        {
            bool isBoxeContentValid = true;
            if (UserLogin.Text.Length == 0)
            {
                isBoxeContentValid = false;
                password.Password = String.Empty;
                UserLogin.Focus();
            }
            return isBoxeContentValid;
        }
    }
}