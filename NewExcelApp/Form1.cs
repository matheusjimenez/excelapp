using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary;

namespace NewExcelApp
{
    public partial class Form1 : Form
    {
        ReadFile _readfile = new ReadFile();
        public Form1()
        {
            InitializeComponent();
        }

        public class PessoaTabela1
        {
            public string Nome;
            public string Email;
            public string CarimboHora;
            public string InstituicaoEnsino;
            public string Curso;
        }

        public class PessoaTabela2
        {
            public string Nome;
            public string Email;
            public string NumeroIncricao;
            public string NumeroPedido;
        }

        public class PessoaCompleta
        {
            public string Nome;
            public string Email;
            public string CarimboHora;
            public string InstituicaoEnsino;
            public string Curso;
            public string NumeroIncricao;
            public string NumeroPedido;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            string NomePlanilha1 = ShowDialog("Nome planilha GoogleForms", "");
            string NomePlanilha2 = ShowDialog("Nome planilha Sympla", "");
            //string localizacaoColunasTabela1 = ShowDialog("Nome,Email,CarimboHora,Curso,InstituicaoEnsino", "Posição dos campos tabela 1");
            //string localizacaoColunasTabela2 = ShowDialog("Nome,Email,NumeroIncricao,NumeroPedido", "Posição dos campos tabela 2");


            if (NomePlanilha1 == string.Empty || 
                NomePlanilha2 == string.Empty
                )
                return;

            //var arrayLocalizacaoColunasTabela1 = localizacaoColunasTabela1.Split(',');
            //var arrayLocalizacaoColunasTabela2 = localizacaoColunasTabela2.Split(',');

            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = _readfile.ReadExcel(filePath, fileExt, NomePlanilha1); //read excel file  
                        dataGridView1.Visible = true;
                        dataGridView1.DataSource = dtExcel;

                        var emailListTable1 = dtExcel.AsEnumerable().Select(x =>
                            new PessoaTabela1
                            {
                                Nome = x[2].ToString(),
                                Email = x[1].ToString(),
                                CarimboHora = x[0].ToString(),
                                Curso = x[4].ToString(),
                                InstituicaoEnsino = x[3].ToString()
                            }).ToList();
                        ////////////////////////////////////////////////////////////////////////////
                        DataTable dtExcel2 = new DataTable();
                        dtExcel2 = _readfile.ReadExcel(filePath, fileExt, NomePlanilha2); //read excel file
                        dataGridView2.Visible = true;
                        dataGridView2.DataSource = dtExcel2;

                        var emailListTable2 = dtExcel2.AsEnumerable().Select(x => 
                            new PessoaTabela2 
                            {
                                Nome = x[2].ToString() + x[3].ToString(),
                                Email =  x[7].ToString(),
                                NumeroIncricao = x[1].ToString(),
                                NumeroPedido = x[6].ToString()
                            }).ToList();
                        //var emailListTable2 = dtExcel2.AsEnumerable().Select(x => new Pessoa { Email = x[8].ToString(), Nome = x[2].ToString() + x[3].ToString() }).ToList();

                        List<PessoaCompleta> resultingMailList = new List<PessoaCompleta>();

                        

                        foreach(var item in emailListTable1)
                        {
                            if(emailListTable2.Exists(x => x.Email.ToLower() == item.Email.ToLower()))
                            {
                                resultingMailList.Add(new PessoaCompleta
                                {
                                    Nome = item.Nome,
                                    Email = item.Email,
                                    Curso = item.Curso,
                                    CarimboHora = item.CarimboHora,
                                    InstituicaoEnsino = item.InstituicaoEnsino,
                                    NumeroIncricao = emailListTable2.Find(x => x.Email.ToLower() == item.Email.ToLower()).NumeroIncricao,
                                    NumeroPedido = emailListTable2.Find(x => x.Email.ToLower() == item.Email.ToLower()).NumeroPedido
                                });
                            }
                        }

                        var count = 0;
                        DataTable dt = new DataTable();
                        dt.Columns.AddRange(new DataColumn[8] { 
                            new DataColumn("id"), new DataColumn("Nome", typeof(string)), new DataColumn("Email", typeof(string)), 
                            new DataColumn("Curso", typeof(string)), new DataColumn("CarimboHora ", typeof(string)), new DataColumn("InstituicaoEnsino ", typeof(string)),
                            new DataColumn("NumeroIngresso ", typeof(string)), new DataColumn("NumeroPedido", typeof(string))
                        });
                        foreach(var item in resultingMailList)
                        {
                            dt.Rows.Add(count, item.Nome, item.Email, item.Curso, item.CarimboHora, item.InstituicaoEnsino, item.NumeroIncricao, item.NumeroPedido);
                            count++;
                        }

                        dataGridView3.Visible = true;
                        dataGridView3.DataSource = dt;

                        DataSet ds = new DataSet("New_DataSet");

                        ds.Tables.Add(dt);


                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop).ToString();

                        CreateFolder(desktopPath+"\\Excel");
                        ExcelLibrary.DataSetHelper.CreateWorkbook(desktopPath + "\\Excel\\ListaDeParticipantes"+DateTime.Now.TimeOfDay.Minutes.ToString()+".xls", ds);;
                        //ExcelLibrary.DataSetHelper.CreateWorkbook("C:\\Users\\Burned\\Desktop\\excel\\MyExcelFile.xls", ds);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        public static string ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 50, Top = 20, Text = text, Width = 300 };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 250, Width = 100, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        public static void CreateFolder(string path)
        {
            // Specify the directory you want to manipulate.
            

            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    Console.WriteLine("That path exists already.");
                    return;
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
        }
        private void close_Click(object sender, EventArgs e)
        {
            this.Close(); //to close the window(Form1)
        }
    }
}
