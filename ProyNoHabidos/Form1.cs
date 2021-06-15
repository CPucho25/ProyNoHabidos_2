using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using ExcelDataReader;
using SpreadsheetLight;
using System.Diagnostics;

namespace ProyNoHabidos
{
    public partial class Form1 : Form
    {
        private static IWebDriver driver;

        //Ruta de Descarga
        //string pathDL = Environment.CurrentDirectory + @"\1-NoHabidos-Downloads\";
        static string path1_DL = Environment.CurrentDirectory + @"\1-NoHabidos-Downloads\";
        static string path2_Unzip = Environment.CurrentDirectory + @"\2-NoHabidos-Unzip\";
        static string path3_Xlsx = Environment.CurrentDirectory + @"\3-NoHabidos-Xlsx\";
        static string path4_Final = Environment.CurrentDirectory + @"\4-NoHabidos-Final\";

        public Form1()
        {
            InitializeComponent();

            createPaths();

            //option por default para descarga multiple
            var options = new ChromeOptions();
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
            options.AddUserProfilePreference("download.default_directory",path1_DL);

            driver = new ChromeDriver(options);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            cboMes.Items.Add("enero");
            cboMes.Items.Add("febrero");
            cboMes.Items.Add("marzo");
            cboMes.Items.Add("abril");
            cboMes.Items.Add("mayo");
            cboMes.Items.Add("junio");
            cboMes.Items.Add("julio");
            cboMes.Items.Add("agosto");
            cboMes.Items.Add("setiembre");
            cboMes.Items.Add("octubre");
            cboMes.Items.Add("noviembre");
            cboMes.Items.Add("diciembre");

            string nowMes = DateTime.Now.ToString("MMMM");

            cboMes.SelectedItem = nowMes.ToLower();

            int nowYear = DateTime.Now.Year;

            for (int i = nowYear; i > nowYear-10; --i)
            {
                cboAnio.Items.Add(i);
            }

            cboAnio.SelectedItem = nowYear;
        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToString();

            string mes = cboMes.SelectedItem.ToString();
            string mesNum = "";
            string anio = cboAnio.SelectedItem.ToString();

            //Mes en texto segun lo seleccionado en el cbox
            mesNum = mes == "enero" ? "01" : mes == "febrero" ? "02" : mes == "marzo" ? "03" : mes == "abril" ? "04" : mes == "mayo" ? "05" : mes == "junio" ? "06" : mes == "julio" ? "07" : mes == "agosto" ? "08" : mes == "setiembre" ? "09" : mes == "octubre" ? "10" : mes == "noviembre" ? "11" : mes == "diciembre" ? "12" : "0";


            // 	          h ttps://www.sunat.gob.pe/orientacion/Nohallados/descargas/NoHabidos/  2021     /   abril   /nohabido  2021      04     .html
            string url = "https://www.sunat.gob.pe/orientacion/Nohallados/descargas/NoHabidos/" + anio + "/" + mes + "/nohabido" + anio + mesNum +".html";

            driver.Url = url;

            #region Descarga Zip IPCN, LIMA, Otras dependencias

            Thread.Sleep(2000);

            IWebElement iIPCN = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[10]/td/font/a"));
            iIPCN.Click();

            Thread.Sleep(2000);

            IWebElement iLima = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[12]/td/font/a"));
            iLima.Click();

            Thread.Sleep(2000);

            IWebElement iODepend = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[14]/td/font/b/a"));
            iODepend.Click();

            //Captura la fecha dd/MM/yyyy
            string strFecha ;
            IWebElement iFecha = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[3]/td[1]/center/font/b"));
            strFecha =  iFecha.Text.Substring(22,10);

            #endregion

            Thread.Sleep(13000);

            #region Descomprimir Zip

            //Ruta donde se descomprimira los archivos ZIP
            string pathUnzip = Environment.CurrentDirectory + @"\2-NoHabidos-Unzip\";

            if (!Directory.Exists(pathUnzip))
            { Directory.CreateDirectory(pathUnzip); }

            DirectoryInfo di = new DirectoryInfo(path1_DL);

            foreach (var fi in di.GetFiles())
            {
                ZipFile.ExtractToDirectory(path1_DL + fi.Name, pathUnzip);

                Thread.Sleep(1000);
            }

            #endregion

            Thread.Sleep(7500);

            driver.Close();

            #region Conversion Xls a Xlsx (Excel)

            // PATH USADO EN EL SCRIPT PARA LA CONVERSION DE EXCEL: XLS A XLSX

            //strPath = Application.ActiveWorkbook.Path & "\2-NoHabidos-Unzip\"
            //strPathSalida = Application.ActiveWorkbook.Path & "\3-NoHabidos-Xlsx\"

            convertExcel();

            string[] dir2 = Directory.GetFiles(pathUnzip);

            int cantFiles = dir2.Length;

            #endregion

            int seg = 15000 * cantFiles;

            Thread.Sleep(seg);

            // Copiar archivo txt de capeta "2-NoHabidos-Unzip"a carpeta "3-NoHabidos-Xlsx"
            DirectoryInfo diSearchTxt = new DirectoryInfo(path2_Unzip);
            foreach (var miFi in diSearchTxt.GetFiles())
            {
                string miFile = miFi.Name;
                int lengthFile = miFile.Length;
                if (lengthFile > 25) {
                    miFi.CopyTo(path3_Xlsx + miFile);
                }
            }

            #region Agregando la 4ta Columna con fecha de Publicacion

            //string pathAddC = Environment.CurrentDirectory + @"\3-NoHabidos-Xlsx\";

            int tip_contrib = 0;
            string strgDAte = strFecha;//capturado con Selenium de la web de Sunat(strFecha)

            DateTime dtimen = DateTime.ParseExact(strgDAte, "dd/MM/yyyy", null);

            DirectoryInfo din = new DirectoryInfo(path3_Xlsx);

            string strTipoFile = "";

            foreach (var fin in din.GetFiles())
            {
                strTipoFile = fin.Name.Substring(13, 4);
                int indexDif = fin.Name.Length;

                if (indexDif < 25)
                {
                    if (strTipoFile.Equals("ipcn"))
                    {
                        tip_contrib = 1;
                    }
                    if (strTipoFile.Equals("lima"))
                    {
                        tip_contrib = 2;
                    }

                    addColumn(path3_Xlsx + fin.Name, tip_contrib, dtimen);
                } 
            }

            #endregion

            #region Convertir Excel a un solo archivo txt

            string fechaPublic = anio + mesNum;

            convertToTXT(fechaPublic);

            #endregion

            Thread.Sleep(8000);

            convertToTXT2(fechaPublic);

            Thread.Sleep(2000);

            vaciarCarpetas1_2_3();

            label2.Text = DateTime.Now.ToString();

        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            //driver.Close();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Estas seguro?", "Eliminar Archivos", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                vaciarCarpetas1_2_3();

                DirectoryInfo di4 = new DirectoryInfo(path4_Final);

                foreach (var fi4 in di4.GetFiles())
                {
                    File.Delete(path4_Final + fi4.Name);
                }
                MessageBox.Show("Listo!");
            }
            else if (dialogResult == DialogResult.No)
            {
                //
            }
        }

        // V O I D     FINAL CONVERTER XLS A XLSX(MACRO)
        public void convertExcel() {

            string path111 = Environment.CurrentDirectory + @"\Book12.xlsm";

            System.Diagnostics.Process.Start(path111);

        }

        // V O I D     AGREGAR COLUMNAS A EXCEL(SpreadsheetLight Libreria)
        public void addColumn(string pathAddC, int tipo_contr, DateTime datetime)
        {
            if (tipo_contr == 1 || tipo_contr == 2)
            {
                #region Agrega Columna a Excel
                //USANDO EXCEL DATA READER ------------------------------------------------------
                var stream = File.Open(pathAddC, FileMode.Open, FileAccess.Read);
                var reader = ExcelReaderFactory.CreateReader(stream);
                var result = reader.AsDataSet();
                var tables = result.Tables.Cast<DataTable>();
                DataTable dt = tables.First();

                int totalFilas = dt.Rows.Count - 7;
                stream.Close();

                //------------------------------------------------------------------------------
                SLDocument sl = new SLDocument(pathAddC);

                SLStyle style = sl.CreateStyle();
                style.FormatCode = "dd/MM/yyyy";

                sl.SetColumnStyle(6, style);

                sl.SetCellValue("B7", "NUM_DOC_IDEN");
                sl.SetCellValue("C7", "NOMBRE_DEL_CONTRIBUYENTE");
                sl.SetCellValue("D7", "FEC_ADQ_COND_NO_HAB");
                sl.SetCellValue("E7", "TIPO_CONTR");
                sl.SetCellValue("F7", "FEC_PUB");

                for (int i = 8; i <= totalFilas + 7; ++i)
                {
                    sl.SetCellValue("E" + i, tipo_contr);

                    sl.SetCellValue("F" + i, datetime);
                }

                sl.Save();

                #endregion
            }
        }

        // V O I D     CONVERTER EXCEL y TXT a  TXT_FINAL
        public void convertToTXT(string fechaPublicacion) {

            #region Convertir xlsx a TXT (Usando libreria ExcelDataReader)

            DirectoryInfo di2 = new DirectoryInfo(path3_Xlsx);

            DataTable dt2 = new DataTable();
            int contDT = 0;

            #region Formateo de Excel a DataTable(uniendo varios DT a uno solo)
            foreach (var fi2 in di2.GetFiles())
            {
                if (fi2.Name.Length < 25) {

                    var stream = File.Open(path3_Xlsx + fi2.Name, FileMode.Open, FileAccess.Read);
                    //var stream = File.Open(path12, FileMode.Open, FileAccess.Read);
                    var reader = ExcelReaderFactory.CreateReader(stream);
                    //var reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    var result = reader.AsDataSet();
                    var tables = result.Tables.Cast<DataTable>();

                    DataTable dt = tables.First();

                    stream.Close();

                    #region Eliminando Filas y columna (extraidas de excel)

                    #region Elimina 7 primeras Filas
                    List<DataRow> rowsToDelete = new List<DataRow>();

                    int contRow = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        if (contRow < 7)
                        {
                            rowsToDelete.Add(row);
                        }
                        contRow++;
                    }

                    foreach (DataRow row in rowsToDelete)
                    {
                        dt.Rows.Remove(row);
                    }
                    #endregion

                    //Elimina primera Columna vacia
                    dt.Columns.RemoveAt(0);

                    dt.AcceptChanges();

                    Thread.Sleep(1000);

                    #endregion

                    #region Uniendo Tablas en uno solo
                        dt2.Merge(dt);
                    #endregion

                    Thread.Sleep(500);

                    contDT++;
                    
                    dt.Clear();

                    Thread.Sleep(500);
                }
            }
            #endregion

            //Pasar DataTable final para escribir al archivo TXT
            #region Grabando el DataTable(de excel) al txt

            /*string newContenido = "";*/

            string strRuc = "";
            string strNomCont = "";
            string strFechAd = "";
            string strTipo = "";
            string strFechPub = "";

            int totalFilas = dt2.Rows.Count;

            //label3.Text = totalFilas.ToString();
            /*int newCorredor = 1;*/

            var stopwatch = new Stopwatch(); // 2) Nueva Linea

            stopwatch.Start(); // 2) Nueva Linea

            using (StreamWriter streamwriter = File.AppendText(path4_Final + fechaPublicacion + "_FINAL.txt"))
            {
                streamwriter.WriteLine("NUM_DOC_IDEN|NOMBRE_DEL_CONTRIBUYENTE|FEC_ADQ_COND_NO_HAB|TIPO_CONTR|FEC_PUB");

                #region EXCEL A TXT FINAL
                foreach (DataRow row in dt2.Rows)
                {
                    //strRuc = row[0].ToString();
                    //strNomCont = row[1].ToString();
                    strFechAd = row[2].ToString();
                    //strTipo = row[3].ToString();
                    strFechPub = row[4].ToString().Substring(0, 10);

                    string value = strFechPub.Substring(9, 1);

                    if (value == " ")
                    {
                        strFechPub = "0" + row[4].ToString().Substring(0, 9);
                    }

                    DateTime dtFechaAdq = DateTime.ParseExact(strFechAd, "dd/MM/yyyy", null);
                    strFechAd = dtFechaAdq.ToString("yyyy-MM-dd");

                    DateTime dtFechaPub = DateTime.ParseExact(strFechPub, "dd/MM/yyyy", null);
                    strFechPub = dtFechaPub.ToString("yyyy-MM-dd");

                    streamwriter.WriteLine(row[0].ToString() + "|" + row[1].ToString() + "|" + strFechAd + "|" + row[3].ToString() + "|" + strFechPub);// 2) Nueva Linea

                }
                #endregion

                Thread.Sleep(1000);
            }
            stopwatch.Stop();

            #endregion

            Thread.Sleep(2000);

            #endregion
        }

        public void convertToTXT2(string fechaPublicacion) {

            var stopwatch2 = new Stopwatch(); // 2) Nueva Linea

            stopwatch2.Start();

            using (StreamWriter streamwriter1 = File.AppendText(path4_Final + fechaPublicacion + "_FINAL.txt"))
            {

                DirectoryInfo di3 = new DirectoryInfo(path3_Xlsx);

                #region SUMANDO EL TXT(otras depedencias) AL TXT FINAL!!!

                foreach (var fi3 in di3.GetFiles())
                {
                    int tipofile3 = fi3.Name.Length;

                    if (tipofile3 > 25)
                    {
                        string[] linesTxt = File.ReadAllLines(path3_Xlsx + fi3.Name, Encoding.Default);//Leera todos los registros del txt
                        int totLine = linesTxt.Length;
                      
                        string strgDateAdq;
                        DateTime DateAdq2 = DateTime.ParseExact("01/01/1900", "dd/MM/yyyy", null);

                        // Captura FECHA MAYOR de Adquicion para usar como fecha de publicación
                        foreach (var lineaFecha in linesTxt)
                        {
                            int indexFechaAdq = lineaFecha.Length - 11;
                            strgDateAdq = lineaFecha.Substring(indexFechaAdq, 10);

                            DateTime fechaAdq1 = DateTime.ParseExact(strgDateAdq, "dd/MM/yyyy", null);

                            int fechaResultado = DateTime.Compare(fechaAdq1, DateAdq2);

                            if (fechaResultado < 0) { DateAdq2 = DateAdq2; }
                            else if (fechaResultado == 0) { DateAdq2 = DateAdq2; }
                            else { DateAdq2 = fechaAdq1; }
                        }

                        //Formatea las fechas y agrega las columnas "Tipo comprobante" y "Fecha publicación"
                        foreach (var lineTxt in linesTxt)
                        {
                            int index2 = lineTxt.Length - 11;
                            string newLine = lineTxt.Substring(0, index2);
                            string fechaAdq2 = lineTxt.Substring(index2, 10);
                            string tipo_contr = "3";
                            DateTime fechaAdq22 = DateTime.ParseExact(fechaAdq2, "dd/MM/yyyy", null);

                            fechaAdq2 = fechaAdq22.ToString("yyyy-MM-dd");
                            string fechaPublic = DateAdq2.ToString("yyyy-MM-dd");

                            streamwriter1.WriteLine(newLine + fechaAdq2 + "|" + tipo_contr + "|" + fechaPublic);

                        }
                    }
                }
                #endregion

            }
            stopwatch2.Stop();
        }

        public void createPaths() {
           
            string path1 = Environment.CurrentDirectory + @"\1-NoHabidos-Downloads\"; 
            string path2 = Environment.CurrentDirectory + @"\2-NoHabidos-Unzip\";
            string path3 = Environment.CurrentDirectory + @"\3-NoHabidos-Xlsx\";
            string path4 = Environment.CurrentDirectory + @"\4-NoHabidos-Final\";

            if (!Directory.Exists(path1))
            {
                Directory.CreateDirectory(path1);
            } 
            if (!Directory.Exists(path2))
            {
                Directory.CreateDirectory(path2);
            }
            if (!Directory.Exists(path3))
            {
                Directory.CreateDirectory(path3);
            }
            if (!Directory.Exists(path4))
            {
                Directory.CreateDirectory(path4);
            }
        }

        public void vaciarCarpetas1_2_3()
        {
            DirectoryInfo di1 = new DirectoryInfo(path1_DL);

            foreach (var fi1 in di1.GetFiles())
            {
                File.Delete(path1_DL + fi1.Name);
            }

            DirectoryInfo di2 = new DirectoryInfo(path2_Unzip);

            foreach (var fi2 in di2.GetFiles())
            {
                File.Delete(path2_Unzip + fi2.Name);
            }

            DirectoryInfo di3 = new DirectoryInfo(path3_Xlsx);

            foreach (var fi3 in di3.GetFiles())
            {
                File.Delete(path3_Xlsx + fi3.Name);
            }
        }

        
    }
}
