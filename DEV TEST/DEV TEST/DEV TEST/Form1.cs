using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DEV_TEST
{
    public partial class Form1 : Form
    {

        List<string> listarutas = new List<string>();
        string ruta = @"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo";
        string rutaExcel = @"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\Procesado\";
        string rutaOtros= @"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\No Aplicable\";


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            GetFiles();
        }

        private void GetFiles()
            {


            lstArchivos.Items.Clear();
            listarutas.Clear();
            string[] ext = new string[1] { "*.*" };

            foreach (string found in ext)
            {

                string[] extracted = Directory.GetFiles(ruta, found, SearchOption.AllDirectories);
                int cont=0;
                foreach (string file in extracted)
                {

                    if (Path.GetFileName(file).Contains(".xlsx"))
                    {
                        cont++;
                        clsLogica ex = new clsLogica(ruta + "\\" + Path.GetFileName(file), 1);
                        string[,] read = ex.ReadRange(1, 1, 100, 5);
                        ex.close();

                        clsLogica ex1 = new clsLogica(@"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\LibroMaestro\LibroMaestro", cont);
                        ex1.SelectWorksheet(cont);
                        ex1.WriteRange(1, 1, 100, 5, read);
                        ex1.createNewSheet();
                        ex1.save();
                        ex1.close();
                        File.Move(file, rutaExcel + Path.GetFileName(file));
                    
                    }
                    else
                    {
                        File.Move(file, rutaOtros + Path.GetFileName(file));

                    }
                        


                    lstArchivos.Items.Add(Path.GetFileName(file));
                    listarutas.Add(file);



                }
  
        }

    }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ruta = dialog.SelectedPath;
                    GetFiles();
                }
            }
        }
    }
}
