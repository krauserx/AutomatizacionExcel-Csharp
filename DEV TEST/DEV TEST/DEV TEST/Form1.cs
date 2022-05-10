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
        string ruta = "";
        string rutaExcel = @"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\Procesado\";
        string rutaOtros= @"C:\Users\kraus\Desktop\Pruebas Genpact\DEV TEST\CarpetaMonitoreo\No Aplicable\";


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
           
        }

        private void GetFiles()
            {
            int cont = 1;

            lstArchivos.Items.Clear();
            listarutas.Clear();
            string[] ext = new string[1] { "*.*" };

            foreach (string found in ext)
            {

                string[] extracted = Directory.GetFiles(ruta, found, SearchOption.AllDirectories);
                foreach (string file in extracted)
                {

                    if (Path.GetFileName(file).Contains(".xlsx"))
                    {
                      

                        clsLogica ex2 = new clsLogica(ruta + "\\" + Path.GetFileName(file), 1,cont);
                        ex2.CountRows();
                           File.Move(file, rutaExcel + Path.GetFileName(file));
                        cont++;

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
