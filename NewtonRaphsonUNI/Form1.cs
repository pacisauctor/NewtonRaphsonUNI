using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Calculus;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace NewtonRaphsonUNI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Calculo analizadorFunciones = new Calculo();
            if (analizadorFunciones.Sintaxis(tbFormula.Text, 'x'))
            {
                MessageBox.Show("Función bien digitada!!!", "TODO OK",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Information);

            }
            else
            {
                MessageBox.Show("Función mal digitada!!!", "ERROR 505",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Error);
            }
        }

        private void bSuma_Click(object sender, EventArgs e)
        {
            AgregarFormula("+");
        }

        private void bResta_Click(object sender, EventArgs e)
        {
            AgregarFormula("-");
        }

        private void bProducto_Click(object sender, EventArgs e)
        {
            AgregarFormula("*");

        }

        private void bDivision_Click(object sender, EventArgs e)
        {
            AgregarFormula("/");

        }

        private void bBorrar_Click(object sender, EventArgs e)
        {
            tbFormula.Text = tbFormula.Text.Remove(tbFormula.Text.Length - 1);
        }

        private void bCalcular_Click(object sender, EventArgs e)
        {
            tbFormula.Text = "";
        }



        private void AgregarFormula(string txt)
        {
            tbFormula.Text += txt;
        }

        private void b7_Click(object sender, EventArgs e)
        {
            AgregarFormula("7");
        }

        private void b8_Click(object sender, EventArgs e)
        {
            AgregarFormula("8");
        }

        private void b9_Click(object sender, EventArgs e)
        {
            AgregarFormula("9");
        }

        private void b4_Click(object sender, EventArgs e)
        {
            AgregarFormula("4");
        }

        private void b5_Click(object sender, EventArgs e)
        {
            AgregarFormula("5");
        }

        private void b6_Click(object sender, EventArgs e)
        {
            AgregarFormula("6");
        }

        private void b1_Click(object sender, EventArgs e)
        {
            AgregarFormula("1");
        }

        private void b2_Click(object sender, EventArgs e)
        {
            AgregarFormula("2");
        }

        private void b3_Click(object sender, EventArgs e)
        {
            AgregarFormula("3");
        }

        private void gbOpBas_Enter(object sender, EventArgs e)
        {

        }

        private void b0_Click(object sender, EventArgs e)
        {
            AgregarFormula("0");
        }

        private void bpunto_Click(object sender, EventArgs e)
        {
            AgregarFormula(".");
        }

        private void b00_Click(object sender, EventArgs e)
        {
            AgregarFormula("00");
        }

        private void bSeno_Click(object sender, EventArgs e)
        {
            AgregarFormula("sen(");
        }

        private void bCoseno_Click(object sender, EventArgs e)
        {
            AgregarFormula("cos(");
        }

        private void bTangente_Click(object sender, EventArgs e)
        {
            AgregarFormula("tan(");
        }

        private void bSenInv_Click(object sender, EventArgs e)
        {
            AgregarFormula("asen(");
        }

        private void bCosInv_Click(object sender, EventArgs e)
        {
            AgregarFormula("acos(");
        }

        private void bTanInv_Click(object sender, EventArgs e)
        {
            AgregarFormula("atan(");
        }

        private void bPotencia_Click(object sender, EventArgs e)
        {
            AgregarFormula("^");
        }

        private void bLn_Click(object sender, EventArgs e)
        {
            AgregarFormula("ln(");
        }

        private void bLog_Click(object sender, EventArgs e)
        {
            AgregarFormula("log(");
        }

        private void bAbroParentesis_Click(object sender, EventArgs e)
        {
            AgregarFormula("(");
        }

        private void bCierroParentesis_Click(object sender, EventArgs e)
        {
            AgregarFormula(")");
        }

        private void bX_Click(object sender, EventArgs e)
        {
            AgregarFormula("x");
        }
        private void Calcular(string text)
        {
            Calculo analizadorFunciones = new Calculo();
            if(analizadorFunciones.Sintaxis(text, 'x'))
            {
                try
                {
                    double valorInicial = double.Parse(tbVI.Text);
                    metodoNewtonRaphson(text, valorInicial);
                }
                catch (FormatException)
                {
                    MessageBox.Show("Valor Inicial mal digitado!!!", "ERROR 504",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("Función mal digitada!!!", "ERROR 505",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Error);
            }
        }

        private void metodoNewtonRaphson(string text, double valorInicial)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "3M2Sistemas";
                excelPackage.Workbook.Properties.Title = "NewtonRaphson";
                excelPackage.Workbook.Properties.Subject = "Exportado desde Visual";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Tabla");
                worksheet.Cells["B2"].Value = "f(x): " + text;
                worksheet.Cells["D2"].Value = "Valor inicial: ";
                worksheet.Cells["E2"].Value = valorInicial;
                worksheet.Cells["F2"].Value = "Error: ";
                worksheet.Cells["G2"].Value = 0.001;


                worksheet.Cells["B3"].Value = "IT";
                worksheet.Cells["C3"].Value = "n";
                worksheet.Cells["D3"].Value = "f(x)";
                worksheet.Cells["E3"].Value = "f'(x)";
                worksheet.Cells["F3"].Value = "Xn+1";
                worksheet.Cells["G3"].Value = "ERROR";

                worksheet.Cells["B3:G3"].Style.Font.Bold = true;
                Calculo analizadorFunciones = new Calculo();
                analizadorFunciones.Sintaxis(text, 'x');
                double error = 0.001, errorCalculado = 0, funcion, funcionDerivada, Xnmas1, valor= valorInicial;
                int contador = 0, iteracion, n;
                
                do
                {
                    iteracion = contador + 1;
                    n = contador;
                    funcion = Math.Round(analizadorFunciones.EvaluaFx(valor),4);
                    funcionDerivada = Math.Round(analizadorFunciones.Dx(valor),4);
                    Xnmas1 = Math.Round(valor - (funcion / funcionDerivada),4);
                    errorCalculado = Math.Abs(valor - Xnmas1);

                    //You could also use [line, column] notation:
                    worksheet.Cells[contador + 4, 2].Value = iteracion;
                    worksheet.Cells[contador + 4, 3].Value = n;
                    worksheet.Cells[contador + 4, 4].Value = funcion;
                    worksheet.Cells[contador + 4, 5].Value = funcionDerivada;
                    string formula = (contador == 0) ? "E2-(D4/E4)" :
                        "F" + (contador + 3) + "-(D" + (contador + 4) + "/E" + (contador + 4)+ ")";
                    worksheet.Cells[contador + 4, 6].Formula = formula;

                    formula = (contador == 0) ? "abs(F4-E2)" : "abs(F"+(contador+4)+"-F"+(contador+3)+")";
                    worksheet.Cells[contador + 4, 7].Formula = formula;

                    valor = Xnmas1;
                    contador++;
                } while (errorCalculado > error);
                worksheet.Cells[contador + 3, 6].Style.Font.Bold= true;

                ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets.Add("Gráfico");
                worksheet2.Cells["B1"].Value = "f(x): " + text;
                worksheet2.Cells["A1"].Value = "x";
                double x = -5.0; int cont = 2;
                while (x<5.0)
                {
                    worksheet2.Cells[cont, 1].Value = x;
                    worksheet2.Cells[cont, 2].Value = analizadorFunciones.EvaluaFx(x);
                    x = Math.Round(x + 0.1,2);
                    cont++;
                }
                var chart = worksheet2.Drawings.AddChart(text, eChartType.Line);
                chart.SetPosition(0, 600);
                chart.SetSize(600, 600);
                var serie = chart.Series.Add(worksheet2.Cells["B2:B" + (cont-1)], worksheet2.Cells["A2:A" + (cont-1)]);
                serie.Header = text;
                worksheet2.Calculate();
                //Save your file
                FileInfo fi = new FileInfo(@"newton.xlsx");
                
                excelPackage.SaveAs(fi);
                MessageBox.Show("Excel generado exitosamente!!! Guardado en: " + fi.Directory, "Wi",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Exclamation);
                tbFormula.Text = "";
                tbVI.Text = "";
            }
        }

        private void bArchivoExcel_Click(object sender, EventArgs e)
        {
            Calcular(tbFormula.Text);
        }
    }
}
