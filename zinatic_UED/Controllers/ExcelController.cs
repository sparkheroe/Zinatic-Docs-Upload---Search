using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using zinatic_UED.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace zinatic_UED.Controllers
{
    public class ExcelController : Controller
    {
        
        Excel.Workbook workbook;
        Excel.Worksheet worksheet;
        private int ultimaFila = 1;

        // GET: Excel
        public ActionResult Index()
        {
            CreateExcel();
            return View();
        }


        //Datos Prueba
        private IEnumerable<SeccionModels> CargaDatosCabecera() {

            List<SeccionModels> listaDatos = new List<SeccionModels>();

            SeccionModels seccion = new SeccionModels();
            seccion.Titulo = "APELLIDOS Y NOMBRES:";
            seccion.ValorDato = "KEYNER JARA SORIA";
            seccion.NumeroOrden = 1;
            listaDatos.Add(seccion);

            SeccionModels seccion2 = new SeccionModels();
            seccion2.Titulo = "AREA:";
            seccion2.ValorDato = "SISTEMAS";
            seccion2.NumeroOrden = 2;
            listaDatos.Add(seccion2); 

            SeccionModels seccion3 = new SeccionModels();
            seccion3.Titulo = "CARGO:";
            seccion3.ValorDato = "GERENTE DE PROYECTOS";
            seccion3.NumeroOrden = 3;
            listaDatos.Add(seccion3);

            SeccionModels seccion4 = new SeccionModels();
            seccion4.Titulo = "FECHA DE EVALUACIÓN:";
            seccion4.ValorDato = "2017-04-04";
            seccion4.NumeroOrden = 4;
            listaDatos.Add(seccion4);

            SeccionModels seccion5 = new SeccionModels();
            seccion5.Titulo = "FECHA DE INGRESO:";
            seccion5.ValorDato = "2016-01-01";
            seccion5.NumeroOrden = 5;
            listaDatos.Add(seccion5);

            SeccionModels seccion6 = new SeccionModels();
            seccion6.Titulo = "DNI:";
            seccion6.ValorDato = "47647606";
            seccion6.NumeroOrden = 6;
            listaDatos.Add(seccion6);

            SeccionModels seccion7 = new SeccionModels();
            seccion7.Titulo = "DNI:";
            seccion7.ValorDato = "47647606";
            seccion7.NumeroOrden = 7;
            listaDatos.Add(seccion7);

            SeccionModels seccion8 = new SeccionModels();
            seccion8.Titulo = "DNI:";
            seccion8.ValorDato = "47647606";
            seccion8.NumeroOrden = 8;
            listaDatos.Add(seccion8);

            return listaDatos;
        }

        private void SetCellFirstWordBold(Microsoft.Office.Interop.Excel.Range rng, char wordsSeparator)
        {
            string cellString = rng.Text;

            int firstWordEndIdx = cellString.IndexOf(wordsSeparator);
            this.SetCellBoldPartial(rng, 0, firstWordEndIdx);
        }

        private void SetCellBoldPartial(Microsoft.Office.Interop.Excel.Range rng, int boldStartIndex, int boldEndIndex)
        {
            rng.Characters[boldStartIndex, boldEndIndex].Font.Bold = 1;
        }


        public Excel.Worksheet DefineTipoSeccion(Excel.Worksheet worksheet, int tipoSeccion, IEnumerable<SeccionModels> datosSeccion, string Titulo, int CantidadColumnas, int UltimaColumna)
        {
            int cantidadFilas = datosSeccion.Count() / CantidadColumnas;
            //Excel.Range contenidoSeccion;
            //Define que tipo de Seccion es el contenido de Excel
            switch (tipoSeccion)
            {
                case 1:
                    #region Seccion Tipo Cabecera
                    //Seccion de Cabecera
                    /*if (UltimaColumna < 4) {
                        UltimaColumna = 4;
                    }*/

                    //La cantidad de columnas para este tipo de seccion se multiplica por dos por los pares de nombre->valor que se mostraran
                    //CantidadColumnas = CantidadColumnas * 2;

                    //Coloca Titulo de Sección
                    worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, UltimaColumna]].Merge();
                    worksheet.Range["A1"].Interior.Color = Color.LightBlue;
                    worksheet.Cells[1, 1] = Titulo;
                    SetCellFirstWordBold(worksheet.Cells[1, 1], ':');
                    worksheet.Cells.Font.Size = 8;
                    ultimaFila++;

                    
                    //Crea Columnas de Contenido
                    List<int> NumerosColumnaParaSeparar = new List<int>();
                    //Agrega la primera base a tomar en cuenta para la creación de los bordes de las columnas
                    NumerosColumnaParaSeparar.Add(1);

                    int ColumnaSeparadoraInicial = 1;
                    int ColummaSeparadoraFinal = 0;
                    //Obtiene cada cuantas columnas se hara un borde en la fila
                    int CantidadSumarColumnas = UltimaColumna / CantidadColumnas;
                    //Incializamos la Columna Separadora Final
                    ColummaSeparadoraFinal = CantidadSumarColumnas;
                    //falta verificar si es para o no para la cantidad de columnas actuales y si se puede dividir entre cantidad de columnas requeridas
                    List<SeccionModels> listaEjecucion = new List<SeccionModels>();
                    int contadorDatos = 1;
                    //Agrega por tipo de dato 
                    for (int i = 0; i < cantidadFilas; i++)
                    {                        
                        int contadorDatosPorColumnas = 0;
                        for (int j = 0; j < CantidadColumnas; j++)
                        {
                            if (contadorDatosPorColumnas < CantidadColumnas)
                            {
                                //rangocolumnasvisibles.BorderAround2(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
                                //Coloca el borde sobre las columnas visibles
                                //SeccionModels sec = new SeccionModels();
                                foreach (var obj in datosSeccion)
                                {

                                    SeccionModels seccion = obj;
                                    if (seccion.NumeroOrden == contadorDatos && contadorDatosPorColumnas < CantidadColumnas)
                                    {
                                        //define el rango para crear las columnas visibles
                                        Excel.Range rangocolumnasvisibles = worksheet.Range[worksheet.Cells[ultimaFila, ColumnaSeparadoraInicial], worksheet.Cells[ultimaFila, ColummaSeparadoraFinal]];
                                        worksheet.Cells[ultimaFila, ColummaSeparadoraFinal] = seccion.Titulo + " " + seccion.ValorDato;
                                        Excel.Borders border = rangocolumnasvisibles.Borders;
                                        border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        SetCellFirstWordBold(worksheet.Cells[ultimaFila, ColummaSeparadoraFinal], ':');

                                        //incrementa valores para la proxima columna visible
                                        ColumnaSeparadoraInicial = ColummaSeparadoraFinal;
                                        ColummaSeparadoraFinal = ColummaSeparadoraFinal + CantidadSumarColumnas;
                                        contadorDatos++;
                                        contadorDatosPorColumnas++;
                                    }
                                    if (contadorDatosPorColumnas >= CantidadColumnas)
                                        break;
                                }
                            }
                            
                            else {
                                //contadorDatos = 1;
                                contadorDatosPorColumnas = 0;
                                break;
                            }
                            
                        }
                        ColumnaSeparadoraInicial = 1;
                        ColummaSeparadoraFinal = CantidadSumarColumnas;
                        ultimaFila++;
                    }
                    Excel.Range rangoAjustarContenido = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[ultimaFila, CantidadColumnas]];
                    rangoAjustarContenido.Columns.AutoFit();
                    #endregion
                    break;
                case 2:
                    break;

            }
            return worksheet;
        }

        public void CreateExcel()
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Add(Type.Missing);

            worksheet = (Excel.Worksheet)workbook.ActiveSheet;


            #region Revisión de Desempeño
            worksheet.Name = "Revisión De Desempeño";
            #endregion

            //Define sección Cabecera
            string TituloCabecera = "1.- DATOS DEL TRABAJADOR";
            worksheet = DefineTipoSeccion(worksheet, 1, CargaDatosCabecera(), TituloCabecera, 2, 2);

            //Nombre aleatorio basado en Timestamp Actual
            String timeStamp = GetTimestamp(DateTime.Now);

            workbook.SaveAs("C:\\Users\\DitaMonster\\Documents\\Visual Studio 2017\\Projects\\Zinatic_UED\\Zinatic-Docs-Upload---Search\\zinatic_UED\\files\\" + timeStamp + ".xlsx"); ;
            workbook.Close();
            excel.Quit();
        }
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }
    }
}