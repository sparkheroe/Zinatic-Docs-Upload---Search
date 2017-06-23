using Microsoft.Office.Interop.Excel;
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
        private int UltimaColumna = 1;

        int ColumnaSeparadoraInicial = 1;
        int ColummaSeparadoraFinal = 0;

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

            return listaDatos;
        }

        private IEnumerable<SeccionModels> DatosInstrucciones() {
            List<SeccionModels> datosInstrucciones = new List<SeccionModels>();

            SeccionModels Puntaje1 = new SeccionModels();
            Puntaje1.NumeroOrden = 1;
            Puntaje1.Titulo = "DEBAJO DE LO ESPERADO";
            Puntaje1.Descripcion = "No cumple con las expectativas y/o perfil de puesto";
            datosInstrucciones.Add(Puntaje1);

            SeccionModels Puntaje2 = new SeccionModels();
            Puntaje2.NumeroOrden = 2;
            Puntaje2.Titulo = "REQUIERE MEJORAR";
            Puntaje2.Descripcion = "Cumple parcialmente las expectativas y/o requerimientos del puesto. Presenta tendencia a mejora.";
            datosInstrucciones.Add(Puntaje2);

            SeccionModels Puntaje3 = new SeccionModels();
            Puntaje3.NumeroOrden = 3;
            Puntaje3.Titulo = "CUMPLE LAS EXPECTATIVAS";
            Puntaje3.Descripcion = "Cumple con lo esperado y/orequerimiento del puesto.";
            datosInstrucciones.Add(Puntaje3);

            SeccionModels Puntaje4 = new SeccionModels();
            Puntaje4.NumeroOrden = 4;
            Puntaje4.Titulo = "SUPERA LAS EXPECTATIVAS";
            Puntaje4.Descripcion = "Supera la expectativas y/o requerimientos del puesto. Presenta un desempeño sobresaliente.";
            datosInstrucciones.Add(Puntaje4);

            return datosInstrucciones;
        }

        //Metodo para poder separar un string por un tipo de caracter unico
        private void SetCellFirstWordBold(Microsoft.Office.Interop.Excel.Range rng, char wordsSeparator)
        {
            string cellString = rng.Text;

            int firstWordEndIdx = cellString.IndexOf(wordsSeparator);
            this.SetCellBoldPartial(rng, 0, firstWordEndIdx);
        }

        //Metodo para poner una seccion de una celda con letras oscura (Bold)
        private void SetCellBoldPartial(Microsoft.Office.Interop.Excel.Range rng, int boldStartIndex, int boldEndIndex)
        {
            rng.Characters[boldStartIndex, boldEndIndex].Font.Bold = 1;
        }


        public Excel.Worksheet DefineTipoSeccion(Excel.Worksheet worksheet, int tipoSeccion, IEnumerable<SeccionModels> datosSeccion, string Titulo, string Descripcion, int CantidadColumnas)
        {
            int cantidadFilas = 0;
            int ColumnaSeparadoraInicial = 1;
            int ColummaSeparadoraFinal = 0;
            //Excel.Range contenidoSeccion;
            //Define que tipo de Seccion es el contenido de Excel
            switch (tipoSeccion)
            {
                
                case 1:

                    #region Seccion Tipo Cabecera
                    //Seccion de Cabecera

                    if (datosSeccion != null)
                    {
                        cantidadFilas = datosSeccion.Count() / CantidadColumnas;
                    }

                    if (UltimaColumna < 4) {
                        UltimaColumna = 4;
                        
                    }

                    //La cantidad de columnas para este tipo de seccion se multiplica por dos por los pares de nombre->valor que se mostraran
                    //CantidadColumnas = CantidadColumnas * 2;

                    //Coloca Titulo de Sección
                    worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, UltimaColumna]].Merge();
                    worksheet.Range["A1"].Interior.Color = Color.LightBlue;
                    worksheet.Cells[1, 1] = Titulo;
                    SetCellFirstWordBold(worksheet.Cells[1, 1], ':');
                    worksheet.Cells.Font.Size = 8;
                    ultimaFila++;                    
                    
                    //Obtiene cada cuantas columnas se hara un borde en la fila
                    int CantidadSumarColumnas = UltimaColumna / CantidadColumnas;
                    //Incializamos la Columna Separadora Final
                    AgregarDatosFilasColumnas(worksheet, CantidadColumnas, cantidadFilas, datosSeccion, 1);
                    /*Excel.Range rangoAjustarContenido = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[ultimaFila, CantidadColumnas]];
                    rangoAjustarContenido.Columns.AutoFit();*/
                    #endregion
                    break;
                case 2:
                    #region Tipo Instrucciones 2.1                   
                   
                    //Coloca el Titulo de la Seccion
                    Excel.Range columnas = worksheet.Range[worksheet.Cells[ultimaFila, 1], worksheet.Cells[ultimaFila, UltimaColumna]];
                    columnas.Merge();
                    columnas.Interior.Color = Color.LightBlue;
                    worksheet.Cells[ultimaFila, 1] = Titulo;
                    ultimaFila++;

                    //Coloca Descripcion
                    Excel.Range columnasDescripcion = worksheet.Range[worksheet.Cells[ultimaFila, 1], worksheet.Cells[ultimaFila, UltimaColumna]];
                    columnasDescripcion.Merge();             
                    columnasDescripcion.Columns.AutoFit();
                    worksheet.Cells[ultimaFila, 1] = Descripcion;
                    ultimaFila++;

                    #endregion
                    break;

                case 3:
                    #region Tipo Intrucciones 2.2

                    //Establece cantidad de filas a colocar (incluyendo puntajes y su descriptivo, menos se colocaran 2 filas)
                    if (datosSeccion != null)
                    {
                        if (datosSeccion.Count() % 4 != 0)
                        {
                            if (datosSeccion.Count() < 4)
                            {
                                cantidadFilas = 1;
                            }
                            else
                            {
                                cantidadFilas = (datosSeccion.Count() / 4) +1;
                            }
                        }
                        else
                        {
                            if (datosSeccion.Count() / 4 == 1)
                            {
                                cantidadFilas = 1;
                            }
                            else
                            {
                                cantidadFilas = (datosSeccion.Count() / 4);
                            }
                            
                        }
                    }

                    //Establece cantidad de columnas 
                    //Debido a Formato la cantidad de columas visibles se multiplican * 2  
                    CantidadColumnas = CantidadColumnas * 2;            
                    
                    int cantidadColumnasAgregar = 0;

                    //Define cantidad de bloques que seran creados
                    int cantidadBloques =cantidadFilas*2;

                    int numeroFilaInicial = 0;
                    //Verifica si cantidad de columnas actuales es menor que la cantidad requerida
                    if (UltimaColumna < CantidadColumnas)
                    {
                        cantidadColumnasAgregar = CantidadColumnas - UltimaColumna;
                        //Verificar si la cantidad de columnas a agregar es par
                        if (cantidadColumnasAgregar % 2 != 0)
                        {
                            cantidadColumnasAgregar = (cantidadColumnasAgregar / 2) + 1;
                        }

                        //Agrega columnas si es necesario
                        worksheet = AgregadorColumnasEquitativas(worksheet, cantidadColumnasAgregar);
                        cantidadColumnasAgregar = 0;
                    }

                    //Agrega valores a la primera sección del bloque
                    AgregarDatosFilasColumnas(worksheet, CantidadColumnas, cantidadFilas, datosSeccion, 3);

                    //Agrega valores a la segunda sección del bloque
                    AgregarDatosFilasColumnas(worksheet, CantidadColumnas, cantidadFilas, datosSeccion, 4);
                    #endregion
                    break;

            }
            return worksheet;
        }

        private Worksheet AgregarDatosFilasColumnas(Worksheet worksheet, int CantidadColumnas, int cantidadFilas, IEnumerable<SeccionModels> datosSeccion, int tipoSeccion) {

            if (tipoSeccion == 4)
            {
                CantidadColumnas = CantidadColumnas / 2;
            }

            //Obtiene cada cuantas columnas se hara un borde en la fila
            int CantidadSumarColumnas = UltimaColumna / CantidadColumnas;
            //Incializamos la Columna Separadora Final
            ColummaSeparadoraFinal = CantidadSumarColumnas;
            //falta verificar si es para o no para la cantidad de columnas actuales y si se puede dividir entre cantidad de columnas requeridas
            int contadorDatos = 1;
            //Agrega por tipo de dato 
            for (int i = 0; i < cantidadFilas; i++)
            {
                int contadorDatosPorColumnas = 0;
                for (int j = 0; j < CantidadColumnas; j++)
                {
                    if (contadorDatosPorColumnas < CantidadColumnas)
                    {
                        //Coloca el borde sobre las columnas visibles
                        foreach (var obj in datosSeccion)
                        {
                            SeccionModels seccion = obj;
                            if (seccion.NumeroOrden == contadorDatos)
                            {
                                //define el rango para crear las columnas visibles
                                
                                
                                switch (tipoSeccion) {
                                    case 1:
                                        Excel.Range rangocolumnasvisibles = worksheet.Range[worksheet.Cells[ultimaFila, ColumnaSeparadoraInicial], worksheet.Cells[ultimaFila, ColummaSeparadoraFinal]];
                                        rangocolumnasvisibles.Merge();
                                        rangocolumnasvisibles.Columns.AutoFit();
                                        worksheet.Cells[ultimaFila, ColummaSeparadoraFinal - 1] = seccion.Titulo + " " + seccion.ValorDato;
                                        Excel.Borders border = rangocolumnasvisibles.Borders;
                                        border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        SetCellFirstWordBold(worksheet.Cells[ultimaFila, ColummaSeparadoraFinal - 1], ':');
                                        break;
                                    case 2:
                                        break;
                                    case 3:
                                        //Asigna valores
                                        worksheet.Cells[ultimaFila, ColummaSeparadoraFinal] = seccion.NumeroOrden;
                                        worksheet.Cells[ultimaFila, ColummaSeparadoraFinal+1] = seccion.Titulo;
                                        
                                        //Establece Estilos
                                        contadorDatosPorColumnas++;
                                        ColummaSeparadoraFinal++;
                                        Excel.Range rango221 = worksheet.Range[worksheet.Cells[ultimaFila, ColummaSeparadoraFinal-1], worksheet.Cells[ultimaFila, ColummaSeparadoraFinal-1]];
                                        Excel.Range rango222 = worksheet.Range[worksheet.Cells[ultimaFila, ColummaSeparadoraFinal], worksheet.Cells[ultimaFila, ColummaSeparadoraFinal]];
                                        rango221.Columns.AutoFit();
                                        rango221.Merge();
                                        rango222.Columns.AutoFit();
                                        rango222.Merge();

                                        //Coloca Bordes
                                        Excel.Borders border221 = rango221.Borders;
                                        Excel.Borders border222 = rango222.Borders;
                                        border221.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        border222.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        break;
                                    case 4:                                       
                                        
                                        Excel.Range celdaDescripcion = worksheet.Range[worksheet.Cells[ultimaFila, ColumnaSeparadoraInicial], worksheet.Cells[ultimaFila, ColummaSeparadoraFinal]];
                                        celdaDescripcion.Merge();
                                        
                                        //Establece Estilos
                                        celdaDescripcion.Style.WrapText = true;
                                        worksheet.Cells[ultimaFila, ColummaSeparadoraFinal - 1] = seccion.Descripcion;
                                        celdaDescripcion.EntireRow.AutoFit();
                                        celdaDescripcion.EntireRow.RowHeight = 60;
                                        celdaDescripcion.VerticalAlignment = XlVAlign.xlVAlignCenter;

                                        //Coloca bordes
                                        Excel.Borders border223= celdaDescripcion.Borders;
                                        border223.LineStyle = Excel.XlLineStyle.xlContinuous;

                                        break;

                                }
                                

                                //incrementa valores para la proxima columna visible
                                ColumnaSeparadoraInicial = ColummaSeparadoraFinal + 1;
                                ColummaSeparadoraFinal = ColummaSeparadoraFinal + CantidadSumarColumnas;
                                contadorDatos++;
                                contadorDatosPorColumnas++;
                            }
                            if (contadorDatosPorColumnas >= CantidadColumnas)
                                break;
                        }
                    }

                    else
                    {
                        //contadorDatos = 1;
                        contadorDatosPorColumnas = 0;
                        break;
                    }

                }
                ColumnaSeparadoraInicial = 1;
                ColummaSeparadoraFinal = CantidadSumarColumnas;
                ultimaFila++;
            }
            ColummaSeparadoraFinal = 0;
            ColumnaSeparadoraInicial = 1;


            return worksheet;
        }


        //Metodo para Agregar una cantidad N de Filas en una Fila especifica
        private Worksheet AgregadorFilas(Worksheet worksheet, int cantidadFilasAgregar, int NumeroFila,bool bordeadoCompleto)
        {
            int FilaActual = NumeroFila;
            if (cantidadFilasAgregar > 0)
            {
                Range rangoFila = worksheet.Range[worksheet.Cells[FilaActual+1, 1], worksheet.Cells[FilaActual+1, UltimaColumna]];
                for (int i = 0; i < cantidadFilasAgregar; i++)
                {
                    rangoFila.Insert(XlInsertShiftDirection.xlShiftDown, false);
                    if (bordeadoCompleto == true)
                    {                        
                        Range filaNueva = worksheet.Range[worksheet.Cells[FilaActual -i, 1], worksheet.Cells[FilaActual -i, UltimaColumna]];
                        Excel.Borders borderCaso221 = rangoFila.Borders;
                        borderCaso221.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    ultimaFila++;
                }
            }
            return worksheet;
        }


        //Metodo para Agregar Columnas de manera equitativa en el Excel        
        private Excel.Worksheet AgregadorColumnasEquitativas(Excel.Worksheet worksheet, int cantidadColumnasAgregar) {

            //determina cuantas columnas se agregan por cada extremo
            int cantidadPorCadaLado = cantidadColumnasAgregar / 2;

            /*
             * Variables de Columna y Fila Finales actuales: 
             * UltimaColumna / UltimaFila
             */

            //Agrega Columnas en Pares
            for (int i = 0; i < cantidadPorCadaLado; i++) {
                //Añade al principio de las columnas (Fila 1, Columna 2)
                Excel.Range rangoInicial = (Excel.Range)worksheet.Cells[1, 2];                
                rangoInicial.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                        XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                UltimaColumna++;


                //Añade al final de las columnas (Fila 1, Columna n)
                Excel.Range rangoFinal = (Excel.Range)worksheet.Cells[1, UltimaColumna];
                rangoFinal.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight,
                        XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                UltimaColumna++;
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
            worksheet = DefineTipoSeccion(worksheet, 1, CargaDatosCabecera(), TituloCabecera,null, 2);

            //Define seccion Instrucciones 2.1
            string TituloInstrucciones21 = "2.- INSTRUCCIONES";
            string DescripcionInstrucciones21 = "Tomando en cuenta el desempeño del evaluado durante el periodo establecido, siga las siguientes instrucciones para realizar la calificación.";
            worksheet = DefineTipoSeccion(worksheet, 2, null, TituloInstrucciones21, DescripcionInstrucciones21, 8);

            //Define seccion Instrucciones 2.2
            //string TituloInstrucciones22 = "2.- INSTRUCCIONES";
            //string DescripcionInstrucciones22 = "Tomando en cuenta el desempeño del evaluado durante el periodo establecido, siga las siguientes instrucciones para realizar la calificación.";
            worksheet = DefineTipoSeccion(worksheet, 3, DatosInstrucciones(), null, null, 4);

            //Nombre aleatorio basado en Timestamp Actual
            String timeStamp = GetTimestamp(DateTime.Now);

            workbook.SaveAs("C:\\Users\\DitaMonster\\Documents\\Visual Studio 2017\\Projects\\Zinatic_UED\\Zinatic-Docs-Upload---Search\\zinatic_UED\\files\\" + timeStamp + ".xlsx"); ;
            workbook.Close(true);
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }
    }
}