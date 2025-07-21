using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XCeedWordInspeccion
{
    internal class Program
    {
        
        private const int MAX_VIAS = 5;
        
        // private const int IdOT = 81445; // ID de la OT para pruebas
        // private const string NumOs = "250529.16"; // Número de OS para pruebas
        
        private const int IdOTC = 6931; // ID de la OT para pruebas
        private const int IdOT = 82233; // ID de la OT para pruebas
        private const string NumOs = "250622.01"; // Número de OS para pruebas
        
        public static void Main(string[] args)
        {
            
            string filename = "Inspecciones.docx";
            string templatePath = @"C:\Users\ljauregui\RiderProjects\XCeedWord\XCeedWord\bin\Debug\PlantillaAC.docx";
            // string templatePath = @"C:\Users\LUIS\RiderProjects\XCeedWordInspecciones\XCeedWordInspeccion\bin\Debug\PlantillaAC.docx";
            
            
            File.Copy(templatePath, filename, true);

            using (DocX document = DocX.Load(filename))
            {
                
                
                SqlRepository repository = new SqlRepository();

                List<Model.Ensayo> ensayos = repository.ObtenerEnsayos<Model.Ensayo>(IdOT, 5, 2).ToList();
                List<Model.Via> vias = repository.ObtenerVias<Model.Via>(IdOT, 2, 5).ToList();
                List<Model.ViaResultado> viaResultado = repository.ViasResultados<Model.ViaResultado>(IdOT, 2).ToList();
                List<Model.CodigoVia> codigoVias = repository.ObtenerCodigoVias<Model.CodigoVia>(NumOs).ToList();
                List<Model.MuestraCls> muestras = repository.ObtenerMuestras<Model.MuestraCls>(IdOT, 2).ToList();
                // Model.LoteCls lote = repository.ObtenerLote<Model.LoteCls > (IdOT);
                
                // Configurar márgenes
                
                document.MarginLeft   = 25; // 2.5 cm
                document.MarginRight  = 20; // 2 cm
                document.MarginTop    = 0; // 2 cm
                document.MarginBottom = 0; // 2 cm
                
                // Crear Tabla
                
                // Hacer un filtro para que no se repita filas con el mismo código de vía
                
                codigoVias = codigoVias
                    .GroupBy(c => c.CodigoInterno)
                    .Select(g => g.First())
                    .ToList();
                
                // Pero si mi codigo via es por ejemplo "M1", "M2" como puedo ordenarlo ascendenten

                codigoVias = codigoVias
                    // Ordenamos por la parte numérica del CodigoInterno
                    .OrderBy(c => int.Parse(c.CodigoInterno.Substring(1)))
                    .ThenBy(c => c.ProductoCodigo) // Mantenemos el segundo nivel de orden si lo necesitas
                    .ToList();
                
                CrearTablaMuestrasExtraidasMicrobiologia(document, repository);
                document.InsertParagraph();
                CrearTablaMuestrasExtraidasFisicoSensorial(document, repository);
                document.InsertParagraph();
                CrearTablaMuestreoParaAnalisisMicrobiologicos(document, repository);
                document.InsertParagraph();
                CrearTablaExamenesSensorialesSecoSalado(document, repository);
                
                // CrearTablaLaboratorioMuestrasDirimentes(document, codigoVias, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaEsterilidadComercial(document, ensayos, codigoVias, viaResultado);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaIndicadoresParasitologicos(document, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaEvaluacionDobleCierre(document, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaDeterminacionPresionVacio(document, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaHistamina(document, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaMetalesPesados(document, repository);
                // document.InsertParagraph().SpacingAfter(10);
                // CrearTablaExtensionesSensoriales(document, codigoVias, repository);
                
                // CrearTablaK(document, ensayos, codigoVias);
                // CrearTablaB(document, vias, codigoVias);
                // CrearTablaDeterminacionVacia(document, codigoVias);
                // CrearTablaHistamina(document, codigoVias);
                // CrearTablaMetalesPesados(document, codigoVias);
                // CrearTablaExtensionesSensoriales(document, codigoVias);
                
                // var totalTables = (int) Math.Ceiling(vias.Count / (double)MAX_VIAS);
                //
                // for (int i = 0; i < totalTables; i++)
                // {
                //
                //     List<Model.Via> rangeVias =
                //         vias.GetRange((i * MAX_VIAS), Math.Min(MAX_VIAS, vias.Count - i * MAX_VIAS));
                //
                //     var rangeResultados =
                //         viaResultado
                //             .OrderBy(v => v.IdAnalisis)
                //             .ThenBy(v => v.CodPrecinto.Substring(1))
                //             .ThenBy(v => int.Parse(v.Muestra.Substring(1)))
                //             .ToList();
                //     
                //     CreateTableB(document, rangeVias, ensayos, rangeResultados, i, vias.Count);
                //
                //     // CreateTableJ(document, codigoVias);
                //     // CreateTableG(document, rangeVias, ensayos, rangeResultados, codigoVias, i, vias.Count);
                //     document.InsertParagraph().SpacingAfter(1);
                // }

                document.Save();
                
                // Abrir documento
                
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filename,
                    UseShellExecute = true
                });
                
                Console.WriteLine("Documento creado exitosamente: " + Path.GetFullPath(filename));
            }

        }

        public static void CreateTableA(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, List<Model.CodigoVia> codigoVias, int iTable, int numVias)
        {
            int headerRows = 3;
            int headerColumns = 6;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 100, 40, 40, 45, 45 };
            
            for (int i = 0; i < tableColumns; i++)
            {

                if (i <= columnWidths.Length - 1)
                {
                    table.SetColumnWidth(i, columnWidths[i]);
                }
                
                // Vías dinámicas

                if (i >= 5 && i < tableColumns - 1)
                {
                    table.SetColumnWidth(i, 35);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 0, headerRows - 1);
            
            table.MergeCellsInColumn(1, 0, 1);
            table.MergeCellsInColumn(2, 0, 1);
            
            table.MergeCellsInColumn(3, 0, 1);
            table.MergeCellsInColumn(4, 0, 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 0, headerRows - 1);
            
            // Microorganismo
            
            FormatTableCell(table.Rows[0].Cells[0], "MICROORGANISMO", 3, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[0].MergeCells(1, 2);
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[0].Cells[1], "PLAN DE EVALUACIÓN", 3, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", 3, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", 3, true, Alignment.center);
            
            // Limites
            
            table.Rows[0].MergeCells(2, 3);
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[0].Cells[2], "LIMITES", 3, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[3], "m", 3, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[4], "M", 3, true, Alignment.center);
            
            // Distribución de muestras
            
            table.Rows[0].MergeCells(3, 3 + vias.Count - 1);
            FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", 3, true, Alignment.center);
            
            // Código de Vías (Dinámicas)
            
            AgruparYFormatearVias(table, vias, codigoVias, 1, 5);
                
            // for (int i = 0, aux = 0; i < vias.Count;)
            // {
            //     string currentVia = vias[i].Presentacion;
            //     int startCol = 5 + i - aux;
            //     int j = i + 1;
            //
            //     // Buscar cuántas 'vias' consecutivas tienen la misma presentación
            //     while (j < vias.Count && vias[j].Presentacion == currentVia)
            //     {
            //         j++;
            //         aux++;
            //     }
            //
            //     int endCol =  5 + j - (i == 0 ? 1 : aux);
            //
            //     // Formatear y/o fusionar celdas según cantidad de columnas iguales
            //     if (j - i > 1)
            //     {
            //         table.Rows[1].MergeCells(startCol - 2, endCol - 2); 
            //         // table.Rows[2].MergeCells(startCol, endCol);
            //     }
            //
            //
            //     var productoCodigo = codigoVias.Find(x => x.CodigoInterno == currentVia).ProductoCodigo;
            //     FormatTableCell(table.Rows[1].Cells[startCol - 2], productoCodigo, 3, true, Alignment.center);
            //     
            //     // FormatTableCell(table.Rows[2].Cells[startCol], currentVia, 4, true, Alignment.center); // Codigo
            //
            //     i = j; // Saltar al siguiente grupo
            // }
            
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, 4, true, Alignment.center);
            }
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 4, false, Alignment.left);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), 4, false, Alignment.center);

                int j = 5;
                
                // Resultado por cada ensayo
                
                var resultados = 
                    viaResultados
                        .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis).ToList()
                        .GetRange((iTable * MAX_VIAS), Math.Min(MAX_VIAS, viaResultados.Count - iTable * MAX_VIAS));

                foreach (var via in resultados)
                {
                    
                    // bool match = (ensayos[i].IdProducto == via.IdProducto && ensayos[i].IdAnalisis == via.IdAnalisis && via.CodigoInterno ==);

                    if (true)
                    {
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 4, true, Alignment.center);
                    }

                    j++;

                }
                    
            }
            
            // Conclusión

            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", 3, true, Alignment.center);

            // Guardar
            
            document.InsertTable(table);
        }

        public static void CreateTableB(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, int iTable)
        {
            int headerRows = 4;
            int headerColumns = 7;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.left;
            
            // Agregar titulo
            
            AgregarTitulo(table, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS MICROBIOLÓGICAS) - PERÚ Y OTROS PAÍSES. NUMERAL 1.2.8");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 100, 30, 30, 50, 50, 45 };
            
            for (int i = 0; i < tableColumns; i++)
            {

                if (i <= columnWidths.Length - 1)
                {
                    table.SetColumnWidth(i, columnWidths[i]);
                }
                
                // Vías dinámicas

                if (i >= 5 && i < tableColumns - 1)
                {
                    table.SetColumnWidth(i, 40);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
            }
            
            // Combinar filas
            
            table.MergeCellsInColumn(0, 2, headerRows - 1);
            
            table.MergeCellsInColumn(3, 2, headerRows - 1);
            table.MergeCellsInColumn(4, 2, headerRows - 1);
            
            table.MergeCellsInColumn(5, 2, headerRows - 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 2, headerRows - 1);
            
            table.Rows[0].MergeCells(0, table.ColumnCount - 1); // Titulo
            
            // Esterilidad comercial
            
            table.Rows[1].MergeCells(0, table.Rows[1].Cells.Count - 1);
            FormatTableCell(table.Rows[1].Cells[0], "ESTERILIDAD COMERCIAL", 7, true, Alignment.center);
            
            // Analisis
            
            FormatTableCell(table.Rows[2].Cells[0], "ANALISIS", 7, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[2].MergeCells(1, 2);
            FormatTableCell(table.Rows[2].Cells[1], "PLAN DE EVALUACIÓN", 6, true, Alignment.center);
            
            FormatTableCell(table.Rows[3].Cells[1], "n", 7, true, Alignment.center);
            FormatTableCell(table.Rows[3].Cells[2], "c", 7, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[2].Cells[2], "ACEPTACIÓN", 7, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[2].Cells[3], "RECHAZO", 7, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[2].Cells[4], "CÓDIGO", 7, true, Alignment.center);
            
            // Número de Vías (Dinámicas)
            
            for (int i = 0, aux = 0; i < vias.Count;)
            {
                
                string currentVia = vias[i].Presentacion;
                int startCol = 5 + i - aux;
                int j = i + 1;

                // Buscar cuántas 'vias' consecutivas tienen la misma presentación
                while (j < vias.Count && vias[j].Presentacion == currentVia)
                {
                    j++;
                    aux++;
                }

                int endCol =  5 + j - (i == 0 ? 1 : aux);

                // Formatear y/o fusionar celdas según cantidad de columnas iguales
                if (j - i > 1)
                {
                    table.Rows[2].MergeCells(startCol, endCol);
                }

                FormatTableCell(table.Rows[2].Cells[startCol], currentVia, 7, true, Alignment.center);

                i = j; // Saltar al siguiente grupo
            }

            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 6; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[3].Cells[iCellIndex], vias[i].Muestra, 7, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[2].Cells[table.Rows[2].Cells.Count - 1], "CONCLUSIÓN", 7, true, Alignment.center);
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 7, false, Alignment.center, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], "5", 7, false, Alignment.center, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[2], "0", 7, false, Alignment.center, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[3], "Estéril comercialmente", 7, false, Alignment.center, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[4], "No estéril comercialmente", 7, false, Alignment.center, false);

                int j = 6;
                
                // Resultado por cada ensayo
                
                var resultados = 
                    viaResultados
                        .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis).ToList()
                        .GetRange((iTable * MAX_VIAS), Math.Min(MAX_VIAS, viaResultados.Count - iTable * MAX_VIAS));

                foreach (var via in resultados)
                {
                    
                    // bool match = (ensayos[i].IdProducto == via.IdProducto && ensayos[i].IdAnalisis == via.IdAnalisis && via.CodigoInterno ==);

                    if (true)
                    {
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 6, false, Alignment.center, false);
                    }

                    j++;

                }
                    
            }
            
            // Si no hay descripción eliminamos la ultima fila
            
            table.Rows.Last().Remove();
            
            // Guardar
            
            document.InsertTable(table);
            
        }
        
        public static void CreateTableC(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, List<Model.CodigoVia> codigoVias, int iTable, int numVias)
        {
            int headerRows = 4;
            int headerColumns = 6;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.left;
            
            // Agregar titulo
            
            AgregarTitulo(table, "CARACTERISTICAS QUIMICAS (PERU Y OTROS PAISES): 1.2.4-TABLA N°03 (item 2)");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 100, 30, 30, 30, 30 };
            
            for (int i = 0; i < tableColumns; i++)
            {

                if (i <= columnWidths.Length - 1)
                {
                    table.SetColumnWidth(i, columnWidths[i]);
                }
                
                // Vías dinámicas

                if (i >= 5 && i < tableColumns - 1)
                {
                    table.SetColumnWidth(i, 30);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 60);
                }
                
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            
            table.MergeCellsInColumn(1, 1, 2);
            table.MergeCellsInColumn(2, 1, 2);
            
            table.MergeCellsInColumn(3, 1, 2);
            table.MergeCellsInColumn(4, 1, 2);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 1, headerRows - 1);
            
            table.Rows[0].MergeCells(0, table.ColumnCount - 1); // Titulo
            
            // Determinación
            
            FormatTableCell(table.Rows[1].Cells[0], "DETERMINACIÓN", 8, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[1].MergeCells(1, 2);
            table.Rows[2].MergeCells(1, 2);
            FormatTableCell(table.Rows[1].Cells[1], "PLAN DE EVALUACIÓN", 8, true, Alignment.center);
            
            FormatTableCell(table.Rows[3].Cells[1], "n", 8, true, Alignment.center);
            FormatTableCell(table.Rows[3].Cells[2], "c", 8, true, Alignment.center);
            
            // Limites
            
            table.Rows[1].MergeCells(2, 3);
            table.Rows[2].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "LIMITES (mg/kg)", 8, true, Alignment.center);
            
            FormatTableCell(table.Rows[3].Cells[3], "m", 8, true, Alignment.center);
            FormatTableCell(table.Rows[3].Cells[4], "M", 8, true, Alignment.center);
            
            // Distribución de muestras
            
            table.Rows[1].MergeCells(3, 3 + vias.Count - 1);
            FormatTableCell(table.Rows[1].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", 8, true, Alignment.center);
            
            // Código de Vías (Dinámicas)
            
            AgruparYFormatearVias(table, vias, codigoVias, 2,5);
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[3].Cells[iCellIndex], vias[i].Muestra, 8, true, Alignment.center);
            }
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 7, false, Alignment.left, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), 7, false, Alignment.center, false);

                int j = 5;
                
                // Resultado por cada ensayo
                
                var resultados = 
                    viaResultados
                        .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis).ToList()
                        .GetRange((iTable * MAX_VIAS), Math.Min(MAX_VIAS, viaResultados.Count - iTable * MAX_VIAS));

                foreach (var via in resultados)
                {
                    
                    // bool match = (ensayos[i].IdProducto == via.IdProducto && ensayos[i].IdAnalisis == via.IdAnalisis && via.CodigoInterno ==);

                    if (true)
                    {
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 8, false, Alignment.center, false);
                    }

                    j++;

                }
                    
            }
                
            // Conclusión

            FormatTableCell(table.Rows[1].Cells[table.Rows[1].Cells.Count - 1], "CONCLUSIÓN", 8, true, Alignment.center);
            
            table.RemoveRow(); // Removemos la fila para la descripción
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableD(DocX document, int totalEnsayos)
        {
            int headerRows = 1;
            int headerColumns = 6;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.left;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 80, 50, 100, 100, 100, 80 };

            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // M
            
            FormatTableCell(table.Rows[0].Cells[0], "M", 8, true, Alignment.center);
            
            // n
            
            FormatTableCell(table.Rows[0].Cells[1], "n", 8, true, Alignment.center);
            
            // Elementos
            
            FormatTableCell(table.Rows[0].Cells[2], "ELEMENTOS", 8, true, Alignment.center);
            
            // Contenido máximo
            
            FormatTableCell(table.Rows[0].Cells[3], "CONTENIDO MÁXIMO", 8, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(table.Rows[0].Cells[4], "RESULTADO", 8, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[0].Cells[5], "CONCLUSIÓN", 8, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableE(DocX document, int vias)
        {
            int headerRows = 2;
            int headerColumns = 5;

            int tableRows = headerRows + vias;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 80, 80, 100, 100, 100 };

            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Indicadores parasitologicos
            
            table.Rows[0].MergeCells(0, table.Rows[0].Cells.Count - 1);
            FormatTableCell(table.Rows[0].Cells[0], "INDICADORES PARASITOLOGICOS", 6, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", 5, true, Alignment.center);
            
            // Vías (n)
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS (n)", 5, true, Alignment.center);
            
            // Plan de evaluación
            
            FormatTableCell(table.Rows[1].Cells[2], "PLAN DE EVALUACIÓN", 5, true, Alignment.center);
            
            // Resultados
            
            FormatTableCell(table.Rows[1].Cells[3], "RESULTADOS", 5, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[4], "CONCLUSIÓN", 5, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableF(DocX document, int ensayos)
        {
            int headerRows = 2;
            int headerColumns = 7;

            int tableRows = headerRows + ensayos;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 80, 60, 60, 80, 80, 80, 80 };

            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Metales pesados
            
            table.Rows[0].MergeCells(0, table.Rows[0].Cells.Count - 1);
            FormatTableCell(table.Rows[0].Cells[0], "METALES PESADOS", 6, true, Alignment.center);
            
            // Análisis
            
            FormatTableCell(table.Rows[1].Cells[0], "ANÁLISIS", 5, true, Alignment.center);
            
            // Códigos
            
            FormatTableCell(table.Rows[1].Cells[1], "CÓDIGOS", 5, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[2], "VÍAS", 5, true, Alignment.center);
            
            // Contenido máximo (mg/kg peso fresco)
            
            FormatTableCell(table.Rows[1].Cells[3], "CONTENIDO MÁXIMO (mg/kg peso fresco)", 5, true, Alignment.center);
            
            // Unidades
            
            FormatTableCell(table.Rows[1].Cells[4], "UNIDADES", 5, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(table.Rows[1].Cells[5], "RESULTADO", 5, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[6], "CONCLUSIÓN", 5, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableG(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, List<Model.CodigoVia> codigoVias, int iTable, int numVias)
        {
            int headerRows = 3;
            int headerColumns = 7;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.left;
            
            // Agregar titulo
            
            AgregarTitulo(table, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERISTICAS MICROBIOLÓGICAS): R.M. 591-2008/MINSA. NUMERAL XIX.1");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 45, 20, 20, 50, 50, 50 };
            
            for (int i = 0; i < tableColumns; i++)
            {

                if (i <= columnWidths.Length - 1)
                {
                    table.SetColumnWidth(i, columnWidths[i]);
                }
                
                // Vías dinámicas

                if (i >= 6 && i < tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            
            table.MergeCellsInColumn(3, 1, headerRows - 1);
            table.MergeCellsInColumn(4, 1, headerRows - 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 1, headerRows - 1);
            
            table.Rows[0].MergeCells(0, table.ColumnCount - 1); // Titulo
            
            // Ensayo
            
            FormatTableCell(table.Rows[1].Cells[0], "ENSAYO", 6, true, Alignment.center);
            
            // Plan de muestreo
            
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[1].Cells[1], "PLAN DE MUESTREO", 6, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", 7, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[1].Cells[2], "ACEPTACIÓN", 6, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[1].Cells[3], "RECHAZO", 6, true, Alignment.center);
            
            // Resultados
            
            table.Rows[1].MergeCells(4, 4 + vias.Count);
            FormatTableCell(table.Rows[1].Cells[4], "RESULTADOS", 6, true, Alignment.center);
            
            // Lote
            
            FormatTableCell(table.Rows[2].Cells[5], "LOTE", 7, true, Alignment.center);
            
            // Vias (Dinámicas)
            
            for (int i = 0, iCellIndex= 6; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, 7, true, Alignment.center);
            }
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 6, false, Alignment.left, false);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), 6, false, Alignment.center, false);

                int j = 6;
                
                // Resultado por cada ensayo
                
                var resultados = 
                    viaResultados
                        .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis).ToList()
                        .GetRange((iTable * MAX_VIAS), Math.Min(MAX_VIAS, viaResultados.Count - iTable * MAX_VIAS));

                foreach (var via in resultados)
                {
                    
                    // bool match = (ensayos[i].IdProducto == via.IdProducto && ensayos[i].IdAnalisis == via.IdAnalisis && via.CodigoInterno ==);

                    if (true)
                    {
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 6, false, Alignment.center, false);
                    }

                    j++;

                }
                    
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[table.Rows[1].Cells.Count - 1], "CONCLUSIÓN", 6, true, Alignment.center);
            
            // Descripción
            
            AgregarDescripcion(table, "1 Norma o referencia: AOAC Official Method 972.44, 22nd Edition: 2023: Sterility (Commercial) of foods (Canned, low acid).");
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableH(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 3;
            int headerColumns = 11;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 50, 50, 30, 30, 40, 40, 50, 50, 50, 50, 50 };
            
            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            table.MergeCellsInColumn(1, 1, headerRows - 1);
            
            table.MergeCellsInColumn(6, 1, headerRows - 1);
            table.MergeCellsInColumn(7, 1, headerRows - 1);
            table.MergeCellsInColumn(8, 1, headerRows - 1);
            table.MergeCellsInColumn(9, 1, headerRows - 1);
            table.MergeCellsInColumn(10, 1, headerRows - 1);
            
            // Examenes sensoriales
            
            table.Rows[0].MergeCells(0, table.Rows[0].Cells.Count - 1);
            FormatTableCell(table.Rows[0].Cells[0], "EXAMENES SENSORIALES", 8, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", 7, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS", 7, true, Alignment.center);
            
            // Numeración de Aceptación
            
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[2], "N*", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[3], "(c)*", 7, true, Alignment.center);
            
            // Aspecto
            
            table.Rows[1].MergeCells(3, 4);
            FormatTableCell(table.Rows[1].Cells[3], "ASPECTO", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[4], "EXTERIOR", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[5], "INTERIOR", 7, true, Alignment.center);
            
            // Olor
            
            FormatTableCell(table.Rows[1].Cells[4], "OLOR", 7, true, Alignment.center);
            
            // Color
            
            FormatTableCell(table.Rows[1].Cells[5], "COLOR", 7, true, Alignment.center);
            
            // Sabor
            
            FormatTableCell(table.Rows[1].Cells[6], "SABOR", 7, true, Alignment.center);
            
            // Textura
            
            FormatTableCell(table.Rows[1].Cells[7], "TEXTURA", 7, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[8], "CONCLUSIÓN", 7, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableI(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 3;
            int headerColumns = 11;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 50, 50, 30, 30, 40, 40, 40, 50, 50, 50, 50 };
            
            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            table.MergeCellsInColumn(1, 1, headerRows - 1);
            
            table.MergeCellsInColumn(6, 1, headerRows - 1);
            table.MergeCellsInColumn(7, 1, headerRows - 1);
            table.MergeCellsInColumn(8, 1, headerRows - 1);
            table.MergeCellsInColumn(9, 1, headerRows - 1);
            table.MergeCellsInColumn(10, 1, headerRows - 1);
            
            // Examenes sensoriales
            
            table.Rows[0].MergeCells(0, table.Rows[0].Cells.Count - 1);
            FormatTableCell(table.Rows[0].Cells[0], "EXAMENES SENSORIALES", 8, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", 7, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS", 7, true, Alignment.center);
            
            // Numeración de Aceptación
            
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[2], "N*", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[3], "(c)*", 7, true, Alignment.center);
            
            // Aspecto
            
            table.Rows[1].MergeCells(3, 4);
            FormatTableCell(table.Rows[1].Cells[3], "ASPECTO", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[4], "EXTERIOR", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[5], "INTERIOR", 7, true, Alignment.center);
            
            // Olor
            
            FormatTableCell(table.Rows[1].Cells[4], "OLOR", 7, true, Alignment.center);
            
            // Color
            
            FormatTableCell(table.Rows[1].Cells[5], "COLOR", 7, true, Alignment.center);
            
            // Sabor
            
            FormatTableCell(table.Rows[1].Cells[6], "SABOR", 7, true, Alignment.center);
            
            // Textura
            
            FormatTableCell(table.Rows[1].Cells[7], "TEXTURA", 7, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[8], "CONCLUSIÓN", 7, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        // public static void CrearTablaEvaluacionDobleCierre(DocX document, SqlRepository repository)
        // {
        //     List<Model.CodigoViaFisicoSensorial> codigoVias = repository.ObtenerCodigoViasFisicoSensorial<Model.CodigoViaFisicoSensorial>(IdOT, 1).ToList();
        //     
        //     codigoVias = codigoVias
        //         // Ordenamos por la parte numérica del CodigoInterno
        //         .OrderBy(c => int.Parse(c.Codigos.Substring(1)))
        //         .ThenBy(c => c.Codigos) // Mantenemos el segundo nivel de orden si lo necesitas
        //         .ToList();
        //     
        //     List<Model.SpGetTablaEvaluacionDobleCierre> resultados =
        //         repository.ObtenerTablaEvaluacion<Model.SpGetTablaEvaluacionDobleCierre>(IdOT, 1).ToList();
        //     
        //     List<Model.UspGetTablaExamenesSensoriales> examenesSensoriales =
        //         repository.ObtenerTablaExamenesSensorial<Model.UspGetTablaExamenesSensoriales>(IdOT, 1).ToList();
        //     
        //     int cabeceraFilas = 4;
        //     // int cabeceraColumnas = 13;
        //     int cabeceraColumnas = 21;
        //
        //     int tablaFilas = cabeceraFilas + (codigoVias.Count * 5); // El usuario indica que siempre van a tener 5 filas
        //     int tablaColumnas = cabeceraColumnas;
        //     
        //     Table tabla = document.AddTable(tablaFilas, tablaColumnas);
        //     tabla.Alignment = Alignment.left;
        //     
        //     // Titulo
        //     
        //     AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS FÍSICO-SENSORIALES) - PERÚ Y OTROS PAÍSES. NUMERAL 1.2.6-TABLA N°04");
        //     
        //     // Encabezado
        //     
        //     // Ancho de columnas
        //     
        //     int[] columnWidths = { 25, 20, 20, 35, 40, 40, 20,20,20, 20,20,20, 20,20,20, 20,20,20, 32, 30, 50 };
        //     
        //     for (int i = 0; i < tablaColumnas; i++)
        //     {
        //         tabla.SetColumnWidth(i, columnWidths[i]);
        //     }
        //     
        //     // Combinar filas
        //     
        //     tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1); // Ùltima columnna
        //     tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1);
        //     
        //     tabla.Rows[2].MergeCells(6, 8); // Compacidad
        //     tabla.Rows[2].MergeCells(7, 9); // Penetracion
        //     tabla.Rows[2].MergeCells(8, 10); // Traslape
        //     tabla.Rows[2].MergeCells(9, 11); // Traslape teorico
        //     
        //     tabla.Rows[3].MergeCells(6, 8); // Compacidad
        //     tabla.Rows[3].MergeCells(7, 9); // Penetracion
        //     tabla.Rows[3].MergeCells(8, 10); // Traslape
        //     tabla.Rows[3].MergeCells(9, 11); // Traslape teorico
        //
        //     tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(1, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(3, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(4, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(5, 2, cabeceraFilas - 1);
        //     
        //     tabla.MergeCellsInColumn(6, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(7, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(8, 2, cabeceraFilas - 1);
        //     tabla.MergeCellsInColumn(9, 2, cabeceraFilas - 1);
        //     
        //     // Requisitos
        //     
        //     tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
        //     FormatTableCell(tabla.Rows[1].Cells[0], "REQUISITOS PARA LA EVALUACION DEL DOBLE CIERRE EN ENVASES DE HOJALATA", 7, true, Alignment.center, true);
        //     
        //     // Código
        //
        //     tabla.Rows[2].Height = 80;
        //     tabla.Rows[3].Height = 20;
        //     FormatTableCell(tabla.Rows[2].Cells[0], "CODIGO", 7, true, Alignment.center, true, TextDirection.btLr);
        //     
        //     // Vías (n)
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[1], "VIAS (n)", 7, true, Alignment.center, true, TextDirection.btLr);
        //     
        //     // Tolerancia
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[2], "TOLERANCIA", 7, true, Alignment.center, true, TextDirection.btLr);
        //     
        //     // Ganchos
        //
        //     var celdaGanchos = tabla.Rows[2].Cells[3];
        //     
        //     Paragraph pGanchos = celdaGanchos.Paragraphs.First();
        //     
        //     pGanchos.Append("Ganchos de cuerpo y tapa").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pGanchos.AppendLine();
        //     pGanchos.Append("Uniformes en su perímetro").FontSize(4).Font("Arial");
        //
        //     pGanchos.Alignment = Alignment.center;
        //     celdaGanchos.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaGanchos.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Borde
        //     
        //     var celdaBordes = tabla.Rows[2].Cells[4];
        //     
        //     Paragraph pBordes = celdaBordes.Paragraphs.First();
        //     
        //     pBordes.Append("Borde superior e inferior del doble cierre").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pBordes.AppendLine();
        //     pBordes.Append("Lisos y sin irregularidades").FontSize(4).Font("Arial");
        //
        //     pBordes.Alignment = Alignment.center;
        //     celdaBordes.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaBordes.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Compuesto
        //     
        //     var celdaCompuesto = tabla.Rows[2].Cells[5];
        //     
        //     Paragraph pCompuesto = celdaCompuesto.Paragraphs.First();
        //     
        //     pCompuesto.Append("Compuesto sellador").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pCompuesto.AppendLine();
        //     pCompuesto.Append("Debe cubrir los espacios libres internos del doble cierre").FontSize(4).Font("Arial");
        //
        //     pCompuesto.Alignment = Alignment.center;
        //     celdaCompuesto.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaCompuesto.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Compacidad
        //     
        //     var celdaCompacidad = tabla.Rows[2].Cells[6];
        //     
        //     Paragraph pCompacidad = celdaCompacidad.Paragraphs.First();
        //     
        //     pCompacidad.Append("Compacidad (%)").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pCompacidad.AppendLine();
        //     pCompacidad.Append("Envases redondos: Mayor o igual al 75%.").FontSize(4).Font("Arial");
        //     pCompacidad.AppendLine();
        //     pCompacidad.Append("Envases de forma: Mayor o igual al 60%").FontSize(4).Font("Arial");
        //
        //     pCompacidad.Alignment = Alignment.center;
        //     celdaCompacidad.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaCompacidad.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Penetracion
        //     
        //     var celdaPenetracion = tabla.Rows[2].Cells[7];
        //     
        //     Paragraph pPenetracion = celdaPenetracion.Paragraphs.First();
        //     
        //     pPenetracion.Append("Penetración de gancho de cuerpo (%)").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pPenetracion.AppendLine();
        //     pPenetracion.Append("Mayor o igual al 70%").FontSize(4).Font("Arial");
        //
        //     pPenetracion.Alignment = Alignment.center;
        //     celdaPenetracion.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaPenetracion.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Traslape
        //     
        //     var celdaTraslape = tabla.Rows[2].Cells[8];
        //     
        //     Paragraph pTraslape = celdaTraslape.Paragraphs.First();
        //     
        //     pTraslape.Append("Traslape (%)").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pTraslape.AppendLine();
        //     pTraslape.Append("Mayor o igual al 45%").FontSize(4).Font("Arial");
        //
        //     pTraslape.Alignment = Alignment.center;
        //     celdaTraslape.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaTraslape.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Traslape teorico
        //     
        //     var celdaTraslapeTeorico = tabla.Rows[2].Cells[9];
        //     
        //     Paragraph pTraslapeTeorico = celdaTraslapeTeorico.Paragraphs.First();
        //     
        //     pTraslapeTeorico.Append("Traslape teórico (mm)").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     pTraslapeTeorico.AppendLine();
        //     pTraslapeTeorico.Append("Mayor o igual a 1mm").FontSize(4).Font("Arial");
        //
        //     pTraslapeTeorico.Alignment = Alignment.center;
        //     celdaTraslapeTeorico.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaTraslapeTeorico.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     // Arrugas
        //     
        //     var celdaArrugas = tabla.Rows[2].Cells[10];
        //     
        //     Paragraph pArrugasTitulo = celdaArrugas.Paragraphs.First();
        //     
        //     pArrugasTitulo.Append("Arrugas (grado de apriente)").FontSize(5).Font("Arial").UnderlineStyle(UnderlineStyle.singleLine).Bold();
        //     
        //     pArrugasTitulo.Alignment = Alignment.center;
        //
        //     Paragraph pArrugas = celdaArrugas.InsertParagraph();
        //     
        //     pArrugas.Append("Envases redondos: La arruga no debe tener una longitud que represente mas del 25% de la longitud del gancho de tapa (grado de apriete mayor al 75%)").FontSize(4).Font("Arial");
        //     pArrugas.AppendLine();
        //     pArrugas.Append("Envases irregulares: La peor arruga no debe tener una longitud que represente mas del 40% de la longitud del gancho de tapa (grado de apriete mayor al 60%)").FontSize(4).Font("Arial");
        //
        //     pArrugas.Alignment = Alignment.left;
        //     
        //     celdaArrugas.FillColor = Color.FromArgb(234, 241, 221);
        //     celdaArrugas.VerticalAlignment = VerticalAlignment.Center;
        //     
        //     tabla.Rows[2].MergeCells(10, 11);
        //     FormatTableCell(tabla.Rows[3].Cells[10], "PLANCHADO (%)", 4, true, Alignment.center);
        //     FormatTableCell(tabla.Rows[3].Cells[11], "ARRUGAS (%)", 4, true, Alignment.center);
        //
        //     // Conclusión
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[11], "CONCLUSIÓN", 7, true, Alignment.center);
        //     
        //     // Muestras
        //     
        //     for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias += 5)
        //     {
        //         
        //         string codigoInterno = codigoVias[i].Codigos;
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[0], codigoInterno, 5, true, Alignment.center, false);
        //         
        //         tabla.MergeCellsInColumn(0, inicioVias, inicioVias + 4);
        //         
        //         tabla.MergeCellsInColumn(tabla.ColumnCount - 1, inicioVias, inicioVias + 4);
        //         
        //         tabla.MergeCellsInColumn(2, inicioVias, inicioVias + 4);
        //         
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[2], "0", 5, true, Alignment.center, false);
        //         
        //         // Vias
        //         
        //         // TODO: revisar el numeros de vias del for tiene que traerde la base de datos
        //
        //         for (int j = 0; j < 5; j++) 
        //         {
        //             FormatTableCell(tabla.Rows[inicioVias + j].Cells[1], (j + 1).ToString(), 6, false, Alignment.center, false);
        //
        //             if (resultados.Count > 0)
        //             {
        //                 var resultado = resultados.FirstOrDefault(x => x.Codigos == codigoVias[i].Codigos && x.Vias == j + 1);
        //                 var resultadoExamenSensorial = examenesSensoriales.FirstOrDefault(x => x.Codigos == codigoVias[i].Codigos && x.Vias == j + 1);
        //                 
        //                 if (resultado != null)
        //                 {
        //                     // Ganchos de cuerpo y tapa
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[3], resultadoExamenSensorial.EnvaseInterno, 5, false, Alignment.center, false);
        //                     
        //                     // Borde superior e inferior del doble cierre
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[4], resultadoExamenSensorial.EnvaseExterno, 5, false, Alignment.center, false);
        //                     
        //                     // Compuesto sellador
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[5], resultadoExamenSensorial.EnvaseInterno, 5, false, Alignment.center, false);
        //                     
        //                     // Compacidad
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[6], FormatearResultadoNumerico(resultado.Compacidad1), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[7], FormatearResultadoNumerico(resultado.Compacidad2), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[8], FormatearResultadoNumerico(resultado.Compacidad3), 5, false, Alignment.center, false);
        //                     
        //                     // Penetracion
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[9], FormatearResultadoNumerico(resultado.PenetracionDeGanchoDeCuerpo1), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[10], FormatearResultadoNumerico(resultado.PenetracionDeGanchoDeCuerpo2), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[11], FormatearResultadoNumerico(resultado.PenetracionDeGanchoDeCuerpo3), 5, false, Alignment.center, false);
        //                     
        //                     // Traslape
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[12], FormatearResultadoNumerico(resultado.Traslape1), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[13], FormatearResultadoNumerico(resultado.Traslape2), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[14], FormatearResultadoNumerico(resultado.Traslape3), 5, false, Alignment.center, false);
        //                     
        //                     // Traslape teorico
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[15], RedondearValorCustom(FormatearResultadoNumerico(resultado.Traslapem1)), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[16], RedondearValorCustom(FormatearResultadoNumerico(resultado.Traslapem2)), 5, false, Alignment.center, false);
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[17], RedondearValorCustom(FormatearResultadoNumerico(resultado.Traslapem3)), 5, false, Alignment.center, false);
        //                     
        //                     // Planchados
        //                     
        //                     //TODO: falta el resultado o en todo caso preguntar como calcular
        //                     
        //                     // Arrugas
        //                     
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[19], resultado.Arrugas, 5, false, Alignment.center, false);
        //                     
        //                     
        //                 }
        //                 else
        //                 {
        //                     FormatTableCell(tabla.Rows[inicioVias + j].Cells[4], "N/A", 6, false, Alignment.center, false);
        //                 }
        //             }
        //             
        //         }
        //         
        //     }
        //     
        //     tabla.Rows.Last().Remove();
        //
        //     // Guardar
        //
        //     document.InsertTable(tabla);
        //     
        // }

//         public static void CrearTablaLaboratorioMuestrasDirimentes(DocX document, List<Model.CodigoVia> codigoVias, SqlRepository repository)
//         {
//             
//             List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentesMB = repository.ObtenerMuestrasLaboratorioDirimente<Model.UspGetMuestraLaboratorioDirimente>(NumOs, 1).ToList();
//             List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentesFS = repository.ObtenerMuestrasLaboratorioDirimenteFS<Model.UspGetMuestraLaboratorioDirimente>(NumOs).ToList();
//             List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentesFQ = repository.ObtenerMuestrasLaboratorioDirimente<Model.UspGetMuestraLaboratorioDirimente>(NumOs, 2).ToList();
//
//             var resultado = DividirMuestraLaboratorioFQ(muestraLaboratorioDirimentesFQ);
//
//             List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentesFQHistamina = resultado.histamina;
//             List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentesFQMetalesPesados =
//                 resultado.metalesPesados;
//                 
//             // Por defecto, asumimos que no hay un punto de corte.
//             int indiceDeCorte = -1;
//
//             for (int i = 1; i < muestraLaboratorioDirimentesFQ.Count; i++)
//             {
//                 int numeroAnterior = int.Parse(muestraLaboratorioDirimentesFQ[i - 1].CodInterno.Substring(1));
//                 int numeroActual = int.Parse(muestraLaboratorioDirimentesFQ[i].CodInterno.Substring(1));
//
//                 // Si el número actual es menor o igual que el anterior, ¡hemos encontrado el reinicio!
//                 if (numeroActual <= numeroAnterior)
//                 {
//                     indiceDeCorte = i; // Guardamos el índice del primer elemento de la segunda lista
//                     break; // Salimos del bucle porque ya encontramos el punto que buscábamos
//                 }
//             }
//
//             if (indiceDeCorte != -1)
//             {
//                 muestraLaboratorioDirimentesFQHistamina = muestraLaboratorioDirimentesFQ.Take(indiceDeCorte).ToList();
//                 muestraLaboratorioDirimentesFQMetalesPesados =
//                     muestraLaboratorioDirimentesFQ.Skip(indiceDeCorte).ToList();
//             }
//             else
//             {
//                 muestraLaboratorioDirimentesFQHistamina = muestraLaboratorioDirimentesFQ;
//             }
//                 
//             
//             int cabeceraFilas = 2;
//             int cabeceraColumnas = 7;
//
//             int tablaFilas = cabeceraFilas + ( codigoVias.Count * 2 );
//             int tablaColumnas = cabeceraColumnas;
//
//             Table tabla = document.AddTable(tablaFilas, tablaColumnas);
//             tabla.Alignment = Alignment.left;
//
//             // Ancho de columnas
//
//             int[] columnWidths = { 30, 110, 75, 75, 75, 75, 75 };
//
//             for (int i = 0; i < tablaColumnas; i++)
//             {
//                 if (i <= columnWidths.Length - 1)
//                 {
//                     tabla.SetColumnWidth(i, columnWidths[i]);
//                 }
//             }
//
//             tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
//             
//             // Titulo
//             
//             FormatTableCell(tabla.Rows[0].Cells[0], "PRECINTOS ASIGNADOS A LAS MUESTRAS PARA LABORATORIO Y MUESTRAS DIRIMENTES", 7, true, Alignment.center);
//             
//             // Aux
//             
//             FormatTableCell(tabla.Rows[1].Cells[0], "M", 7, true, Alignment.center);
//
//             // M
//             
//             FormatTableCell(tabla.Rows[1].Cells[1], "", 7, true, Alignment.center);
//
//             // Microbiologico
//
//             FormatTableCell(tabla.Rows[1].Cells[2], "Microbiologico", 7, true, Alignment.center);
//
//             // Fisicosensorial
//
//             FormatTableCell(tabla.Rows[1].Cells[3], "Fisicosensorial", 7, true, Alignment.center);
//
//             // Cierre
//
//             FormatTableCell(tabla.Rows[1].Cells[4], "Cierre", 7, true, Alignment.center);
//
//             // Histamina
//
//             FormatTableCell(tabla.Rows[1].Cells[5], "Histamina", 7, true, Alignment.center);
//
//             // Metales pesados
//
//             FormatTableCell(tabla.Rows[1].Cells[6], "Metales Pesados", 7, true, Alignment.center);
//
//             // Codigo Vias (Dinámicas)
//             
//             for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias += 2)
//             {
//                 string ensayoLabel = codigoVias[i].CodigoInterno;
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[0], ensayoLabel, 7, true, Alignment.center, false);
//                 
//                 // Merge de las columnas
//                 
//                 tabla.MergeCellsInColumn(0, inicioVias, inicioVias + 1);
//                 
//                 // Muestras
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[1], "Muestras para laboratorio", 6, false, Alignment.right, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[1], "Muestras dirimentes", 6, false, Alignment.right, false);
//                 
//                 // Microbiologia
//                 
//                 var muestraMb = muestraLaboratorioDirimentesMB.FirstOrDefault(x => x.CodInterno == ensayoLabel);
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[2], muestraMb.MuestraLaboratorio, 5, false, Alignment.center, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[2], muestraMb.MuestraDirimente, 5, false, Alignment.center, false);
//                 
//                 // Fisico sensorial y cierre
//                 
//                 var muestraFs = muestraLaboratorioDirimentesFS.FirstOrDefault(x => x.CodInterno == ensayoLabel);
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[3], muestraFs.MuestraLaboratorio, 5, false, Alignment.center, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[3], muestraFs.MuestraDirimente, 5, false, Alignment.center, false);
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[4], muestraFs.MuestraLaboratorioCierre, 5, false, Alignment.center, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[4], muestraFs.MuestraDirimenteCierre, 5, false, Alignment.center, false);
//                 
//                 // Histamina y metales pesados
//                 
//                 var muestraFq = muestraLaboratorioDirimentesFS.FirstOrDefault(x => x.CodInterno == ensayoLabel);
//                 
//                 var muestraFqHistamina = muestraLaboratorioDirimentesFQHistamina.FirstOrDefault(x => x.CodInterno == ensayoLabel);
//                 var muestraFqMetalesPesados = muestraLaboratorioDirimentesFQMetalesPesados.FirstOrDefault(x => x.CodInterno == ensayoLabel);
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[5], muestraFqHistamina.MuestraLaboratorio, 5, false, Alignment.center, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[5], muestraFqHistamina.MuestraDirimente, 5, false, Alignment.center, false);
//                 
//                 FormatTableCell(tabla.Rows[inicioVias].Cells[6], muestraFqMetalesPesados.MuestraLaboratorio, 5, false, Alignment.center, false);
//                 FormatTableCell(tabla.Rows[inicioVias + 1].Cells[6], muestraFqMetalesPesados.MuestraDirimente, 5, false, Alignment.center, false);
//                 
//             }
//             
//             tabla.Rows[1].MergeCells(0, 1);
//
//             // Descripción
//
//             tabla.InsertRow();
//             
//             // No mover el tabulado a continuación, es necesario para el formato correcto de la tabla.
//             
//             string textNotas = @"Notas:
// - Para la extracción de muestras de análisis de Microbiológico (MB), Fisicosensorial (FS), Histamina (HS), Cierre (C) y Metales Pesados (MP) se aplicó la Norma Técnica Peruana 700.002 ""Lineamientos y 
//   Procedimientos de Muestreo del Pescado y Productos Pesqueros para Inspección"" 2ª Edición, del 04 de julio de 2012. Nivel de Inspección I, NCA = 6,5. Se procedió a la toma de muestras con fines de ensayo 
//   (Muestras  para laboratorio) y muestras dirimentes, en igual cantidad, bajo la misma metodología y con precinto propio. Las muestras dirimentes serán conservadas bajo condiciones adecuadas de 
//   almacenamiento y custodia por un período de 180 días, conforme a los procedimientos internos vigentes.";
//
//             AgregarDescripcion(tabla, textNotas, Alignment.both);
//
//             document.InsertTable(tabla);
//         }
        
        
        // public static void CrearTablaIndicadoresParasitologicos(DocX document, SqlRepository repository)
        // {
        //     List<Model.CodigoViaFisicoSensorial> codigoVias = repository.ObtenerCodigoViasFisicoSensorial<Model.CodigoViaFisicoSensorial>(IdOT, 1).ToList();
        //     
        //     codigoVias = codigoVias
        //         // Ordenamos por la parte numérica del CodigoInterno
        //         .OrderBy(c => int.Parse(c.Codigos.Substring(1)))
        //         .ThenBy(c => c.Codigos) // Mantenemos el segundo nivel de orden si lo necesitas
        //         .ToList();
        //     
        //     int cabeceraFilas = 3;
        //     int cabeceraColumnas = 5;
        //
        //     int tablaFilas = cabeceraFilas + codigoVias.Count;
        //     int tablaColumnas = cabeceraColumnas;
        //
        //     Table tabla = document.AddTable(tablaFilas, tablaColumnas);
        //     tabla.Alignment = Alignment.left;
        //     
        //     // Agregar titulo
        //     
        //     AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS FÍSICO-SENSORIALES) - PERÚ Y OTROS PAÍSES. NUMERAL 1.2.5");
        //
        //     // Ancho de columnas
        //
        //     int[] columnWidths = { 80, 80, 150, 150, 50 };
        //
        //     for (int i = 0; i < tablaColumnas; i++)
        //     {
        //         if (i <= columnWidths.Length - 1)
        //         {
        //             tabla.SetColumnWidth(i, columnWidths[i]);
        //         }
        //     }
        //
        //     tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
        //     
        //     // Encabezados
        //     
        //     tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
        //     FormatTableCell(tabla.Rows[1].Cells[0], "INDICADORES PARASITOLOGICOS", 7, true, Alignment.center);
        //     
        //     // Codigo
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[0], "CÓDIGO", 7, true, Alignment.center);
        //
        //     // Vias (n)
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[1], "VÍAS (n)", 7, true, Alignment.center);
        //
        //     // Plan de evaluación
        //
        //     FormatTableCell(tabla.Rows[2].Cells[2], "PLAN DE EVALUACIÓN", 7, true, Alignment.center);
        //
        //     // Resultados
        //
        //     FormatTableCell(tabla.Rows[2].Cells[3], "RESULTADOS", 7, true, Alignment.center);
        //
        //     // Conclusión
        //
        //     FormatTableCell(tabla.Rows[2].Cells[4], "CONCLUSION", 7, true, Alignment.center);
        //     
        //     // Codigo Vias (Dinámicas)
        //     
        //     for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias++)
        //     {
        //         string codigoInterno = codigoVias[i].Codigos;
        //         string rangoVias = codigoVias[i].RangoVias;
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[0], codigoInterno, 6, true, Alignment.center, false);
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[1], rangoVias, 6, false, Alignment.center, false);
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[2], "Ausencia de parásitos visibles", 6, false, Alignment.center, false);
        //     }
        //     
        //     // Si no hay descripción eliminamos la ultima fila
        //     
        //     tabla.Rows.Last().Remove();
        //
        //     document.InsertTable(tabla);
        // }
        
        
        // public static void CrearTablaDeterminacionPresionVacio(DocX document, SqlRepository repository)
        // {
        //     List<Model.UspGetTablaExamenesSensoriales> tablaResultados =
        //         repository.ObtenerTablaExamenesSensorial<Model.UspGetTablaExamenesSensoriales>(IdOT, 1).ToList();
        //     
        //     List<Model.CodigoViaFisicoSensorial> codigoVias = repository.ObtenerCodigoViasFisicoSensorial<Model.CodigoViaFisicoSensorial>(IdOT, 1).ToList();
        //     
        //     codigoVias = codigoVias
        //         // Ordenamos por la parte numérica del CodigoInterno
        //         .OrderBy(c => int.Parse(c.Codigos.Substring(1)))
        //         .ThenBy(c => c.Codigos) // Mantenemos el segundo nivel de orden si lo necesitas
        //         .ToList();
        //     
        //     int cabeceraFilas = 3;
        //     int cabeceraColumnas = 6;
        //
        //     int tablaFilas = cabeceraFilas + (codigoVias.Count * 5); // El usuario indica que siempre van a tener 5 filas
        //     int tablaColumnas = cabeceraColumnas;
        //     
        //     Table tabla = document.AddTable(tablaFilas, tablaColumnas);
        //     tabla.Alignment = Alignment.left;
        //     
        //     // Titulo
        //     
        //     AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS FÍSICO-SENSORIALES) - PERÚ Y OTROS PAÍSES. NUMERAL 1.2.7");
        //     
        //     // Encabezado
        //     
        //     // Ancho de columnas
        //     
        //     int[] columnWidths = { 25, 25, 20, 300, 100, 60 };
        //     
        //     for (int i = 0; i < tablaColumnas; i++)
        //     {
        //         tabla.SetColumnWidth(i, columnWidths[i]);
        //     }
        //     
        //     // Combinar filas
        //     
        //     // tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
        //     // tabla.MergeCellsInColumn(1, 2, cabeceraFilas - 1);
        //     // tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
        //     // tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
        //     // tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
        //     // tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1); // Ùltima columnna
        //     
        //     tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
        //     
        //     // Requisitos
        //     
        //     tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
        //     FormatTableCell(tabla.Rows[1].Cells[0], "REQUISITOS PARA LA DETERMINACION DE VACIO", 7, true, Alignment.center, true);
        //     
        //     // Código
        //
        //     tabla.Rows[2].Height = 80;
        //     FormatTableCell(tabla.Rows[2].Cells[0], "CODIGO", 7, true, Alignment.center, true, TextDirection.btLr);
        //     
        //     // Vías (n)
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[1], "VIAS (n)", 7, true, Alignment.center);
        //     
        //     // Tolerancia
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[2], "TOLERANCIA", 7, true, Alignment.center, true, TextDirection.btLr);
        //     
        //     // Requisitos
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[3], "REQUISITOS", 7, true, Alignment.center, true);
        //     
        //     // Resultados
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[4], "RESULTADOS (mm Hg)", 7, true, Alignment.center, true);
        //     
        //     // Conclusion
        //     
        //     FormatTableCell(tabla.Rows[2].Cells[5], "CONCLUSIÓN", 7, true, Alignment.center, true);
        //     
        //     // Muestras
        //     
        //     for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias += 5)
        //     {
        //         
        //         string codigoInterno = codigoVias[i].Codigos;
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[0], codigoInterno, 5, true, Alignment.center, false);
        //         
        //         tabla.MergeCellsInColumn(0, inicioVias, inicioVias + 4);
        //         tabla.MergeCellsInColumn(2, inicioVias, inicioVias + 4);
        //         tabla.MergeCellsInColumn(3, inicioVias, inicioVias + 4);
        //         tabla.MergeCellsInColumn(5, inicioVias, inicioVias + 4);
        //         
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[2], "0", 6, false, Alignment.center, false);
        //
        //         string requisitosTexto =
        //             "- El vacio mínimo en envases de hojalata cilíndricos con capacidad de hasta 370ml deberá ser no menor a 76.2mmHg (3 pulgadas de Hg)." +
        //             "\n- Para los envases rectangulares, el vacio mínimo deberá ser de 40mm Hg (1.6 pulgadas de Hg)." +
        //             "\n- El vacio mínimo en envases de vidrio, deberá ser no menor de 140mm Hg (5.5 pulgadas de Hg). " +
        //             "\n  El vacio mínimo en envases de hojalata con capacidad mayor a 370ml hasta 500ml deberá ser no menor a 150mm Hg (6 pulgadas de hg).";
        //         
        //         FormatTableCell(tabla.Rows[inicioVias].Cells[3], requisitosTexto, 5, false, Alignment.left, false);
        //         
        //         // Vias y Resultados
        //
        //         for (int j = 0; j < 5; j++) 
        //         {
        //             // Vias
        //             FormatTableCell(tabla.Rows[inicioVias + j].Cells[1], (j + 1).ToString(), 6, false, Alignment.center, false);
        //             // Resulados
        //             string resultado = tablaResultados.Find(r => r.Codigos == codigoInterno && r.Vias == j + 1).PresionDeVacioMmHg;
        //             if (resultado is null) resultado = "0";
        //             FormatTableCell(tabla.Rows[inicioVias + j].Cells[4], resultado, 6, false, Alignment.center, false);
        //         }
        //         
        //     }
        //     
        //     tabla.Rows.Last().Remove();
        //
        //     // Guardar
        //
        //     document.InsertTable(tabla);
        //     
        // }
        
        public static void CrearTablaHistamina(DocX document, SqlRepository repository)
        {
            List<Model.UspGetReporteInspeccionTablaHistamina> tablaResultados =
                repository.ObtenerTablaHistamina<Model.UspGetReporteInspeccionTablaHistamina>(IdOT, 3).ToList();
            
            // Obtener los codigos de vias de la tabla de resultados 

            List<string> codigoVias = tablaResultados
                .Select(x => x.CodPrecinto) // Selecciona solo la propiedad CodPrecinto
                .Distinct() // Elimina duplicados para obtener valores únicos
                .OrderBy(x => x) // Ordena los valores alfabéticamente de forma ascendente
                .ToList(); // Convierte el resultado en una List<string>
            
            int NUMERO_VIAS = 9; // El usuario indica que siempre van a tener 5 filas
            
            int cabeceraFilas = 4;
            int cabeceraColumnas = 7;

            int tablaFilas = cabeceraFilas + (codigoVias.Count * NUMERO_VIAS);
            int tablaColumnas = cabeceraColumnas;
            
            Table tabla = document.AddTable(tablaFilas, tablaColumnas);
            tabla.Alignment = Alignment.left;
            
            // Titulo
            
            AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS QUÍMICAS) - PERÚ Y OTROS PAÍSES:  NUMERAL 1.2.9.1");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 20, 25, 20, 150, 150, 100, 60 };
            
            for (int i = 0; i < tablaColumnas; i++)
            {
                tabla.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Combinar filas
            
            tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(1, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(5, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(6, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1); // Ùltima columnna
            
            tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
            
            // Histamina
            
            tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
            FormatTableCell(tabla.Rows[1].Cells[0], "HISTAMINA", 7, true, Alignment.center, true);
            
            // Código

            tabla.Rows[2].Height = 80;
            tabla.Rows[3].Height = 20;
            FormatTableCell(tabla.Rows[2].Cells[0], "CODIGO", 7, true, Alignment.center, true, TextDirection.btLr);
            
            // Vías (n)
            
            FormatTableCell(tabla.Rows[2].Cells[1], "VIAS (n)", 7, true, Alignment.center);
            
            // Tolerancia
            
            FormatTableCell(tabla.Rows[2].Cells[2], "TOLERANCIA (c)", 7, true, Alignment.center, true, TextDirection.btLr);
            
            // Limites de tolerancia
            
            tabla.Rows[2].MergeCells(3, 4);
            
            var celda = tabla.Rows[2].Cells[3];
            var parrafo = celda.Paragraphs.First();
            
            parrafo.Append("LÍMITES DE TOLERANCIA (ppm)").FontSize(6).Font("Calibri (Cuerpo)").Bold();
            parrafo.Append("1").Script(Script.superscript).FontSize(8).Font("Calibri (Cuerpo)").Bold();

            parrafo.Alignment = Alignment.center;
            celda.FillColor = Color.FromArgb(234, 241, 221);
            celda.VerticalAlignment = VerticalAlignment.Center;
            
            // FormatTableCell(tabla.Rows[2].Cells[3], "LIMITES DE TOLERANCIA (ppm)1", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[3], "m", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[4], "M", 7, true, Alignment.center);
            
            // Resultados
            
            FormatTableCell(tabla.Rows[2].Cells[4], "RESULTADOS (mm Hg)", 7, true, Alignment.center);
            
            // Conclusion
            
            FormatTableCell(tabla.Rows[2].Cells[5], "CONCLUSIÓN", 7, true, Alignment.center);
            
            // Muestras
            
            for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias += NUMERO_VIAS)
            {
                
                string codigoInterno = codigoVias[i];
                FormatTableCell(tabla.Rows[inicioVias].Cells[0], codigoInterno, 5, true, Alignment.center, false);
                
                tabla.MergeCellsInColumn(0, inicioVias, inicioVias + NUMERO_VIAS - 1);
                tabla.MergeCellsInColumn(2, inicioVias, inicioVias + NUMERO_VIAS - 1);
                tabla.MergeCellsInColumn(3, inicioVias, inicioVias + NUMERO_VIAS - 1);
                tabla.MergeCellsInColumn(4, inicioVias, inicioVias + NUMERO_VIAS - 1);
                tabla.MergeCellsInColumn(6, inicioVias, inicioVias + NUMERO_VIAS - 1);

                FormatTableCell(tabla.Rows[inicioVias].Cells[2], "2", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[inicioVias].Cells[3], "100", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[inicioVias].Cells[4], "200", 6, false, Alignment.center, false);
                
                // Vias

                for (int j = 0; j < NUMERO_VIAS; j++) 
                {
                    // Via
                    FormatTableCell(tabla.Rows[inicioVias + j].Cells[1], (j + 1).ToString(), 6, false, Alignment.center, false);
                    
                    // Resultados
                    
                    string resultado = tablaResultados.Find(r => r.CodPrecinto == codigoInterno && r.NroVia == j + 1)?.Resultado ?? "0";
                    FormatTableCell(tabla.Rows[inicioVias + j].Cells[5], resultado, 6, false, Alignment.center, false);
                }
                
            }

            AgregarDescripcion(tabla, "1ppm = 1mg/kg");
            
            // Guardar

            document.InsertTable(tabla);
            
        }
        
        public static void CrearTablaMetalesPesados(DocX document, SqlRepository repository)
        {
            List<Model.Ensayo> analisis = repository.ObtenerEnsayos<Model.Ensayo>(IdOT, 1, 3).ToList();
            List<Model.ViaResultado> tablaResultados = repository.ViasResultados<Model.ViaResultado>(IdOT, 3).ToList();
            
            List<string> codigoVias = tablaResultados
                .Select(x => x.CodPrecinto) // Selecciona solo la propiedad CodPrecinto
                .Distinct() // Elimina duplicados para obtener valores únicos
                .OrderBy(x => x) // Ordena los valores alfabéticamente de forma ascendente
                .ToList(); // Convierte el resultado en una List<string>
            
            // Obtener los ensayos a excepción de histamina
            
            analisis = analisis.Where(e => e.Analisis != "Histamina").ToList();
            
            int cabeceraFilas = 3;
            int cabeceraColumnas = 7;

            int tablaFilas = cabeceraFilas + analisis.Count;
            int tablaColumnas = cabeceraColumnas;
            
            Table tabla = document.AddTable(tablaFilas, tablaColumnas);
            tabla.Alignment = Alignment.left;
            
            // Titulo
            
            AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS QUÍMICAS) - PERÚ Y OTROS PAÍSES:  NUMERAL 1.3.2.1-TABLA N°06");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] anchoColumnas = { 60, 80, 80, 100, 60, 60, 60 };
            
            for (int i = 0; i < tablaColumnas; i++)
            {
                tabla.SetColumnWidth(i, anchoColumnas[i]);
            }
            
            // Combinar filas
            
            // tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(1, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(2, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(5, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(6, 2, cabeceraFilas - 1);
            // tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1); // Ùltima columnna
            
            tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
            
            // Metales pesados
            
            tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
            FormatTableCell(tabla.Rows[1].Cells[0], "METALES PESADOS", 7, true, Alignment.center);
            
            // Analisis

            FormatTableCell(tabla.Rows[2].Cells[0], "ANÁLISIS", 7, true, Alignment.center);
            
            // Vías (n)
            
            FormatTableCell(tabla.Rows[2].Cells[1], "CÓDIGOS", 7, true, Alignment.center);
            
            // Vias
            
            FormatTableCell(tabla.Rows[2].Cells[2], "VÍAS", 7, true, Alignment.center);
            
            // Contenido maximo
            
            FormatTableCell(tabla.Rows[2].Cells[3], "CONTENIDO MAXIMO (MG/KG PESO FRESCO)", 7, true, Alignment.center);
            
            // Unidades
            
            FormatTableCell(tabla.Rows[2].Cells[4], "UNIDADES", 7, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(tabla.Rows[2].Cells[5], "RESULTADO", 7, true, Alignment.center);
            
            // Conclusion
            
            FormatTableCell(tabla.Rows[2].Cells[6], "CONCLUSION", 7, true, Alignment.center);
            
            
            tabla.MergeCellsInColumn(1, cabeceraFilas, tabla.RowCount - 1);
            tabla.MergeCellsInColumn(2, cabeceraFilas, tabla.RowCount - 1);
            tabla.MergeCellsInColumn(tabla.ColumnCount - 1, cabeceraFilas, tabla.RowCount - 1);
            
            // Codigo
            
            string codigoConcatenado = string.Join("\n ", codigoVias.Select(via => via));
                
            FormatTableCell(tabla.Rows[3].Cells[1], codigoConcatenado, 6, true, Alignment.center, false);
            FormatTableCell(tabla.Rows[3].Cells[2], "1", 6, false, Alignment.center, false);
            
            // Muestras
            
            for (int i = 0, inicioAnalisis = cabeceraFilas; i < analisis.Count; i++, inicioAnalisis++)
            {
                
                var celda = tabla.Rows[inicioAnalisis];
                celda.Height = 10;
                
                string analisisLabel = analisis[i].Analisis;
                
                FormatTableCell(tabla.Rows[inicioAnalisis].Cells[0], analisisLabel, 5, false, Alignment.center, false);
                
                // Solo mostramos el primer codigo de via, ya que es donde se muestra en el Syslab
                
                string resultado = tablaResultados.Find(r => r.CodPrecinto == "M1" && r.Muestra == "n1" && r.IdAnalisis == analisis[i].IdAnalisis).Resultado;
                string unidadMedida = tablaResultados.Find(r => r.CodPrecinto == "M1" && r.Muestra == "n1" && r.IdAnalisis == analisis[i].IdAnalisis).UnidMedida;

                if (resultado is null) resultado = "0";
                
                if (analisisLabel == "Estaño")
                {
                    // Estaño tiene un contenido maximo de 200 mg/kg (valor fijo)
                    FormatTableCell(tabla.Rows[inicioAnalisis].Cells[3], "200", 5, false, Alignment.center, false);
                }

                FormatTableCell(tabla.Rows[inicioAnalisis].Cells[4], unidadMedida, 5, false, Alignment.center, false);
                
                FormatTableCell(tabla.Rows[inicioAnalisis].Cells[5], resultado, 5, false, Alignment.center, false);
                

            }
            
            tabla.Rows.Last().Remove();

            // Guardar

            document.InsertTable(tabla);
            
        }
        
        public static void CrearTablaExtensionesSensoriales(DocX document, List<Model.CodigoVia> codigoVias, SqlRepository repository)
        {
            
            List<Model.UspGetTablaExamenesSensoriales> tablaResultados =
                repository.ObtenerTablaExamenesSensorial<Model.UspGetTablaExamenesSensoriales>(IdOT, 1).ToList();

            if (tablaResultados.Count == 0)
            {
                throw new Exception("No se encontraron resultados para la tabla de extensiones sensoriales.");
            }
            
            int cabeceraFilas = 4;
            int cabeceraColumnas = 11;

            int tablaFilas = cabeceraFilas + codigoVias.Count;
            int tablaColumnas = cabeceraColumnas;
            
            Table tabla = document.AddTable(tablaFilas, tablaColumnas);
            tabla.Alignment = Alignment.left;
            
            // Titulo
            
            AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS FÍSICO-SENSORIALES). ÍTEMS: 4.1.1; 4.1.8; 4.1.9; 4.1.10 Y 4.1.11");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 50, 50, 30, 30, 40, 40, 55, 55, 55, 55, 55 };
            
            for (int i = 0; i < tablaColumnas; i++)
            {
                tabla.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Combinar filas
            
            tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(1, 2, cabeceraFilas - 1);
            
            tabla.MergeCellsInColumn(6, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(7, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(8, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(9, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1); // Ùltima columnna
            
            tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
            
            // Examenes sensoriales
            
            tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
            FormatTableCell(tabla.Rows[1].Cells[0], "EXAMENES SENSORIALES", 7, true, Alignment.center);
            
            // Codigo

            FormatTableCell(tabla.Rows[2].Cells[0], "CÓDIGO", 7, true, Alignment.center);
            
            // Vías (n)
            
            FormatTableCell(tabla.Rows[2].Cells[1], "VÍAS (n)", 7, true, Alignment.center);
            
            // Numero de aceptacion

            tabla.Rows[2].MergeCells(2, 3);
            FormatTableCell(tabla.Rows[2].Cells[2], "NÚMERO DE ACEPTACIÓN", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[2], "N°", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[3], "(c)*", 7, true, Alignment.center);
            
            // Aspecto
            
            tabla.Rows[2].MergeCells(3, 4);
            FormatTableCell(tabla.Rows[2].Cells[3], "ASPECTO", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[4], "EXTERIOR", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[5], "INTERIOR", 7, true, Alignment.center);
            
            // Olor
            
            FormatTableCell(tabla.Rows[2].Cells[4], "OLOR", 7, true, Alignment.center);
            
            // Color
            
            FormatTableCell(tabla.Rows[2].Cells[5], "COLOR", 7, true, Alignment.center);
            
            // Sabor
            
            FormatTableCell(tabla.Rows[2].Cells[6], "SABOR", 7, true, Alignment.center);
            
            // Textura
            
            FormatTableCell(tabla.Rows[2].Cells[7], "TEXTURA", 7, true, Alignment.center);
            
            // conclusion
            
            FormatTableCell(tabla.Rows[2].Cells[8], "CONCLUSION", 7, true, Alignment.center);
            
            
            // Muestras
            
            for (int i = 0, inicioVias = cabeceraFilas; i < codigoVias.Count; i++, inicioVias++)
            {
                string codigoInterno = codigoVias[i].CodigoInterno;
                FormatTableCell(tabla.Rows[inicioVias].Cells[0], codigoInterno, 5, true, Alignment.center, false);
                
                // Vias

                var resultado = tablaResultados.FirstOrDefault(x => x.Codigos == codigoVias[i].CodigoInterno);

                if (resultado != null)
                {
                    
                    // TODO: Hacer una revisión para verificar si existe un no conforme en las vias
                    
                    // Vias
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[1], resultado.RangoVias, 6, false, Alignment.center, false);
                    
                    // Exterior
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[4], resultado.EnvaseExterno, 6, false, Alignment.center, false);
                    
                    // Interior
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[5], resultado.EnvaseInterno, 6, false, Alignment.center, false);
                    
                    // Olor
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[6], resultado.Olor, 6, false, Alignment.center, false);
                    
                    // Color
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[7], resultado.Color, 6, false, Alignment.center, false);
                    
                    // Sabor
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[8], resultado.Sabor, 6, false, Alignment.center, false);
                    
                    // Textura
                    
                    FormatTableCell(tabla.Rows[inicioVias].Cells[9], resultado.Textura, 6, false, Alignment.center, false);
                    
                }
                
            }
            
            AgregarDescripcion( tabla, "(*) El paréntesis en el número de aceptación (c) indica el número de aceptación para descomposición \n Nota: M1(n10): Sabor/ No Conforme/ Ligeramente picante asociado a histamina");
            
            // Guardar

            document.InsertTable(tabla);
            
        }
        
        // public static void CreateTableA(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, List<Model.CodigoVia> codigoVias, int iTable, int numVias)
        // {
        //     int headerRows = 3;
        //     int headerColumns = 6;
        //
        //     int tableRows = headerRows + ensayos.Count;
        //     int tableColumns = headerColumns + vias.Count;
        //     
        //     Table table = document.AddTable(tableRows, tableColumns);
        //     table.Alignment = Alignment.center;
        //     
        //     // Encabezado
        //     
        //     // Ancho de columnas
        //     
        //     int[] columnWidths = { 100, 40, 40, 45, 45 };
        //     
        //     for (int i = 0; i < tableColumns; i++)
        //     {
        //
        //         if (i <= columnWidths.Length - 1)
        //         {
        //             table.SetColumnWidth(i, columnWidths[i]);
        //         }
        //         
        //         // Vías dinámicas
        //
        //         if (i >= 5 && i < tableColumns - 1)
        //         {
        //             table.SetColumnWidth(i, 35);
        //         }
        //         
        //         // Última columna
        //
        //         if (i == tableColumns - 1)
        //         {
        //             table.SetColumnWidth(i, 50);
        //         }
        //         
        //     }
        //     
        //     // Combinas filas
        //     
        //     table.MergeCellsInColumn(0, 0, headerRows - 1);
        //     
        //     table.MergeCellsInColumn(1, 0, 1);
        //     table.MergeCellsInColumn(2, 0, 1);
        //     
        //     table.MergeCellsInColumn(3, 0, 1);
        //     table.MergeCellsInColumn(4, 0, 1);
        //     
        //     table.MergeCellsInColumn(table.ColumnCount - 1, 0, headerRows - 1);
        //     
        //     // Microorganismo
        //     
        //     FormatTableCell(table.Rows[0].Cells[0], "MICROORGANISMO", 3, true, Alignment.center);
        //     
        //     // Plan de evaluación
        //     
        //     table.Rows[0].MergeCells(1, 2);
        //     table.Rows[1].MergeCells(1, 2);
        //     FormatTableCell(table.Rows[0].Cells[1], "PLAN DE EVALUACIÓN", 3, true, Alignment.center);
        //     
        //     FormatTableCell(table.Rows[2].Cells[1], "n", 3, true, Alignment.center);
        //     FormatTableCell(table.Rows[2].Cells[2], "c", 3, true, Alignment.center);
        //     
        //     // Limites
        //     
        //     table.Rows[0].MergeCells(2, 3);
        //     table.Rows[1].MergeCells(2, 3);
        //     FormatTableCell(table.Rows[0].Cells[2], "LIMITES", 3, true, Alignment.center);
        //     
        //     FormatTableCell(table.Rows[2].Cells[3], "m", 3, true, Alignment.center);
        //     FormatTableCell(table.Rows[2].Cells[4], "M", 3, true, Alignment.center);
        //     
        //     // Distribución de muestras
        //     
        //     table.Rows[0].MergeCells(3, 3 + vias.Count - 1);
        //     FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", 3, true, Alignment.center);
        //     
        //     // Código de Vías (Dinámicas)
        //     
        //     AgruparYFormatearVias(table, vias, codigoVias, 1, 5);
        //         
        //     // for (int i = 0, aux = 0; i < vias.Count;)
        //     // {
        //     //     string currentVia = vias[i].Presentacion;
        //     //     int startCol = 5 + i - aux;
        //     //     int j = i + 1;
        //     //
        //     //     // Buscar cuántas 'vias' consecutivas tienen la misma presentación
        //     //     while (j < vias.Count && vias[j].Presentacion == currentVia)
        //     //     {
        //     //         j++;
        //     //         aux++;
        //     //     }
        //     //
        //     //     int endCol =  5 + j - (i == 0 ? 1 : aux);
        //     //
        //     //     // Formatear y/o fusionar celdas según cantidad de columnas iguales
        //     //     if (j - i > 1)
        //     //     {
        //     //         table.Rows[1].MergeCells(startCol - 2, endCol - 2); 
        //     //         // table.Rows[2].MergeCells(startCol, endCol);
        //     //     }
        //     //
        //     //
        //     //     var productoCodigo = codigoVias.Find(x => x.CodigoInterno == currentVia).ProductoCodigo;
        //     //     FormatTableCell(table.Rows[1].Cells[startCol - 2], productoCodigo, 3, true, Alignment.center);
        //     //     
        //     //     // FormatTableCell(table.Rows[2].Cells[startCol], currentVia, 4, true, Alignment.center); // Codigo
        //     //
        //     //     i = j; // Saltar al siguiente grupo
        //     // }
        //     
        //     
        //     // Vias (Dinámicas)
        //         
        //     for (int i = 0, iCellIndex= 5; i < vias.Count; i++, iCellIndex++)
        //     {
        //         FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, 4, true, Alignment.center);
        //     }
        //     
        //     // Ensayos
        //     
        //     for (int i = 0; i < ensayos.Count; i++)
        //     {
        //
        //         string ensayoLabel = ensayos[i].Analisis;
        //         
        //         FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 4, false, Alignment.left);
        //         FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), 4, false, Alignment.center);
        //
        //         int j = 5;
        //         
        //         // Resultado por cada ensayo
        //         
        //         var resultados = 
        //             viaResultados
        //                 .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis).ToList()
        //                 .GetRange((iTable * MAX_VIAS), Math.Min(MAX_VIAS, viaResultados.Count - iTable * MAX_VIAS));
        //
        //         foreach (var via in resultados)
        //         {
        //             
        //             // bool match = (ensayos[i].IdProducto == via.IdProducto && ensayos[i].IdAnalisis == via.IdAnalisis && via.CodigoInterno ==);
        //
        //             if (true)
        //             {
        //                 FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 4, true, Alignment.center);
        //             }
        //
        //             j++;
        //
        //         }
        //             
        //     }
        //     
        //     // Conclusión
        //
        //     FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", 3, true, Alignment.center);
        //
        //     // Guardar
        //     
        //     document.InsertTable(table);
        // }
        
        private static void FormatTableCell(Cell cell, string text, int fontSize, bool isBold, Alignment alignment, bool setColor = true, TextDirection textDirection = TextDirection.right )
        {
            cell.Paragraphs[0].Append(text)
                .Font("Calibri (Cuerpo)")
                .Bold(isBold)
                .FontSize(fontSize)
                .Alignment = alignment;

            cell.VerticalAlignment = VerticalAlignment.Center;
            cell.TextDirection = textDirection;

            if (setColor)
            {
                cell.FillColor = Color.FromArgb(234, 241, 221);
            }
            
        }
        
        private static void AgruparYFormatearVias(Table table, List<Model.Via> vias, List<Model.CodigoVia> codigoVias, int rowIndex,
            int startColumnOffset)
        {
            for (int i = 0, aux = 0; i < vias.Count;)
            {
                string currentVia = vias[i].Presentacion;
                int startCol = startColumnOffset + i - aux;
                
                int j = i + 1;
        
                for (int k = 0; j < vias.Count && vias[j].Presentacion == currentVia; j++, k++)
                {
                    aux++;
                    // if (k == 0)
                    // {
                    //     aux++;
                    // }
                }
                
                int endCol =  startColumnOffset + j - (i == 0 ? 1 : aux - 1);
                
                if (j - i > 1)
                {
                    table.Rows[rowIndex].MergeCells(startCol - 2, endCol - 2);
                }
        
                var productoCodigo = codigoVias.Find(x => x.CodigoInterno == currentVia)?.ProductoCodigo ?? "";
                FormatTableCell(table.Rows[rowIndex].Cells[startCol - 2], productoCodigo, 8, true,
                    Alignment.center);
        
                i = j;
            }
        }
        
        public static void CrearTablaEsterilidadComercial(DocX document, List<Model.Ensayo> ensayos, List<Model.CodigoVia> codigoVias, List<Model.ViaResultado> viaResultados)
        {
            int cabeceraFilas = 4;
            int cabeceraColumnas = 12;

            int tablaFilas = cabeceraFilas + codigoVias.Count;
            int tablaColumnas = cabeceraColumnas;
            
            Table tabla = document.AddTable(tablaFilas, tablaColumnas);
            tabla.Alignment = Alignment.left;
            
            // Agregar titulo
            
            AgregarTitulo(tabla, "INSPECCIÓN DE LOTES POR MUESTREO (CARACTERÍSTICAS MICROBIOLÓGICAS) - PERÚ Y OTROS PAÍSES. NUMERAL 1.2.8");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] anchoColumnas = { 80, 22, 22, 60, 60, 35, 35, 35, 35, 35, 35, 60 };
            
            for (int i = 0; i < tablaColumnas; i++)
            {
                tabla.SetColumnWidth(i, anchoColumnas[i]);
            }
            
            // Combinar filas
            
            tabla.MergeCellsInColumn(0, 2, cabeceraFilas - 1);
            
            tabla.MergeCellsInColumn(3, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(4, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(5, 2, cabeceraFilas - 1);
            tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 2, cabeceraFilas - 1);
            
            tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1); // Titulo
            
            // Esterilidad comercial
            
            tabla.Rows[1].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
            FormatTableCell(tabla.Rows[1].Cells[0], "ESTERILIDAD COMERCIAL", 7, true, Alignment.center);
            
            // Analisis

            tabla.Rows[2].Height = 20;
            tabla.Rows[3].Height = 10;
            FormatTableCell(tabla.Rows[2].Cells[0], "ANALISIS", 7, true, Alignment.center);
            
            // Plan de evaluación
            
            tabla.Rows[2].MergeCells(1, 2);
            FormatTableCell(tabla.Rows[2].Cells[1], "PLAN DE EVALUACIÓN", 6, true, Alignment.center);
            
            FormatTableCell(tabla.Rows[3].Cells[1], "n", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[2], "c", 7, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(tabla.Rows[2].Cells[2], "ACEPTACIÓN", 7, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(tabla.Rows[2].Cells[3], "RECHAZO", 7, true, Alignment.center);
            
            // Código
            
            FormatTableCell(tabla.Rows[2].Cells[4], "CÓDIGO", 7, true, Alignment.center);
            
            // Número de Vías (Dinámicas)
            
            tabla.Rows[2].MergeCells(5, 9);
            FormatTableCell(tabla.Rows[2].Cells[5], "NÚMERO DE VÍAS", 7, true, Alignment.center);
            
            // Recordar que el usuario indico que las 5 vias son fijas, por lo que no es necesario hacer un bucle dinámico para las vias
            
            FormatTableCell(tabla.Rows[3].Cells[6], "n1", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[7], "n2", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[8], "n3", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[9], "n4", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[3].Cells[10], "n5", 7, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(tabla.Rows[2].Cells[tabla.Rows[2].Cells.Count - 1], "CONCLUSIÓN", 7, true, Alignment.center);
            
            // Ensayos
            
            for (int i = 0; i < codigoVias.Count; i++)
            {

                string ensayo = ensayos[0].Analisis; // REVIEW: ¿Por qué siempre se usa el primer ensayo? ¿No debería ser dinámico?
                string codigoInterno = codigoVias[i].CodigoInterno;
                
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[0], ensayo, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[1], "5", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[2], "0", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[3], "Estéril comercialmente", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[4], "No estéril comercialmente", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[5], codigoInterno, 6, true, Alignment.center, false);

                // Resultado por cada ensayo
                
                int inicioVias = 6; // Las vias empiezan en la columna 6
                
                string[] resultados = new string[5];

                for (int k = 0, nvia = 1; k < 5; k++, nvia++)
                {
                    string via = "n" + nvia;

                    string resultado = viaResultados.Find(r => r.CodPrecinto == codigoInterno && r.Muestra == via)
                        .Resultado;

                    if (resultado is null) resultado = "0";
                    
                    resultados[k] = resultado;
                    
                    FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[inicioVias + k], resultado, 6, false, Alignment.center, false);
                    
                }

                string conclusion = resultados.Any(r => r != "Estéril") ? "No conforme" : "Conforme";

                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[tabla.ColumnCount - 1], conclusion, 6, false, Alignment.center, false);
                
            }
            
            // Si no hay descripción eliminamos la ultima fila
            
            tabla.Rows.Last().Remove();
            
            // Guardar
            
            document.InsertTable(tabla);
            
        }
        
        public static void CrearTablaMuestrasExtraidasMicrobiologia(DocX document, SqlRepository repository)
        {
            List<Model.CodigoVia> codigoVias = repository.ObtenerCodigoVias<Model.CodigoVia>(NumOs).ToList();
            List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentes = repository.ObtenerMuestrasLaboratorioDirimente<Model.UspGetMuestraLaboratorioDirimente>(IdOTC, 1).ToList();
            
            int encabezadoFilas = 1;

            int cantidadFilas = encabezadoFilas + ( 4 * codigoVias.Count);
            int cantidadColumnas = 2;

            Table tabla = document.AddTable(cantidadFilas, cantidadColumnas);
            tabla.Alignment = Alignment.center;

            int[] anchoColumnas = { 100, 100 };

            for (int i = 0; i < cantidadColumnas; i++)
            {
                if (i <= anchoColumnas.Length - 1)
                {
                    tabla.SetColumnWidth(i, anchoColumnas[i]);
                }
            }

            FormatTableCell(tabla.Rows[0].Cells[0], "Lote", 8, true, Alignment.center);
            FormatTableCell(tabla.Rows[0].Cells[1], "Muestras extraídas para ensayo microbiológico", 8, true,
                Alignment.center);

            for (int i = 0, aux = encabezadoFilas; aux < (codigoVias.Count * 4) - 1; i++, aux+= 3)
            {
                string codigoInterno = codigoVias[i].CodigoInterno;
                string precintoMuestra = muestraLaboratorioDirimentes.Find(m => m.CodInterno == codigoInterno).MuestraLaboratorio;
                string precintoDirimente = muestraLaboratorioDirimentes.Find(m => m.CodInterno == codigoInterno).MuestraDirimente;
                
                tabla.Rows[aux + i].MergeCells(0, tabla.ColumnCount - 1);
                FormatTableCell(tabla.Rows[aux + i].Cells[0], codigoInterno, 8, true, Alignment.center, true);
                
                // Vias
                
                FormatTableCell(tabla.Rows[aux + i + 1].Cells[0], "", 8, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[aux + i + 1].Cells[1], "n1,n2,n3,n4,n5", 8, false, Alignment.center, false);
                
                // Precinto Muestras
                
                FormatTableCell(tabla.Rows[aux + i + 2].Cells[0], "Precinto de muestras", 8, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[aux + i + 2].Cells[1], precintoMuestra, 8, false, Alignment.center, false);
                
                // Precinto Dirimientes
                
                FormatTableCell(tabla.Rows[aux + i + 3].Cells[0], "Precinto de dirimencias", 8, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[aux + i + 3].Cells[1], precintoDirimente, 8, false, Alignment.center, false);
                
            }

            tabla.InsertRow();

            AgregarDescripcion(tabla,
                "Se tomaron muestras dirimentes con la misma metodología de extracción, en la misma cantidad, con precinto propio y sin requerimiento de ensayo");

            document.InsertTable(tabla);
            document.InsertParagraph();

        }
        
        public static void CrearTablaMuestrasExtraidasFisicoSensorial(DocX document, SqlRepository repository)
        {
            List<Model.CodigoViaFisicoSensorial> codigoVias =
                repository.ObtenerCodigoViasFisicoSensorialSecoSalado<Model.CodigoViaFisicoSensorial>(IdOTC).ToList();

            List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentes = repository.ObtenerMuestrasLaboratorioDirimente<Model.UspGetMuestraLaboratorioDirimente>(IdOTC, 3).ToList();

            int encabezadoFilas = 1;

            int cantidadFilas = encabezadoFilas + (4 * codigoVias.Count);
            int cantidadColumnas = 2;

            Table tabla = document.AddTable(cantidadFilas, cantidadColumnas);
            tabla.Alignment = Alignment.center;

            int[] anchoColumnas = { 100, 100 };

            for (int i = 0; i < cantidadColumnas; i++)
            {
                if (i <= anchoColumnas.Length - 1)
                {
                    tabla.SetColumnWidth(i, anchoColumnas[i]);
                }
            }
            
            FormatTableCell(tabla.Rows[0].Cells[0], "Lote", 8, true, Alignment.center);
            FormatTableCell(tabla.Rows[0].Cells[1], "Muestras extraídas para ensayo físico sensorial", 8, true,
                Alignment.center);


            for (int i = 0, aux = encabezadoFilas; aux < codigoVias.Count * 4 - 1; i++, aux += 3)
            {
                string codigoInterno = codigoVias[i].CodigoInterno;
                string precintoMuestra = muestraLaboratorioDirimentes.Find(m => m.CodInterno == codigoInterno).MuestraLaboratorio;
                string precintoDirimente = muestraLaboratorioDirimentes.Find(m => m.CodInterno == codigoInterno).MuestraDirimente;
                
                // Codigo
                
                tabla.Rows[aux + i].MergeCells(0, tabla.ColumnCount - 1);
                FormatTableCell(tabla.Rows[aux + i].Cells[0], codigoInterno, 8, true, Alignment.center);
                
                // Vias
                
                FormatTableCell(tabla.Rows[aux + i + 1].Cells[0], "", 8, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[aux + i + 1].Cells[1], codigoInterno, 8, false, Alignment.center, false);
                
                // Precinto de muestras

                FormatTableCell(tabla.Rows[aux + i + 2].Cells[0], "Precintos de muestras", 8, false, Alignment.left, false);
                FormatTableCell(tabla.Rows[aux + i + 2].Cells[1], precintoMuestra,8, false, Alignment.center, false);
                
                // Precinto Dirimentes

                FormatTableCell(tabla.Rows[aux + i + 3].Cells[0], "Precintos de dirimencias", 8, false, Alignment.left, false);
                FormatTableCell(tabla.Rows[aux + i + 3].Cells[1], precintoDirimente, 8, false, Alignment.center, false);
            }

            tabla.InsertRow();

            AgregarDescripcion(tabla, "Se tomaron muestras dirimentes con la misma metodología de extracción, en la misma cantidad, con precinto propio y sin requerimiento de ensayo");
        
            document.InsertTable(tabla);
            document.InsertParagraph();

        }
        
        public static void CrearTablaMuestreoParaAnalisisMicrobiologicos(DocX document, SqlRepository repository)
        {
            List<Model.CodigoVia> codigoVias = repository.ObtenerCodigoVias<Model.CodigoVia>("250622.01").ToList();
            List<Model.Ensayo> ensayos = repository.ObtenerEnsayos<Model.Ensayo>(80633, 5, 2).ToList();
            List<Model.ViaResultado> viaResultado = repository.ViasResultados<Model.ViaResultado>(80633, 2).ToList();
        
            int encabezadoFilas = 4;
            int totalColumnas = 12;
            
            int totalFilas = encabezadoFilas + ensayos.Count;
        
            foreach (var codigoVia in codigoVias)
            {
                Table tabla = document.AddTable(totalFilas, totalColumnas);
                tabla.Alignment = Alignment.center;

                int[] anchoColumnas = { 40, 20, 20, 40, 20, 20, 20, 20, 20, 20, 20, 40 };
        
                for (int i = 0; i < totalColumnas; i++)
                {
                    if (i <= anchoColumnas.Length - 1)
                    {
                        tabla.SetColumnWidth(i, anchoColumnas[i]);
                    }
                }
                
                // Combinar filas
                
                int ultimaFila = encabezadoFilas - 1;
                
                tabla.MergeCellsInColumn(0, 1, ultimaFila); // Microorganismo
                tabla.MergeCellsInColumn(1, 2, ultimaFila); // n 
                tabla.MergeCellsInColumn(2, 2, ultimaFila); // c
                tabla.MergeCellsInColumn(3, 1, ultimaFila); //categoria
                tabla.MergeCellsInColumn(4, 2, ultimaFila); // m
                tabla.MergeCellsInColumn(5, 2, ultimaFila); // m
                
                tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 0, ultimaFila);
                
                // Titulo Tabla
                
                tabla.Rows[0].MergeCells(0, tabla.ColumnCount - 1);
                FormatTableCell(tabla.Rows[0].Cells[0], "PLANES DE MUESTREO PARA ANALISIS MICROBIOLOGICOS", 3, true, Alignment.center);
                
                // Microorganismo
                
                FormatTableCell(tabla.Rows[1].Cells[0], "MICROORGANISMO", 3, true, Alignment.center);
                
                // Plan de evaluación
                
                tabla.Rows[1].MergeCells(1, 2);
                FormatTableCell(tabla.Rows[1].Cells[1], "PLAN DE EVALUACION", 3, true, Alignment.center);
                
                FormatTableCell(tabla.Rows[2].Cells[1], "n", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[2].Cells[2], "c", 3, true, Alignment.center);
                
                // Categoria
                
                FormatTableCell(tabla.Rows[1].Cells[2], "CATEGORIA", 3, true, Alignment.center);
                
                // Limites
                
                tabla.Rows[1].MergeCells(3, 4);
                FormatTableCell(tabla.Rows[1].Cells[3], "LIMITES", 3, true, Alignment.center);
                
                FormatTableCell(tabla.Rows[2].Cells[4], "m", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[2].Cells[5], "M", 3, true, Alignment.center);
                
                // Resultados
                
                tabla.Rows[1].MergeCells(4, 8);
                tabla.Rows[2].MergeCells(6, 10);
                FormatTableCell(tabla.Rows[1].Cells[4], "RESULTADOS", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[2].Cells[6], codigoVia.CodigoInterno, 3, true, Alignment.center);
                
                FormatTableCell(tabla.Rows[3].Cells[6], "n1", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[3].Cells[7], "n2", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[3].Cells[8], "n3", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[3].Cells[9], "n4", 3, true, Alignment.center);
                FormatTableCell(tabla.Rows[3].Cells[10], "n5", 3, true, Alignment.center);
                
                // Conclusion
                FormatTableCell(tabla.Rows[1].Cells[5], "CONCLUSION", 3, true, Alignment.center);
                
                // Ensayos
                
                for (int i = 0; i < ensayos.Count; i++)
                {

                    string ensayoLabel = ensayos[i].Analisis;
                    
                    // Información por cada ensayo fija

                    if (ensayoLabel == "ENUMERACIÓN DE MICROORGANISMOS A 30º C")
                    {
                        ensayoLabel = "Aerobios mesófilos (30°C)";
                        
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[2], "3", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[3], "1", 3, false, Alignment.center, false);
                        
                        var celda = tabla.Rows[encabezadoFilas + i].Cells[4];
                        var parrafo = celda.Paragraphs.First();
                        parrafo.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Append("4").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Alignment = Alignment.center;
                        celda.VerticalAlignment = VerticalAlignment.Center;
                        
                        var celda2 = tabla.Rows[encabezadoFilas + i].Cells[5];
                        var parrafo2 = celda2.Paragraphs.First();
                        parrafo2.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Append("4").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo2.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Alignment = Alignment.center;
                        celda2.VerticalAlignment = VerticalAlignment.Center;
                    }
                    
                    if (ensayoLabel == "ENUMERACIÓN DE BACTERIAS ANAEROBIAS SULFITO REDUCTORES")
                    {
                        ensayoLabel = "Anaerobio sulfito reductores (**)";
                        
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[2], "2", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[3], "5", 3, false, Alignment.center, false);
                        
                        var celda = tabla.Rows[encabezadoFilas + i].Cells[4];
                        var parrafo = celda.Paragraphs.First();
                        parrafo.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Append("2").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Alignment = Alignment.center;
                        celda.VerticalAlignment = VerticalAlignment.Center;
                        
                        var celda2 = tabla.Rows[encabezadoFilas + i].Cells[5];
                        var parrafo2 = celda2.Paragraphs.First();
                        parrafo2.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Append("3").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo2.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Alignment = Alignment.center;
                        celda2.VerticalAlignment = VerticalAlignment.Center;
                    }
                    
                    if (ensayoLabel == "SALMONELLA")
                    {
                        ensayoLabel = "Salmonella spp";
                        
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[2], "0", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[3], "10", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[4], "Ausencia/25g", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[5], "-", 3, false, Alignment.center, false);
                    }
                    
                    if (ensayoLabel == "Enumeración de Enterobacteriaceae")
                    {
                        ensayoLabel = "Enterobacterias";
                        
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[2], "2", 3, false, Alignment.center, false);
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[3], "5", 3, false, Alignment.center, false);

                        var celda = tabla.Rows[encabezadoFilas + i].Cells[4];
                        var parrafo = celda.Paragraphs.First();
                        parrafo.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Append("3").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo.Alignment = Alignment.center;
                        celda.VerticalAlignment = VerticalAlignment.Center;
                        
                        var celda2 = tabla.Rows[encabezadoFilas + i].Cells[5];
                        var parrafo2 = celda2.Paragraphs.First();
                        parrafo2.Append("10").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Append("4").Script(Script.superscript).FontSize(3).Font("Calibri (Cuerpo)").Bold();
                        parrafo2.Append(" UFC/g").FontSize(3).Font("Calibri (Cuerpo)").Bold(false);
                        parrafo2.Alignment = Alignment.center;
                        celda2.VerticalAlignment = VerticalAlignment.Center;
                    }
                    
                    FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[0], ensayoLabel, 3, false, Alignment.center, false);
                    FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[1], "5", 3, false, Alignment.center, false);
                
                    int j = 6;
                    
                    // Resultado por cada ensayo

                    var resultados =
                        viaResultado
                            .Where(v => v.IdAnalisis == ensayos[i].IdAnalisis && v.CodPrecinto == codigoVia.CodigoInterno).ToList();
                    
                    foreach (var via in resultados)
                    {
                        FormatTableCell(tabla.Rows[encabezadoFilas + i].Cells[j], via.Resultado, 3, false, Alignment.center, false);
                        j++;
                    }
                    
                }
        
                document.InsertTable(tabla);
                document.InsertParagraph();
            }
        }
        
        public static void CrearTablaExamenesSensorialesSecoSalado(DocX document, SqlRepository repository)
        {
            
            // List<Model.UspGetTablaExamenesSensoriales> tablaResultados =
            //     repository.ObtenerTablaExamenesSensorial<Model.UspGetTablaExamenesSensoriales>(IdOT, 1).ToList();
            
            List<Model.CodigoViaFisicoSensorial> codigoVias =
                repository.ObtenerCodigoViasFisicoSensorialSecoSalado<Model.CodigoViaFisicoSensorial>(IdOTC).ToList();

            List<Model.UspGetListarAnalisisCNNuevo> examenesSensorial = repository
                .ObtenerExamenesSensorialSecoSalado<Model.UspGetListarAnalisisCNNuevo>(IdOTC, 1).ToList();

            int cabeceraFilas = 3;
            int cabeceraColumnas = 11;

            int tablaFilas = cabeceraFilas + codigoVias.Count;
            int tablaColumnas = cabeceraColumnas;

            Table tabla = document.AddTable(tablaFilas, tablaColumnas);
            tabla.Alignment = Alignment.left;

            // Encabezado

            // Ancho de columnas

            int[] anchoColumnas = { 25, 30, 23, 23, 40, 60, 60, 80, 55, 55, 55 };

            for (int i = 0; i < tablaColumnas; i++)
            {
                tabla.SetColumnWidth(i, anchoColumnas[i]);
            }

            // Combinar filas
            
            int ultimaFila = cabeceraFilas - 1;

            tabla.MergeCellsInColumn(0, 1, ultimaFila); // Codigos
            tabla.MergeCellsInColumn(1, 1, ultimaFila); // Vias
            tabla.MergeCellsInColumn(tabla.ColumnCount - 1, 1, ultimaFila); // Ùltima columnna

            // Examenes sensoriales

            tabla.Rows[0].MergeCells(0, tabla.Rows[1].Cells.Count - 1);
            FormatTableCell(tabla.Rows[0].Cells[0], "EXAMENES SENSORIALES", 7, true, Alignment.center);

            // Codigo

            FormatTableCell(tabla.Rows[1].Cells[0], "CÓDIGO", 6, true, Alignment.center, true, TextDirection.btLr);

            // Vías (n)

            FormatTableCell(tabla.Rows[1].Cells[1], "VIAS (n)", 6, true, Alignment.center);

            // Numero de aceptacion

            tabla.Rows[1].MergeCells(2, 3);
            FormatTableCell(tabla.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", 6, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[2], "N°", 6, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[3], "C", 6, true, Alignment.center);

            // Especie

            FormatTableCell(tabla.Rows[1].Cells[3], "ESPECIE", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[4], "Corresponde a la declarada por el exportador", 5, false, Alignment.center);

            // Presentacion

            FormatTableCell(tabla.Rows[1].Cells[4], "PRESENTACION", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[5], "Corresponde a la declarada por el exportador y debe incluir todos los aspectos señalados por éste (ejemplo: tipo de corte, tipo de empaque, entre otros)", 5, false, Alignment.both);

            // Aspecto

            FormatTableCell(tabla.Rows[1].Cells[5], "ASPECTO", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[6], "Normal. Ausencia de materias extrañas. No existen zonas micóticas, Ni moho Alófilo. Ausencia de quemaduras por excesivo calentamiento durante el secado evidenciadas por una piel viscosa o pegajosa", 5, false, Alignment.both);

            // Olor

            FormatTableCell(tabla.Rows[1].Cells[6], "OLOR", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[7], "Propio.Característico. Ausencia de olores objetables, persistentes e inconfundibles que sean signos de descomposición (olor ácido, pútrido, etc) o de contaminación por sustancias extrañas (combustibles, productos de limpieza, etc)", 5, false, Alignment.both);

            // Color

            FormatTableCell(tabla.Rows[1].Cells[7], "COLOR", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[8], "Natural, típico y uniforme. No se permite la presencia de manchas Rojizas o verdosas ni decoloración amarilla o naranja amarillenta", 5, false, Alignment.both);
            
            // Textura

            FormatTableCell(tabla.Rows[1].Cells[8], "TEXTURA", 7, true, Alignment.center);
            FormatTableCell(tabla.Rows[2].Cells[9], "Típica de acuerdo al producto. Ausencia de carne con textura caracterizada por agrietamiento generalizado en mas de dos tercios de superficie, desgarrada o rota", 5, false, Alignment.center);

            // conclusion

            FormatTableCell(tabla.Rows[1].Cells[9], "CONCLUSION", 7, true, Alignment.center);
            
            // Vias

            for (int i = 0; i < codigoVias.Count; i++)
            {
                string codigoInterno = codigoVias[i].CodigoInterno;
                string rangoVias = codigoVias[i].RangoVias;
                
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[0], codigoInterno, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[1], rangoVias, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[2], "2", 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[3], "(1)", 6, false, Alignment.center, false);
                
                var aspecto = examenesSensorial
                    .Where(e => e.Codigos == codigoInterno)
                    .All(e => e.Aspecto == "BUENO") ? "Bueno" : "No Bueno";
                
                var color = examenesSensorial
                    .Where(e => e.Codigos == codigoInterno)
                    .All(e => e.Color == "BUENO") ? "Bueno" : "No Bueno";
                
                var olor = examenesSensorial
                    .Where(e => e.Codigos == codigoInterno)
                    .All(e => e.Olor == "BUENO") ? "Bueno" : "No Bueno";
                
                var textura = examenesSensorial
                    .Where(e => e.Codigos == codigoInterno)
                    .All(e => e.Textura == "BUENO") ? "Bueno" : "No Bueno";
                    
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[6], aspecto, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[7], olor, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[8], color, 6, false, Alignment.center, false);
                FormatTableCell(tabla.Rows[cabeceraFilas + i].Cells[9], textura, 6, false, Alignment.center, false);
                
            }

            tabla.InsertRow();

            AgregarDescripcion(tabla,
                "(*) El paréntesis en el número de aceptación (c) indica el número de aceptación para descomposición \n METODO DE ENSAYO (ANALISIS SENSORIAL  : ISO 4121. SECOUND EDITION. Item 5.2, 6.3.2: 2003: Sensory analysis. Guidelines for the use of quantitative response scales");
            
            // Guardar

            document.InsertTable(tabla);
            
        }
            
        private static void AgregarTitulo(Table table, string titulo)
        {
            table.InsertRow();
            var firstRow = table.Rows[0];
            
            firstRow.Cells[0].SetBorder(TableCellBorderType.InsideH, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            firstRow.Cells[0].SetBorder(TableCellBorderType.InsideV, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            firstRow.Cells[0].SetBorder(TableCellBorderType.Top, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            firstRow.Cells[0].SetBorder(TableCellBorderType.Bottom, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            firstRow.Cells[0].SetBorder(TableCellBorderType.Left, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            firstRow.Cells[0].SetBorder(TableCellBorderType.Right, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            
            FormatTableCell(firstRow.Cells[0], titulo, 7, true, Alignment.left, false);
        }
        
        private static void AgregarDescripcion(Table table, string descripcion, Alignment alignment = Alignment.left)
        {
            var lastRow = table.Rows[table.Rows.Count - 1];
            lastRow.MergeCells(0, table.ColumnCount - 1);
            
            lastRow.Cells[0].SetBorder(TableCellBorderType.InsideH, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.InsideV, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Top, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Bottom, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Left, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Right, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            
            FormatTableCell(lastRow.Cells[0], descripcion, 6, false, alignment, false);
        }
        
        private static string FormatearResultadoNumerico(string valor)
        {
            // Usamos CultureInfo.InvariantCulture para asegurarnos de que el punto "."
            // siempre se reconozca como el separador decimal, sin importar la configuración del servidor.
            if (Decimal.TryParse(valor, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal numero))
            {
                // "G29" es un especificador de formato estándar que significa "formato general"
                // y es la forma más segura de eliminar los ceros finales sin perder precisión.
                return numero.ToString("G29", CultureInfo.InvariantCulture);
            }

            // Si el valor no se puede convertir a número (ej. es "N/A" o texto),
            // simplemente devolvemos el valor original sin cambios.
            return valor;
        }

        public static string RedondearValorCustom(string valorComoTexto)
        {
            
            // Paso 1: Intentar convertir el string a decimal (esta lógica no cambia).
            bool esNumeroValido = decimal.TryParse(
                valorComoTexto,
                NumberStyles.Any,
                CultureInfo.InvariantCulture, // Espera un '.' como separador decimal en la entrada
                out decimal valorNumerico
            );

            if (esNumeroValido)
            {
                // Paso 2: Aplicar el redondeo matemático (esta lógica no cambia).
                decimal valorRedondeado = Math.Round(valorNumerico, 1, MidpointRounding.AwayFromZero);

                // Paso 3: ¡NUEVO! Convertir el resultado de vuelta a string con el formato deseado.
                // Usamos "F1" para asegurar que siempre tenga un decimal y CultureInfo.InvariantCulture
                // para asegurar que el separador sea un punto '.'.
                return valorRedondeado.ToString("F1", CultureInfo.InvariantCulture);
            }
            else
            {
                // Devolvemos un string por defecto que sea coherente con el formato de salida.
                return "0.0";
            }
        }
        
        public static (List<Model.UspGetMuestraLaboratorioDirimente> histamina, List<Model.UspGetMuestraLaboratorioDirimente> metalesPesados)
            DividirMuestraLaboratorioFQ(List<Model.UspGetMuestraLaboratorioDirimente> muestraLaboratorioDirimentes)
        {
            int indiceDeCorte = -1;

            for (int i = 1; i < muestraLaboratorioDirimentes.Count; i++)
            {
                try
                {
                    int numeroAnterior = int.Parse(muestraLaboratorioDirimentes[i - 1].CodInterno.Substring(1));
                    int numeroActual = int.Parse(muestraLaboratorioDirimentes[i].CodInterno.Substring(1));

                    if (numeroActual <= numeroAnterior)
                    {
                        indiceDeCorte = i;
                        break;
                    }
                }
                catch (FormatException)
                {
                    continue;
                }
            }

            if (indiceDeCorte != -1)
            {
                var primeraParte = muestraLaboratorioDirimentes.Take(indiceDeCorte).ToList();
                var segundaParte = muestraLaboratorioDirimentes.Skip(indiceDeCorte).ToList();
                return (primeraParte, segundaParte);
            }
            else
            {
                return (muestraLaboratorioDirimentes, new List<Model.UspGetMuestraLaboratorioDirimente>());
            }
        }
        
    }
}
