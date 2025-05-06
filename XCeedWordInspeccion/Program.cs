using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XCeedWordInspeccion
{
    internal class Program
    {
        
        private const int MAX_VIAS = 5;
        
        public static void Main(string[] args)
        {
            
            string filename = "Inspecciones.docx";
            string templatePath = @"C:\Users\ljauregui\RiderProjects\XCeedWord\XCeedWord\bin\Debug\PlantillaAC.docx";
            
            File.Copy(templatePath, filename, true);

            using (DocX document = DocX.Load(filename))
            {
                
                SqlRepository repository = new SqlRepository();

                List<Model.Ensayo> ensayos = repository.ObtenerEnsayos<Model.Ensayo>(79107, 5, 2).ToList();
                List<Model.Via> vias = repository.ObtenerVias<Model.Via>(79107, 2, 5).ToList();
                List<Model.ViaResultado> viaResultado = repository.ViasResultados<Model.ViaResultado>(79107, 2).ToList();
                List<Model.CodigoVia> codigoVias = repository.ObtenerCodigoVias<Model.CodigoVia>("250317.23").ToList();
                List<Model.MuestraCls> muestras = repository.ObtenerMuestras<Model.MuestraCls>(79822, 1).ToList();
                
                // Configurar márgenes
                
                document.MarginLeft   = 25; // 2.5 cm
                document.MarginRight  = 20; // 2 cm
                document.MarginTop    = 0; // 2 cm
                document.MarginBottom = 0; // 2 cm
                
                // Crear Tabla
                
                // CreateTableA(document, 3, 3);

                var totalTables = (int) Math.Ceiling(vias.Count / (double)MAX_VIAS);

                for (int i = 0; i < totalTables; i++)
                {

                    List<Model.Via> rangeVias =
                        vias.GetRange((i * MAX_VIAS), Math.Min(MAX_VIAS, vias.Count - i * MAX_VIAS));

                    var rangeResultados =
                        viaResultado
                            .OrderBy(v => v.IdAnalisis)
                            .ThenBy(v => v.CodPrecinto.Substring(1))
                            .ThenBy(v => int.Parse(v.Muestra.Substring(1)))
                            .ToList();

                    CreateTableA(document, rangeVias, ensayos, rangeResultados, codigoVias, i, vias.Count);
                    document.InsertParagraph().SpacingAfter(1);
                }
                
                
                // CreateTableC(document, 9, 1);
                // CreateTableD(document, 3);
                // CreateTableE(document, 3);
                // CreateTableF(document, 3);
                // CreateTableG(document, 5, 1);
                // CreateTableH(document, 5, 1);

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
            
            FormatTableCell(table.Rows[0].Cells[0], "MICROORGANISMO", "Arial", 3, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[0].MergeCells(1, 2);
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[0].Cells[1], "PLAN DE EVALUACIÓN", "Arial", 3, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", "Arial", 3, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", "Arial", 3, true, Alignment.center);
            
            // Limites
            
            table.Rows[0].MergeCells(2, 3);
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[0].Cells[2], "LIMITES", "Arial", 3, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[3], "m", "Arial", 3, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[4], "M", "Arial", 3, true, Alignment.center);
            
            // Distribución de muestras
            
            table.Rows[0].MergeCells(3, 3 + vias.Count - 1);
            FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", "Arial", 3, true, Alignment.center);
            
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
            //     FormatTableCell(table.Rows[1].Cells[startCol - 2], productoCodigo, "Arial", 3, true, Alignment.center);
            //     
            //     // FormatTableCell(table.Rows[2].Cells[startCol], currentVia, "Arial", 4, true, Alignment.center); // Codigo
            //
            //     i = j; // Saltar al siguiente grupo
            // }
            
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, "Arial", 4, true, Alignment.center);
            }
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, "Arial", 4, false, Alignment.left);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), "Arial", 4, false, Alignment.center);

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
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, "Arial", 4, true, Alignment.center);
                    }

                    j++;

                }
                    
            }
            
            // Conclusión

            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", "Arial", 3, true, Alignment.center);

            // Guardar
            
            document.InsertTable(table);
        }

        public static void CreateTableB(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, int iTable, int numVias)
        {
            int headerRows = 3;
            int headerColumns = 7;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
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
                    table.SetColumnWidth(i, 35);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
            }
            
            // Combinar filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            
            table.MergeCellsInColumn(3, 1, headerRows - 1);
            table.MergeCellsInColumn(4, 1, headerRows - 1);
            
            table.MergeCellsInColumn(5, 1, headerRows - 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 1, headerRows - 1);
            
            // Esterilidad comercial
            
            table.Rows[0].MergeCells(0, table.Rows[0].Cells.Count - 1);
            FormatTableCell(table.Rows[0].Cells[0], "ESTERILIDAD COMERCIAL", "Arial", 5, true, Alignment.center);
            
            // Analisis
            
            FormatTableCell(table.Rows[1].Cells[0], "ANALISIS", "Arial", 4, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[1].Cells[1], "PLAN DE EVALUACIÓN", "Arial", 4, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", "Arial", 4, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", "Arial", 4, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[1].Cells[2], "ACEPTACIÓN", "Arial", 4, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[1].Cells[3], "RECHAZO", "Arial", 4, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[4], "CÓDIGO", "Arial", 4, true, Alignment.center);
            
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
                    table.Rows[1].MergeCells(startCol, endCol);
                }

                FormatTableCell(table.Rows[1].Cells[startCol], currentVia, "Arial", 4, true, Alignment.center);

                i = j; // Saltar al siguiente grupo
            }

            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 6; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, "Arial", 4, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[table.Rows[1].Cells.Count - 1], "CONCLUSIÓN", "Arial", 4, true, Alignment.center);
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, "Arial", 4, false, Alignment.left);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), "Arial", 4, false, Alignment.center);

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
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, "Arial", 4, true, Alignment.center);
                    }

                    j++;

                }
                    
            }
            
            
            // Guardar
            
            document.InsertTable(table);
            
        }
        
        public static void CreateTableC(DocX document, List<Model.Via> vias, List<Model.Ensayo> ensayos, List<Model.ViaResultado> viaResultados, List<Model.CodigoVia> codigoVias, int iTable, int numVias)
        {
            int headerRows = 3;
            int headerColumns = 6;

            int tableRows = headerRows + ensayos.Count;
            int tableColumns = headerColumns + vias.Count;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
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
                    table.SetColumnWidth(i, 40);
                }
                
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 0, headerRows - 1);
            
            table.MergeCellsInColumn(1, 0, 1);
            table.MergeCellsInColumn(2, 0, 1);
            
            table.MergeCellsInColumn(3, 0, 1);
            table.MergeCellsInColumn(4, 0, 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 0, headerRows - 1);
            
            // Determinación
            
            FormatTableCell(table.Rows[0].Cells[0], "DETERMINACIÓN", "Arial", 4, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[0].MergeCells(1, 2);
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[0].Cells[1], "PLAN DE EVALUACIÓN", "Arial", 4, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", "Arial", 4, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", "Arial", 4, true, Alignment.center);
            
            // Limites
            
            table.Rows[0].MergeCells(2, 3);
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[0].Cells[2], "LIMITES (mg/kg)", "Arial", 4, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[3], "m", "Arial", 4, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[4], "M", "Arial", 4, true, Alignment.center);
            
            // Distribución de muestras
            
            table.Rows[0].MergeCells(3, 3 + vias.Count - 1);
            FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", "Arial", 4, true, Alignment.center);
            
            // Código de Vías (Dinámicas)
            
            AgruparYFormatearVias(table, vias, codigoVias, 1,5);
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, "Arial", 4, true, Alignment.center);
            }
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, "Arial", 4, false, Alignment.left);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), "Arial", 4, false, Alignment.center);

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
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, "Arial", 4, true, Alignment.center);
                    }

                    j++;

                }
                    
            }
                
            // Conclusión

            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", "Arial", 4, true, Alignment.center);

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
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 80, 50, 100, 100, 100, 80 };

            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // M
            
            FormatTableCell(table.Rows[0].Cells[0], "M", "Arial", 6, true, Alignment.center);
            
            // n
            
            FormatTableCell(table.Rows[0].Cells[1], "n", "Arial", 6, true, Alignment.center);
            
            // Elementos
            
            FormatTableCell(table.Rows[0].Cells[2], "ELEMENTOS", "Arial", 6, true, Alignment.center);
            
            // Contenido máximo
            
            FormatTableCell(table.Rows[0].Cells[3], "CONTENIDO MÁXIMO", "Arial", 6, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(table.Rows[0].Cells[4], "RESULTADO", "Arial", 6, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[0].Cells[5], "CONCLUSIÓN", "Arial", 6, true, Alignment.center);
            
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
            FormatTableCell(table.Rows[0].Cells[0], "INDICADORES PARASITOLOGICOS", "Arial", 6, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", "Arial", 5, true, Alignment.center);
            
            // Vías (n)
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS (n)", "Arial", 5, true, Alignment.center);
            
            // Plan de evaluación
            
            FormatTableCell(table.Rows[1].Cells[2], "PLAN DE EVALUACIÓN", "Arial", 5, true, Alignment.center);
            
            // Resultados
            
            FormatTableCell(table.Rows[1].Cells[3], "RESULTADOS", "Arial", 5, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[4], "CONCLUSIÓN", "Arial", 5, true, Alignment.center);
            
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
            FormatTableCell(table.Rows[0].Cells[0], "METALES PESADOS", "Arial", 6, true, Alignment.center);
            
            // Análisis
            
            FormatTableCell(table.Rows[1].Cells[0], "ANÁLISIS", "Arial", 5, true, Alignment.center);
            
            // Códigos
            
            FormatTableCell(table.Rows[1].Cells[1], "CÓDIGOS", "Arial", 5, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[2], "VÍAS", "Arial", 5, true, Alignment.center);
            
            // Contenido máximo (mg/kg peso fresco)
            
            FormatTableCell(table.Rows[1].Cells[3], "CONTENIDO MÁXIMO (mg/kg peso fresco)", "Arial", 5, true, Alignment.center);
            
            // Unidades
            
            FormatTableCell(table.Rows[1].Cells[4], "UNIDADES", "Arial", 5, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(table.Rows[1].Cells[5], "RESULTADO", "Arial", 5, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[6], "CONCLUSIÓN", "Arial", 5, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableG(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 2;
            int headerColumns = 7;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns + totalVias;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 70, 40, 40, 50, 50, 50 };
            
            for (int i = 0; i < tableColumns; i++)
            {

                if (i <= columnWidths.Length - 1)
                {
                    table.SetColumnWidth(i, columnWidths[i]);
                }
                
                // Vías dinámicas

                if (i >= 6 && i < tableColumns - 1)
                {
                    table.SetColumnWidth(i, 30);
                }
                
                // Última columna

                if (i == tableColumns - 1)
                {
                    table.SetColumnWidth(i, 50);
                }
                
            }
            
            // Combinas filas
            
            table.MergeCellsInColumn(0, 0, headerRows - 1);
            
            table.MergeCellsInColumn(3, 0, headerRows - 1);
            table.MergeCellsInColumn(4, 0, headerRows - 1);
            
            table.MergeCellsInColumn(table.ColumnCount - 1, 0, headerRows - 1);
            
            // Ensayo
            
            FormatTableCell(table.Rows[0].Cells[0], "ENSAYO", "Calibri", 7, true, Alignment.center);
            
            // Plan de muestreo
            
            table.Rows[0].MergeCells(1, 2);
            FormatTableCell(table.Rows[0].Cells[1], "PLAN DE MUESTREO", "Calibri", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[1].Cells[1], "n", "Calibri", 8, true, Alignment.center);
            FormatTableCell(table.Rows[1].Cells[2], "c", "Calibri", 8, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[0].Cells[2], "ACEPTACIÓN", "Calibri", 7, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[0].Cells[3], "RECHAZO", "Calibri", 7, true, Alignment.center);
            
            // Resultados
            
            table.Rows[0].MergeCells(4, 4 + totalVias);
            FormatTableCell(table.Rows[0].Cells[4], "RESULTADOS", "Calibri", 7, true, Alignment.center);
            
            // Lote
            
            FormatTableCell(table.Rows[1].Cells[5], "LOTE", "Calibri", 7, true, Alignment.center);
            
            // Vias (Dinámicas)
            
            for (int i = 0, iCellIndex= 6; i < totalVias; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[1].Cells[iCellIndex], $"n{i+1}", "Calibri", 7, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", "Calibri", 7, true, Alignment.center);
            
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
            FormatTableCell(table.Rows[0].Cells[0], "EXAMENES SENSORIALES", "Calibri", 8, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", "Calibri", 7, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS", "Calibri", 7, true, Alignment.center);
            
            // Numeración de Aceptación
            
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", "Calibri", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[2], "N*", "Calibri", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[3], "(c)*", "Calibri", 7, true, Alignment.center);
            
            // Aspecto
            
            table.Rows[1].MergeCells(3, 4);
            FormatTableCell(table.Rows[1].Cells[3], "ASPECTO", "Calibri", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[4], "EXTERIOR", "Calibri", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[5], "INTERIOR", "Calibri", 7, true, Alignment.center);
            
            // Olor
            
            FormatTableCell(table.Rows[1].Cells[4], "OLOR", "Calibri", 7, true, Alignment.center);
            
            // Color
            
            FormatTableCell(table.Rows[1].Cells[5], "COLOR", "Calibri", 7, true, Alignment.center);
            
            // Sabor
            
            FormatTableCell(table.Rows[1].Cells[6], "SABOR", "Calibri", 7, true, Alignment.center);
            
            // Textura
            
            FormatTableCell(table.Rows[1].Cells[7], "TEXTURA", "Calibri", 7, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[8], "CONCLUSIÓN", "Calibri", 7, true, Alignment.center);
            
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
            FormatTableCell(table.Rows[0].Cells[0], "EXAMENES SENSORIALES", "Calibri", 8, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CÓDIGO", "Calibri", 7, true, Alignment.center);
            
            // Vías
            
            FormatTableCell(table.Rows[1].Cells[1], "VÍAS", "Calibri", 7, true, Alignment.center);
            
            // Numeración de Aceptación
            
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", "Calibri", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[2], "N*", "Calibri", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[3], "(c)*", "Calibri", 7, true, Alignment.center);
            
            // Aspecto
            
            table.Rows[1].MergeCells(3, 4);
            FormatTableCell(table.Rows[1].Cells[3], "ASPECTO", "Calibri", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[4], "EXTERIOR", "Calibri", 7, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[5], "INTERIOR", "Calibri", 7, true, Alignment.center);
            
            // Olor
            
            FormatTableCell(table.Rows[1].Cells[4], "OLOR", "Calibri", 7, true, Alignment.center);
            
            // Color
            
            FormatTableCell(table.Rows[1].Cells[5], "COLOR", "Calibri", 7, true, Alignment.center);
            
            // Sabor
            
            FormatTableCell(table.Rows[1].Cells[6], "SABOR", "Calibri", 7, true, Alignment.center);
            
            // Textura
            
            FormatTableCell(table.Rows[1].Cells[7], "TEXTURA", "Calibri", 7, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[8], "CONCLUSIÓN", "Calibri", 7, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        private static void FormatTableCell(Cell cell, string text, string fontName, int fontSize, bool isBold, Alignment alignment, TextDirection textDirection = TextDirection.right )
        {
            cell.Paragraphs[0].Append(text)
                .Font(fontName)
                .Bold(isBold)
                .FontSize(fontSize)
                .Alignment = alignment;

            cell.VerticalAlignment = VerticalAlignment.Center;
            cell.TextDirection = textDirection;
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
                
                int endCol =  5 + j - (i == 0 ? 1 : aux - 1);
                
                if (j - i > 1)
                {
                    table.Rows[rowIndex].MergeCells(startCol - 2, endCol - 2);
                }
        
                var productoCodigo = codigoVias.Find(x => x.CodigoInterno == currentVia)?.ProductoCodigo ?? "";
                FormatTableCell(table.Rows[rowIndex].Cells[startCol - 2], productoCodigo, "Arial", 3, true,
                    Alignment.center);
        
                i = j;
            }
        }
        
    }
}
