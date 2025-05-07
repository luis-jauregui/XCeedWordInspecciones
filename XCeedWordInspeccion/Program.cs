using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Drawing;
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

                List<Model.Ensayo> ensayos = repository.ObtenerEnsayos<Model.Ensayo>(79107, 1, 3).ToList();
                List<Model.Via> vias = repository.ObtenerVias<Model.Via>(79107, 3, 1).ToList();
                List<Model.ViaResultado> viaResultado = repository.ViasResultados<Model.ViaResultado>(79107, 3).ToList();
                List<Model.CodigoVia> codigoVias = repository.ObtenerCodigoVias<Model.CodigoVia>("250317.23").ToList();
                List<Model.MuestraCls> muestras = repository.ObtenerMuestras<Model.MuestraCls>(79822, 1).ToList();
                
                // Configurar márgenes
                
                document.MarginLeft   = 25; // 2.5 cm
                document.MarginRight  = 20; // 2 cm
                document.MarginTop    = 0; // 2 cm
                document.MarginBottom = 0; // 2 cm
                
                CreateTableJ(document, codigoVias);

                document.InsertParagraph().SpacingAfter(1);
                
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
                
                    CreateTableC(document, rangeVias, ensayos, rangeResultados, codigoVias, i, vias.Count);
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
            FormatTableCell(table.Rows[0].Cells[0], "ESTERILIDAD COMERCIAL", 5, true, Alignment.center);
            
            // Analisis
            
            FormatTableCell(table.Rows[1].Cells[0], "ANALISIS", 4, true, Alignment.center);
            
            // Plan de evaluación
            
            table.Rows[1].MergeCells(1, 2);
            FormatTableCell(table.Rows[1].Cells[1], "PLAN DE EVALUACIÓN", 4, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[1], "n", 4, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[2], "c", 4, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[1].Cells[2], "ACEPTACIÓN", 4, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[1].Cells[3], "RECHAZO", 4, true, Alignment.center);
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[4], "CÓDIGO", 4, true, Alignment.center);
            
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

                FormatTableCell(table.Rows[1].Cells[startCol], currentVia, 4, true, Alignment.center);

                i = j; // Saltar al siguiente grupo
            }

            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 6; i < vias.Count; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], vias[i].Muestra, 4, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[table.Rows[1].Cells.Count - 1], "CONCLUSIÓN", 4, true, Alignment.center);
            
            // Ensayos
            
            for (int i = 0; i < ensayos.Count; i++)
            {

                string ensayoLabel = ensayos[i].Analisis;
                
                FormatTableCell(table.Rows[headerRows + i].Cells[0], ensayoLabel, 4, false, Alignment.left);
                FormatTableCell(table.Rows[headerRows + i].Cells[1], numVias.ToString(), 4, false, Alignment.center);

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
                        FormatTableCell(table.Rows[headerRows + i].Cells[j], via.Resultado, 4, true, Alignment.center);
                    }

                    j++;

                }
                    
            }
            
            
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
            table.Alignment = Alignment.center;
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 80, 50, 100, 100, 100, 80 };

            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // M
            
            FormatTableCell(table.Rows[0].Cells[0], "M", 6, true, Alignment.center);
            
            // n
            
            FormatTableCell(table.Rows[0].Cells[1], "n", 6, true, Alignment.center);
            
            // Elementos
            
            FormatTableCell(table.Rows[0].Cells[2], "ELEMENTOS", 6, true, Alignment.center);
            
            // Contenido máximo
            
            FormatTableCell(table.Rows[0].Cells[3], "CONTENIDO MÁXIMO", 6, true, Alignment.center);
            
            // Resultado
            
            FormatTableCell(table.Rows[0].Cells[4], "RESULTADO", 6, true, Alignment.center);
            
            // Conclusión
            
            FormatTableCell(table.Rows[0].Cells[5], "CONCLUSIÓN", 6, true, Alignment.center);
            
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
            
            FormatTableCell(table.Rows[0].Cells[0], "ENSAYO", 7, true, Alignment.center);
            
            // Plan de muestreo
            
            table.Rows[0].MergeCells(1, 2);
            FormatTableCell(table.Rows[0].Cells[1], "PLAN DE MUESTREO", 7, true, Alignment.center);
            
            FormatTableCell(table.Rows[1].Cells[1], "n", 8, true, Alignment.center);
            FormatTableCell(table.Rows[1].Cells[2], "c", 8, true, Alignment.center);
            
            // Aceptación
            
            FormatTableCell(table.Rows[0].Cells[2], "ACEPTACIÓN", 7, true, Alignment.center);
            
            // Rechazo
            
            FormatTableCell(table.Rows[0].Cells[3], "RECHAZO", 7, true, Alignment.center);
            
            // Resultados
            
            table.Rows[0].MergeCells(4, 4 + totalVias);
            FormatTableCell(table.Rows[0].Cells[4], "RESULTADOS", 7, true, Alignment.center);
            
            // Lote
            
            FormatTableCell(table.Rows[1].Cells[5], "LOTE", 7, true, Alignment.center);
            
            // Vias (Dinámicas)
            
            for (int i = 0, iCellIndex= 6; i < totalVias; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[1].Cells[iCellIndex], $"n{i+1}", 7, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", 7, true, Alignment.center);
            
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
        
        public static void CreateTableJ(DocX document, List<Model.CodigoVia> codigoVias)
        {
            int headerRows = 3;
            int headerColumns = 11;

            int tableRows = headerRows + codigoVias.Count;
            int tableColumns = headerColumns;
            
            Table table = document.AddTable(tableRows, tableColumns);
            table.Alignment = Alignment.left;
            
            // Titulo
            
            AgregarTitulo(table, "CARACTERISTICAS FISICO-SENSORIALES: 5.6.6.1.1-TABLA 27");
            
            // Encabezado
            
            // Ancho de columnas
            
            int[] columnWidths = { 25, 35, 25, 25, 60, 60, 60, 60, 60, 60, 60 };
            
            for (int i = 0; i < tableColumns; i++)
            {
                table.SetColumnWidth(i, columnWidths[i]);
            }
            
            // Combinar filas
            
            table.MergeCellsInColumn(0, 1, headerRows - 1);
            table.MergeCellsInColumn(1, 1, headerRows - 1);
            table.MergeCellsInColumn(table.ColumnCount - 1, 1, headerRows - 1);
            table.Rows[0].MergeCells(0, table.ColumnCount - 1); // Titulo
            
            // Código
            
            FormatTableCell(table.Rows[1].Cells[0], "CODIGO", 8, true, Alignment.center, true, TextDirection.btLr);
            
            // Vías (n)
            
            FormatTableCell(table.Rows[1].Cells[1], "#VIAS (n)", 7, true, Alignment.center);
            
            // Numero de aceptación
            
            table.Rows[1].MergeCells(2, 3);
            FormatTableCell(table.Rows[1].Cells[2], "NÚMERO DE ACEPTACIÓN", 6, true, Alignment.center);
            
            FormatTableCell(table.Rows[2].Cells[2], "No", 6, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[3], "(c)*", 6, true, Alignment.center);
            
            // Especie
            
            FormatTableCell(table.Rows[1].Cells[3], "ESPECIE", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[4], "Corresponde a la declarada por el exportador", 6, false, Alignment.center);
            
            // Presentación
            
            FormatTableCell(table.Rows[1].Cells[4], "PRESENTACIÓN", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[5], "Corresponde a la declarada por el exportador y debe incluir todos los aspectos señalados por éste (ejemplo: tipo de corte, tipo de empaque, entre otros)", 6, false, Alignment.center);
            
            // Aspecto
            
            FormatTableCell(table.Rows[1].Cells[5], "ASPECTO", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[6], "Normal. Ausencia de materias extrañas, no existen zonas micóticas, ni moho halófilo. Ausencia de quemaduras por excesivo calentamiento durante el secado evidenciadas por una piel viscosa o pegajosa.", 6, false, Alignment.center);
            
            // Olor
            
            FormatTableCell(table.Rows[1].Cells[6], "OLOR", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[7], "Propio, característico. Ausencia de olores objetables, persistentes e inconfundibles que sean signos de descomposición (olor acido, pútrido. Etc.) o de contaminación pos sustancias extrañas (comestibles, productos de limpieza, etc).", 6, false, Alignment.center);
            
            // Color
            
            FormatTableCell(table.Rows[1].Cells[7], "COLOR", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[8], "Natural, típico y uniforme. No se permite la presencia de manchas rojizas o verdosas ni decoloración amarilla o naranja amarillenta.", 6, false, Alignment.center);
            
            // Textura
            
            FormatTableCell(table.Rows[1].Cells[8], "TEXTURA", 5, true, Alignment.center);
            FormatTableCell(table.Rows[2].Cells[9], "Típica de acuerdo al producto. Ausencia de carne con textura caracterizada por agnetamiento generalizado en más de dos tercios de la superficie, desgarrada o rota.", 6, false, Alignment.center);

            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[9], "CONCLUSIÓN", 8, true, Alignment.center);
            
            // Muestras
            
            for (int i = 0, iRowIndex= headerRows; i < codigoVias.Count; i++, iRowIndex++)
            {
                
                // Codigo
                
                FormatTableCell(table.Rows[iRowIndex].Cells[0], codigoVias[i].CodigoInterno, 8, true, Alignment.center, false);
                FormatTableCell(table.Rows[iRowIndex].Cells[1], codigoVias[i].Vias, 7, true, Alignment.center, false);
                
            }
            
            // Descripción
            
            AgregarDescripcion(table, "El paréntesis en el número de aceptación (c) indica el número de aceptación para descomposición");

            // Guardar

            document.InsertTable(table);
            
        }
        
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
        
        private static void AgregarDescripcion(Table table, string descripcion)
        {
            var lastRow = table.Rows[table.Rows.Count - 1];
            lastRow.MergeCells(0, table.ColumnCount - 1);
            
            lastRow.Cells[0].SetBorder(TableCellBorderType.InsideH, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.InsideV, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Top, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Bottom, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Left, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            lastRow.Cells[0].SetBorder(TableCellBorderType.Right, new Border(BorderStyle.Tcbs_none, 0, 0, Color.Transparent));
            
            FormatTableCell(lastRow.Cells[0], descripcion, 6, false, Alignment.left, false);
        }
        
    }
}
