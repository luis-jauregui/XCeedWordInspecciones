using System;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace XCeedWordInspeccion
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            
            string filename = "Inspecciones.docx";
            string templatePath = @"C:\Users\ljauregui\RiderProjects\XCeedWord\XCeedWord\bin\Debug\PlantillaAC.docx";
            
            File.Copy(templatePath, filename, true);

            using (DocX document = DocX.Load(filename))
            {
                
                // Configurar márgenes
                
                document.MarginLeft   = 25; // 2.5 cm
                document.MarginRight  = 20; // 2 cm
                document.MarginTop    = 0; // 2 cm
                document.MarginBottom = 0; // 2 cm
                
                // Crear Tabla
                
                // CreateTableA(document, 3, 3);
                // CreateTableB(document, 5, 3);
                CreateTableC(document, 9, 1);

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

        public static void CreateTableA(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 3;
            int headerColumns = 6;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns + totalVias;
            
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
            
            table.Rows[0].MergeCells(3, 3 + totalVias - 1);
            table.Rows[1].MergeCells(3, 3 + totalVias - 1);
            FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", "Arial", 3, true, Alignment.center);
            FormatTableCell(table.Rows[1].Cells[3], "(1)", "Arial", 3, true, Alignment.center);
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < totalVias; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], $"n{i+1}", "Arial", 3, true, Alignment.center);
            }
            
            // Conclusión

            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", "Arial", 3, true, Alignment.center);

            // Guardar
            
            document.InsertTable(table);
        }

        public static void CreateTableB(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 3;
            int headerColumns = 7;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns + totalVias;
            
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
            
            // Combinas filas
            
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
            
            // Número de Vías

            Console.WriteLine(table.Rows[1].Cells.Count);
            
            table.Rows[1].MergeCells(5, 5 + totalVias - 1);
            FormatTableCell(table.Rows[1].Cells[5], "NÚMERO DE VÍAS", "Arial", 4, true, Alignment.center);
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 6; i < totalVias; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], $"n{i+1}", "Arial", 4, true, Alignment.center);
            }
            
            // Conclusión
            
            FormatTableCell(table.Rows[1].Cells[table.Rows[1].Cells.Count - 1], "CONCLUSIÓN", "Arial", 4, true, Alignment.center);
            
            // Guardar
            
            document.InsertTable(table);
        }
        
        public static void CreateTableC(DocX document, int totalVias, int totalEnsayos)
        {
            int headerRows = 3;
            int headerColumns = 6;

            int tableRows = headerRows + totalEnsayos;
            int tableColumns = headerColumns + totalVias;
            
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
            
            table.Rows[0].MergeCells(3, 3 + totalVias - 1);
            table.Rows[1].MergeCells(3, 3 + totalVias - 1);
            FormatTableCell(table.Rows[0].Cells[3], "DISTRIBUCIÓN DE MUESTRAS", "Arial", 4, true, Alignment.center);
            FormatTableCell(table.Rows[1].Cells[3], "(1)", "Arial", 4, true, Alignment.center);
            
            // Vias (Dinámicas)
                
            for (int i = 0, iCellIndex= 5; i < totalVias; i++, iCellIndex++)
            {
                FormatTableCell(table.Rows[2].Cells[iCellIndex], $"n{i+1}", "Arial", 4, true, Alignment.center);
            }
            
            // Conclusión

            FormatTableCell(table.Rows[0].Cells[table.Rows[0].Cells.Count - 1], "CONCLUSIÓN", "Arial", 4, true, Alignment.center);

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
    }
}