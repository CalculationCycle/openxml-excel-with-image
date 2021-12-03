using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

//using Cotecna.Voc.Web.Models;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SkiaSharp;

//namespace Cotecna.Voc.Web.Common
namespace Custom.Excel
{
    public static class ExcelGenearator
    {
        /// <summary>
        /// Generate an excel file with the information of certificates
        /// </summary>
        /// <param name="dataSource">The list of certificates</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream GenerateExcel()
        {
            MemoryStream ms = new MemoryStream();
            
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
                //create the new workbook
                WorkbookPart workbookPart = document.AddWorkbookPart();
                Workbook workbook = new Workbook();
                workbookPart.Workbook = workbook;

                //  If we don't add a "WorkbookStylesPart", OLEDB will refuse to connect to this .xlsx file !
                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");

                //get and save the stylesheet
                Stylesheet stylesheet = StyleSheet();
                workbookStylesPart.Stylesheet = stylesheet;
                workbookStylesPart.Stylesheet.Save();

                //add the new workseet
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                Worksheet worksheet = new Worksheet();
                SheetData sheetData1 = new SheetData();

                Sheets sheets = new Sheets();

                //get the string name of the columns
                string[] excelColumnNamesTitle = new string[5];
                for (int n = 0; n < 5; n++)
                    excelColumnNamesTitle[n] = GetExcelColumnName(n);

                Row rowTitle = new Row() { RowIndex = (UInt32Value)1 };
                AppendTextCell(excelColumnNamesTitle[0], "Image", rowTitle, 1);
                AppendTextCell(excelColumnNamesTitle[1], "Style", rowTitle, 1);
                AppendTextCell(excelColumnNamesTitle[2], "Color", rowTitle, 1);
                AppendTextCell(excelColumnNamesTitle[3], "Gender", rowTitle, 1);
                AppendTextCell(excelColumnNamesTitle[4], "ArticleNumber", rowTitle, 1);
                sheetData1.Append(rowTitle);

                rowTitle = new Row() { RowIndex = (UInt32Value)2 };
                AppendTextCell(excelColumnNamesTitle[0], "<no img>", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[1], "Sudo Pjuck", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[2], "Navy Green", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[3], "Men/Women", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[4], "F0001001", rowTitle, 0);
                sheetData1.Append(rowTitle);

                rowTitle = new Row() { RowIndex = (UInt32Value)3 };
                AppendTextCell(excelColumnNamesTitle[0], "<img!>", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[1], "Sudo Pjuck", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[2], "Yellow", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[3], "Men/Women", rowTitle, 0);
                AppendTextCell(excelColumnNamesTitle[4], "F0001002", rowTitle, 0);
                sheetData1.Append(rowTitle);

                var rowToSetHeightFor = sheetData1.Elements<Row>().ElementAt(2);
                rowToSetHeightFor.CustomHeight = true;
                rowToSetHeightFor.Height = 60.0;

                // Set column width
                var columns = new Columns();
                columns.Append(new Column() { Min = 1, Max = 1, CustomWidth = true, Width = 20.0 });
                columns.Append(new Column() { Min = 2, Max = 2, CustomWidth = true, Width = 18.0 });
                columns.Append(new Column() { Min = 3, Max = 3, CustomWidth = true, Width = 18.0 });
                columns.Append(new Column() { Min = 4, Max = 4, CustomWidth = true, Width = 18.0 });
                columns.Append(new Column() { Min = 5, Max = 5, CustomWidth = true, Width = 24.0 });
                worksheet.Append(columns);

                /*
                var colToSetWidthFor = sheetData1.Elements<Column>().ElementAt(0);
                colToSetWidthFor.CustomWidth = true;
                colToSetWidthFor.Width = 120.0;
                
                colToSetWidthFor = sheetData1.Elements<Column>().ElementAt(1);
                colToSetWidthFor.CustomWidth = true;
                colToSetWidthFor.Width = 100.0;

                colToSetWidthFor = sheetData1.Elements<Column>().ElementAt(2);
                colToSetWidthFor.CustomWidth = true;
                colToSetWidthFor.Width = 100.0;

                colToSetWidthFor = sheetData1.Elements<Column>().ElementAt(3);
                colToSetWidthFor.CustomWidth = true;
                colToSetWidthFor.Width = 80.0;
                */

                //add the information of the current sheet
                worksheet.Append(sheetData1);

                Drawing drawing = AddLogo("monkey_logo_200x200.png", worksheetPart);

                //add merged cells
                //worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                worksheet.Append(drawing);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                //create the new sheet for this report
                Sheet sheet = new Sheet() { Name = "NameOfSheet", SheetId = (UInt32Value)1, Id = workbookPart.GetIdOfPart(worksheetPart) };
                sheets.Append(sheet);
                //add the new sheet to the report
                workbook.Append(sheets);
                //save all report
                workbook.Save();
                //close the stream.
                
                document.Close();
            }

            return ms;
        }

        #region Helper funcs
        /// <summary>
        /// Append text in a cell
        /// </summary>
        /// <param name="cellReference">Reference</param>
        /// <param name="cellStringValue">Value</param>
        /// <param name="excelRow">Excel row</param>
        /// <param name="styleIndex">Style</param>
        private static void AppendNumberCell(string cellReference, string cellStringValue, Row excelRow, UInt32Value styleIndex)
        {
            //  Add a new Excel Cell to our Row
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.Number };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue;
            cell.StyleIndex = styleIndex;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }

        /// <summary>
        /// Append text in a cell
        /// </summary>
        /// <param name="cellReference">Reference</param>
        /// <param name="cellStringValue">Value</param>
        /// <param name="excelRow">Excel row</param>
        /// <param name="styleIndex">Style</param>
        private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow, UInt32Value styleIndex)
        {
            //  Add a new Excel Cell to our Row
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue;
            cell.StyleIndex = styleIndex;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }

        /// <summary>
        /// Create a column and set the height of the column
        /// </summary>
        /// <param name="StartColumnIndex">Initial index</param>
        /// <param name="EndColumnIndex">End index</param>
        /// <param name="ColumnWidth">Column width</param>
        /// <returns></returns>
        private static Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)
        {
            Column column;
            column = new Column();
            column.Min = StartColumnIndex;
            column.Max = EndColumnIndex;
            column.Width = ColumnWidth;
            column.CustomWidth = true;
            return column;
        }

        /// <summary>
        /// Get an excel column
        /// </summary>
        /// <param name="columnIndex">index</param>
        /// <returns></returns>
        private static string GetExcelColumnName(int columnIndex)
        {
            //  Convert a zero-based column index into an Excel column reference  (A, B, C.. Y, Y, AA, AB, AC... AY, AZ, B1, B2..)
            //
            //  eg  GetExcelColumnName(0) should return "A"
            //      GetExcelColumnName(1) should return "B"
            //      GetExcelColumnName(25) should return "Z"
            //      GetExcelColumnName(26) should return "AA"
            //      GetExcelColumnName(27) should return "AB"
            //      ..etc..
            //
            if (columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            char firstChar = (char)('A' + (columnIndex / 26) - 1);
            char secondChar = (char)('A' + (columnIndex % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }

        /// <summary>
        /// Update the value of a existing cell
        /// </summary>
        /// <param name="cellReference">Address of the cell</param>
        /// <param name="cellStringValue">Value to update</param>
        /// <param name="excelRow">Row of the cell</param>
        private static void UpdateNumberCellValue(string cellReference, string cellStringValue, Row excelRow, UInt32Value styleIndex = null)
        {
            Cell currentCell = excelRow.Elements<Cell>().First(cell => cell.CellReference.Value == cellReference);
            currentCell.CellValue = new CellValue(cellStringValue);
            currentCell.DataType = new EnumValue<CellValues>(CellValues.Number);
            if (styleIndex != null)
                currentCell.StyleIndex = styleIndex;
        }

        /// <summary>
        /// Update the value of a existing cell
        /// </summary>
        /// <param name="cellReference">Address of the cell</param>
        /// <param name="cellStringValue">Value to update</param>
        /// <param name="excelRow">Row of the cell</param>
        private static void UpdateStringCellValue(string cellReference, string cellStringValue, Row excelRow, UInt32Value styleIndex = null)
        {
            Cell currentCell = excelRow.Elements<Cell>().First(cell => cell.CellReference.Value == cellReference);
            currentCell.CellValue = new CellValue(cellStringValue);
            currentCell.DataType = new EnumValue<CellValues>(CellValues.String);
            if (styleIndex != null)
                currentCell.StyleIndex = styleIndex;
        }
        #endregion Helper funcs

        #region AddLogo function
        /// <summary>
        /// Add the logo of the system
        /// </summary>
        /// <param name="logoPath">Path of the logo</param>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <returns>Drawing</returns>
        private static Drawing AddLogo(string logoPath, WorksheetPart worksheetPart)
        {
            string sImagePath = logoPath;
            DrawingsPart dp = worksheetPart.AddNewPart<DrawingsPart>();
            ImagePart imgp = dp.AddImagePart(ImagePartType.Png, worksheetPart.GetIdOfPart(dp));
            using (FileStream fs = new FileStream(sImagePath, FileMode.Open, FileAccess.Read))
            {
                imgp.FeedData(fs);
            }

            DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties nvdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties();
            nvdp.Id = 1025;
            nvdp.Name = "Picture 1";
            nvdp.Description = "logo";
            DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
            picLocks.NoChangeAspect = true;
            picLocks.NoChangeArrowheads = true;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties nvpdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties();
            nvpdp.PictureLocks = picLocks;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties nvpp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties();
            nvpp.NonVisualDrawingProperties = nvdp;
            nvpp.NonVisualPictureDrawingProperties = nvpdp;

            DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

            DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
            DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip();
            blip.Embed = dp.GetIdOfPart(imgp);
            blip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
            blipFill.Blip = blip;
            blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
            blipFill.Append(stretch);

            DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
            DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
            offset.X = 0;
            offset.Y = 0;
            t2d.Offset = offset;

            var imageFileStream = new FileStream(logoPath, FileMode.Open, System.IO.FileAccess.Read);
            SKBitmap bm = SKBitmap.Decode(imageFileStream);
            
            DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            extents.Cx = (long)bm.Width * 9525; //English Metric Units
            extents.Cy = (long)bm.Height * 9525; //English Metric Units
            bm.Dispose();

            //Bitmap bm = new Bitmap(sImagePath);
            //DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            //extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
            //extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
            //bm.Dispose();

            t2d.Extents = extents;
            DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
            sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
            sp.Transform2D = t2d;
            DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
            prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
            prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
            sp.Append(prstGeom);
            sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

            DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
            picture.NonVisualPictureProperties = nvpp;
            picture.BlipFill = blipFill;
            picture.ShapeProperties = sp;

            DocumentFormat.OpenXml.Drawing.Spreadsheet.Position pos = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Position();
            pos.X = 18 * 914400 / 72;
            pos.Y = 28 * 914400 / 72;

            Extent ext = new Extent();
            ext.Cx = extents.Cx;
            ext.Cy = extents.Cy;
            AbsoluteAnchor anchor = new AbsoluteAnchor();
            anchor.Position = pos;
            anchor.Extent = ext;

            anchor.Append(picture);
            anchor.Append(new ClientData());
            WorksheetDrawing wsd = new WorksheetDrawing();
            wsd.Append(anchor);
            Drawing drawing = new Drawing();
            drawing.Id = dp.GetIdOfPart(imgp);

            wsd.Save(dp);
            return drawing;
        }
        #endregion AddLogo function

        #region Stylesheet
        /// <summary>
        /// Create an stylesheet to use in excel files
        /// </summary>
        /// <returns></returns>
        private static Stylesheet StyleSheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            Fonts fonts = new Fonts();
            // 0 - normal fonts
            DocumentFormat.OpenXml.Spreadsheet.Font myFont = new DocumentFormat.OpenXml.Spreadsheet.Font()
            {
                FontSize = new FontSize() { Val = 11 },
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                FontName = new FontName() { Val = "Arial" }
            };
            fonts.Append(myFont);

            //1 - font bold
            myFont = new DocumentFormat.OpenXml.Spreadsheet.Font()
            {
                Bold = new Bold(),
                FontSize = new FontSize() { Val = 11 },
                Color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                FontName = new FontName() { Val = "Arial" }
            };
            fonts.Append(myFont);


            Fills fills = new Fills();
            //default fill
            Fill fill = new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            };
            fills.Append(fill);

            Borders borders = new Borders();
            //normal borders
            Border border = new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);

            CellFormats cellFormats = new CellFormats();
            //0- normal
            CellFormat cellFormat = new CellFormat()
            {
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyFill = false
            };
            cellFormats.Append(cellFormat);

            //1 -title
            cellFormat = new CellFormat()
            {
                FontId = 1,
                FillId = 0,
                BorderId = 0,
                ApplyFill = false
            };
            cellFormats.Append(cellFormat);

            styleSheet.Append(fonts);
            styleSheet.Append(fills);
            styleSheet.Append(borders);
            styleSheet.Append(cellFormats);

            return styleSheet;
        }
        #endregion Stylesheet
    }
}
