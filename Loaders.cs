using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace AutocadUpLoad
{
    interface loader
    {
        void upload();
    }

    //upload class incapsulates logic of creating new book
    //and uploading to that book item numbers, item names and pictures
    class Upload : loader
    {
        protected Excel.Range[] _artCells;
        protected Excel.Range _defCol;
        protected Excel.Range _picCol;
        protected Excel.Range _quanCol;
        protected string _name;
        protected string _path;
        protected Excel.Workbook _uploadBook;
        protected Excel.Worksheet _uploadSheet;
        protected Excel.Worksheet _originalSheet;

        //by default
        public Upload()
        {
            _artCells = null;
            _defCol = null;
            _picCol = null;
            _quanCol = null;
            _name = null;
            _path = null;
            _uploadBook = null;
            _uploadSheet = null;
            _originalSheet = null;
        }

        public Upload(Excel.Range artCells, Excel.Range[] cellsToDelete, Excel.Range defCol, Excel.Range picCol,
            Excel.Range quanCol, string name, string path, Excel.Worksheet originalSheet)
        {
            //создаем массив ячеек, по которым будем копировать
            if (cellsToDelete != null)
            {
                int newSize = artCells.Count - cellsToDelete.Length;
                _artCells = new Excel.Range[newSize];
                int i = 0;
                for (int j = 1; j <= artCells.Count; ++j)
                {
                    bool flag = true;
                    foreach (Excel.Range cell in cellsToDelete)
                    {
                        if (artCells.Item[j].Row == cell.Row)
                        {
                            flag = false;
                        }
                    }
                    if (flag)
                    {
                        _artCells[i] = artCells.Item[j];
                        ++i;
                    }
                }
            }
            else
            {
                int newSize = artCells.Count;
                _artCells = new Excel.Range[newSize];
                for (int j = 1; j <= artCells.Count; ++j)
                {
                    _artCells[j - 1] = artCells.Item[j];
                }
            }

            _defCol = defCol;
            _picCol = picCol;
            _quanCol = quanCol;
            _name = name;
            _path = path;
            _uploadBook = null;
            _uploadSheet = null;
            _originalSheet = originalSheet;

        }

        //method builds workbook to upload
        protected virtual void buildBook()
        {
            //Создаём книгу и форматируем 
            Excel.Workbook uploadBook = ThisAddIn.thisApp.Workbooks.Add();
            Excel.Worksheet uploadSheet = (Excel.Worksheet)uploadBook.Worksheets.get_Item(1);
            for (int i = 1; i <= _artCells.Length; ++i)
            {
                uploadSheet.Cells[i, 1].RowHeight = 80;
                uploadSheet.Cells[i, 1].WrapText = true;
            }
            uploadSheet.Cells[1, 1].ColumnWidth = 50;
            uploadSheet.Cells[1, 2].ColumnWidth = 20;
            uploadSheet.Cells[1, 3].ColumnWidth = 30;

            _uploadBook = uploadBook;
            _uploadSheet = uploadSheet;

            //Удаляем лишнее
            if ((int)uploadBook.Worksheets.Count > (int)1)
            {
                int listCount = uploadBook.Worksheets.Count;
                for (int i = listCount; i >= 2; --i)
                {
                    Excel.Worksheet temp = (Excel.Worksheet)uploadBook.Worksheets.get_Item(i);
                    temp.Delete();
                }
            }
        }

        //method accepts 2 cells and return string containing formula CONCATENATE values of cells
        protected string contRanges(Excel.Range range1, Excel.Range range2)
        {
            var val1 = range1.Value;
            var val2 = range2.Value;
            string result;
            if ((range1.Value == null) && (range2.Value == null))
            {
                result = $"=СЦЕПИТЬ({(char)34}{(char)34})";
            }
            else if (range1.Value == null)
            {
                val2 = val2.ToString().Replace($"{(char)34}", "");
                result = $"=СЦЕПИТЬ({(char)34}{val2.ToString()}{(char)34})";
            }
            else if (range2.Value == null)
            {
                val1 = val1.ToString().Replace($"{(char)34}", "");
                result = $"=СЦЕПИТЬ({(char)34}{val1.ToString()}{(char)34})";
            }
            else if (range1.Value.ToString() == range2.Value.ToString())
            {
                val1 = val1.ToString().Replace($"{(char)34}", "");
                result = $"=СЦЕПИТЬ({(char)34}{val1.ToString()}{(char)34})";
            }
            else
            {
                val1 = val1.ToString().Replace($"{(char)34}", "");
                val2 = val2.ToString().Replace($"{(char)34}", "");
                result = $"=СЦЕПИТЬ({(char)34}{val1.ToString()}{(char)34};" +
                    $"{(char)34}{(char)10}{(char)34};{(char)34}{val2.ToString()}{(char)34})";
            }
            return result;
        }

        //method accepts name of imageformat and return asced imagecodecInfo object
        protected virtual ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }

        //method accepts cell and saves cell as picture
        protected void saveTempPic(Excel.Range pic)
        {
            Clipboard.Clear();
            _originalSheet.Activate();
            pic.Select();
            pic.Copy();
            if (Clipboard.ContainsImage())
            {
                ImageCodecInfo myImageCodecInfo = GetEncoderInfo("image/gif");
                System.Drawing.Imaging.Encoder myEncoder =
                        System.Drawing.Imaging.Encoder.ColorDepth;
                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 8L);
                EncoderParameters myEncoderParameters = new EncoderParameters(1);
                myEncoderParameters.Param[0] = myEncoderParameter;
                System.Drawing.Image imageToSave = null;
                imageToSave = Clipboard.GetImage();
                imageToSave.Save(_path + "temp154810.gif", myImageCodecInfo, myEncoderParameters);
            }
        }

        //method accepts cell and iterator of main cycle
        //edits formatting of cell, saves cell as picture, restores previous formatting using command object
        protected void copyImg(Excel.Range cell, int iterator)
        {
            commandBordersReset commandCell = new commandBordersReset(cell);
            saveTempPic(cell);
            _uploadSheet.Activate();
            System.Drawing.Image oImage = System.Drawing.Image.FromFile(_path + "temp154810.gif");
            System.Windows.Forms.Clipboard.SetDataObject(oImage, true);
            Excel.Range newCell = _uploadSheet.Cells[iterator + 1, 3];
            newCell.Select();
            _uploadSheet.Paste();
            Excel.Shape newImg = null;
            for (int i = 1; i <= _uploadSheet.Shapes.Count; i++)
            {
                if (_uploadSheet.Shapes.Item(i).TopLeftCell.Address == newCell.Address)
                {
                    newImg = _uploadSheet.Shapes.Item(i);
                }
            }
            newImg.LockAspectRatio = Office.MsoTriState.msoTrue;
            if ((newImg.Height / newImg.Width) >= ((float)newCell.Height / (float)newCell.Width))
            {
                newImg.Height = (float)newCell.Height - 10f;
                newImg.Left = (float)newCell.Left + (((float)newCell.Width - newImg.Width) / 2f);
                newImg.Top = (float)newCell.Top + (((float)newCell.Height - newImg.Height) / 2f);
            }
            else
            {
                newImg.Width = (float)newCell.Width - 10f;
                newImg.Left = (float)newCell.Left + (((float)newCell.Width - newImg.Width) / 2f);
                newImg.Top = (float)newCell.Top + (((float)newCell.Height - newImg.Height) / 2f);
            }
            commandCell.reset();
            Clipboard.Clear();
        }

        //method with main upload cycle
        public void upload()
        {
            //create book
            ThisAddIn.thisApp.ActiveWindow.DisplayGridlines = false;
            ThisAddIn.thisApp.ActiveWindow.View = XlWindowView.xlNormalView;
            this.buildBook();
            _uploadBook.Activate();
            _uploadSheet.Activate();
            _uploadSheet.Cells[1, 1].Select();
            //upload
            for (int i = 0; i < _artCells.Length; ++i)
            {
                string formula = contRanges(_artCells[i], _originalSheet.Cells[_artCells[i].Row, _defCol.Column]);
                _uploadSheet.Cells[i + 1, 1].FormulaLocal = formula;
                if (_originalSheet.Cells[_artCells[i].Row, _quanCol.Column].Value != null)
                {
                    _uploadSheet.Cells[i + 1, 2].Value = _originalSheet.Cells[_artCells[i].Row, _quanCol.Column].Value;
                }
                this.copyImg(_originalSheet.Cells[_artCells[i].Row, _picCol.Column], i);
            }
            //edit printrange and save
            string firstAdd = _uploadSheet.Cells[1, 3].Address;
            string lastAdd = _uploadSheet.Cells[(_artCells.Length), 3].Address;
            string printRange = firstAdd + ":" + lastAdd;
            string fullName = _path + _name + ".xlsx";
            _uploadSheet.PageSetup.PrintArea = printRange;
            _uploadBook.SaveAs(fullName);
            //create PDF
            string fullNamePDF = _path + _name + ".pdf";
            _uploadBook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, fullNamePDF,
                Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, true, Type.Missing);
            _originalSheet.Activate();
            ThisAddIn.thisApp.ActiveWindow.DisplayGridlines = true;
            File.Delete(_path + "temp154810.gif");
        }
    }
}
