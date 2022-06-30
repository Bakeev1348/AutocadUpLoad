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
using System.Drawing;
using System.IO;

namespace AutocadUpLoad
{
    //Functor classes save formatting of cell needed to mark,
    //edit it and can restore previous formatting
    interface command
    {
        void reset();
    }

    class commandCellsToUpload : command
    {
        protected Excel.Range _range;
        protected System.Drawing.Color _bordersColor;
        protected System.Drawing.Color _newBordersColor;
        protected string _comment;
        protected int[] _bordersStyle;
        public commandCellsToUpload()
        {
            _range = null;
            _bordersColor = System.Drawing.Color.Black;
            _bordersStyle = new int[4];
            _newBordersColor = System.Drawing.Color.Black;
            _comment = null;
        }
        public commandCellsToUpload(Excel.Range range, string comment)
        {
            _range = range;
            _bordersColor = System.Drawing.ColorTranslator.FromOle((int)_range.Borders.Color);
            _bordersStyle = new int[4];
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _bordersStyle[i] = (int)_range.Borders.Item[currentBorder].LineStyle;
            }
            _newBordersColor = System.Drawing.Color.Green;
            _comment = comment;
            this.formatCell();
        }
        protected virtual void formatCell()
        {
            _range.Item[1].AddComment(_comment);
            _range.Item[1].Comment.Shape.TextFrame.Characters().Font.Size = 30;
            _range.Item[1].Comment.Shape.Width = 600;
            _range.Item[1].Comment.Shape.Height = 100;
            _range.Borders.Color = _newBordersColor;
            _range.Borders.Weight = 4;
        }
        public virtual void reset()
        {
            _range.Item[1].ClearComments();
            _range.Borders.Color = _bordersColor;
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _range.Borders.Item[currentBorder].Weight = 2;
                _range.Borders.Item[currentBorder].LineStyle = _bordersStyle[i];
            }
        }
        public Excel.Range getCell()
        {
            return _range;
        }
    }

    class commandCellsNOTtoUpload : commandCellsToUpload, command
    {
        public commandCellsNOTtoUpload(Excel.Range range)
        {
            _range = range;
            _bordersColor = System.Drawing.ColorTranslator.FromOle((int)_range.Borders.Color);
            _bordersStyle = new int[4];
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _bordersStyle[i] = (int)_range.Borders.Item[currentBorder].LineStyle;
            }
            _newBordersColor = System.Drawing.Color.Red;
            _comment = $"Выгрузка в AutoCad:{(char)10}Ячейка НЕ будет выгружена";
            this.formatCell();
        }
    }

    class commandBordersReset : commandCellsToUpload, command
    {
        public commandBordersReset(Excel.Range range)
        {
            _range = range;
            _bordersColor = System.Drawing.ColorTranslator.FromOle((int)_range.Borders.Color);
            _bordersStyle = new int[4];
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _bordersStyle[i] = (int)_range.Borders.Item[currentBorder].LineStyle;
            }
            _newBordersColor = System.Drawing.Color.White;
            this.formatCell();
        }
        protected override void formatCell()
        {
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _range.Borders.Item[currentBorder].LineStyle = -4142;
            }
        }

        public override void reset()
        {
            _range.Borders.Color = _bordersColor;
            for (int i = 0; i < 4; i++)
            {
                Excel.XlBordersIndex currentBorder = (Excel.XlBordersIndex)(i + 7);
                _range.Borders.Item[currentBorder].Weight = 2;
                _range.Borders.Item[currentBorder].LineStyle = _bordersStyle[i];
            }
        }
    }
}
