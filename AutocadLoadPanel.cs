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
    public partial class AutocadLoadPanel
    {
        //arrays of commands
        private List<commandCellsNOTtoUpload> commandsCellsNOTtoUpload;
        private List<commandCellsToUpload> commandsCellsToUpload;

        //initialise arrays of commands as new empty lists
        private void AutocadLoadPanel_Load(object sender, RibbonUIEventArgs e)
        {
            commandsCellsNOTtoUpload = new List<commandCellsNOTtoUpload>();
            commandsCellsToUpload = new List<commandCellsToUpload>();
        }

        //method checks elemets of interface and retun string containing result
        //if ckeck is passed return null
        private string checkElements()
        {
            //check interface
            string message = null;
            RibbonEditBox[] editBoxesToCheck = {editBoxArtFirst, editBoxArtLast, editBoxDef,
                editBoxPic, editBoxQuan, editBoxName, editBoxPath};
            bool flag = false;
            foreach (RibbonEditBox currentBox in editBoxesToCheck)
            {
                if (currentBox.Text == "")
                {
                    flag = true;
                    break;
                }
            }
            if (flag)
            {
                message = $"Необходимо заполнить:{(char)10}";
                foreach (RibbonEditBox currentBox in editBoxesToCheck)
                {
                    if (currentBox.Text == "")
                    {
                        message += $"{(char)10}- {currentBox.Label}";
                    }
                }
                return message;
            }
            //check item number and item name columns
            Excel.Range artCells = ThisAddIn.activeWorksheet.get_Range(this.editBoxArtFirst.Text + ":"
                + this.editBoxArtLast.Text);
            if (artCells.Item[1].Column != artCells.Item[artCells.Count].Column)
            {
                message = $"Ячейки в разделе {(char)34}Первый столбец с текстом{(char)34} должны быть из одного столбца";
                return message;
            }
            //check cell not to upload
            Excel.Range[] cellsToDelete = getRanges(this.editBoxCellsNotToCopy.Text);
            if (cellsToDelete != null)
            {
                int count = 0;
                for (int i = 1; i <= artCells.Count; ++i)
                {
                    foreach (Excel.Range cell in cellsToDelete)
                    {
                        if (artCells.Item[i].Row == cell.Row)
                        {
                            ++count;
                        }
                    }
                }
                if (count != cellsToDelete.Length)
                {
                    message = $"Ячейки в разделе {(char)34}Не выгружать{(char)34} должны быть в строках " +
                        $"между первой и последней строкой раздела {(char)34}Первый столбец с текстом{(char)34}";
                }
            }
            return message;
        }

        //method builds instance of the loader class using editboxes'es values
        private loader buildLoader()
        {
            Excel.Range artCells = ThisAddIn.activeWorksheet.get_Range(this.editBoxArtFirst.Text + ":"
                + this.editBoxArtLast.Text);
            Excel.Range[] cellsToDelete = getRanges(this.editBoxCellsNotToCopy.Text);
            Excel.Range defRange = ThisAddIn.activeWorksheet.get_Range(this.editBoxDef.Text);
            Excel.Range picRange = ThisAddIn.activeWorksheet.get_Range(this.editBoxPic.Text);
            Excel.Range quanRange = ThisAddIn.activeWorksheet.get_Range(this.editBoxQuan.Text);
            string name = this.editBoxName.Text;
            string path = this.editBoxPath.Text;

            loader ldr = new Upload(artCells, cellsToDelete, defRange, picRange,
                quanRange, name, path, ThisAddIn.activeWorksheet);
            return ldr;
        }

        //method accepts string containing addresses, parses it and returns array of Range objects
        private Excel.Range[] getRanges(string addresses)
        {
            int size = 0;
            Excel.Range[] addRange = new Excel.Range[size];
            if (addresses == "")
            {
                addRange = null;
            }
            else
            {
                do
                {
                    Excel.Range[] temp = addRange;
                    ++size;
                    addRange = new Excel.Range[size];
                    if (temp.Length > 0)
                    {
                        for (int i = 0; i < temp.Length; ++i)
                        {
                            addRange[i] = temp[i];
                        }
                    }
                    string currentAddress = addresses.Remove(addresses.IndexOf(" "),
                        addresses.Length - addresses.IndexOf(" "));
                    addRange[addRange.Length - 1] = ThisAddIn.activeWorksheet.get_Range(currentAddress);
                    addresses = addresses.Remove(0, addresses.IndexOf(" ") + 1);
                } while (addresses.Length != 0);
            }
            return addRange;
        }

        //on click on adding address button this method subscribes to ckange select worksheet event
        //adds address to needed editbox and initialise instance of needed command
        public void fillAddress()
        {
            ThisAddIn.thisApp.ScreenUpdating = false;
            string previousAddress = ThisAddIn.sender.getEditBox().Text;
            if (ThisAddIn.sender.getAddress() != null)
            {
                if (ThisAddIn.sender.adding())
                    ThisAddIn.sender.getEditBox().Text += ThisAddIn.sender.getAddress().Replace("$", "") + " ";
                else
                    ThisAddIn.sender.getEditBox().Text = ThisAddIn.sender.getAddress().Replace("$", "");
            }
            else
            {
                MessageBox.Show("Выберите ячейку для добавления адреса", "Error");
            }
            ThisAddIn.sender.getToggleButton().Checked = false;
            if (ThisAddIn.sender.getEditBox().Name == this.editBoxCellsNotToCopy.Name)
            {
                this.buildCommandCellsNOTtoUpload(ThisAddIn.sender.getAddress());
            }
            else
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(ThisAddIn.sender.getAddress());
                if (previousAddress != "")
                {
                    this.buildCommandCellsToUpload(ThisAddIn.activeWorksheet.get_Range(previousAddress));
                }
                commandsCellsToUpload.Add(new commandCellsToUpload(cell, 
                    $"Выгрузка в AutoCad:{(char)10}{ThisAddIn.sender.getEditBox().Label}"));
            }
            ThisAddIn.sender.hasAddress -= this.fillAddress;
            ThisAddIn.sender.disable();
            ThisAddIn.thisApp.ScreenUpdating = true;
        }

        //method builds command
        private void buildCommandCellsNOTtoUpload(string address)
        {
            ThisAddIn.thisApp.ScreenUpdating = false;
            Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(address);
            commandCellsNOTtoUpload command = new commandCellsNOTtoUpload(cell);
            commandsCellsNOTtoUpload.Add(command);
            ThisAddIn.thisApp.ScreenUpdating = true;
        }

        //method builds command
        private void buildCommandCellsToUpload(Excel.Range cell)
        {
            ThisAddIn.thisApp.ScreenUpdating = false;
            commandsCellsToUpload.ForEach(delegate (commandCellsToUpload command)
            {
                if (command.getCell().Address == cell.Address)
                {
                    command.reset();
                    commandsCellsToUpload.Remove(command);
                }
            });
            ThisAddIn.thisApp.ScreenUpdating = true;
        }

        //method unchecks all togglebuttons
        private void uncheckButtons(RibbonToggleButton thisButton)
        {
            ThisAddIn.sender.disable();
            ThisAddIn.sender.hasAddress -= this.fillAddress;
            RibbonToggleButton[] buttons = { toggleButtonArtFirst, toggleButtonArtLast,
                toggleButtonDef, toggleButtonPic};
            foreach (RibbonToggleButton button in buttons)
            {
                if (button != thisButton) button.Checked = false;
            }
        }

        //UPLOAD BUTTON
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string exMessage = this.checkElements();
                if (exMessage == null)
                {
                    const string message = "Создать файл выгрузки ?";
                    const string caption = "Выгрузка";
                    var result = MessageBox.Show(message, caption,
                                                 MessageBoxButtons.YesNo,
                                                 MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        loader loader = this.buildLoader();
                        loader.upload();
                    }
                }
                else
                {
                    MessageBox.Show(exMessage);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //adding addresses buttuns
        //enables ThisAddin.sender and subscribes this.fillAddress to ThisAddin.sender.hasAddress event
        private void toggleButtonArtFirst_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonArtFirst);
            if (toggleButtonArtFirst.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxArtFirst, this.toggleButtonArtFirst, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void toggleButtonArtLast_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonArtLast);
            if (toggleButtonArtLast.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxArtLast, this.toggleButtonArtLast, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void toggleButtonDef_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonDef);
            if (toggleButtonDef.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxDef, this.toggleButtonDef, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void toggleButtonPic_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonPic);
            if (toggleButtonPic.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxPic, this.toggleButtonPic, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void toggleButtonQuan_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonQuan);
            if (toggleButtonQuan.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxQuan, this.toggleButtonQuan, false);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }
        private void toggleButtonCellsNotToCopy_Click(object sender, RibbonControlEventArgs e)
        {
            this.uncheckButtons(toggleButtonCellsNotToCopy);
            if (toggleButtonCellsNotToCopy.Checked)
            {
                ThisAddIn.sender.enable(this.editBoxCellsNotToCopy, this.toggleButtonCellsNotToCopy, true);
                ThisAddIn.sender.hasAddress += this.fillAddress;
            }
            else
            {
                ThisAddIn.sender.disable();
                ThisAddIn.sender.hasAddress -= this.fillAddress;
            }
        }

        //clear buttons
        private void toggleButtonArtFirstClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxArtFirst.Text != "")
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(this.editBoxArtFirst.Text);
                buildCommandCellsToUpload(cell);
                this.editBoxArtFirst.Text = "";
            }
            toggleButtonArtFirstClear.Checked = false;
        }
        private void toggleButtonArtLastClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxArtLast.Text != "")
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(this.editBoxArtLast.Text);
                buildCommandCellsToUpload(cell);
                this.editBoxArtLast.Text = "";
            }
            toggleButtonArtLastClear.Checked = false;
        }
        private void toggleButtonDefClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxDef.Text != "")
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(this.editBoxDef.Text);
                buildCommandCellsToUpload(cell);
                this.editBoxDef.Text = "";
            }
            toggleButtonDefClear.Checked = false;
        }
        private void toggleButtonPicClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxPic.Text != "")
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(this.editBoxPic.Text);
                buildCommandCellsToUpload(cell);
                this.editBoxPic.Text = "";
            }
            toggleButtonPicClear.Checked = false;
        }
        private void toggleButtonQuanClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxQuan.Text != "")
            {
                Excel.Range cell = ThisAddIn.activeWorksheet.get_Range(this.editBoxQuan.Text);
                buildCommandCellsToUpload(cell);
                this.editBoxQuan.Text = "";
            }
            toggleButtonQuanClear.Checked = false;
        }

        //resets all commands in commandsCellsNOTtoUpload list and clears list
        private void toggleButtonCellsNotToCopyClear_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.editBoxCellsNotToCopy.Text != "")
            {
                const string message = "Удалить адреса ячеек, которые не надо выгружать ?";
                const string caption = "Очистить";
                var result = MessageBox.Show(message, caption,
                                             MessageBoxButtons.YesNo,
                                             MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    ThisAddIn.thisApp.ScreenUpdating = false;
                    this.editBoxCellsNotToCopy.Text = "";
                    commandsCellsNOTtoUpload.ForEach(delegate (commandCellsNOTtoUpload command)
                    {
                        command.reset();
                    });
                    commandsCellsNOTtoUpload = new List<commandCellsNOTtoUpload>();
                    ThisAddIn.thisApp.ScreenUpdating = true;
                }
            }
            toggleButtonCellsNotToCopyClear.Checked = false;
        }
        private void toggleButtonPathClear_Click(object sender, RibbonControlEventArgs e)
        {
            this.editBoxPath.Text = "";
            toggleButtonPathClear.Checked = false;
        }

        //buttons adding name and path of new book to uplead
        private void buttonCreateName_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (this.editBoxPath.Text == "")
                {
                    this.editBoxPath.Text = ThisAddIn.thisWorkbook.Path + (char)92;
                }
                string name;
                if (this.editBoxName.Text == "")
                {
                    name = "ВЫГРУЗКА " + ThisAddIn.thisWorkbook.Name
                        .Replace("Коммерческое предложение от ", "")
                        .Replace(".xlsm", "")
                        .Replace(".xlsx", "");
                }
                else
                {
                    name = this.editBoxName.Text;
                }
                if (File.Exists(this.editBoxPath.Text + name + ".xlsx"))
                {
                    int fileIndex = 1;
                    do
                    {
                        if (File.Exists(this.editBoxPath.Text + name + $" ({fileIndex}).xlsx"))
                        {
                            ++fileIndex;
                            continue;
                        }
                        else
                        {
                            this.editBoxName.Text = name + $" ({fileIndex})";
                            break;
                        }

                    } while (true);
                }
                else
                {
                    this.editBoxName.Text = name;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void buttonCreatePath_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                FolderBrowserDialog folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
                DialogResult result = folderBrowserDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    this.editBoxPath.Text = folderBrowserDialog1.SelectedPath + (char)92;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //TEST button, not visible
        private void TEST_button_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
