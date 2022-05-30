
namespace AutocadUpLoad
{
    partial class AutocadLoadPanel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AutocadLoadPanel()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBoxArtFirst = this.Factory.CreateRibbonEditBox();
            this.toggleButtonArtFirst = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonArtFirstClear = this.Factory.CreateRibbonToggleButton();
            this.editBoxArtLast = this.Factory.CreateRibbonEditBox();
            this.toggleButtonArtLast = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonArtLastClear = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.editBoxDef = this.Factory.CreateRibbonEditBox();
            this.toggleButtonDef = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonDefClear = this.Factory.CreateRibbonToggleButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.editBoxPic = this.Factory.CreateRibbonEditBox();
            this.toggleButtonPic = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonPicClear = this.Factory.CreateRibbonToggleButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editBoxQuan = this.Factory.CreateRibbonEditBox();
            this.toggleButtonQuan = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonQuanClear = this.Factory.CreateRibbonToggleButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.editBoxCellsNotToCopy = this.Factory.CreateRibbonEditBox();
            this.toggleButtonCellsNotToCopy = this.Factory.CreateRibbonToggleButton();
            this.toggleButtonCellsNotToCopyClear = this.Factory.CreateRibbonToggleButton();
            this.editBoxName = this.Factory.CreateRibbonEditBox();
            this.editBoxPath = this.Factory.CreateRibbonEditBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.buttonCreateName = this.Factory.CreateRibbonButton();
            this.buttonCreatePath = this.Factory.CreateRibbonButton();
            this.toggleButtonPathClear = this.Factory.CreateRibbonToggleButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.buttonUPLOAD = this.Factory.CreateRibbonButton();
            this.TEST_button = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Label = "Выгрузка ведомости в AutoCad";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBoxArtFirst);
            this.group1.Items.Add(this.toggleButtonArtFirst);
            this.group1.Items.Add(this.toggleButtonArtFirstClear);
            this.group1.Items.Add(this.editBoxArtLast);
            this.group1.Items.Add(this.toggleButtonArtLast);
            this.group1.Items.Add(this.toggleButtonArtLastClear);
            this.group1.Label = "Первый столбец с текстом";
            this.group1.Name = "group1";
            // 
            // editBoxArtFirst
            // 
            this.editBoxArtFirst.Enabled = false;
            this.editBoxArtFirst.Label = "Первый столбец c текстом, начало";
            this.editBoxArtFirst.Name = "editBoxArtFirst";
            this.editBoxArtFirst.ScreenTip = "Адрес первой ячейки первого столбца с текстом";
            this.editBoxArtFirst.ShowLabel = false;
            this.editBoxArtFirst.Text = null;
            // 
            // toggleButtonArtFirst
            // 
            this.toggleButtonArtFirst.Label = "Указать начало";
            this.toggleButtonArtFirst.Name = "toggleButtonArtFirst";
            this.toggleButtonArtFirst.ScreenTip = "Указать адрес первой ячейки первого столбца с текстом";
            this.toggleButtonArtFirst.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Для тог" +
    "о, чтобы выйти из режима выбора ячейки, нажмите кнопку ещё раз.";
            this.toggleButtonArtFirst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonArtFirst_Click);
            // 
            // toggleButtonArtFirstClear
            // 
            this.toggleButtonArtFirstClear.Label = "Очистить";
            this.toggleButtonArtFirstClear.Name = "toggleButtonArtFirstClear";
            this.toggleButtonArtFirstClear.ScreenTip = "Очистить";
            this.toggleButtonArtFirstClear.SuperTip = "Удаляет адрес первой ячейки первого столбца с текстом.";
            this.toggleButtonArtFirstClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonArtFirstClear_Click);
            // 
            // editBoxArtLast
            // 
            this.editBoxArtLast.Enabled = false;
            this.editBoxArtLast.Label = "Первый столбец с текстом, конец";
            this.editBoxArtLast.Name = "editBoxArtLast";
            this.editBoxArtLast.ScreenTip = "Адрес последней ячейки первого столбца с текстом";
            this.editBoxArtLast.ShowLabel = false;
            this.editBoxArtLast.Text = null;
            // 
            // toggleButtonArtLast
            // 
            this.toggleButtonArtLast.Label = "Указать конец";
            this.toggleButtonArtLast.Name = "toggleButtonArtLast";
            this.toggleButtonArtLast.ScreenTip = "Указать адрес последней ячейки первого столбца с текстом";
            this.toggleButtonArtLast.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Для тог" +
    "о, чтобы выйти из режима выбора ячейки, нажмите кнопку ещё раз.";
            this.toggleButtonArtLast.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonArtLast_Click);
            // 
            // toggleButtonArtLastClear
            // 
            this.toggleButtonArtLastClear.Label = "Очистить";
            this.toggleButtonArtLastClear.Name = "toggleButtonArtLastClear";
            this.toggleButtonArtLastClear.ScreenTip = "Очистить";
            this.toggleButtonArtLastClear.SuperTip = "Удаляет адрес последней ячейки первого столбца с текстом.";
            this.toggleButtonArtLastClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonArtLastClear_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.editBoxDef);
            this.group2.Items.Add(this.toggleButtonDef);
            this.group2.Items.Add(this.toggleButtonDefClear);
            this.group2.Label = "Второй столбец с текстом";
            this.group2.Name = "group2";
            // 
            // editBoxDef
            // 
            this.editBoxDef.Enabled = false;
            this.editBoxDef.Label = "Второй столбец с текстом";
            this.editBoxDef.Name = "editBoxDef";
            this.editBoxDef.ScreenTip = "Адрес ячейки из второго столбца с текстом";
            this.editBoxDef.ShowLabel = false;
            this.editBoxDef.Text = null;
            // 
            // toggleButtonDef
            // 
            this.toggleButtonDef.Label = "Указать";
            this.toggleButtonDef.Name = "toggleButtonDef";
            this.toggleButtonDef.ScreenTip = "Указать адрес ячейки из второго столбца с текстом";
            this.toggleButtonDef.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Для тог" +
    "о, чтобы выйти из режима выбора ячейки, нажмите кнопку ещё раз.";
            this.toggleButtonDef.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonDef_Click);
            // 
            // toggleButtonDefClear
            // 
            this.toggleButtonDefClear.Label = "Очистить";
            this.toggleButtonDefClear.Name = "toggleButtonDefClear";
            this.toggleButtonDefClear.ScreenTip = "Очистить";
            this.toggleButtonDefClear.SuperTip = "Удаляет адрес ячейки из второго столбца с текстом.";
            this.toggleButtonDefClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonDefClear_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.editBoxPic);
            this.group3.Items.Add(this.toggleButtonPic);
            this.group3.Items.Add(this.toggleButtonPicClear);
            this.group3.Label = "Столбец с картинками";
            this.group3.Name = "group3";
            // 
            // editBoxPic
            // 
            this.editBoxPic.Enabled = false;
            this.editBoxPic.Label = "Столбец с картинками";
            this.editBoxPic.Name = "editBoxPic";
            this.editBoxPic.ScreenTip = "Адрес ячейки из столбца с картинками";
            this.editBoxPic.ShowLabel = false;
            this.editBoxPic.Text = null;
            // 
            // toggleButtonPic
            // 
            this.toggleButtonPic.Label = "Указать";
            this.toggleButtonPic.Name = "toggleButtonPic";
            this.toggleButtonPic.ScreenTip = "Указать адрес ячейки из столбца с картинками";
            this.toggleButtonPic.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Для тог" +
    "о, чтобы выйти из режима выбора ячейки, нажмите кнопку ещё раз.";
            this.toggleButtonPic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPic_Click);
            // 
            // toggleButtonPicClear
            // 
            this.toggleButtonPicClear.Label = "Очистить";
            this.toggleButtonPicClear.Name = "toggleButtonPicClear";
            this.toggleButtonPicClear.ScreenTip = "Очистить";
            this.toggleButtonPicClear.SuperTip = "Удаляет адрес ячейки из столбца с картинками.";
            this.toggleButtonPicClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPicClear_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.editBoxQuan);
            this.group4.Items.Add(this.toggleButtonQuan);
            this.group4.Items.Add(this.toggleButtonQuanClear);
            this.group4.Label = "Столбец с количеством";
            this.group4.Name = "group4";
            // 
            // editBoxQuan
            // 
            this.editBoxQuan.Enabled = false;
            this.editBoxQuan.Label = "Столбец с количеством";
            this.editBoxQuan.Name = "editBoxQuan";
            this.editBoxQuan.ScreenTip = "Адрес ячейки из столбца с количеством";
            this.editBoxQuan.ShowLabel = false;
            this.editBoxQuan.Text = null;
            // 
            // toggleButtonQuan
            // 
            this.toggleButtonQuan.Label = "Указать";
            this.toggleButtonQuan.Name = "toggleButtonQuan";
            this.toggleButtonQuan.ScreenTip = "Указать адрес ячейки из столбца с количеством";
            this.toggleButtonQuan.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Для тог" +
    "о, чтобы выйти из режима выбора ячейки, нажмите кнопку ещё раз.";
            this.toggleButtonQuan.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonQuan_Click);
            // 
            // toggleButtonQuanClear
            // 
            this.toggleButtonQuanClear.Label = "Очистить";
            this.toggleButtonQuanClear.Name = "toggleButtonQuanClear";
            this.toggleButtonQuanClear.ScreenTip = "Очистить";
            this.toggleButtonQuanClear.SuperTip = "Удаляет адрес ячейки из столбца с количеством.";
            this.toggleButtonQuanClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonQuanClear_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.editBoxCellsNotToCopy);
            this.group5.Items.Add(this.toggleButtonCellsNotToCopy);
            this.group5.Items.Add(this.toggleButtonCellsNotToCopyClear);
            this.group5.Items.Add(this.editBoxName);
            this.group5.Items.Add(this.editBoxPath);
            this.group5.Items.Add(this.label1);
            this.group5.Items.Add(this.buttonCreateName);
            this.group5.Items.Add(this.buttonCreatePath);
            this.group5.Items.Add(this.toggleButtonPathClear);
            this.group5.Label = "Настройки";
            this.group5.Name = "group5";
            // 
            // editBoxCellsNotToCopy
            // 
            this.editBoxCellsNotToCopy.Enabled = false;
            this.editBoxCellsNotToCopy.Label = "Не выгружать";
            this.editBoxCellsNotToCopy.Name = "editBoxCellsNotToCopy";
            this.editBoxCellsNotToCopy.ScreenTip = "Адреса ячеек, которые не нужно выгружать";
            this.editBoxCellsNotToCopy.SizeString = "111111111111111111";
            this.editBoxCellsNotToCopy.Text = null;
            // 
            // toggleButtonCellsNotToCopy
            // 
            this.toggleButtonCellsNotToCopy.Label = "Указать";
            this.toggleButtonCellsNotToCopy.Name = "toggleButtonCellsNotToCopy";
            this.toggleButtonCellsNotToCopy.ScreenTip = "Указать адреса ячеек, которые не нужно выгружать";
            this.toggleButtonCellsNotToCopy.SuperTip = "Нажмите на кнопку, после этого кликните на ячейку, которую нужно указать. Можно у" +
    "казать несколько ячеек.  Для того, чтобы выйти из режима выбора ячейки, нажмите " +
    "кнопку ещё раз.";
            this.toggleButtonCellsNotToCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonCellsNotToCopy_Click);
            // 
            // toggleButtonCellsNotToCopyClear
            // 
            this.toggleButtonCellsNotToCopyClear.Label = "Очистить";
            this.toggleButtonCellsNotToCopyClear.Name = "toggleButtonCellsNotToCopyClear";
            this.toggleButtonCellsNotToCopyClear.ScreenTip = "Очистить";
            this.toggleButtonCellsNotToCopyClear.SuperTip = "Удаляет адреса ячеек, которые не нужно выгружать.";
            this.toggleButtonCellsNotToCopyClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonCellsNotToCopyClear_Click);
            // 
            // editBoxName
            // 
            this.editBoxName.Label = "Имя файла";
            this.editBoxName.Name = "editBoxName";
            this.editBoxName.ScreenTip = "Имя файла для сохранения";
            this.editBoxName.SizeString = "1111111111111111111111111111111111111";
            this.editBoxName.SuperTip = "Редактируется. Не следует указывать символы, которые не могут быть в названии фай" +
    "ла в Windows.";
            this.editBoxName.Text = null;
            // 
            // editBoxPath
            // 
            this.editBoxPath.Enabled = false;
            this.editBoxPath.Label = "Путь сохранения";
            this.editBoxPath.Name = "editBoxPath";
            this.editBoxPath.ScreenTip = "Путь сохранения файла";
            this.editBoxPath.SizeString = "1111111111111111111111111111111111111";
            this.editBoxPath.Text = null;
            // 
            // label1
            // 
            this.label1.Label = " ";
            this.label1.Name = "label1";
            // 
            // buttonCreateName
            // 
            this.buttonCreateName.Label = "Сформировать имя и путь";
            this.buttonCreateName.Name = "buttonCreateName";
            this.buttonCreateName.ScreenTip = "Сформировать имя и путь сохранения по умолчанию";
            this.buttonCreateName.SuperTip = "Имя файла создаётся на основании имени оригинального файла. Путь по умолчанию - п" +
    "апка, в которой находится оригинальный файл.";
            this.buttonCreateName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateName_Click);
            // 
            // buttonCreatePath
            // 
            this.buttonCreatePath.Label = "Указать";
            this.buttonCreatePath.Name = "buttonCreatePath";
            this.buttonCreatePath.ScreenTip = "Указать другой путь сохранения";
            this.buttonCreatePath.SuperTip = "Открывает окно выбора директории.";
            this.buttonCreatePath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreatePath_Click);
            // 
            // toggleButtonPathClear
            // 
            this.toggleButtonPathClear.Label = "Очистить путь";
            this.toggleButtonPathClear.Name = "toggleButtonPathClear";
            this.toggleButtonPathClear.ScreenTip = "Очистить путь";
            this.toggleButtonPathClear.SuperTip = "Удаляет путь сохранения.";
            this.toggleButtonPathClear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleButtonPathClear_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.buttonUPLOAD);
            this.group6.Items.Add(this.TEST_button);
            this.group6.Name = "group6";
            // 
            // buttonUPLOAD
            // 
            this.buttonUPLOAD.Label = "Выгрузить";
            this.buttonUPLOAD.Name = "buttonUPLOAD";
            this.buttonUPLOAD.ScreenTip = "Выгрузить";
            this.buttonUPLOAD.SuperTip = "Создаёт новый файл Excel, в который выгружается содержимое в соответствие с настр" +
    "ойками.";
            this.buttonUPLOAD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // TEST_button
            // 
            this.TEST_button.Label = "TEST";
            this.TEST_button.Name = "TEST_button";
            this.TEST_button.Visible = false;
            this.TEST_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TEST_button_Click);
            // 
            // AutocadLoadPanel
            // 
            this.Name = "AutocadLoadPanel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AutocadLoadPanel_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxArtFirst;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonArtFirst;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonArtFirstClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxArtLast;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonArtLast;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonArtLastClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxDef;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonDef;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonDefClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxPic;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPic;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPicClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxQuan;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonQuan;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonQuanClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxCellsNotToCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonCellsNotToCopy;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonCellsNotToCopyClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxName;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUPLOAD;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreatePath;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButtonPathClear;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TEST_button;
    }

    partial class ThisRibbonCollection
    {
        internal AutocadLoadPanel AutocadLoadPanel
        {
            get { return this.GetRibbon<AutocadLoadPanel>(); }
        }
    }
}
