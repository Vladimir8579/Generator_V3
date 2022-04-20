using ExcelDataReader;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Generator_V3
{
    public partial class Generator : Form
    {
        private string filename = string.Empty;//Путь к файлу Excel
        private string filename2 = string.Empty;//Путь к файлу Word        
        public DataTableCollection tableCollection = null;

        public Generator()
        {
            InitializeComponent();
        }

        void GenerationButtonCheked(object sender, EventArgs e)//Включение кнопки Генерация если 3 поля заполнены
        {
            if (textBox1SelectWord.Text != "")
                if (textBoxSelectExcel.Text != "")
                    if (textBoxSelectPathSave.Text != "")
                        Generation.Enabled = true;
        }

        private void OpenExcelFile(string path)//Чтение файла Excel 
        {
            try
            {
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                //
                //=================================================================
                //
                try
                {
                    toolStripComboBox2.Items.Clear();
                    tableCollection = db.Tables;
                    toolStripComboBox1.Items.Clear();
                    foreach (DataTable tbl in tableCollection)
                    {
                        toolStripComboBox1.Items.Add(tbl.TableName);
                    }
                    toolStripComboBox1.SelectedIndex = 0;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "В Excel нет листа номер 1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //
                //=================================================================
                //
                try
                {
                    toolStripComboBox2.Items.Clear();
                    tableCollection = db.Tables;
                    toolStripComboBox2.Items.Clear();
                    if (db.Tables.Count > 1)
                    {
                        foreach (DataTable tbl in tableCollection)
                        {
                            toolStripComboBox2.Items.Add(tbl.TableName);
                        }

                        toolStripComboBox2.SelectedIndex = 1;
                    }
                    if (db.Tables.Count < 2)
                    {
                        toolStripComboBox2.Items.Clear();
                        toolStripComboBox2.Items.Add("");
                        toolStripComboBox2.SelectedIndex = 0;
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "В Excel нет листа номер 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void SelectWord_Click(object sender, EventArgs e)//Выбрать и прочитать шаблон Word
        {
            try
            {
                DialogResult res = openFileDialog2.ShowDialog();
                if (res == DialogResult.OK)
                {
                    filename2 = openFileDialog2.FileName;
                    textBox1SelectWord.Text = filename2;
                    Word.Application app = new Word.Application
                    {
                        Visible = false
                    };
                    Object missing = Type.Missing;
                    app.Documents.Open(filename2);
                    //
                    //Получение имени закладки и создание соответствующе именованных чекбоксов
                    //
                    int NumberBookmarksEnd = app.ActiveDocument.Bookmarks.Count;
                    checkedListBox1.Items.Clear();
                    for (int i = 1; i <= NumberBookmarksEnd; i++)
                    {
                        string bookmarks = app.ActiveDocument.Bookmarks[i].Name;
                        checkedListBox1.Items.Add(bookmarks);
                    }
                    //
                    //Получение количества таблиц в документе и создание чекбоксов
                    //
                    int TableNumber = app.ActiveDocument.Tables.Count;
                    checkedListBox2.Items.Clear();
                    for (int i = 1; i <= TableNumber; i++)
                    {
                        checkedListBox2.Items.Add("Таблица " + i + " " + app.ActiveDocument.Tables[i].Title);
                    }

                    app.Documents.Close();
                    app.Quit();
                }

                else
                {
                    throw new Exception("Файл не выбран");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Закройте шаблон документа и процесс в диспетчере", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectExcel_Click(object sender, EventArgs e)//Открыть файл Excel 
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;
                    Text = filename;
                    textBoxSelectExcel.Text = filename.ToString();
                    OpenExcelFile(filename);
                }
                else
                {
                    throw new Exception("Файл не выбран");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)// Обработка значений из 1-го листа Exсel
        {
            try
            {
                DataTable tbl = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];

                dataGridView1.DataSource = tbl;

                int NumberOfColumn = dataGridView1.Columns.GetColumnCount(0);
                comboBox1.Items.Clear();
                for (int l = 0; l < NumberOfColumn; l++)
                {
                    comboBox1.Items.Add(dataGridView1.Columns[l].HeaderText.ToString());
                    //comboBox1.Items.Add(tbl.Columns[l].ColumnName);                   
                }

                comboBox1.SelectedIndex = 0;
                LblStatus.Text = "Будет создано " + dataGridView1.Rows.GetLastRow(0).ToString() + " комплект(a)(ов) документов";
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)// Обработка значений из 2-го листа Exсel
        {
            try
            {
                DataTable table2 = tableCollection[Convert.ToString(toolStripComboBox2.SelectedItem)];
                dataGridView2.DataSource = table2;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectPathSave_Click(object sender, EventArgs e)// Обработка выбора пути сохранения готовых файлов
        {
            try
            {
                FolderBrowserDialog FBD = new FolderBrowserDialog
                {
                    ShowNewFolderButton = false
                };
                if (FBD.ShowDialog() == DialogResult.OK)

                {
                    textBoxSelectPathSave.Text = FBD.SelectedPath;
                }
                else
                {
                    throw new Exception("Папка куда сохранить файлы не выбрана");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Generation_Click(object sender, EventArgs e)// Генерация актов
        {
            //Получение количества столбцов и строк для внесение имен переменных
            //и использовании индексов в дальнейшем
            //
            int CountColumnsDGV1 = dataGridView1.Columns.GetColumnCount(0);
            int CountRowDGV1 = dataGridView1.Rows.GetLastRow(0);

            int CountColumnsDGV2 = dataGridView2.Columns.GetColumnCount(0);
            int CountRowDGV2 = dataGridView2.Rows.GetLastRow(0);
            //
            //Инициализация переменной для счётчика итераций
            //
            int Counter = 0;
            //
            //
            //
            ProgressBar.Maximum = dataGridView1.Rows.GetLastRow(0);
            Word.Application app = new Word.Application();

            //Переменная содержит путь куда складывать готовые файлы
            string PathFolder = textBoxSelectPathSave.Text;

            try
            {
                for (int m = 0; m < CountRowDGV1; m++)
                {
                    //Создаём переменную, которая указывает с какого столбца брать имена для файлов,
                    //далее загоняем переменную в цикл
                    int IndexSelect = comboBox1.SelectedIndex;

                    /*Создаём листы в которых будут храниться имена всех столбцов (наименование переменных)
                    в нашем случае мы выбрали HeaderText в dataGridView1 */
                    ArrayList list = new ArrayList();
                    for (int i = 0; i < CountColumnsDGV1; i++)
                    {
                        list.Add(dataGridView1.Columns[i].HeaderText.ToString());
                    }

                    ArrayList list2 = new ArrayList();
                    for (int i = 0; i < CountColumnsDGV2; i++)
                    {
                        list2.Add(dataGridView2.Columns[i].HeaderText.ToString());
                    }
                    //
                    //Задаём имя новому файлу
                    //                    
                    object fileNameEkz1Docx = PathFolder + "\\" + "Экз №1 " + dataGridView1.Rows[m].Cells[IndexSelect].Value.ToString() + ".docx";
                    object fileNameEkz2Docx = PathFolder + "\\" + "Экз №2 " + dataGridView1.Rows[m].Cells[IndexSelect].Value.ToString() + ".docx";

                    object fileNameEkz1Pdf = PathFolder + "\\" + "Экз №1 " + dataGridView1.Rows[m].Cells[IndexSelect].Value.ToString() + ".pdf";
                    object fileNameEkz2Pdf = PathFolder + "\\" + "Экз №2 " + dataGridView1.Rows[m].Cells[IndexSelect].Value.ToString() + ".pdf";

                    object oMissing = System.Reflection.Missing.Value;
                    object oEndOfDoc = "\\endofdoc"; /* \endofdoc это предопределенная закладка */
                    //
                    //Путь до файла шаблона Word
                    //
                    string fileName2 = textBox1SelectWord.Text;
                    Object missing = Type.Missing;
                    app.Documents.Open(fileName2);
                    //
                    //Удаление не отмеченных закладок
                    //
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        string BookmarkName = checkedListBox1.Items[i].ToString();
                        if (checkedListBox1.GetItemChecked(i) == false)
                        {
                            app.ActiveDocument.Bookmarks[BookmarkName].Range.Delete();
                        }
                    }
                    //
                    //Получаем номер таблицы в открытом документе WORD и отмеченную в checkedListBox2,
                    //которую нужно почистить от пустых строк и проверяем построчно ячейки 2, 3, 4
                    // если три ячейки подряд в строке пустые удаляем строку
                    //
                    int NumberCheckTable = checkedListBox2.Items.Count;

                    for (int i = 0; i < NumberCheckTable; i++)
                    {
                        if (checkedListBox2.GetItemChecked(i) == true)
                        {
                            Word.Table tbl1 = app.ActiveDocument.Tables[i + 1];
                            int NumberRowsTable = tbl1.Rows.Count;

                            for (int z = NumberRowsTable; z > 0; z--)
                            {
                                int NumberCellTable = tbl1.Rows[z].Cells.Count;

                                if ((NumberCellTable > 4) == true)
                                {

                                    if (string.IsNullOrEmpty(tbl1.Rows[z].Cells[2].Range.Text.Replace("\r\a", "").Trim()) == true)
                                    {

                                        if (string.IsNullOrEmpty(tbl1.Rows[z].Cells[3].Range.Text.Replace("\r\a", "").Trim()) == true)
                                        {

                                            if (string.IsNullOrEmpty(tbl1.Rows[z].Cells[4].Range.Text.Replace("\r\a", "").Trim()) == true)
                                            {
                                                tbl1.Rows[z].Delete();
                                            }

                                        }

                                    }

                                }

                            }

                        }

                    }
                    //
                    //Срост данных таблицы и шаблона документа Word
                    //
                    object replace = Word.WdReplace.wdReplaceAll;
                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Word.Find find = app.Selection.Find;
                    object fileformat = Word.WdSaveFormat.wdFormatPDF;
                    int SectionCount = app.ActiveDocument.Sections.Count;
                    //
                    //Блок 1 поиск и замена одиночных переменных по документу и колонтитулам
                    //
                    {
                        for (int index = 0; index < CountColumnsDGV1; index++)
                        {
                            find.Text = "{$" + (string)list[index] + "$}";// что меняем переменные в шаблоне
                            find.Replacement.Text = dataGridView1.Rows[m].Cells[(string)list[index]].Value.ToString();// на что меняем значение переменных из Excel
                            find.Execute(FindText: Type.Missing, Wrap: wrap, ReplaceWith: missing, Replace: replace);

                            object FindTextFooter = "{$" + (string)list[index] + "$}";// что меняем
                            object ReplaceWithFooter = dataGridView1.Rows[m].Cells[(string)list[index]].Value.ToString(); // на что меняем
                            for (int i = 1; i <= SectionCount; i++)
                            {
                                app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute(FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                                app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute(FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                                app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute(FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                            }
                        }
                    }
                    //
                    //Блок 2 Поиск и замена списочных переменных по документу
                    //
                    {
                        for (int index = 0; index < CountColumnsDGV2; index++)
                        {
                            find.Text = "[$" + (string)list2[index] + "$]";// что меняем
                            ArrayList spisok = new ArrayList();
                            for (int i = 0; i < CountRowDGV2; i++)
                            {
                                if (dataGridView2.Rows[i].Cells[index].Value.ToString() != "")
                                    spisok.Add(dataGridView2.Rows[i].Cells[index].Value.ToString());
                            }
                            int SizeArraySpisok = spisok.Count - 1;
                            for (int a = 0; a <= SizeArraySpisok; a++)
                            {
                                if (a < SizeArraySpisok)
                                    find.Replacement.Text = (string)spisok[a] + "^p" + "[$" + (string)list2[index] + "$]";// на что меняем
                                else if (a == SizeArraySpisok)
                                {
                                    find.Replacement.Text = (string)spisok[a];
                                }
                                find.Execute(FindText: Type.Missing, Wrap: wrap, ReplaceWith: missing, Replace: replace);
                            }
                        }
                    }

                    object Mirror = app.ActiveDocument.PageSetup.MirrorMargins;
                    if (app.ActiveDocument.Comments.Count > 0)
                    {
                        app.Application.ActiveDocument.DeleteAllComments();
                    }
                    app.ActiveDocument.AcceptAllRevisions();
                    //
                    //Сохранение в выбранном формате и количестве экземпляров с установкой номера экземпляра в колонтитуле
                    //
                    if (((int)numericUpDown1.Value == 1) == true)// Если один экземпляр
                    {
                        object FindTextHeaders = "Экз"; // что меняем
                        object ReplaceWithHeaders = "Экз. №1"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                        }
                        app.ActiveDocument.SaveAs2(ref fileNameEkz1Docx);

                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz1Pdf, fileformat);//Сохраняем в формате PDF

                        Counter++;
                        ProgressBar.Value = Counter;
                        ProgressBar.Update();
                        LblStatus.Text = "Выполнено " + Counter + " из " + CountRowDGV1;
                    }

                    if (((int)numericUpDown1.Value == 2) == true)// Если два экземпляра
                    {
                        object FindTextHeaders = "Экз"; // что меняем
                        object ReplaceWithHeaders = "Экз. №1"; // на что меняем
                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                            (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz1Docx);

                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz1Pdf, fileformat);//Сохраняем в формате PDF

                        object FindTextHeaders2 = "Экз. №1"; // что меняем
                        object ReplaceWithHeaders2 = "Экз. №2"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                            (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);

                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz2Docx);
                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz2Pdf, fileformat);//Сохраняем в формате PDF
                        Counter++;
                        ProgressBar.Value = Counter;
                        ProgressBar.Update();
                        LblStatus.Text = "Выполнено " + Counter + " из " + CountRowDGV1;
                    }
                }

                MessageBox.Show("Готовые файлы находятся " + PathFolder);
                ProgressBar.Value = 0;
                LblStatus.Text = "Processing....";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app?.ActiveDocument.Close(SaveChanges: 0);
                app?.Quit(SaveChanges: 0);
            }
        }

        private void Exit_Click(object sender, EventArgs e)// Завершение работы приложения
        {
            Application.Exit();
        }
    }
}