using ExcelDataReader;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Generator_V3
{
    public partial class Form1 : Form
    {
        private string filename = string.Empty;//Путь к файлу Excel
        private string filename2 = string.Empty;//Путь к файлу Word        
        public DataTableCollection tableCollection = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void OpenExcelFile(string path)
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
                    tableCollection = db.Tables;
                    toolStripComboBox2.Items.Clear();
                    foreach (DataTable tbl in tableCollection)
                    {
                        toolStripComboBox2.Items.Add(tbl.TableName);
                    }
                    toolStripComboBox2.SelectedIndex = 1;
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

            finally
            {

            }
        }

        private void SelectWord_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog2.ShowDialog();
                if (res == DialogResult.OK)
                {
                    filename2 = openFileDialog2.FileName;
                    textBox1SelectWord.Text = filename2;
                    Word.Application app = new Word.Application();
                    Object missing = Type.Missing;
                    app.Documents.Open(filename2);
                    app.Visible = false;
                    //
                    //Получение имени закладки и создание соответствующе именованых чекбоксов
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
                        //checkedListBox2.Items.Add(app.ActiveDocument.Tables[i].Title + i);
                        //if (app.ActiveDocument.Tables[i].Title != null)
                        checkedListBox2.Items.Add("Таблица " + i + " " + app.ActiveDocument.Tables[i].Title);
                    }

                    // Получение текста из нижних и верхних колонтитулов
                    foreach (Word.Section section in app.ActiveDocument.Sections)
                    {
                        // Нижние колонтитулы
                        foreach (Word.HeaderFooter footer in section.Footers)
                        {
                            string FooterText = (footer.Range == null || footer.Range.Text.Replace("\r", "").Trim() == null) ? null : footer.Range.Text.Replace("\r", "").Trim();
                            if (FooterText != null)
                            {
                                //if (FooterText != "")
                                //{
                                //    checkedListBox2.Items.Add(FooterText);
                                //}

                                //string proba = (FooterText.Substring(FooterText.IndexOf("{")));
                                //checkedListBox2.Items.Add(proba);

                                /* Обработка текста */
                            }
                            //checkedListBox2.Items.Add(FooterText);
                        }
                        ArrayList HeaderList = new ArrayList();
                        // Верхние колонтитулы
                        foreach (Word.HeaderFooter header in section.Headers)
                        {
                            string HeaderText = (header.Range == null || header.Range.Text.Replace("\r", "").Trim() == null) ? null : header.Range.Text.Replace("\r", "").Trim();
                            if (HeaderText != null)
                            {
                                //if (HeaderText != "")
                                //{
                                //    checkedListBox2.Items.Add(HeaderText);
                                //}
                                //string proba = (HeaderText.Substring(HeaderText.IndexOf("$"), HeaderText.IndexOf("$")));                             

                                /* Обработка текста */
                            }
                        }
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

        private void SelectExcel_Click(object sender, EventArgs e)
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

        private void ToolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
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
                    //comboBox1.Items.Add(tbl.Columns[l].ColumnName.ToString());
                }

                comboBox1.SelectedIndex = 0;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
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

        private void SelectPathSave_Click(object sender, EventArgs e)
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

        private void Generation_Click(object sender, EventArgs e)
        {
            //
            //Получение количества столбцов и строк для внесение имен переменных
            //и использовании индексов в дальнейшем
            //
            int NumberOfColumnsDGV1 = dataGridView1.Columns.GetColumnCount(0);
            int LastRowIndexDGV1 = dataGridView1.Rows.GetLastRow(0);

            int NumberOfColumnsDGV2 = dataGridView2.Columns.GetColumnCount(0);
            int LastRowIndexDGV2 = dataGridView2.Rows.GetLastRow(0);

            Word.Application app = new Word.Application();

            //Переменная содержит путь куда складывать готовые файлы
            string PathFolder = textBoxSelectPathSave.Text;

            try
            {
                for (int m = 0; m < LastRowIndexDGV1; m++)
                {
                    //Создаём переменную, которая указывает с какого столбца брать имена для файлов,
                    //далее загоняем переменную в цикл
                    int IndexSelect = comboBox1.SelectedIndex;

                    /*Создаём листы в которых будут храниться имена всех столбцов (наименования переменных)
                    в нашем случае мы выбрали HeaderText в dataGridView1 */
                    ArrayList list = new ArrayList();
                    for (int i = 0; i < NumberOfColumnsDGV1; i++)
                    {
                        list.Add(dataGridView1.Columns[i].HeaderText.ToString());
                    }

                    ArrayList list2 = new ArrayList();
                    for (int i = 0; i < NumberOfColumnsDGV2; i++)
                    {
                        list2.Add(dataGridView2.Columns[i].HeaderText.ToString());
                    }

                    //
                    //Задаём имя новому файлу
                    //
                    object fileName = textBoxSelectPathSave.Text + "\\" + dataGridView1.Rows[m].Cells[IndexSelect].Value.ToString() + ".docx";
                    object oMissing = System.Reflection.Missing.Value;
                    object oEndOfDoc = "\\endofdoc"; /* \endofdoc это предопределенная закладка */
                    //
                    //Путь до файла шаблона Word
                    //
                    string fileName2 = textBox1SelectWord.Text;
                    Object missing = Type.Missing;
                    app.Documents.Open(fileName2);

                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        string BookmarkName = checkedListBox1.Items[i].ToString();
                        if (checkedListBox1.GetItemChecked(i) == false)
                        {
                            app.ActiveDocument.Bookmarks[BookmarkName].Range.Delete();
                        }
                    }


                    int NumberCheckTable = checkedListBox2.Items.Count;
                    for (int i = 0; i < NumberCheckTable; i++)
                    {
                        if (checkedListBox2.GetItemChecked(i) == true)
                        {
                            Word.Table tbl1 = app.ActiveDocument.Tables[1];
                            int NumberRowsTable = tbl1.Rows.Count;
                            for (int z = NumberRowsTable; z > 0; z--)
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
                    //
                    //Срост данных таблицы и шаблона документа Word
                    //
                    {
                        for (int index = 0; index < NumberOfColumnsDGV1; index++)
                        {
                            Word.Find find1 = app.Selection.Find;
                            find1.Text = "{$" + (string)list[index] + "$}";
                            find1.Replacement.Text = dataGridView1.Rows[m].Cells[(string)list[index]].Value.ToString();
                            Object wrap = Word.WdFindWrap.wdFindContinue;
                            Object replace = Word.WdReplace.wdReplaceAll;

                            find1.Execute(FindText: Type.Missing,
                                MatchCase: false,
                                MatchWholeWord: false,
                                MatchWildcards: false,

                                MatchSoundsLike: missing,
                                MatchAllWordForms: false,
                                Forward: true,
                                Wrap: wrap,
                                Format: false,
                                ReplaceWith: missing, Replace: replace);
                        }
                    }

                    {
                        for (int index = 0; index < NumberOfColumnsDGV2; index++)
                        {
                            Word.Find find2 = app.Selection.Find;
                            find2.Text = "[$" + (string)list2[index] + "$]";
                            ArrayList spisok = new ArrayList();
                            for (int i = 0; i < LastRowIndexDGV2; i++)
                            {
                                if (dataGridView2.Rows[i].Cells[index].Value.ToString() != "")
                                    spisok.Add(dataGridView2.Rows[i].Cells[index].Value.ToString());
                            }
                            int SizeArraySpisok = spisok.Count - 1;
                            for (int a = 0; a <= SizeArraySpisok; a++)
                            {
                                if (a < SizeArraySpisok)
                                    find2.Replacement.Text = (string)spisok[a] + "^p" + "[$" + (string)list2[index] + "$]";
                                else if (a == SizeArraySpisok)
                                {
                                    find2.Replacement.Text = (string)spisok[a];
                                }

                                Object wrap2 = Word.WdFindWrap.wdFindContinue;
                                Object replace2 = Word.WdReplace.wdReplaceAll;
                                find2.MatchPhrase = false;
                                find2.Execute(FindText: Type.Missing,
                                    MatchCase: false,
                                    MatchWholeWord: false,
                                    MatchWildcards: false,
                                    MatchSoundsLike: missing,
                                    MatchAllWordForms: false,
                                    Forward: true,
                                    Wrap: wrap2,
                                    Format: false,
                                    ReplaceWith: missing, Replace: replace2);
                            }
                        }
                    }                   

                    app.ActiveDocument.AcceptAllRevisions();
                    app.ActiveDocument.SaveAs2(ref fileName);
                }

                MessageBox.Show("Готовые Файлы находятся " + PathFolder);
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
    }
}
