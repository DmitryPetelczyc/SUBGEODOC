using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excell = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace DOCSUBGE
{
    internal class Replacer
    {
        private FileInfo doc_file;
        private Workbooks _workbooks;
        private Workbook _workbook;

        public Replacer()
        {

        }

        internal void ReplaceExcell(string file_name, Form1.excell_items_struct[] items, string path)
        {
            string worddoc_path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\")) + "\\WORD\\";

            try
            {
                //File.Copy(Environment.CurrentDirectory + "\\WORD\\" + file_name, Path.Combine(path, file_name), true);
                File.Copy(worddoc_path + file_name, Path.Combine(path, file_name), true);
            }
            catch(System.IO.DirectoryNotFoundException ex)
            {
                MessageBox.Show(ex.Message);
            }
            

            string doc_name = path + "\\" + file_name;
            object missing = Type.Missing;
            Excell.Application app = new Excell.Application();

            app.DisplayAlerts = false;
            app.ScreenUpdating = false;
            app.Visible = false;
            app.UserControl = false;
            app.Interactive = false;

            if (File.Exists(doc_name))
            {
                doc_file = new FileInfo(doc_name);
            }
            else
            {
                MessageBox.Show("FILE: " + doc_name + " \nNOT FOUND");
                return;
            }

            try
            {
                _workbooks = app.Workbooks;
                _workbook = _workbooks.Open(doc_name);

                for (int i = 0; i <= 9; i++)
                {
                    ((Worksheet)app.ActiveSheet).Cells[items[i].row, items[i].col] = items[i].data; //.value
                }

                Object newFileName = Path.Combine(path, doc_file.Name);
                
                _workbook.Save();
                
                _workbook.Close(SaveChanges: false);
                _workbooks.Close();
                app.Quit();
                Marshal.FinalReleaseComObject(_workbook);
                Marshal.FinalReleaseComObject(_workbooks);
                Marshal.FinalReleaseComObject(app);

                newFileName = null;

                killExcell();
            }            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void killExcell()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0) { PK.Kill(); }
            }
        }

        internal void ReplaceWord(Dictionary<string, string> items, string file_name, string path, string alternative_name = "")
        {
            string worddoc_path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\")) + "\\WORD\\";
            string doc_name = worddoc_path + file_name; //Environment.CurrentDirectory + "\\WORD\\" + file_name;
            Word.Application app = new Word.Application();

            if (File.Exists(doc_name))
            {
                doc_file = new FileInfo(doc_name);
            }
            else
            {
                MessageBox.Show("FILE: " + doc_name + " \nNOT FOUND");
                return;
            }

            try
            {
                app = new Word.Application();
                Object file = doc_file.FullName;
                Object missing = Type.Missing;

                app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                                MatchCase: false,
                                MatchWholeWord: false,
                                MatchSoundsLike: missing,
                                MatchWildcards: false,
                                MatchAllWordForms: false,
                                Forward: true,
                                Wrap: wrap,
                                Format: false,
                                ReplaceWith: missing,
                                Replace: replace);
                }

                if (alternative_name == "")
                    alternative_name = doc_file.Name;

                Object newFileName = Path.Combine(path, alternative_name);

                app.ActiveDocument.SaveAs2(newFileName,
                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Bląd! Jeden z plikow 'MS Word' zostal otwarty przez uwytkownika. Zamkni te plik(i) i spróbuj ponownie. \nPodejrzewany plik: " + doc_name);
                CloseWord(app);
            }
            catch (Exception ex)
            {
                CloseWord(app);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                CloseWord(app);
            }
        }

        private void CloseWord(Word.Application app)
        {
            if (app != null)
            {
                app.ActiveDocument.Close();
                app.Quit();
            }
        }

    }
}
