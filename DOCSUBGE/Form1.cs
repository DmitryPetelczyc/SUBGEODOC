using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Text;

namespace DOCSUBGE
{
    public partial class Form1 : Form
    {
        internal struct excell_items_struct
        {
            public int row;
            public string col;
            public string data;
        }
        private Dictionary<string, string> word_items;
        private Replacer replacer = new Replacer();
        private excell_items_struct[] ex = new excell_items_struct[10];
        private int doc_count = 0;
        private readonly string[] file_list =
        [
            "1,2 operat techniczny.docx",
            "1,2 operat techniczny_.docx",
            "3 sprawozdanie - wyrys.docx",
            "3 sprawozdanie - wyrys_.docx",
            "4 Mapa porównania.docx",
            "5 Szkic z pomiaru.docx",
            "6 wykaz współrzędnych.docx",
            "8 mapa do celów projektowych.docx",
            "8 wykaz.xlsx",
            "11 MAPA INWENTARYZACYJNA.docx"
        ];

        private string tbl_svnth_plc_txt = "Arkusz danych dotyczących budynku";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progress_label.Text = "";
            string imgs_path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..\\..\\..\\")) + "\\Images\\";
            pictureBox1.ImageLocation = imgs_path + "full_logo.png";
            pictureBox2.ImageLocation = imgs_path + "full_logo.png";
            button2.BackgroundImage = Image.FromFile(imgs_path + "trashcan.png");
        }

        private void fill_tag_list()
        {
            word_items = new Dictionary<string, string>
            {
                { "<MIE>", textBox1.Text},
                { "<OBR>", textBox2.Text},
                { "<JDE>", textBox3.Text},
                { "<JEI>", textBox4.Text},
                { "<WOJ>", textBox5.Text},
                { "<NRO>", textBox6.Text},
                { "<ADRB>", textBox7.Text},
                { "<NKA>", textBox8.Text},
                { "<MZR>", textBox9.Text},
                { "<OBJ>", textBox10.Text},
                { "<OBO>", comboBox6.Text},
                { "<GMI>", comboBox7.Text},
                { "<IDB>", textBox13.Text},
                { "<IDE>", textBox14.Text},
                { "<JER>", textBox16.Text},
                { "<PPZ>", textBox17.Text},
                { "<DZN>", textBox18.Text},
                { "<NU7>", textBox19.Text},
                { "<NU8>", textBox20.Text},
                { "<OBI>", textBox21.Text},
                { "<PPZD>", textBox22.Text},
                { "<KST>", comboBox1.Text},
                { "<MGN>", comboBox2.Text},
                { "<POW>", comboBox3.Text},
                { "<KSTO>", comboBox5.Text},
                { "<TD7>", tbl_svnth_plc_txt},
                { "<DT1>", dateTimePicker1.Value.ToString("dd.MM.yyyy")},
                { "<DT2>", dateTimePicker2.Value.ToString("dd.MM.yyyy")},
                { "<DT3>", dateTimePicker3.Value.ToString("dd.MM.yyyy")}
            };
        }

        private void fill_excel_table(excell_items_struct[] ex)
        {
            ex[0].row = 1; ex[0].col = "K"; ex[0].data = textBox5.Text;
            ex[1].row = 2; ex[1].col = "K"; ex[1].data = comboBox3.Text;
            ex[2].row = 3; ex[2].col = "K"; ex[2].data = textBox3.Text;
            ex[3].row = 4; ex[3].col = "K"; ex[3].data = textBox2.Text;
            ex[4].row = 5; ex[4].col = "K"; ex[4].data = textBox16.Text;
            ex[5].row = 6; ex[5].col = "K"; ex[5].data = textBox8.Text;
            ex[6].row = 7; ex[6].col = "K"; ex[6].data = textBox6.Text;
            ex[7].row = 19; ex[7].col = "B"; ex[7].data = textBox14.Text;
            ex[8].row = 19; ex[8].col = "H"; ex[8].data = textBox14.Text;
            ex[9].row = 60; ex[9].col = "D"; ex[9].data = "Data " + dateTimePicker2.Value.ToString("dd.MM.yyyy") + " r";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!checkBox_1i.Checked && !checkBox_1d.Checked && !checkBox_3i.Checked && !checkBox_3d.Checked &&
                !checkBox_4.Checked && !checkBox_5.Checked && !checkBox_6.Checked && !checkBox_8i.Checked &&
                 !checkBox_8d.Checked && !checkBox_11.Checked && !checkBox_7w.Checked && !checkBox_7a.Checked)
            {
                MessageBox.Show("Żaden z dokumentów nie jest zaznaczony. \nPrzejdź do zakładki USTAWIENIA i wybierz pliki, które chcesz wygenerować.", "Brak wybranych dokumentów");
                return;
            }
            else
            {
                button1.Enabled = false; //Выключаем кнопку на время работы блока

                fill_tag_list();
                fill_excel_table(ex);
                generate_documents();
                resetProgBar();

                button1.Enabled = true;
                progress_label.Text = "";

                Directory.CreateDirectory(Path.Combine(browse_label.Text, "PDF"));
                Directory.CreateDirectory(Path.Combine(browse_label.Text, "MicroMap"));

                if (checkBox_explorer.Checked)
                {
                    System.Diagnostics.Process.Start("explorer", browse_label.Text);
                }

                textBox12.Text = "Operat podpisany elektronicznie / Kierownik - Andrzej Petelczyc, Geodeta Upr Nr 13158, " + textBox8.Text + ", dn. " + dateTimePicker1.Text;
            }
        }

        private void generate_documents()
        {
            if (checkBox_1i.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [Operat techniczny.docx]";
                replacer.ReplaceWord(word_items, "1,2 operat techniczny.docx", browse_label.Text);
            }

            if (checkBox_1d.Checked)
            {
                incProgBar(25);
                progress_label.Text = "Generacja [Operat techniczny.docx]";
                replacer.ReplaceWord(word_items, "1,2 operat techniczny_.docx", browse_label.Text);
            }

            if (checkBox_3i.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [Sprawozdanie.docx]";
                replacer.ReplaceWord(word_items, "3 sprawozdanie - wyrys.docx", browse_label.Text);
            }

            if (checkBox_3d.Checked)
            {
                incProgBar(25);
                progress_label.Text = "Generacja [Sprawozdanie.docx]";
                replacer.ReplaceWord(word_items, "3 sprawozdanie - wyrys_.docx", browse_label.Text);
            }

            if (checkBox_4.Checked)
            {
                incProgBar(25);
                progress_label.Text = "Generacja [Mapa porównania.docx]";
                replacer.ReplaceWord(word_items, "4 Mapa porównania.docx", browse_label.Text);
            }

            if (checkBox_5.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [Szkic z pomiaru.docx]";
                replacer.ReplaceWord(word_items, "5 Szkic z pomiaru.docx", browse_label.Text);
            }

            if (checkBox_6.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [Wykaz współrzędnych.docx]";
                replacer.ReplaceWord(word_items, "6 wykaz współrzędnych.docx", browse_label.Text);
            }

            if (checkBox_7w.Checked || checkBox_7a.Checked)
            {
                if (textBox14.Text.Length > 0)
                {
                    word_items.Add("<IDES>", textBox14.Text.Split('.')[1]);
                }

                if (comboBox4.SelectedIndex == 1)
                {
                    replacer.ReplaceWord(word_items, "7 Wykaz zmian danych budynku.docx", browse_label.Text, textBox15.Text);
                    incProgBar(20);
                    doc_count++;

                    //MessageBox.Show("Wykaz zmian danych budynku zostal stworzony!");

                    resetProgBar();
                }
                else
                {
                    replacer.ReplaceWord(word_items, "7 arkusz_danych_.docx", browse_label.Text, textBox15.Text);
                    incProgBar(20);
                    doc_count++;

                    if (doc_count > 1)
                    {
                        textBox19.Text = "Utworzono " + doc_count.ToString() + " arkusze danych dotyczących budynku";
                    }
                    else
                    {
                        textBox19.Text = "Utworzono arkusz danych dotyczących budynku";
                    }

                    //MessageBox.Show("Arkusz danych zostal stworzony!");

                    textBox15.Text = "7 arkusz_danych_0" + doc_count.ToString();

                    resetProgBar();
                }
            }

            if (checkBox_8d.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [Mapa do celów projektowych.docx]";
                replacer.ReplaceWord(word_items, "8 mapa do celów projektowych.docx", browse_label.Text);
            }

            if (checkBox_8i.Checked)
            {
                incProgBar(25);
                progress_label.Text = "Generacja [Wykaz zmian danych ewid.xslx]";
                replacer.ReplaceExcell("8 wykaz.xlsx", ex, browse_label.Text);
            }

            if (checkBox_11.Checked)
            {
                incProgBar(10);
                progress_label.Text = "Generacja [MAPA INWENTARYZACYJNA.docx]";
                replacer.ReplaceWord(word_items, "11 MAPA INWENTARYZACYJNA.docx", browse_label.Text);
            }

            if (checkBox_sign.Checked)
            {
                create_sign();
            }

            progressBar1.Value = 100;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                browse_label.Text = folderBrowserDialog1.SelectedPath;
                browse_label.ForeColor = Color.Black;
                panel4.Enabled = true;
            }
        }

        private void incProgBar(int value)
        {
            if ((progressBar1.Value + value) <= 100)
            {
                progressBar1.Value += value;
            }
        }

        private void resetProgBar()
        {
            progressBar1.Value = 0;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_8d.Checked)
            {
                textBox20.Text = " oraz wykaz zmian danych ewidencyjnych.";
            }
            else
            {
                textBox20.Text = "";
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedIndex == 0)
            {
                label7.Enabled = false;
                textBox7.Enabled = false;
                textBox15.Text = "7 arkusz_danych_";
                textBox22.Visible = false;
                textBox17.Size = new Size(311, 29);
                textBox17.Location = new Point(73, 133);
                textBox22.PlaceholderText = "";
                textBox17.PlaceholderText = "";
                tbl_svnth_plc_txt = "Arkusz danych dotyczących budynku";
                comboBox5.Visible = false;
                comboBox1.Size = new Size(311, 28);
                comboBox1.Location = new Point(73, 105);
            }
            else
            {
                label7.Enabled = true;
                textBox7.Enabled = true;
                textBox15.Text = "7 Wykaz zmian danych budynku";
                textBox22.Visible = true;
                textBox17.Size = new Size(157, 29);
                textBox17.Location = new Point(227, 133);
                textBox22.PlaceholderText = "Przed";
                textBox17.PlaceholderText = "Po";
                tbl_svnth_plc_txt = "Wykaz zmian danych ewidencyjnych dotyczących budynku";
                comboBox5.Visible = true;
                comboBox1.Size = new Size(157, 28);
                comboBox1.Location = new Point(227, 105);
                comboBox1.DropDownWidth = 450;
                comboBox5.DropDownWidth = 450;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = textBox2.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            checkBox_1i.Checked = true;
            checkBox_1d.Checked = false;
            checkBox_3i.Checked = true;
            checkBox_3d.Checked = false;
            checkBox_4.Checked = true;
            checkBox_5.Checked = true;
            checkBox_6.Checked = true;
            checkBox_8i.Checked = true;
            checkBox_8d.Checked = false;
            checkBox_11.Checked = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            checkBox_1i.Checked = false;
            checkBox_1d.Checked = true;
            checkBox_3i.Checked = false;
            checkBox_3d.Checked = true;
            checkBox_4.Checked = true;
            checkBox_5.Checked = false;
            checkBox_6.Checked = false;
            checkBox_8i.Checked = false;
            checkBox_8d.Checked = true;
            checkBox_11.Checked = false;
        }

        public void create_sign()
        {
            string path = Path.Combine(browse_label.Text, "stopka operat.txt");

            try
            {
                // Create the file, or overwrite if the file exists.
                using (FileStream fs = File.Create(path))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes("Operat podpisany elektronicznie / Kierownik - Andrzej Petelczyc, Geodeta Upr Nr 13158, " + textBox8.Text + ", dn. " + dateTimePicker1.Text);
                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }

                // Open the stream and read it back.
                using (StreamReader sr = File.OpenText(path))
                {
                    string s = "";
                    while ((s = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(s);
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "pdf files (*.pdf)| *.pdf";
            open.Multiselect = true;
            open.Title = "Open Text Files";


            // Open the output document
            PdfDocument outputDocument = new PdfDocument();

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    // Iterate files
                    foreach (string file in open.FileNames)
                    {
                        // Open the document to import pages from it.
                        PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);

                        // Iterate pages
                        int count = inputDocument.PageCount;
                        for (int idx = 0; idx < count; idx++)
                        {
                            // Get the page from the external document...
                            PdfPage page = inputDocument.Pages[idx];
                            // ...and add it to the output document.
                            outputDocument.AddPage(page);
                        }
                    }
                    
                    // Save the document...

                    string filename;
                    if (textBox11.Text != "")
                    {
                        filename = textBox11.Text + ".pdf";
                    }
                    else
                    {
                        filename = "merged.pdf";
                    }

                    string pdf_path = Path.GetDirectoryName(open.FileNames[0]);
                    outputDocument.Save(Path.Combine(pdf_path, filename));
                    //System.Diagnostics.Process.Start("explorer", pdf_path);
                    MessageBox.Show("Sukces!", "Merge PDF");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }           
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = textBox8.Text;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "pdf files (*.pdf)| *.pdf";
            open.Multiselect = false;
            open.Title = "Open Text Files";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    // Create a new PDF document
                    PdfDocument document = PdfReader.Open(open.FileName, PdfDocumentOpenMode.Modify);

                    // Create a font
                    XFont font = new XFont("Sans", 8);

                    PdfPage page;
                    XGraphics gfx;

                    int count = document.PageCount;
                    // Create some more pages
                    for (int idx = 0; idx < count; idx++)
                    {
                        page = document.Pages[idx];
                        //gfx.Dispose();
                        gfx = XGraphics.FromPdfPage(page);

                        string text = (idx + 1) + " / " + count;
                        gfx.DrawString(text, font, XBrushes.Black, page.Width / 2, 20, XStringFormats.Default);
                    }

                    string filename;
                    if (textBox11.Text != "")
                    {
                        filename = textBox11.Text + " str" + ".pdf";
                    }
                    else
                    {
                        filename = "merged str.pdf";
                    }
                    // Save the document...
                    string pdf_path = Path.GetDirectoryName(open.FileName);
                 
                    document.Save(Path.Combine(pdf_path, filename));
                    //System.Diagnostics.Process.Start("explorer", pdf_path);
                    MessageBox.Show("Sukces!", "Page numbers");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "pdf files (*.pdf)| *.pdf";
            open.Multiselect = false;
            open.Title = "Open Text Files";

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    // Create a new PDF document
                    PdfDocument document = PdfReader.Open(open.FileName, PdfDocumentOpenMode.Modify);

                    // Create a font
                    XFont font = new XFont("Sans", 8);

                    PdfPage page;
                    XGraphics gfx;

                    int count = document.PageCount;
                    // Create some more pages
                    for (int idx = 0; idx < count; idx++)
                    {
                        page = document.Pages[idx];
                        gfx = XGraphics.FromPdfPage(page);

                        string text = textBox12.Text;
                        gfx.DrawString(text, font, XBrushes.Blue, (page.Width / 2) - 220, page.Height - 20, XStringFormats.Default);
                    }
                    
                    string filename;
                    if (textBox11.Text != "")
                    {
                        filename = textBox11.Text + " " + dateTimePicker1.Text + ".pdf";
                    }
                    else
                    {
                        filename = "merged str st.pdf";
                    }
                    // Save the document...
                    string pdf_path = Path.GetDirectoryName(open.FileName);
                    document.Save(Path.Combine(pdf_path, filename));
                    System.Diagnostics.Process.Start("explorer", pdf_path);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void checkBox_7a_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_7a.Checked || checkBox_7w.Checked)
            {
                panel1.Enabled = true;
            }
            else
            {
                panel1.Enabled = false;
            }
        }

        private void checkBox_7w_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_7a.Checked || checkBox_7w.Checked)
            {
                panel1.Enabled = true;
            }
            else
            {
                panel1.Enabled = false;
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            doc_count = 0;
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Text = "podlaskie";
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            textBox13.Clear();
            textBox14.Clear();
            //textBox15.Clear();
            textBox16.Text = "G-";
            textBox17.Clear();
            textBox18.Clear();
            textBox21.Clear();
            textBox22.Clear();

            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "sokólski";
            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox7.Text = "";

            if (comboBox4.SelectedIndex == 1)
            {
                textBox15.Text = "7 Wykaz zmian danych budynku";
            }
            else
            {
                textBox15.Text = "7 arkusz_danych_";
            }

            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            panel7.Location = new Point((this.Size.Width - panel7.Size.Width) / 2, (this.Size.Height - panel7.Size.Height) / 2);
            tabControl1.Size = new Size(770, 772);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //this.Size = new Size(785, 764);
            //panel7.Location = new Point(0, 0);
        }
    }
}