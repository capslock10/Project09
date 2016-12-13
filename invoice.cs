using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace invoice_sys
{
    public partial class invoice : Form
    {

        Form1 f1;
        database_function df = new database_function();
        static int csd_int = 1;
        System.Data.DataTable dt;
        public static int contract_no_sys = 0;

        public invoice()
        {

            InitializeComponent();
            WindowState = FormWindowState.Maximized;
            //Read the xml file to get the contract_no
            if (File.Exists("data.xml"))
            {
                XmlSerializer xs = new XmlSerializer(typeof(information));
                FileStream read = new FileStream("data.xml", FileMode.Open, FileAccess.Read, FileShare.Read);
                information info = (information)xs.Deserialize(read);
                contract_no_sys = Convert.ToInt32(info.Contract_no);
                read.Close();
            }
            DateTime dateTime = DateTime.Now.Date;
            invoices_tb.Text = "C" + dateTime.ToString("yyyyMM") + contract_no_sys.ToString("00000");
            dt = new System.Data.DataTable();
            
            List<String> values = df.ins_combobox_value();

            //Insert the institution values to the combo box.
            foreach(String str in values)
                ins_cb.Items.Add(str);

            Debug.WriteLine(dateTime);
            to_cb.Items.Add("Correctional Services Department");
            to_cb.SelectedIndex = 0;
            total_amount_tb.Text = "$0";
            date_tb.Text = dateTime.ToString("yyyy/MM/dd");
            

      
            
        }

        private void invoice_Load(object sender, EventArgs e)
        {
            
        }


        public void set_form1(Form1 f1)
        {
            this.f1 = f1;
        }

        private void back_btn_Click(object sender, EventArgs e)
        {
            f1.Show();
            this.Visible = false;

        }

        private void invoice_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void ins_cb_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void add_item_btn_Click(object sender, EventArgs e)
        {
            add_itemToInvoice aiti = new add_itemToInvoice();
            aiti.set_invoice_form(this);
            aiti.Show();

        }

        private void item_grid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        public void set_item_grid(List<String> values)
        {
            DateTime dateTime = DateTime.Now.Date;
            Debug.WriteLine(dateTime);
            String[] detail;
            foreach (String str in values)
            {
                String[] split_str = str.Split('|');
               detail = df.get_item_detail(split_str[0]);
                //insert item to invoice grid view. delivery date, job_no, csd, item_code, tic_code, item_des, qty, unit, unit_price, amount
                this.item_grid.Rows.Add(new Object[] {dateTime.ToString("yyyy/MM/dd"), "", "CSD" + csd_int.ToString("000"), split_str[0], "", detail[0], 1, detail[1], df.get_price(split_str[0]), 1 * df.get_price(split_str[0]) });
                csd_int++;
            }

            this.set_total_amount_textbox();

        }

        private void del_item_btn_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> selectedRows = item_grid.SelectedRows.Cast<DataGridViewRow>().ToList();
            int selectedRow = selectedRows.Count;
          
            foreach (DataGridViewRow row in selectedRows)
            {
                if (item_grid.Rows[row.Index].Cells["item_code_col"].Value == null)
                    selectedRow--;
                  
            }

           
            if (MessageBox.Show(string.Format("Do you want to delete {0} rows?", selectedRow), "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
                foreach (DataGridViewRow row in selectedRows)
                {

                    if (item_grid.Rows[row.Index].Cells["item_code_col"].Value != null)
                        this.item_grid.Rows.RemoveAt(row.Index);
                }

            this.set_total_amount_textbox();
        }

        private void item_grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }


        //When the grid cell end edited, update the total amount
        private void item_grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (item_grid.Rows[e.RowIndex].Cells["item_code_col"].Value != null)
            {
                int n;
                //Check if qunatity is it a number
                if (int.TryParse(item_grid.Rows[e.RowIndex].Cells["quantity_col"].Value.ToString(), out n))
                {
                    double new_price = Convert.ToInt32(item_grid.Rows[e.RowIndex].Cells["quantity_col"].Value.ToString()) * df.get_price(item_grid.Rows[e.RowIndex].Cells["item_code_col"].Value.ToString());
                    item_grid.Rows[e.RowIndex].Cells["item_cost"].Value = new_price;
                    set_total_amount_textbox();
                }
                else
                {
                    MessageBox.Show("Please input a number!");
                    item_grid.Rows[e.RowIndex].Cells["quantity_col"].Value = "1";
                    double new_price = Convert.ToInt32(item_grid.Rows[e.RowIndex].Cells["quantity_col"].Value.ToString()) * df.get_price(item_grid.Rows[e.RowIndex].Cells["item_code_col"].Value.ToString());
                    item_grid.Rows[e.RowIndex].Cells["item_cost"].Value = new_price;
                    set_total_amount_textbox();
                  
                }
                    
            }
            else
                item_grid.Rows[e.RowIndex].Cells["quantity_col"].Value = "";

        }

        
        //Set the total amount of the invoice
        public double set_total_amount_textbox()
        {
            int rowCount = item_grid.RowCount;
            double total_amount = 0.0;
            for (int i = 0; i < rowCount; i++)
            {
                if(item_grid.Rows[i].Cells["item_cost"].Value != null)
                total_amount += Convert.ToDouble(item_grid.Rows[i].Cells["item_cost"].Value.ToString());
            }

            total_amount_tb.Text = "$" + total_amount.ToString();
            return total_amount;
        }


        //Save the invoice
        private void save_in_btn_Click(object sender, EventArgs e)
        {
            if (!check_field())
            {
                if (!df.get_invoice_existed(invoices_tb.Text.ToString()))
                {
                    if (!check_field())
                    {
                        save_invoice(1);
                        MessageBox.Show("Invoice saved successfully.");
                    }
                }
                else
                {
                    df.set_invoice_form(this);
                    df.update_invovice(contract_no_tb.Text.ToString(), invoices_tb.Text.ToString(), docket_no_tb.Text.ToString(), date_tb.Text.ToString(), df.get_ins_id(ins_cb.Text.ToString()), set_total_amount_textbox());
                    df.delete_invoice_item(df.get_invoice_id(invoices_tb.Text.ToString()));

                    int rowCount = item_grid.RowCount;
                    List<String> item_detail;
                    for (int i = 0; i < rowCount; i++)
                    {
                        if (item_grid.Rows[i].Cells["item_cost"].Value != null)
                        {
                            item_detail = new List<string>();
                            //if(!item_grid.Rows[i].Cells["item_cost"].Value.ToString().Equals(""))
                            item_detail.Add(item_grid.Rows[i].Cells["delivery_date_col"].Value.ToString());
                            item_detail.Add(item_grid.Rows[i].Cells["job_no_col"].Value.ToString());
                            item_detail.Add(item_grid.Rows[i].Cells["csd_col"].Value.ToString());
                            item_detail.Add(df.get_item_id(item_grid.Rows[i].Cells["item_code_col"].Value.ToString()).ToString());
                            item_detail.Add(item_grid.Rows[i].Cells["tic_code_col"].Value.ToString());
                            item_detail.Add(df.get_invoice_id(invoices_tb.Text.ToString()).ToString());
                            item_detail.Add(item_grid.Rows[i].Cells["quantity_col"].Value.ToString());
                            item_detail.Add(item_grid.Rows[i].Cells["item_cost"].Value.ToString());
                            df.insert_inoviceItem(item_detail, 0);
                        }

                    }
                }


            }  
        }


        public void save_invoice(int update)
        {
            df.set_invoice_form(this);
            df.insert_inovice(contract_no_tb.Text.ToString(), invoices_tb.Text.ToString(), docket_no_tb.Text.ToString(), date_tb.Text.ToString(), df.get_ins_id(ins_cb.Text.ToString()), set_total_amount_textbox());

            int rowCount = item_grid.RowCount;
            List<String> item_detail;
            for (int i = 0; i < rowCount; i++)
            {
                if (item_grid.Rows[i].Cells["item_cost"].Value != null)
                {
                    item_detail = new List<string>();
                    //if(!item_grid.Rows[i].Cells["item_cost"].Value.ToString().Equals(""))
                    item_detail.Add(item_grid.Rows[i].Cells["delivery_date_col"].Value.ToString());
                    item_detail.Add(item_grid.Rows[i].Cells["job_no_col"].Value.ToString());
                    item_detail.Add(item_grid.Rows[i].Cells["csd_col"].Value.ToString());
                    item_detail.Add(df.get_item_id(item_grid.Rows[i].Cells["item_code_col"].Value.ToString()).ToString());
                    item_detail.Add(item_grid.Rows[i].Cells["tic_code_col"].Value.ToString());
                    item_detail.Add(df.get_invoice_id(invoices_tb.Text.ToString()).ToString());
                    item_detail.Add(item_grid.Rows[i].Cells["quantity_col"].Value.ToString());
                    item_detail.Add(item_grid.Rows[i].Cells["item_cost"].Value.ToString());
                    df.insert_inoviceItem(item_detail, update);
                }

            }
        }

        //Return the contract_no to database.cs
        public int get_contract_no()
        {
            return contract_no_sys;
        }

        private void op_in_btn_Click(object sender, EventArgs e)
        {
            load_invoice li = new load_invoice();
            li.set_invoice_form(this);
            li.Show();
        }

        private void export_in_btn_Click(object sender, EventArgs e)
        {

            if (!check_field())
            {
                try
                {
                    string fileTest = invoices_tb.Text + ".xlsx";
                    Boolean exist = true;
                    int count_exist = 1;
                    while (exist)
                    {
                        if (File.Exists(fileTest))
                        {
                            fileTest = invoices_tb.Text + "(" + count_exist + ").xlsx";
                        }
                        else
                            exist = false;
                    }


                    //The row and col number
                    int rowCount = item_grid.RowCount;
                    int colCount = item_grid.ColumnCount;
                    int sheet_page = 1;
                    int total_cell_num = 0;
                    int start_row = 0;
                    int end_row = 18;

                    Boolean more_than_table = false;
                    Excel.Application oApp;
                    Excel.Worksheet oSheet;
                    Excel.Workbook oBook;

                    //Check if the row records are more then the table row.
                    if (rowCount > 19)
                        more_than_table = true;

                    oApp = new Excel.Application();
                    oBook = oApp.Workbooks.Add();
                   
                  
                        
                    oSheet = (Excel.Worksheet)oBook.Worksheets.Add(After: oBook.Sheets[oBook.Sheets.Count]);
                       
                    oSheet.Name = invoices_tb.Text;
                    oApp.Application.UserName = invoices_tb.Text;
                    oApp.Visible = true;
                    //Set up the excel page style, orientation and the center alignment
                    //oSheet.PageSetup.Zoom = false;
                    //oSheet.PageSetup.Zoom = false;
                    
                    oSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    oSheet.PageSetup.CenterHorizontally = true;
                    oSheet.PageSetup.FitToPagesWide = 1;
                    oSheet.PageSetup.LeftMargin = 0.01f;
                    oSheet.PageSetup.RightMargin = 0.01f;
                    oSheet.PageSetup.TopMargin = 0.25f;
                    oSheet.PageSetup.BottomMargin = 0.50f;
                    oSheet.PageSetup.FooterMargin = 0.25;
                    
                    //oSheet.PageSetup.FitToPagesTall = 1;
                    //Set the page number
                    oSheet.PageSetup.CenterFooter = "Page &P of &N";


                    //Set the header text
                    oSheet.Cells[1, 1] = "Tai Kow Lo Transportation Co., Ltd";
                    oSheet.Cells[2, 1] = "G/F., 43 Fuk Chak Street, Tai Kok Tsui, Kowloon";
                    oSheet.Cells[3, 1] = "Tel: 2392 7108    Fax: 2391 9595";

                    //Set the invoice info text.
                    oSheet.Cells[5, 1] = "To";
                    oSheet.Cells[6, 1] = "Contract No";
                    oSheet.Cells[7, 1] = "Date";
                    oSheet.Cells[8, 1] = "Invoices";
                    oSheet.Cells[9, 1] = "institution";

                    //Set the value of the symbol ":"
                    oSheet.Cells[5, 2] = ": ";
                    oSheet.Cells[6, 2] = ": ";
                    oSheet.Cells[7, 2] = ": ";
                    oSheet.Cells[8, 2] = ": ";
                    oSheet.Cells[9, 2] = ": ";

                    //Set the invoice information
                    oSheet.Cells[5, 3] = to_cb.Text;
                    oSheet.Cells[6, 3] = contract_no_tb.Text;
                    oSheet.Cells[7, 3] = Convert.ToDateTime(date_tb.Text.ToString()).ToString("yyyy/MM/dd");
                    oSheet.Cells[7, 3].NumberFormat = "yyyy/MM/dd";
                    oSheet.Cells[8, 3] = invoices_tb.Text;
                    oSheet.Cells[9, 3] = ins_cb.Text;

                    //Set the docket label
                    oSheet.Cells[9, 6] = "Docket No:";
                    oSheet.Cells[9, 7] = docket_no_tb.Text;

                   

                    
                    //Set the header font style
                    Excel.Range oRange2 = oSheet.Range["A1", "A3"];
                    oRange2.Font.Name = "Arial";
                    oRange2.Font.Size = 14;
                    oRange2.Font.Bold = true;
                    oRange2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //Set the invoice info range style
                    Excel.Range oRange3 = oSheet.Range["A5", "A9"];
                    oRange3.Font.Name = "Arial";
                    oRange3.Font.Size = 10;
                    oRange3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange3.Columns.ColumnWidth = 11.00;

                    //Set the range style ":" 
                    Excel.Range oRange4 = oSheet.Range["B5", "B9"];
                    oRange4.Font.Name = "Arial";
                    oRange4.Font.Size = 10;
                    oRange4.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange4.Columns.ColumnWidth = 1.50;

                    //Set docket no Style
                    oSheet.Range["F9"].Font.Size = 10;
                    oSheet.Range["F9"].Font.Name = "Arial";
                    oSheet.Range["F9"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oSheet.Range["F9"].Columns.ColumnWidth = 11.00;

                    oSheet.Range["G9"].Font.Size = 10;
                    oSheet.Range["G9"].Font.Name = "Arial";
                    oSheet.Range["G9"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
      

                    //Merge the header cells.
                    oSheet.Range["A1", "J1"].Merge();
                    oSheet.Range["A2", "J2"].Merge();
                    oSheet.Range["A3", "J3"].Merge();

                    //Set the table header
                    for (int i = 0; i <= colCount - 1; i++)
                        oSheet.Cells[11, i + 1] = item_grid.Columns[i].HeaderText.ToString();

                    //Load the grid view data into the excel
                    for (int c = 0; c <= colCount - 1; c++)
                    {
                     
                        for (int r = 0; r <= rowCount - 1; r++)
                        {
                            
                            //.NumberFormat = "yyyy/MM/dd";
                            if (c == 0)
                            {
                                oSheet.Cells[r + 12, c + 1] = item_grid.Rows[r].Cells[c].Value;
                                oSheet.Cells[r + 12, c + 1].NumberFormat = "yyyy/MM/dd";
                            }
                            else
                            {
                                oSheet.Cells[r + 12, c + 1] = item_grid.Rows[r].Cells[c].Value;
                            }

                            if (r == rowCount - 1)
                            {
                                //Set total
                                
                                total_cell_num = r + 12 + 2;
                               
                            }

                        
                        } //end row loop
                            
                    }

                    //Set the table header style
                    Excel.Range oRange5 = oSheet.Range["A11", "J11"];
                    oRange5.Font.Name = "Arial";
                    oRange5.Font.Size = 10;
                    oRange5.Font.Bold = true;
                    oRange5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    oRange5.Columns.ColumnWidth = 12.30;

                    //Determine the data is more than the table row
                    int table_row_number = 0;

                    int table_cell = 29;

                    if (table_cell >= (11 + rowCount))
                        table_cell = 29;
                    else
                        table_cell = 11 + rowCount;
                    //Set the table style
                    Excel.Range oRange6 = oSheet.Range["A" + (12 - 1), oSheet.Cells[table_cell - 1, 5]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6.Font.Name = "Arial";
                    oRange6.Font.Size = 10;
                    oRange6.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange6.Columns.AutoFit();


                    Excel.Range oRange6_2 = oSheet.Range["F" + (12 - 1), oSheet.Cells[table_cell - 1, 6]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6_2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_2.Font.Name = "Arial";
                    oRange6_2.Font.Size = 10;
                    oRange6_2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_2.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_2.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange6_2.Columns.ColumnWidth = 30.00;

                    Excel.Range oRange6_3 = oSheet.Range["G" + (12 - 1), oSheet.Cells[table_cell - 1, 7]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6_3.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_3.Font.Name = "Arial";
                    oRange6_3.Font.Size = 10;
                    oRange6_3.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_3.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_3.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_3.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange6_3.Columns.AutoFit();

                    Excel.Range oRange6_4 = oSheet.Range["H" + (12 - 1), oSheet.Cells[table_cell - 1, 8]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6_4.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_4.Font.Name = "Arial";
                    oRange6_4.Font.Size = 10;
                    oRange6_4.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_4.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_4.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_4.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_4.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange6_4.Columns.ColumnWidth = 12.00;

                    Excel.Range oRange6_5 = oSheet.Range["I" + (12 - 1), oSheet.Cells[table_cell - 1, 9]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6_5.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_5.Font.Name = "Arial";
                    oRange6_5.Font.Size = 10;
                    oRange6_5.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_5.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_5.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_5.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_5.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    oRange6_5.Columns.AutoFit();


                    Excel.Range oRange6_1 = oSheet.Range[oSheet.Cells[11, 10], oSheet.Cells[table_cell - 1, 10]];
                    Debug.WriteLine("Row Count: " + rowCount);
                    oRange6_1.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRange6_1.Font.Name = "Arial";
                    oRange6_1.Font.Size = 10;
                    oRange6_1.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    //oRange6_1.Columns.AutoFit();
                    oRange6_1.Columns.ColumnWidth = 10;


                    //Set the invoice info style
                    Excel.Range oRange7 = oSheet.Range["C5", "C9"];
                    oRange7.Font.Name = "Arial";
                    oRange7.Font.Size = 10;
                    oRange7.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    int number_record = 0;
                    oSheet.Cells[table_cell +1, 9] = "Total";
                    oSheet.Cells[table_cell + 1, 9].Font.Name = "Arial";
                    oSheet.Cells[table_cell + 1, 9].Font.Size = 10;

                    oSheet.Cells[table_cell + 1, 10] = total_amount_tb.Text;
                    oSheet.Cells[table_cell + 1, 10].Font.Name = "Arial";
                    oSheet.Cells[table_cell + 1, 10].Font.Size = 10;
                    Excel.Range oRange8 = oSheet.Range[oSheet.Cells[table_cell + 2, 9]];
                    oRange8.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
                    oRange8.Columns.ColumnWidth = 10.00;
                    oRange8.Font.Name = "Arial";
                    oRange8.Font.Size = 10;
                    double how_many_page = Convert.ToDouble(rowCount) / 19.0f ;

                    oSheet.HPageBreaks.Add(oSheet.Cells[32, 10]);

               
                   // oBook.PrintOutEx(1, sheet_page, 1, true, Type.Missing, true, false, Type.Missing);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                    //MessageBox.Show("Excel Closed.");
                }
               
                //Save excel
                //oBook.SaveAs(fileTest);


                /*int rowCount = item_grid.RowCount;
                int colCount = item_grid.ColumnCount;
                Object[,] DataArray = new object[rowCount + 1, colCount + 1];

                for (int c = 0; c <= colCount - 1; c++)
                {
                    for (int r = 0; r <= rowCount - 1; r++)
                    {
                        DataArray[r, c] = item_grid.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop


                //Setup the word
                Word.Application objWord = new Word.Application();
                objWord.Visible = true;

                Word.Document objDoc;
               
                object objMissing = System.Reflection.Missing.Value;
                object name = invoices_tb.Text;
                objDoc = objWord.Documents.Add(ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                
                objDoc.Application.Selection.ParagraphFormat.LineSpacing = 8.5f;
                objDoc.Application.Selection.TypeText("To\t\t: " + to_cb.Text + "\n" +
                                                      "Contract No\t: " + contract_no_tb.Text + "\n" +
                                                      "Date\t\t:" + date_tb.Text + "\n" +
                                                      "Invoices\t:" + invoices_tb.Text + "\n" +
                                                      "Institution\t:" + ins_cb.Text + "\t\t\t\t" +
                                                       "Docket No\t:" + docket_no_tb.Text
                                                       );

                objDoc.Application.Selection.ParagraphFormat.LineSpacing = 12f;
                //Orient the word page direction.
                objDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                int start = 0;
                int end = 5;
                dynamic oRange = objDoc.Content.Application.Selection.Range;

                String oTemp = "";

                //objDoc.Application.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                objDoc.Application.Selection.TypeText("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t  Total: " + total_amount_tb.Text);
             
                //Transfer the objects to the string object.
                for (int x = 0; x <= 11 - 1; x++)
                {
                    for (int y = 0; y <= colCount - 1; y++)
                    {
                        try
                        {
                            oTemp = oTemp + DataArray[x, y] + "\t";
                        }
                        catch (Exception ex)
                        {
                            oTemp = oTemp + " " + "\t";
                        }
                    }
                }

                //Set table format
                oRange.Text = oTemp;
                object separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object applyBorders = true;
                object ApplyShading = true;
                object autoFit = true;
                object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;


                //Convert the word template range to a word table
                oRange.ConvertToTable(ref separator, ref rowCount, ref colCount,
                                      Type.Missing, Type.Missing, applyBorders,
                                      ApplyShading, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref autoFit, ref autoFitBehavior, Type.Missing
                                      );

                oRange.Select();

                objDoc.Application.Selection.Tables[1].Select();
                objDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                objDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                objDoc.Application.Selection.Tables[1].Rows.Select();
                objDoc.Application.Selection.InsertRowsAbove(1);
                objDoc.Application.Selection.Tables[1].Rows[1].Select();

                //Set table header style
                objDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                objDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                objDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 8;
                objDoc.Application.Selection.Tables[1].Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;


                //Add header text
                for (int q = 0; q <= colCount - 1; q++)
                {
                    objDoc.Application.Selection.Tables[1].Cell(1, q + 1).Range.Text = item_grid.Columns[q].HeaderText;

                }

                objDoc.Application.Selection.Tables[1].Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                objDoc.Application.Selection.Tables[1].Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                objDoc.Application.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;


                //Set table content style
                //objDoc.Application.Selection.Tables[1].set_Style(Word.WdBuiltinStyle.wdStyleTableLightShading);
                //objDoc.Application.Selection.Tables[1].Rows[1].Select();
                objDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Word.Section section in objWord.Application.ActiveDocument.Sections)
                {
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "Tai Kow Lo Transportation Co., Ltd" + "\n" + "G/F., 43 Fuk Chak Street, Tai Kok Tsui, Kowloon."
                                        + "\n" + "Tel: 2392 7108" + "\t" + "Fax: 2391 9595";
                    headerRange.Font.Size = 16;
                    headerRange.Font.Name = "Arial";
                    headerRange.Font.Color = Word.WdColor.wdColorBlack;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //Add text to the word

                objDoc = null;*/
                save_invoice(0);
                MessageBox.Show("Invoice saved successfully.");
                new_invoice();
               
                    
            }

        }

        private void ins_lb_Click(object sender, EventArgs e)
        {

        }

        //Check all the require_field is input.
        public Boolean check_field()
        {
            Boolean error = false;
            String errorMsg = "Missing field: Please input ";
            List<String> error_list = new List<string>();
            

            if (to_cb.Text == null || to_cb.Text.ToString() == "")
            {
                error = true;
                error_list.Add("To");
            }
            if (contract_no_tb.Text == null || contract_no_tb.Text.ToString() == "")
            {
                error = true;
                error_list.Add("Contract No");
            }
            if (ins_cb.Text == null || ins_cb.Text.ToString() == "")
            {
                error = true;
                error_list.Add("Institution");
            }
            if (docket_no_tb.Text == null || docket_no_tb.Text.ToString() == "")
            {
                error = true;
                error_list.Add("Docket No");
            }

            for (int i = 0; i < item_grid.RowCount; i++)
            {
                if (item_grid.Rows[i].Cells["item_cost"].Value != null)
                {
                    if (item_grid.Rows[i].Cells["job_no_col"].Value.ToString() == null || item_grid.Rows[i].Cells["job_no_col"].Value.ToString() == "")
                    {
                        error = true;
                        error_list.Add(item_grid.Rows[i].Cells["item_code_col"].Value.ToString() + " Job No");

                    }
                    if (item_grid.Rows[i].Cells["tic_code_col"].Value.ToString() == null || item_grid.Rows[i].Cells["tic_code_col"].Value.ToString() == "")
                    {
                        error = true;
                        error_list.Add(item_grid.Rows[i].Cells["item_code_col"].Value.ToString() + " Tic code");
                    }
                }
            }
            int count = 0;
            foreach (String str in error_list)
            {
                if (count != error_list.Count - 1)
                    errorMsg += str + ", ";
                else
                    errorMsg += str;    
                count++;
            }
            if(error)
             MessageBox.Show(errorMsg);
            return error;
        }


        public void new_invoice()
        {
          
            
            to_cb.SelectedIndex = 0;
            contract_no_tb.Text = "";
            ins_cb.SelectedIndex = -1;
            docket_no_tb.Text = "";
            total_amount_tb.Text = "$0";
            int rowCount = item_grid.RowCount;
            for (int i = 0; i < rowCount; i++)
            {
                //MessageBox.Show(item_grid.Rows[0].Cells["item_code_col"].Value.ToString()+  item_grid.Rows[1].Cells["item_code_col"].Value.ToString());
                if (item_grid.Rows[0].Cells["item_code_col"].Value != null)
                    this.item_grid.Rows.RemoveAt(0);
            }

            if (File.Exists("data.xml"))
            {
                XmlSerializer xs = new XmlSerializer(typeof(information));
                FileStream read = new FileStream("data.xml", FileMode.Open, FileAccess.Read, FileShare.Read);
                information info = (information)xs.Deserialize(read);
                contract_no_sys = Convert.ToInt32(info.Contract_no);
                read.Close();
            }
            DateTime dateTime = DateTime.Now.Date;
            invoices_tb.Text = "C" + dateTime.ToString("yyyyMM") + contract_no_sys.ToString("00000");
            date_tb.Text = dateTime.ToString("yyyy/MM/dd");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you sure to reset the invoice number?", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                df.reset_invoice_no();
                new_invoice();
            }
                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            df.set_invoice_form(this);
            if (df.get_invoice_existed(invoices_tb.ToString()))
            {
                df.add_invoice_no();
                Debug.WriteLine("Fail");
            }
                
            new_invoice(); 
        }


        public void load_invoice(int invoice_id)
        {
            //Reset the invoice form first and 
            df.set_invoice_form(this);
           /* if(df.get_invoice_existed(invoices_tb.ToString()))
                df.add_invoice_no();*/
            new_invoice();

            List<String> invoice_info = df.get_invoice_info(invoice_id);
            List<List<String>> invoice_items = df.get_load_invoice_items(invoice_id);
            //List<String> item_detail = df.get_item_detail_by_id();
            //MessageBox.Show(invoice_id.ToString());

            if (invoice_id != 0)
            {
                to_cb.Text = "Correctionl Services Department";
                contract_no_tb.Text = invoice_info[0];
                date_tb.Text = Convert.ToDateTime(invoice_info[2]).ToString("yyyy/MM/dd");
                docket_no_tb.Text = invoice_info[1];
                invoices_tb.Text = invoice_info[4];
                total_amount_tb.Text = "$" + invoice_info[3];    
                ins_cb.Text = invoice_info[5];
                //MessageBox.Show(invoice_items[0][0]);
                DateTime dateTime = DateTime.Now.Date;
               
                foreach (List<String> items in invoice_items)
                {
                    List<String> item_detail = df.get_item_detail_by_id(Convert.ToInt32(items[3]));

                    this.item_grid.Rows.Add(new Object[] {Convert.ToDateTime(items[0]).ToString("yyyy/MM/dd"), items[1], items[2], item_detail[0], items[4], item_detail[1], items[5], item_detail[2], item_detail[3], items[6] });

                }
            }
        }
    }
}
