using System.Data;
using Microsoft.Data.SqlClient;

using static System.Net.Mime.MediaTypeNames;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;



namespace Tax_Calculator
{
    public partial class texForm : Form
    {



        public void loadData()
        {

            SqlConnection con = new SqlConnection("Data Source=LAPTOP-H6NLRH3P\\SQLEXPRESS;Initial Catalog=practice;Integrated Security=True;Encrypt=True;Trust Server Certificate=True");
            con.Open();


            SqlDataAdapter ad = new SqlDataAdapter("SELECT * FROM TableTaxCalculator", con);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            dgvTaxCalculator.DataSource = dt;


        }
        public texForm()
        {
            InitializeComponent();
           
        }

        private void texForm_Load(object sender, EventArgs e)
        {

            loadData();

        }



        private void addItem_Click(object sender, EventArgs e)
        {
            using (SqlConnection con = new SqlConnection("Data Source=LAPTOP-H6NLRH3P\\SQLEXPRESS;Initial Catalog=practice;Integrated Security=True;Encrypt=True;Trust Server Certificate=True"))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("INSERT INTO TableTaxCalculator ([itemPrice], [taxRate]) VALUES (@itemPrice, @taxRate)", con);
                    cmd.Parameters.AddWithValue("@itemPrice", itemPrice.Text);
                    cmd.Parameters.AddWithValue("@taxRate", taxRate.Text);

                    cmd.ExecuteNonQuery(); 
                    MessageBox.Show("Add success", "Message Title", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    loadData(); 
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


        }

        private void totalCost_Click(object sender, EventArgs e)
        {
            float subtotal = 0;
            float calculatedTaxAmount = 0;
            float calculatedTotal = 0;

            using (SqlConnection con = new SqlConnection("Data Source=LAPTOP-H6NLRH3P\\SQLEXPRESS;Initial Catalog=practice;Integrated Security=True;Encrypt=True;Trust Server Certificate=True"))
            {
                try
                {
                    
                    con.Open();
                    SqlCommand cmd = new SqlCommand("SELECT itemPrice, taxRate FROM TableTaxCalculator", con);
                    SqlDataReader reader = cmd.ExecuteReader();

                    
                    while (reader.Read())
                    {
                       
                        double itemPrice = reader.GetDouble(0);  
                        double taxRate = reader.GetDouble(1);    

                       
                        subtotal += (float)itemPrice;
                        calculatedTaxAmount += (float)(itemPrice * (taxRate / 100));
                        calculatedTotal += (float)(itemPrice + (itemPrice * (taxRate / 100))); 
                    }

                    
                    subTotal.Text = String.Format("{0:C}", subtotal);  
                    taxAmount.Text = String.Format("{0:C}", calculatedTaxAmount); 
                    total.Text = String.Format("{0:C}", calculatedTotal);    

                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error fetching data: " + ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }






        }

        private void dgvTaxCalculator_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvTaxCalculator.Columns[e.ColumnIndex].Name == "itemPrice" && e.Value != null)
            {
                
                float itemPriceValue = Convert.ToSingle(e.Value);
                e.Value = String.Format("${0:F2}", itemPriceValue);
                e.FormattingApplied = true;
            }
            else if (dgvTaxCalculator.Columns[e.ColumnIndex].Name == "taxRate" && e.Value != null)
            {
                
                float taxRateValue = Convert.ToSingle(e.Value);
                e.Value = String.Format("{0:F2}%", taxRateValue);
                e.FormattingApplied = true;
            }
        }

        private void dgvTaxCalculator_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvTaxCalculator.Rows[e.RowIndex];
                itemPrice.Text = row.Cells["itemPrice"].Value?.ToString() ?? string.Empty;
                taxRate.Text = row.Cells["taxRate"].Value?.ToString() ?? string.Empty;
            }
        }

        private void deleteItem_Click(object sender, EventArgs e)
        {
            if (dgvTaxCalculator.SelectedRows.Count > 0)
            {
                try
                {
                  
                    int selectedRowIndex = dgvTaxCalculator.SelectedRows[0].Index;
                    float itemPriceValue = Convert.ToSingle(dgvTaxCalculator.Rows[selectedRowIndex].Cells["itemPrice"].Value);
                    float taxRateValue = Convert.ToSingle(dgvTaxCalculator.Rows[selectedRowIndex].Cells["taxRate"].Value);

                    using (SqlConnection con = new SqlConnection("Data Source=LAPTOP-H6NLRH3P\\SQLEXPRESS;Initial Catalog=practice;Integrated Security=True;Encrypt=True;Trust Server Certificate=True"))
                    {
                        con.Open();

                       
                        SqlCommand cmd = new SqlCommand(
                            "DELETE FROM TableTaxCalculator WHERE itemPrice = @itemPrice AND taxRate = @taxRate",
                            con
                        );
                        cmd.Parameters.AddWithValue("@itemPrice", itemPriceValue);
                        cmd.Parameters.AddWithValue("@taxRate", taxRateValue);

                        int rowsAffected = cmd.ExecuteNonQuery(); 

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Item deleted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            loadData();
                        }
                        else
                        {
                            MessageBox.Show("No matching row found to delete.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting item: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a row to delete.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnExportToPDF_Click(object sender, EventArgs e)
        {



            if (dgvTaxCalculator.Rows.Count > 0)
            {
                SaveFileDialog save = new SaveFileDialog();
                save.Filter = "PDF (*.pdf)|*.pdf";
                save.FileName = "Result.pdf";
                bool ErrorMessage = false;

                if (save.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(save.FileName))
                    {
                        try
                        {
                            File.Delete(save.FileName);
                        }
                        catch (Exception ex)
                        {
                            ErrorMessage = true;
                            MessageBox.Show("Unable to write data to disk: " + ex.Message);
                        }
                    }

                    if (!ErrorMessage)
                    {
                        try
                        {
                            
                            PdfPTable pTable = new PdfPTable(dgvTaxCalculator.Columns.Count);
                            pTable.DefaultCell.Padding = 2;
                            pTable.WidthPercentage = 100;
                            pTable.HorizontalAlignment = Element.ALIGN_LEFT;

                           
                            foreach (DataGridViewColumn col in dgvTaxCalculator.Columns)
                            {
                                PdfPCell pCell = new PdfPCell(new Phrase(col.HeaderText));
                                pTable.AddCell(pCell);
                            }

                            foreach (DataGridViewRow viewRow in dgvTaxCalculator.Rows)
                            {
                                foreach (DataGridViewCell dcell in viewRow.Cells)
                                {
                                    pTable.AddCell(dcell.Value?.ToString() ?? string.Empty);
                                }
                            }

                            using (FileStream fileStream = new FileStream(save.FileName, FileMode.Create))
                            {
                              
                                Document document = new Document(PageSize.A4, 8f, 16f, 16f, 8f);
                                PdfWriter.GetInstance(document, fileStream);
                                document.Open();

                                
                                string imagePath = @"C:\Users\HP\Pictures\logo.png"; 
                                if (File.Exists(imagePath))
                                {
                                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imagePath);
                                    img.ScaleToFit(150f, 150f); 
                                    img.Alignment = Element.ALIGN_RIGHT; 
                                    document.Add(img);
                                }
                                else
                                {
                                    MessageBox.Show("Image not found at " + imagePath, "Error");
                                }

                               
                                document.Add(pTable);

                           
                                document.Close();
                                fileStream.Close();
                            }

                            MessageBox.Show("Data Exported Successfully", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error while exporting data: " + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No Records Found", "Info");
            }



        }

        private void btnExportToExcel_Click(object sender, EventArgs e)
        {

            dgvTaxCalculator.SelectAll();
            DataObject copydata = dgvTaxCalculator.GetClipboardContent();
            if (copydata != null) Clipboard.SetDataObject(copydata);

            Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            xlapp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook xlWbook;
            Microsoft.Office.Interop.Excel.Worksheet xlsheet;
            object miseddata = System.Reflection.Missing.Value;
            xlWbook = xlapp.Workbooks.Add(miseddata);

            xlsheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWbook.Worksheets.get_Item(1);


            xlsheet.Cells[1, 2] = "Item Price";
            xlsheet.Cells[1, 3] = "Tax Rate";

           
            Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)xlsheet.Cells[2, 1];
            xlr.Select();

            
            xlsheet.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }
    }






}
