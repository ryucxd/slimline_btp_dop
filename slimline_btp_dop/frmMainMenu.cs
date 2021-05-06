using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace slimline_btp_dop
{
    public partial class frmMainMenu : Form
    {
        public frmMainMenu()
        {
            InitializeComponent();
            string temp = "";
            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("usp_btp_dop_data", conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    //start messing with data here, add up totals for packed/invoice
                    int counter = dt.Rows.Count + 4;
                    for (int i = dt.Rows.Count; i < counter ; i++) //insert 3 rows ~ one for line break then total packed/invoiced
                    {
                        DataRow newRow = dt.NewRow();
                        dt.Rows.InsertAt(newRow, i);
                    }

                    dt.Rows[dt.Rows.Count - 3][0] = "TOTAL PACKED:";
                    dt.Rows[dt.Rows.Count - 2][0] = "TOTAL INVOICED:";
                    dt.Rows[dt.Rows.Count - 1][0] = "SUM:";

                    //DATA STARTS AT COLUMN 4

                    for (int column = 4; column < dt.Columns.Count;column++) //each column
                    {
                        double packed = 0;
                        double invoiced = 0;
                        for (int row = 0; row < dt.Rows.Count - 4;row++)
                        {
                            if (dt.Rows[row][0].ToString() == "INVOICED") //check for invoice vs packed
                            {
                                if (String.IsNullOrEmpty(dt.Rows[row][column].ToString()) == false)
                                    invoiced = invoiced + Convert.ToDouble(dt.Rows[row][column].ToString());
                            }
                            if (dt.Rows[row][0].ToString() == "PACKED") 
                            {
                                if (String.IsNullOrEmpty(dt.Rows[row][column].ToString()) == false)
                                    packed = packed + Convert.ToDouble(dt.Rows[row][column].ToString());
                            }

                        }
                        //now add them to the last two rows
                        dt.Rows[dt.Rows.Count - 3][column] = packed.ToString();
                        dt.Rows[dt.Rows.Count - 2][column] =  invoiced.ToString();
                        dt.Rows[dt.Rows.Count - 1][column] =  (packed - invoiced).ToString();

                    }


                    dataGridView1.DataSource = dt;
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable; // dont allow the user to sort the columns, will cause big problems 
                        column.DefaultCellStyle.BackColor = Color.DeepSkyBlue;
                    }

                }
                conn.Close();
            
            }
        }

        private void btnExcel_Click(object sender, EventArgs e) //this highlights all of the dgv, copies it, then pastes into a new excel instance
        {
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
            dataGridView1.ClearSelection();

            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            xlWorkSheet.Name = "test";

            //              VVV this changes backcolour 
            int row = dataGridView1.Rows.Count  + 1; //how many rows down the  sum field is
            for (int column = 5; column < dataGridView1.Columns.Count  + 1; column++)
            {
                //read the cell value

                //apply formatting
                double test = Convert.ToDouble((xlWorkSheet.Cells[row, column] as Microsoft.Office.Interop.Excel.Range).Value);
                if (test < 0)
                    xlWorkSheet.Cells[row, column].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleVioletRed);
                else
                    xlWorkSheet.Cells[row, column].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkSeaGreen);
            }
        }

    }
}
