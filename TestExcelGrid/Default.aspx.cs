using Microsoft.Office.Interop;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//test comment
public partial class _Default : System.Web.UI.Page
{
    bool hdr;
    string filepath;
    DataTable dt;
    DataTable tbl;

    public int MyProperty = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            dt = Session["tblTable"] as DataTable;
        }
        else
        {
            if (ViewState["KeepIncrease"] != null)
            {
                MyProperty = Int32.Parse(ViewState["KeepIncrease"].ToString());
            }
            if (ViewState["filepath"] != null)
            {
                filepath = ViewState["filepath"].ToString();
            }
        }
    }
    protected void PreButton(object sender, EventArgs e)
    {
        if (FileUpload1.HasFile)
        {
            String fileExtension = Path.GetExtension(FileUpload1.FileName).ToLower();
            String[] allowedExtensions = { ".xlsx" };
            for (int i = 0; i < allowedExtensions.Length; i++)
            {
                if (fileExtension == allowedExtensions[i])
                {
                    try
                    {
                        if (rbHDR.SelectedItem.Text == "Yes")
                            hdr = true;
                        else
                            hdr = false;

                        string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                        string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                        string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                        string FilePath = Server.MapPath(FolderPath + FileName);

                        GetDataTableFromExcel(FilePath, hdr);
                        Label2.Text = "Data from excel file show below!";
                    }
                    catch
                    {
                        Response.Write("File could not be uploaded.");
                    }

                }
                else
                {
                    Label2.Text = "Please confirm file type is .xls or .xlsx and try again";
                }
            }
        }
        else
        {
            String fileExtension = Path.GetExtension(filepath).ToLower();
            String[] allowedExtensions = { ".xlsx" };
            for (int i = 0; i < allowedExtensions.Length; i++)
            {
                if (fileExtension == allowedExtensions[i])
                {
                    try
                    {
                        if (rbHDR.SelectedItem.Text == "Yes")
                            hdr = true;
                        else
                            hdr = false;

                        GetDataTableFromExcel(filepath, hdr);
                        Label2.Text = "Data from excel file show below!";
                    }
                    catch
                    {
                        Response.Write("File could not be uploaded.");
                    }

                }
                else
                {
                    Label2.Text = "Please confirm file type is .xls or .xlsx and try again";
                }
            }
        }
    }
    protected void UploadID_Click(object sender, EventArgs e)
    {
        if (DropDownList1.SelectedValue != "")
        {
            string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
            string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
            string FilePath = Server.MapPath(FolderPath + FileName);

            if (FileUpload1.HasFile)
            {
                InsertExcelRecords(FilePath);
            }
            else
            {
                if (GridView1.Rows.Count > 0)
                {
                    InsertExcelRecords(filepath);
                }
                else
                {
                    InsertExcelRecords(filepath);
                }
            }
        }
        else
        {
            Label2.Text = "Please select a sheet!";
        }
    }
    protected void ViewDBID_Click(object sender, EventArgs e)
    {
        string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
        using (SqlConnection con = new SqlConnection(constr))
        {
            using (SqlCommand cmd = new SqlCommand("SELECT Name, City, Address, Designation FROM dbo.Employee"))
            {
                using (SqlDataAdapter sda = new SqlDataAdapter())
                {
                    cmd.Connection = con;
                    sda.SelectCommand = cmd;
                    using (dt = new DataTable())
                    {
                        sda.Fill(dt);
                        Session.Add("tblTable", dt);
                        GridView1.DataSource = dt;
                        GridView1.DataBind();
                        if (dt.Rows.Count > 1)
                        {
                            Label2.Text = "All data from this table show below";
                        }
                        else
                            Label2.Text = "This table have no data!";
                    }
                }
            }
        }
    }
    protected void TruncateDBID_Click(object sender, EventArgs e)
    {
        string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
        using (SqlConnection con = new SqlConnection(constr))
        {
            using (SqlCommand cmd = new SqlCommand("TRUNCATE TABLE dbo.Employee"))
            {
                con.Open();
                cmd.Connection = con;
                cmd.ExecuteNonQuery();
                con.Close();
                Label2.Text = "Delete all data in Employee table already!";
                GridView1.DataSource = null;
                GridView1.DataBind();
            }
        }
    }
    protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        GridView1.PageIndex = e.NewPageIndex;
        GridView1.DataSource = Session["tblTable"];
        GridView1.DataBind();
    }
    private void InsertExcelRecords(string FilePath)
    {
        string constr, Query, sqlconn;

        constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", FilePath);
        OleDbConnection connExcel = new OleDbConnection(constr);

        sqlconn = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
        SqlConnection con = new SqlConnection(sqlconn);

        Query = string.Format("SELECT [Name],[City],[Address],[Designation] FROM [{0}]", DropDownList1.SelectedItem.ToString() + "$");
        OleDbCommand Ecom = new OleDbCommand(Query, connExcel);
        connExcel.Open();

        DataSet ds = new DataSet();
        OleDbDataAdapter oda = new OleDbDataAdapter(Query, connExcel);
        connExcel.Close();
        oda.Fill(ds);
        DataTable Exceldt = ds.Tables[0];

        //creating object of SqlBulkCopy    
        SqlBulkCopy objbulk = new SqlBulkCopy(con);
        //assigning Destination table name    
        objbulk.DestinationTableName = "Employee";
        //Mapping Table column    
        objbulk.ColumnMappings.Add("Name", "Name");
        objbulk.ColumnMappings.Add("City", "City");
        objbulk.ColumnMappings.Add("Address", "Address");
        objbulk.ColumnMappings.Add("Designation", "Designation");
        //Set up the event handler to notify after 50 rows
        objbulk.SqlRowsCopied += new SqlRowsCopiedEventHandler(OnSqlRowsCopied);
        objbulk.NotifyAfter = 50;

        //inserting Datatable Records to DataBase    
        con.Open();
        objbulk.WriteToServer(Exceldt);
        con.Close();

        Label2.Text = "Insert data to Employee table already!";

        GridView1.DataSource = null;
        GridView1.DataBind();
    }
    public void GetDataTableFromExcel(string path, bool hasHeader)
    {
        using (var pck = new OfficeOpenXml.ExcelPackage())
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                pck.Load(stream);
            }
            var ws = pck.Workbook.Worksheets[DropDownList1.SelectedItem.ToString()];
            tbl = new DataTable();
            Session.Add("tblTable", tbl);

            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                DataRow row = tbl.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            GridView1.Caption = Path.GetFileName(path);
            GridView1.DataSource = tbl;
            GridView1.DataBind();
        }
    }
    protected void Button1_Click1(object sender, EventArgs e)
    {
        if (MyProperty < 100)
        {
            MyProperty += 10;
        }
        ViewState.Add("KeepIncrease", MyProperty);
        Page.ClientScript.RegisterStartupScript(this.GetType(), "move", "move2();", true);
    }
    protected void Button2_Click(object sender, EventArgs e)
    {
        string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
        string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
        string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
        string FilePath = Server.MapPath(FolderPath + FileName);
        ViewState.Add("filepath", FilePath);

        OleDbConnection OleDbcon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 12.0;");
        OleDbcon.Open();
        DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        OleDbcon.Close();
        DropDownList1.Items.Clear();
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
            sheetName = sheetName.Substring(0, sheetName.Length - 1);
            DropDownList1.Items.Add(sheetName);
        }
    }
    public void OnSqlRowsCopied(object sender, SqlRowsCopiedEventArgs e)
    {
        Label3.Text = "Copied " + e.RowsCopied + " so far...";
    }
}