using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using Aspose.Cells;
using System.Data.SqlClient;
using System.Configuration;

public partial class ST_upload : System.Web.UI.Page
{
    //保存上传文件路径
    public static string savepath = @"UploadFile\ST";

    protected void Page_Load(object sender, EventArgs e)
    {
    }

    protected void uploadcontrol_FileUploadComplete(object sender, DevExpress.Web.FileUploadCompleteEventArgs e)
    {
        string isSubmissionExpired = "Y";

        string resultExtension = Path.GetExtension(e.UploadedFile.FileName);
        //string resultFileName = Path.ChangeExtension(Path.GetRandomFileName(), resultExtension);
        string resultFileName = Path.GetFileNameWithoutExtension(e.UploadedFile.FileName) + DateTime.Now.ToString("yyyyMMddHHmmss") + resultExtension;

        string resultFilePath = MapPath("~") + savepath + @"\" + resultFileName;
        if (!Directory.Exists(MapPath("~") + savepath + @"\"))
        {
            Directory.CreateDirectory(MapPath("~") + savepath + @"\");
        }
        e.UploadedFile.SaveAs(resultFilePath);

        string result = "", DesTableName = "ST_Detail_temp"; 
        SQLHelper.ExSql("truncate table " + DesTableName);
        importdata(resultFilePath, DesTableName, resultFileName, e.UploadedFile.FileName, out result);

        if (result != "")
        {
            File.Delete(resultFilePath);
            e.CallbackData = isSubmissionExpired + "|" + result;
            return;
        }

        try
        {
            string sql = @"exec [Report_ST_upload]";
            DataTable dt_flag = SQLHelper.Query(sql).Tables[0];

            if (dt_flag.Rows[0][0].ToString() == "Y")
            {
                DataTable dt_error = SQLHelper.Query(@"select * from ST_Detail_temp where ISNULL(errorMessage,'')<>'' order by jq").Tables[0];
                for (int i = 0; i < dt_error.Rows.Count; i++)
                {
                    result = result + "工作表[" + dt_error.Rows[i]["jq_name"].ToString() + "],日期" + (Convert.ToDateTime(dt_error.Rows[i]["date"])).ToString("MM月dd日")
                        + ",材料颜色" + dt_error.Rows[i]["material_color"].ToString() + ",产品规格" + dt_error.Rows[i]["product_code"].ToString()
                        + ",批号" + dt_error.Rows[i]["batch_no"].ToString() + ",容器" + dt_error.Rows[i]["cavity"].ToString()
                        + ",error:" + dt_error.Rows[i]["errorMessage"].ToString()
                        + "<br />";
                }
                e.CallbackData = isSubmissionExpired + "|" + result;
                return;
            }


            //临时表 导入正式表
            string sql_prd = @"exec [Report_ST_upload_prd]";
            DataTable dt_flag_prd = SQLHelper.Query(sql_prd).Tables[0];

        }
        catch (Exception ex)
        {
            result = ex.Message;
            e.CallbackData = isSubmissionExpired + "|" + result;
            return;
        }

        isSubmissionExpired = "N";
        e.CallbackData = isSubmissionExpired + "|" + result;
        return;


    }

    public void importdata(string fileName, string DesTableName, string new_filename, string ori_filename, out string result)
    {
        result = ""; 
        try
        {
            DataSet ds = GetExcelData_Table(fileName, out result);
            if (result == "")
            {
                DataTable dt = new DataTable();
                DataColumn col_0 = new DataColumn("jq", typeof(string));
                DataColumn col_1 = new DataColumn("date", typeof(string));
                DataColumn col_2 = new DataColumn("material_color", typeof(string));
                DataColumn col_3 = new DataColumn("product_code", typeof(string));
                DataColumn col_4 = new DataColumn("batch_no", typeof(string));
                DataColumn col_5 = new DataColumn("cavity", typeof(string));
                DataColumn col_6 = new DataColumn("start_weight", typeof(decimal));
                DataColumn col_7 = new DataColumn("D2_weight", typeof(decimal));
                DataColumn col_8 = new DataColumn("D2_remark", typeof(string));
                DataColumn col_9 = new DataColumn("D4_weight", typeof(decimal));
                DataColumn col_10 = new DataColumn("D4_remark", typeof(string));
                DataColumn col_11 = new DataColumn("D7_weight", typeof(decimal));
                DataColumn col_12 = new DataColumn("D7_remark", typeof(string)); 
                DataColumn col_13 = new DataColumn("deal_result", typeof(string));
                DataColumn col_14 = new DataColumn("jq_name", typeof(string));
                DataColumn col_15 = new DataColumn("ori_filename", typeof(string));
                DataColumn col_16 = new DataColumn("new_filename", typeof(string));
                DataColumn col_17 = new DataColumn("upload_time", typeof(DateTime));
                dt.Columns.Add(col_0); dt.Columns.Add(col_1); dt.Columns.Add(col_2); dt.Columns.Add(col_3); dt.Columns.Add(col_4); dt.Columns.Add(col_5);
                dt.Columns.Add(col_6); dt.Columns.Add(col_7); dt.Columns.Add(col_8); dt.Columns.Add(col_9); dt.Columns.Add(col_10); dt.Columns.Add(col_11);
                dt.Columns.Add(col_12); dt.Columns.Add(col_13); dt.Columns.Add(col_14); dt.Columns.Add(col_15); dt.Columns.Add(col_16); dt.Columns.Add(col_17);

                for (int t = 0; t < ds.Tables.Count; t++)
                {
                    DataTable dtExcel = ds.Tables[t];
                    for (int k = 0; k < dtExcel.Rows.Count; k++)
                    {
                        DataRow dr = dtExcel.Rows[k];

                        //过滤空行
                        if (dr["jq"].ToString().Trim() == "" && dr["Date"].ToString().Trim() == "" && dr["material_color"].ToString().Trim() == ""
                             && dr["product_code"].ToString().Trim() == "" && dr["batch_no"].ToString().Trim() == "" && dr["cavity"].ToString().Trim() == ""
                            && dr["start_weight"].ToString().Trim() == "")
                        {
                            continue;
                        }
                        if (dr["Date"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【日期】不可为空！" + "<br />";
                            break;
                        }
                        if (dr["material_color"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【材料颜色】不可为空！" + "<br />";
                            break;
                        }
                        if (dr["product_code"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【产品规格】不可为空！" + "<br />";
                            break;
                        }
                        if (dr["batch_no"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【批号】不可为空！" + "<br />";
                            break;
                        }
                        if (dr["cavity"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【容器】不可为空！" + "<br />";
                            break;
                        }
                        if (dr["start_weight"].ToString().Trim() == "")
                        {
                            result = "工作表【" + dr["jq_name"].ToString().Trim() + "】第" + (k + 3).ToString() + "行" + "【开始重量】不可为空！" + "<br />";
                            break;
                        }

                        DataRow dt_r = dt.NewRow();
                        dt_r["jq"] = dr["jq"].ToString().Trim();
                        dt_r["date"] = dr["date"].ToString().Trim();
                        dt_r["material_color"] = dr["material_color"].ToString().Trim();
                        dt_r["product_code"] = dr["product_code"].ToString().Trim();
                        dt_r["batch_no"] = dr["batch_no"].ToString().Trim();
                        dt_r["cavity"] = dr["cavity"].ToString().Trim();
                        
                        dt_r["start_weight"] = dr["start_weight"].ToString().Trim();
                        dt_r["D2_weight"] = dr["D2_weight"].ToString().Trim();
                        dt_r["D2_remark"] = dr["D2_remark"].ToString().Trim();
                        dt_r["D4_weight"] = dr["D4_weight"].ToString().Trim();
                        dt_r["D4_remark"] = dr["D4_remark"].ToString().Trim();
                        dt_r["D7_weight"] = dr["D7_weight"].ToString().Trim();
                        dt_r["D7_remark"] = dr["D7_remark"].ToString().Trim();
                        dt_r["deal_result"] = dr["deal_result"].ToString().Trim();
                        dt_r["jq_name"] = dr["jq_name"].ToString().Trim();

                        dt_r["ori_filename"] = ori_filename;
                        dt_r["new_filename"] = new_filename;
                        dt_r["upload_time"] = DateTime.Now;

                        dt.Rows.Add(dt_r);
                    }

                    if (result != "")
                    {
                        break;
                    }
                }


                if (result == "")
                {
                    bool bf = DataTableToSQLServer(dt, DesTableName);
                    result = bf == true ? "" : "excel导入临时表失败";
                }
            }
        }
        catch (Exception ex)
        {
            result = "读取excel异常：" + ex.Message;
        }

    }

    public DataSet GetExcelData_Table(string filePath, out string result)
    {
        result = "";
        DataSet ds = new DataSet();

        Workbook book = new Workbook(filePath);
        for (int b = 0; b < book.Worksheets.Count; b++)
        {
            Worksheet sheet = book.Worksheets[b];
            string sheetname = sheet.Name; string sheetno = "";
            if (sheetname.Contains("号机") == true || sheetname.Contains("机") == true || sheetname.Contains("线后氟化") == true)
            {
                Cells cells = sheet.Cells;
                book.CalculateFormula();
                DataTable dt_Import = cells.ExportDataTableAsString(2, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, false);//获取excel中的数据保存到一个datatable中

                if (dt_Import == null || dt_Import.Rows.Count <= 0 || dt_Import.Columns.Count != 19)
                {
                    result = result + "工作表" + sheetname + ":No Data";
                }
                else
                {
                    sheetno = sheetname.Replace("号机", "").Replace("机", "");

                    dt_Import.Columns[0].ColumnName = "Date"; dt_Import.Columns[1].ColumnName = "material_color";
                    dt_Import.Columns[2].ColumnName = "product_code"; dt_Import.Columns[3].ColumnName = "batch_no";
                    dt_Import.Columns[4].ColumnName = "cavity"; dt_Import.Columns[5].ColumnName = "start_weight";
                    dt_Import.Columns[6].ColumnName = "D2_weight"; dt_Import.Columns[7].ColumnName = "D2_loss_weight";
                    dt_Import.Columns[8].ColumnName = "D2_loss_rate"; dt_Import.Columns[9].ColumnName = "D2_remark";
                    dt_Import.Columns[10].ColumnName = "D4_weight"; dt_Import.Columns[11].ColumnName = "D4_loss_weight";
                    dt_Import.Columns[12].ColumnName = "D4_loss_rate"; dt_Import.Columns[13].ColumnName = "D4_remark";
                    dt_Import.Columns[14].ColumnName = "D7_weight"; dt_Import.Columns[15].ColumnName = "D7_loss_weight";
                    dt_Import.Columns[16].ColumnName = "D7_loss_rate"; dt_Import.Columns[17].ColumnName = "D7_remark";
                    dt_Import.Columns[18].ColumnName = "deal_result";

                    dt_Import.Columns.Add("jq", typeof(string)); dt_Import.Columns.Add("jq_name", typeof(string));

                    string Date = "", material_color = "", product_code = "", batch_no = "";
                    for (int i = 0; i < dt_Import.Rows.Count; i++)
                    {
                        dt_Import.Rows[i]["jq_name"] = sheetname;

                        if (dt_Import.Rows[i]["date"].ToString() != "" && dt_Import.Rows[i]["material_color"].ToString() != ""
                              && dt_Import.Rows[i]["product_code"].ToString() != "" && dt_Import.Rows[i]["batch_no"].ToString() != "")
                        {
                            Date = dt_Import.Rows[i]["date"].ToString();
                            material_color = dt_Import.Rows[i]["material_color"].ToString().Replace("\n", ";");
                            product_code = dt_Import.Rows[i]["product_code"].ToString().Replace("\n", ";");
                            batch_no = dt_Import.Rows[i]["batch_no"].ToString();

                            dt_Import.Rows[i]["material_color"] = material_color;
                            dt_Import.Rows[i]["product_code"] = product_code;
                        }

                        if (Date != "" && material_color != "" && product_code != "" && batch_no != "")
                        {
                            if (dt_Import.Rows[i]["start_weight"].ToString() != "")
                            {
                                dt_Import.Rows[i]["date"] = Date;
                                dt_Import.Rows[i]["material_color"] = material_color;
                                dt_Import.Rows[i]["product_code"] = product_code;

                                if (sheetno == "线后氟化")
                                {
                                    dt_Import.Rows[i]["cavity"] = (i + 3).ToString();
                                }
                                else
                                {
                                    dt_Import.Rows[i]["batch_no"] = batch_no;
                                }

                                dt_Import.Rows[i]["jq"] = sheetno;
                            }
                        }
                    }

                    ds.Tables.Add(dt_Import);
                }
            }
        }
        return ds;

    }

    public bool DataTableToSQLServer(DataTable dt, string DesTableName)
    {
        bool bf = true;
        string connectionString = ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString;

        using (SqlConnection destinationConnection = new SqlConnection(connectionString))
        {
            destinationConnection.Open();

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(destinationConnection))
            {
                try
                {

                    bulkCopy.DestinationTableName = DesTableName;//要插入的表的表名
                    bulkCopy.BatchSize = dt.Rows.Count;
                    bulkCopy.ColumnMappings.Add("jq", "jq");//映射字段名 DataTable列名 ,数据库 对应的列名   
                    bulkCopy.ColumnMappings.Add("date", "date");//映射字段名 DataTable列名 ,数据库 对应的列名  
                    bulkCopy.ColumnMappings.Add("material_color", "material_color");//映射字段名 DataTable列名 ,数据库 对应的列名  
                    bulkCopy.ColumnMappings.Add("product_code", "product_code");//映射字段名 DataTable列名 ,数据库 对应的列名  
                    bulkCopy.ColumnMappings.Add("batch_no", "batch_no");
                    bulkCopy.ColumnMappings.Add("cavity", "cavity");
                    bulkCopy.ColumnMappings.Add("start_weight", "start_weight");
                    bulkCopy.ColumnMappings.Add("D2_weight", "D2_weight");
                    bulkCopy.ColumnMappings.Add("D2_remark", "D2_remark");
                    bulkCopy.ColumnMappings.Add("D4_weight", "D4_weight");
                    bulkCopy.ColumnMappings.Add("D4_remark", "D4_remark");
                    bulkCopy.ColumnMappings.Add("D7_weight", "D7_weight");
                    bulkCopy.ColumnMappings.Add("D7_remark", "D7_remark");
                    bulkCopy.ColumnMappings.Add("deal_result", "deal_result");
                    bulkCopy.ColumnMappings.Add("jq_name", "jq_name");
                    bulkCopy.ColumnMappings.Add("ori_filename", "ori_filename");
                    bulkCopy.ColumnMappings.Add("new_filename", "new_filename");
                    bulkCopy.ColumnMappings.Add("upload_time", "upload_time");

                    bulkCopy.WriteToServer(dt);
                }
                catch (Exception ex)
                {
                    bf = false;
                }
            }

            return bf;
        }

    }

}