using DevExpress.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;

namespace DotMercy.custom
{
    public partial class PackingList : System.Web.UI.Page
    {
        public static int GetdataSave_mVarianId = 0;
        public static string GetdataSave_mVarianName = "";

        public static int GetdataSave_mModelId = 0;
        public static string GetdataSave_mModelName = "";

        public static int GetdataSave_mPackingMonth = 0;
        public static string GetdataSave_mPackingMonthName = "";

        public static int GetdataSave_mFileType = 0;
        public static string GetdataSave_mFileTypeName = "";
        
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsCallback && !IsPostBack)
            {

                //VarianId.SelectedIndex = 0;
                //grid.DataBind();
                //grid.DetailRows.ExpandRow(0);
            }
        }
        protected void PlanGrid_DataSelect(object sender, EventArgs e)
        {
            Session["SessionId"] = (sender as ASPxGridView).GetMasterRowKeyValue();
            Session["SessionId2"] = (sender as ASPxGridView).GetMasterRowKeyValue();
            Session["SessionId3"] = (sender as ASPxGridView).GetMasterRowKeyValue();
            Session["SessionId4"] = (sender as ASPxGridView).GetMasterRowKeyValue();
        }


        protected void UploadControl_FileUploadComplete(object sender, FileUploadCompleteEventArgs e)
        {
            //RemoveFileWithDelay(e.UploadedFile.FileNameInStorage, 5);

            //string name = e.UploadedFile.FileName;
            //string url = GetImageUrl(e.UploadedFile.FileNameInStorage);
            //long sizeInKilobytes = e.UploadedFile.ContentLength / 1024;
            //string sizeText = sizeInKilobytes.ToString() + " KB";
            //e.CallbackData = name + "|" + url + "|" + sizeText;
        }

        /*---------------------START IMPORT FILE----------------------------*/
        protected void UcDataUji_FileUploadComplete(object sender, FileUploadCompleteEventArgs e)
        {
            try
            {
                e.CallbackData = this.SavePostedFile(e.UploadedFile);

                int mPackingMonthId = GetdataSave_mPackingMonth;
                int mModelId = GetdataSave_mModelId;
                int mVarianId = GetdataSave_mVarianId;
                int mFileType = GetdataSave_mFileType;

                Process_SaveMaster(mModelId, mVarianId, mPackingMonthId, mFileType);

            }
            catch (Exception ex)
            {
                e.IsValid = false;
                e.ErrorText = ex.Message;
            }
        }

        string SavePostedFile(UploadedFile uploadedFile)
        {
            //// return if File IS NOT VALID
            if (!uploadedFile.IsValid) return String.Empty;


            //=========cek folder Packing Month
            string path = "~/custom/FileUpload/" + GetdataSave_mPackingMonthName;
            if (!Directory.Exists(Server.MapPath(path)))
            {
                Directory.CreateDirectory(Server.MapPath(path));
            }

            //=========cek folder Model
            string pathModel = path + "/" + GetdataSave_mModelName;
            if (!Directory.Exists(Server.MapPath(pathModel)))
            {
                Directory.CreateDirectory(Server.MapPath(pathModel));
            }

            //=========cek folder Varian
            string pathVarian = pathModel + "/" + GetdataSave_mVarianName;
            if (!Directory.Exists(Server.MapPath(pathVarian)))
            {
                Directory.CreateDirectory(Server.MapPath(pathVarian));
            }

            String UploadDir = pathVarian + "/"; // "../custom/FileUpload/";


            FileInfo fileInfo = new FileInfo(uploadedFile.FileName);
            String fileNameOri = uploadedFile.FileName.ToString().Replace(" ", "_");
            String ext = System.IO.Path.GetExtension(uploadedFile.FileName);
            String fileType = uploadedFile.ContentType.ToString();
            if ((fileNameOri.Length - ext.Length) > 16)
            {
                fileNameOri = fileNameOri.Substring(0, 16).ToLower() + ext;
            }
            
            //String fileName = String.Format("PL_{0:yyMMddHHmm}_{1}", DateTime.Now, fileNameOri.ToLower());

            String fileName = fileNameOri;
            String resFileName = Server.MapPath(UploadDir + fileName);
            uploadedFile.SaveAs(resFileName);

            //type file check
            int mFileType = GetdataSave_mFileType;

            _ProcessExcel(ext, resFileName);

            String fileLabel = fileInfo.Name;
            double fileLength = Convert.ToDouble(uploadedFile.ContentLength / 1024); // kilobyte
            //int JumSampUji = (Session[SNAME_LIST_DETUJI] as List<DetailUjiDS>).Count;
            String ret = ""; // String.Format("{0}|{1}|{2}|{3}", fileName, fileLength, fileType, JumSampUji);
            return ret;
        }


        private void _ProcessExcel(string ext, string fileXls)
        {
            System.Data.DataTable dt = null;

            string connString = "";
            string strFileType = ext;
            string path = fileXls;
            //Connection String to Excel Workbook
            if (strFileType.Trim() == ".xls")
            {
                connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (strFileType.Trim() == ".xlsx")
            {
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            string query = "";  //"SELECT * FROM [CAT$]";
            OleDbConnection connExcel = new OleDbConnection(connString);
            if (connExcel.State == ConnectionState.Closed)
                connExcel.Open();

            //---------get sheet name
            int x = 0;
            dt = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            String[] excelSheets = new String[dt.Rows.Count];

            foreach (DataRow row in dt.Rows)
            {
                if (x == 0)
                {
                    excelSheets[x] = row["TABLE_NAME"].ToString();
                    query = "SELECT * FROM [" + excelSheets[x].Replace("'", "") + "]";
                }
                x++;
            }
            //---------------

            OleDbCommand cmdExcel = new OleDbCommand(query, connExcel);
            OleDbDataAdapter daExcel = new OleDbDataAdapter(cmdExcel);
            DataSet dsExcel = new DataSet();

            daExcel.Fill(dsExcel);

            OleDbDataReader rdrExcel;
            rdrExcel = cmdExcel.ExecuteReader();

            int data_Detail = 0;
            int i = 0;
            int j = 0;
            string strValue = "";
            string strCaption = "";

            Import_Delete();

            while (rdrExcel.Read())
            {
                string sqlvalues2 = "";
                string exlcaption = "";

                if (data_Detail >= 0)
                {

                    strValue = "";
                    strCaption = "";

                    for (i = 0; i < rdrExcel.FieldCount; i++)
                    {

                        sqlvalues2 = rdrExcel[i].ToString();
                        exlcaption = rdrExcel.GetName(i);

                        //if (sqlvalues2 != "")
                        //{
                        strValue = sqlvalues2 + "|" + strValue;

                        strCaption = exlcaption + "|" + strCaption;
                        //}

                    }

                }

                data_Detail++;

                if (strValue != "")
                {
                    j++;
                    Import_Proses(strValue, j);
                    //Import_Proses(strValue, strCaption, j);
                }
            }

            rdrExcel.Close();

        }


        private void Import_Delete()
        {
            SqlConnection conn = new SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings["AppDb"].ConnectionString);
            SqlCommand cmd = new SqlCommand();

            bool isError = false;
            string errMsg = "";

            try
            {

                conn.Open();
                cmd.Connection = conn;
                cmd.CommandTimeout = 0;
                cmd.CommandText = "delete PackingListDetails_tmp;";
                string exeMsg = Convert.ToString(cmd.ExecuteScalar());

                //check error
                if (!exeMsg.Trim().Equals(""))
                {
                    isError = true;
                    errMsg = exeMsg;
                }

                cmd.Parameters.Clear();

            }

            catch (Exception ex)
            {
                //Logger.Error(ex.Message);
                isError = true;
                errMsg = ex.Message;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
            }

            if (isError)
            {
                throw new InvalidOperationException(errMsg);
            }

        }

        private void Import_Proses(string strValue, int __no)
        {
            SqlConnection conn = new SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings["AppDb"].ConnectionString);
            SqlCommand cmd = new SqlCommand();

            bool isError = false;
            string errMsg = "";

            try
            {
                string[] strAll = strValue.Split('|');

                int detail_id = Convert.ToInt32(__no);

                int Identification = Convert.ToInt16(strAll[10]);
                string PackingCompany = Convert.ToString(strAll[9]);
                string PlantCode = Convert.ToString(strAll[8]);
                int Year = Convert.ToInt16(strAll[7]);
                int Month = Convert.ToInt16(strAll[6]);
                string Consignment = Convert.ToString(strAll[5]);
                string CountryCode = Convert.ToString(strAll[4]);
                string CountryDescription = Convert.ToString(strAll[3]);
                string Model = Convert.ToString(strAll[2]);
                string ModelDescription = Convert.ToString(strAll[1]);
                string Productionnofrom = Convert.ToString(strAll[0]);

                conn.Open();
                cmd.Connection = conn;
                cmd.CommandTimeout = 0;

                cmd.CommandText = "insert into PackingListDetails_tmp (Id, Identification, PackingCompany, PlantCode, Year, Month, Consignment, CountryCode, CountryDescription, Model, ModelDescription, " +
                                    " Productionnofrom ) " +
                                    " values (" + detail_id + "," + Identification + ", '" + PackingCompany + "','" + PlantCode + "', " +
                                    " " + Year + "," + Month + ",'" + Consignment + "','" + CountryCode + "','" + CountryDescription + "'," +
                                    " '" + Model + "','" + ModelDescription + "','" + Productionnofrom + "')";

                string exeMsg = Convert.ToString(cmd.ExecuteScalar());

                //insert master
                //cmd.CommandText = "";


                //check error
                if (!exeMsg.Trim().Equals(""))
                {
                    //Logger.Error("Execution Query : " + exeMsg);
                    isError = true;
                    errMsg = exeMsg;
                }
                /*else
                {
                    // if success, move file;
                    moveFileSurvTotxtDir(segmentId);
                }*/

                cmd.Parameters.Clear();

            }
            catch (Exception ex)
            {
                //Logger.Error(ex.Message);
                isError = true;
                errMsg = ex.Message;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
            }

            if (isError)
            {
                throw new InvalidOperationException(errMsg);
            }

        }

        //---------end import data excel-----------/\

        protected void btSave_Master(object sender, EventArgs e)
        {
            GetdataSave_mPackingMonth = Convert.ToInt16(PackingMonth.SelectedItem.Value);
            GetdataSave_mModelId = Convert.ToInt16(ModelId.SelectedItem.Value);
            GetdataSave_mVarianId = Convert.ToInt16(VarianId.SelectedItem.Value);
            
        }


        private void Process_SaveMaster(int intModelId, int intVarianId, int intPackingMonth, int intFileType)
        {
            SqlConnection conn = new SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings["AppDb"].ConnectionString);
            SqlCommand cmd = new SqlCommand();

            bool isError = false;
            string errMsg = "";

            try
            {
                conn.Open();
                cmd.Connection = conn;
                cmd.CommandTimeout = 0;

                cmd.CommandText = "insert into PackingLists (PackingMonth, ModelId, VarianId, FileType) " +
                                    " values (" + intPackingMonth + "," + intModelId + ", " + intVarianId + ", " + intFileType + ")";

                string exeMsg = Convert.ToString(cmd.ExecuteScalar());

                //check error
                if (!exeMsg.Trim().Equals(""))
                {
                    //Logger.Error("Execution Query : " + exeMsg);
                    isError = true;
                    errMsg = exeMsg;
                }

                cmd.Parameters.Clear();

            }
            catch (Exception ex)
            {
                //Logger.Error(ex.Message);
                isError = true;
                errMsg = ex.Message;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                conn.Dispose();
            }

            if (isError)
            {
                throw new InvalidOperationException(errMsg);
            }

        }

        protected void FileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetdataSave_mPackingMonth = Convert.ToInt16(PackingMonth.SelectedItem.Value);
            GetdataSave_mModelId = Convert.ToInt16(ModelId.SelectedItem.Value);
            GetdataSave_mVarianId = Convert.ToInt16(VarianId.SelectedItem.Value);
            GetdataSave_mFileType = Convert.ToInt16(FileType.SelectedItem.Value);

            SqlConnection conn = new SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings["AppDb"].ConnectionString);
            conn.Open();

            
            //---get data name Packing Month            
            string sqlPM = string.Empty;
            sqlPM = "select PackingMth from PackingMonths where Id=" + GetdataSave_mPackingMonth + " ";
            SqlCommand selectPM = new SqlCommand(sqlPM, conn);
            selectPM.CommandTimeout = 0;
            SqlDataReader RDPM = selectPM.ExecuteReader();

            while (RDPM.Read())
            {
                GetdataSave_mPackingMonthName = Convert.ToString(RDPM["PackingMth"]);
            }
            RDPM.Close();

            //---get data name Model            
            string sqlMD = string.Empty;
            sqlMD = "select VarianName from Varians where Id=" + GetdataSave_mModelId + " ";
            SqlCommand selectMD = new SqlCommand(sqlMD, conn);
            selectMD.CommandTimeout = 0;
            SqlDataReader RDMD = selectMD.ExecuteReader();

            while (RDMD.Read())
            {
                GetdataSave_mModelName = Convert.ToString(RDMD["VarianName"]);
            }
            RDMD.Close();


            //---get data name Varian
            string sqlVR = string.Empty;
            sqlVR = "select ModelVarian from VarianDetails where Id=" + GetdataSave_mVarianId + " ";
            SqlCommand selectVR = new SqlCommand(sqlVR, conn);
            selectVR.CommandTimeout = 0;
            SqlDataReader RDVR = selectVR.ExecuteReader();

            while (RDVR.Read())
            {
                GetdataSave_mVarianName = Convert.ToString(RDVR["ModelVarian"]);
            }
            RDVR.Close();

            //---get data name File Type
            string sql = string.Empty;
            sql = "select Name from FileType where Id=" + GetdataSave_mFileType + " ";
            SqlCommand select = new SqlCommand(sql, conn);
            select.CommandTimeout = 0;
            SqlDataReader RD = select.ExecuteReader();

            while (RD.Read())
            {
                GetdataSave_mFileTypeName = Convert.ToString(RD["Name"]);
            }
            RD.Close();


        }

        protected void FileType_Callback(object sender, CallbackEventArgsBase e)
        {
            if (String.IsNullOrEmpty(e.Parameter))
               return;

            GetdataSave_mFileType = Convert.ToInt16(e.Parameter);
            //txtFileType.Value = Convert.ToInt16(e.Parameter);

        }
               

    }
}