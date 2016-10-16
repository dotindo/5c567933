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
    public partial class FormFileManager : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        /*
        protected void fileManager_FileUploading(object sender, FileManagerFileUploadEventArgs e)
        {
            ValidateSiteEdit(e);
        }

        protected void fileManager_ItemRenaming(object sender, FileManagerItemRenameEventArgs e)
        {
            ValidateSiteEdit(e);
        }

        protected void fileManager_ItemMoving(object sender, FileManagerItemMoveEventArgs e)
        {
            ValidateSiteEdit(e);
        }

        protected void fileManager_ItemDeleting(object sender, FileManagerItemDeleteEventArgs e)
        {
            ValidateSiteEdit(e);
        }

        protected void fileManager_FolderCreating(object sender, FileManagerFolderCreateEventArgs e)
        {
            ValidateSiteEdit(e);
        }

        
        void ValidateSiteEdit(FileManagerActionEventArgsBase e)
        {
            e.Cancel = Utils.IsSiteMode;
            e.ErrorText = Utils.GetReadOnlyMessageText();
        } */

    }
}