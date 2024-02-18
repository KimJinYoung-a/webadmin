using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class smart_Imageup_Popup : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void uploadbutton_Click(object sender, EventArgs e)
    {
        string tempPath = Server.MapPath("../Temp");

        if (uploadFile.HasFile)
        {
            FileInfo fi = new FileInfo(uploadFile.FileName);
            string fileName = Guid.NewGuid().ToString() + fi.Extension;
            string filePath = string.Format("{0}\\{1}", tempPath, fileName);
            uploadFile.SaveAs(filePath);
            fi = new FileInfo(filePath);

            string script = string.Format("onCompleteUpload('{0}', '{1}', '{2}');", fileName, uploadFile.FileName, fi.Length);
            if (!ClientScript.IsClientScriptBlockRegistered(this.GetType(), "onCompleteUpload"))
            {
                ClientScript.RegisterClientScriptBlock(this.GetType(), "onCompleteUpload", script, true);
            }
        }
    }
}
