using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace GED
{
    public class SPConnection
    {
        public static ClientContext GetSPOLContext(string webUrl)
        {
            try
            {
                var pwd = Environment.GetEnvironmentVariable("SPUserPassword", EnvironmentVariableTarget.Process);


                if (pwd == null)
                {
                    pwd = "DWVOQOEZ2M";
                    //throw new Exception("Key missing on the Function App configuration.");
                }


                SecureString password = new SecureString();
                foreach (char c in pwd.ToCharArray()) password.AppendChar(c);


                string username = Environment.GetEnvironmentVariable("SPUserName", EnvironmentVariableTarget.Process);


                if (username == null)
                {
                    username = "jkeyrouz@ghtpdfr.fr";
                    //throw new Exception("Key missing on the Function App configuration.");
                }


                string siteUrl = Environment.GetEnvironmentVariable("SPSiteUrl", EnvironmentVariableTarget.Process);
                if (siteUrl == null)
                {
                    siteUrl = "https://ghtpdfr.sharepoint.com/sites/ged";
                }


                var ctx = new ClientContext(siteUrl);
                ctx.Credentials = new SharePointOnlineCredentials(username, password);
                ctx.Load(ctx.Web, w => w.Title);


                ctx.ExecuteQuery();


                return ctx;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    }
}
