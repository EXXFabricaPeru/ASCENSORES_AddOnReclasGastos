using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddOnRclsGastos.App
{
    public class Connect
    {

        public static void SetApplication()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = default(SAPbouiCOM.SboGuiApi);
                string sConnectionString = null;
                SboGuiApi = new SAPbouiCOM.SboGuiApi();
                if (Environment.GetCommandLineArgs().Length > 1)
                {
                    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                }
                else
                {
                    sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(0));
                }
                SboGuiApi.Connect(sConnectionString);
                Globals.SBO_Application = SboGuiApi.GetApplication();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static bool ConnectToCompany()
        {
            try
            {
                //Globals.oCompany = (SAPbobsCOM.Company)Globals.SBO_Application.Company.GetDICompany();

                Globals.oCompany = new SAPbobsCOM.Company();

                string cookie = Globals.oCompany.GetContextCookie();

                string conStr = Globals.SBO_Application.Company.GetConnectionContext(cookie);

                if (Globals.oCompany.Connected)
                {
                    Globals.oCompany.Disconnect();
                }

                int ret = Globals.oCompany.SetSboLoginContext(conStr);

                if (ret != 0)
                {
                    throw new Exception("Login context failed");
                }
                Globals.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La;
                ret = Globals.oCompany.Connect();

                Globals.oCompany.GetLastError(out Globals.lRetCode, out Globals.sErrMsg);

                return true;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return false;
            }
        }
    }
}
