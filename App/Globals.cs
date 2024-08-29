using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.App
{
    public class Globals
    {
        public static String ShortName = "(EXX)";
        public static int continuar = -1;
        public static string Query = null;
        public static SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.CompanyService oCmpSrv;
        public static SAPbouiCOM.EventFilters oFilters;
        public static SAPbouiCOM.EventFilter oFilter;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Company oCompanyMirror;
        public static int SAPVersion;
        public static string Addon = null;
        public static string version = null;
        public static string oldversion = "";
        public static bool Actual = false;
        public static int lRetCode;
        public static int sErrCode;
        public static string sErrMsg = null;
        public static string Error = null;

        public static object Release(object objeto)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
            Query = null;
            GC.Collect();
            return null;
        }

        public static void ErrorMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short);
        }

        public static void InformationMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public static void SuccessMessage(String msg)
        {
            Globals.SBO_Application.SetStatusBarMessage(ShortName + ": " + msg, SAPbouiCOM.BoMessageTime.bmt_Short, false);
        }

        public static void MessageBox(String msg)
        {
            Globals.SBO_Application.MessageBox(msg);
        }

        public static SAPbobsCOM.Recordset RunQuery(string Query)
        {
            try
            {
                oRec = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(Query);
                return oRec;
            }
            catch (Exception ex)
            {
                Globals.SBO_Application.MessageBox(ex.Message);
                return null;
            }
        }

        public static string LoadFromXML(ref string FileName)
        {
            System.Xml.XmlDocument oXmlDoc = null;
            string sPath = null;
            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath;
            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }

        public static DateTime ConvertDate(string date)
        {
            if (date.Length == 8)
            {
                date = date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2);
                return Convert.ToDateTime(date);
            }
            else
            {
                throw new Exception("Invalid Date Format");
            }
        }

        public static void StartTransaction()
        {
            try
            {
                Globals.oCompany.StartTransaction();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static bool InTransaction()
        {
            try
            {
                return Globals.oCompany.InTransaction;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void CommitTransaction()
        {
            try
            {
                Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void RollBackTransaction()
        {
            try
            {
                if (Globals.oCompany.InTransaction)
                {
                    Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
