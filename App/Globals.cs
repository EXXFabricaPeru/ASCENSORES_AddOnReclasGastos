using Newtonsoft.Json.Serialization;
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
        public static string Level;
        public static string MonedaLocal;
        public const string LinkedSystemObject = "1";
        public const string LinkedUDO = "2";
        public const string LinkedTable = "3";

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

        public static void LlenarCombo(SAPbouiCOM.Form oForm, string ncombo, string qcombo, bool optLinea, string valLinea)
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(ncombo).Specific;
            Globals.RunQuery(qcombo);
            Globals.oRec.MoveFirst();

            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            if (optLinea) { oCombo.ValidValues.Add(valLinea, valLinea); }

            while (!Globals.oRec.EoF)
            {
                oCombo.ValidValues.Add(Globals.oRec.Fields.Item(0).Value.ToString(), Globals.oRec.Fields.Item(1).Value.ToString());
                Globals.oRec.MoveNext();
            }
            if (oCombo.ValidValues.Count > 0) oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            Globals.Release(oCombo);
            Globals.Release(Globals.oRec);
        }

        public static void LlenarComboMatrix(SAPbouiCOM.Matrix oMatrix, string nColumn, string qcombo, bool optLinea, string valLinea)
        {
            Globals.RunQuery(qcombo);
            Globals.oRec.MoveFirst();

            while (oMatrix.Columns.Item(nColumn).ValidValues.Count > 0)
            {
                oMatrix.Columns.Item(nColumn).ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            if (optLinea) { oMatrix.Columns.Item(nColumn).ValidValues.Add(valLinea, valLinea); }

            while (!Globals.oRec.EoF)
            {
                oMatrix.Columns.Item(nColumn).ValidValues.Add(Globals.oRec.Fields.Item(0).Value.ToString(), Globals.oRec.Fields.Item(1).Value.ToString());
                Globals.oRec.MoveNext();
            }
            //if (oMatrix.Columns.Item(nColumn).ValidValues.Count > 0) oMatrix.Columns.Item(nColumn)(0, SAPbouiCOM.BoSearchKey.psk_Index);
            Globals.Release(Globals.oRec);
        }

        public static bool ValidarMonedaSistema()
        {
            try
            {
                Globals.Query = Properties.Resources.ValidaMonedaSistema;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                string rpta = Globals.oRec.Fields.Item(0).Value.ToString();
                if (rpta == "Y")
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static string ObtenerSL()
        {
            try
            {
                Globals.Query = Properties.Resources.ObtenerConfiguracion;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                    throw new Exception("No se encuentra configurado la IP para service layer, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_ADRG_CONF - Configuración Reclas Gasto \nRegistre la IP de service layer con el código 001");
                else
                {
                    string server = Globals.oRec.Fields.Item("U_EXX_CONF_VALOR").Value.ToString();
                    if (string.IsNullOrEmpty(server))
                        throw new Exception("No se encuentra configurado la IP para service layer, por favor verifique: Herramientas > Ventanas definidas por usuario > EXX_ADRG_CONF - Configuración Reclas Gasto \nRegistre la IP de service layer con el código 001");
                    else
                        return server;
                }
            }
            catch (Exception ex)
            {
                Globals.MessageBox(ex.Message);
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ObtenerLevel()
        {
            try
            {
                Globals.Query = Properties.Resources.ObtenerLevel;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();
                Globals.Level = Globals.oRec.Fields.Item(0).Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ObtenerMonedaLocal()
        {
            try
            {
                Globals.Query = Properties.Resources.ObtenerMonedaLocal;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();
                Globals.MonedaLocal = Globals.oRec.Fields.Item(0).Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }
    }
}
