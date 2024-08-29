using AddOnRclsGastos.DBStructure;
using AddOnRclsGastos.Functionality;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.App
{
    public class Main
    {
        public Main()
        {
            Connect.SetApplication();
            Connect.ConnectToCompany();
            Globals.SAPVersion = Globals.oCompany.Version;
            Globals.SBO_Application.SetStatusBarMessage("Validando estructura de la Base de Datos", SAPbouiCOM.BoMessageTime.bmt_Short, false);

            #region Revisa Versión Cloud
            Globals.Addon = Assembly.GetEntryAssembly().GetName().Name;
            Version version = Assembly.GetEntryAssembly().GetName().Version;
            Globals.version = version.ToString().Replace(".0.0", "");
            #endregion
            #region Estructura
            //Setup oSetup = new Setup();
            //Globals.Actual = oSetup.validarVersion(Globals.Addon, Globals.version);
            if (Globals.Actual == false)
            {
                CreateStructure.CreateStruct();
                //    oSetup.confirmarVersion(Globals.Addon, Globals.version);
                //    oSetup.confirmarVersionUpdate(Globals.Addon, Globals.version);
                Globals.continuar = 0;
            }
            else
                Globals.continuar = 0;
            #endregion
            Globals.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Globals.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            Globals.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            Globals.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            Menu.AddMenuItems();
            Globals.SBO_Application.StatusBar.SetText("El Add-On Reclasificación de Gastos está conectado.", SAPbouiCOM.BoMessageTime.bmt_Short, (SAPbouiCOM.BoStatusBarMessageType)SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            SAPbouiCOM.Form oForm = null;
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        #region SBAControlComp
                        case "EXX_ADRG1":
                            SRF_AsistenteGastos.LoadFormEstr(pVal.MenuUID);
                            break;
                        case "EXX_ADRG2":
                            SRF_HistoricoAsientos.LoadFormEstr(pVal.MenuUID);
                            break;
                        default:
                            break;

                            #endregion
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                if (ex.Message.IndexOf("Form - Not found  [66000-9]") != -1)
                {
                    Globals.Error = "EXX: Activar campos de usuario al crear un documento";
                    Globals.SBO_Application.SetStatusBarMessage(Globals.Error, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    Globals.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx != "0")
                {
                    SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(pVal.FormUID);
                    if (!oForm.Title.ToUpper().Contains("CANCEL"))
                    {
                        switch (pVal.FormTypeEx)
                        {
                            case "EXX_ADRG1":
                                switch (pVal.EventType)
                                {
                                    case BoEventTypes.et_ITEM_PRESSED:
                                        SRF_AsistenteGastos.ItemPressed(ref pVal, oForm, out BubbleEvent); ////
                                        break;
                                    case BoEventTypes.et_COMBO_SELECT:
                                        SRF_AsistenteGastos.ComboSelect(ref pVal, oForm, out BubbleEvent);
                                        break;
                                    case BoEventTypes.et_LOST_FOCUS:
                                        SRF_AsistenteGastos.LostFocus(ref pVal, oForm, out BubbleEvent);
                                        break;
                                }
                                break;
                            case "EXX_ADRG2":
                                switch (pVal.EventType)
                                {
                                    case BoEventTypes.et_ITEM_PRESSED:
                                        SRF_HistoricoAsientos.ItemPressed(ref pVal, oForm, out BubbleEvent); ////
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                if (ex.Message == "Form - Invalid Form")
                    return;

                if (ex.Message.IndexOf("Form - Not found  [66000-9]") != -1)
                {
                    Globals.Error = "EXX: Activar campos de usuario al crear un documento";
                    Globals.SBO_Application.SetStatusBarMessage(Globals.Error, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    Globals.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //SAPbouiCOM.Form oForm = Globals.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                Globals.SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown)
                System.Windows.Forms.Application.Exit();
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
                System.Windows.Forms.Application.Exit();
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged)
                System.Windows.Forms.Application.Exit();
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_FontChanged)
                System.Windows.Forms.Application.Exit();
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition)
                System.Windows.Forms.Application.Exit();
        }
    }
}
