using AddOnFacturador.App;
using AddOnRclsGastos.App;
using AddOnRclsGastos.Entity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.Functionality
{
    public class SRF_AsistenteGastos
    {
        public static List<string> oCecos = new List<string>();
        public static List<string> oCuentas = new List<string>();
        public static List<string> oMonedas = new List<string>();
        public static List<string> oProyectos = new List<string>();
        public static List<int> oListaIndices = new List<int>();
        public static OJDT oAsiento = new OJDT();
        public static List<JDT1> oListaxDistribuir = new List<JDT1>();
        public static List<OPRC> oDistribucion = new List<OPRC>();
        public static List<OPRJ> oProyectosD = new List<OPRJ>();
        public static void LoadFormEstr(string Frm)
        {
            SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
            try
            {
                oForm = Globals.SBO_Application.Forms.Item(Frm);
                Globals.SBO_Application.MessageBox("El formulario ya se encuentra abierto.");
                oForm.Visible = true;
            }
            catch
            {
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = (SAPbouiCOM.FormCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = Frm;
                fcp.UniqueID = Frm;
                string FormName = "\\Views\\frmAsistenteGastos.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);

                Globals.ObtenerMonedaLocal();
                Globals.ObtenerLevel();
                BuildCFLCuenta(oForm, "CFL_1");
                BuildCFLCuenta(oForm, "CFL_2");
                Globals.Query = Properties.Resources.ListarDimensiones;
                Globals.LlenarCombo(oForm, "cb1_1", Globals.Query, false, "");
                CargaGrid("grid1_2", oForm);
                CargaGrid("grid1_3", oForm);
                oForm.Visible = true;
                if (Globals.ValidarMonedaSistema())
                    Globals.MessageBox("En configuración de Asientos debe desmarcar la casilla 'Bloquear tratamiento de totales en moneda del sistema'. Por favor revise:\nMódulos > Gestión > Inicialización sistema > Parametrizaciones de documento");
            }
        }

        private static void BuildCFLCuenta(Form oForm, string CFL)
        {
            SAPbouiCOM.ChooseFromList oCFL_BP;
            oCFL_BP = oForm.ChooseFromLists.Item(CFL);

            SAPbouiCOM.Conditions oCons;
            SAPbouiCOM.Condition oCon;

            oCons = oCFL_BP.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "Levels";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = Globals.Level;

            //oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            //oCon = oCons.Add();
            //oCon.Alias = "Finanse";
            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            //oCon.CondVal = "Y";

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            oCon = oCons.Add();
            oCon.Alias = "U_EXX_ADRG_CTAGASTO";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";

            oCFL_BP.SetConditions(oCons);
        }

        public static void ItemPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.BeforeAction)
                    switch (pVal.ItemUID)
                    {
                        case "btnEjec":
                        case "btnAnt":
                        case "btnSig":
                            oForm.Items.Item("etAux").Click();
                            break;
                    }


                if (pVal.ActionSuccess)
                    switch (pVal.ItemUID)
                    {
                        case "btn1":
                            ListarCuentas(pVal, oForm, out BubbleEvent);
                            break;
                        case "btnEjec":
                            Ejecutar(pVal, oForm, out BubbleEvent);
                            break;
                        case "btnAnt":
                            Anterior(pVal, oForm, out BubbleEvent);
                            break;
                        case "btnSig":
                            Siguiente(pVal, oForm, out BubbleEvent);
                            break;
                    }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void ListarCuentas(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                CargaGrid("grid1_2", oForm);
                Globals.SuccessMessage("Lista de cuentas obtenida correctamente.");
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
        }

        private static void Ejecutar(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Button btnEjecutar = (Button)oForm.Items.Item("btnEjec").Specific;
            Button btnAnterior = (Button)oForm.Items.Item("btnAnt").Specific;
            Button btnSiguiente = (Button)oForm.Items.Item("btnSig").Specific;

            try
            {
                if (Globals.ValidarMonedaSistema())
                    Globals.MessageBox("En configuración de Asientos debe desmarcar la casilla 'Bloquear tratamiento de totales en moneda del sistema'. Por favor revise:\nMódulos > Gestión > Inicialización sistema > Parametrizaciones de documento");
                else
                {
                    if ((Globals.SBO_Application.MessageBox("¿Esta seguro de crear el asiento distribuido?.", 1, "Si", "No") == 1))
                    {
                        btnEjecutar.Item.Enabled = false;
                        btnAnterior.Item.Enabled = false;
                        btnSiguiente.Item.Enabled = false;


                        #region SDK
                        //SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific;
                        //SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_4");
                        //int Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);

                        //Globals.StartTransaction();
                        //SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        //oJE.Memo = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Value;
                        //oJE.ReferenceDate = Globals.ConvertDate(((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Value);
                        //oJE.TransactionCode = "RCSF";

                        //for (int i = 0; i < oLista.Rows.Count; i++)
                        //{
                        //    if (i != 0) oJE.Lines.Add();

                        //    oJE.Lines.BPLID = 2;
                        //    oJE.Lines.AccountCode = oLista.GetValue("Col_0", i).ToString();
                        //    if (Convert.ToDouble(oLista.GetValue("Col_3", i)) > 0)
                        //    {
                        //        oJE.Lines.Debit = Convert.ToDouble(oLista.GetValue("Col_3", i));
                        //        oJE.Lines.Credit = 0;
                        //    }
                        //    if (Convert.ToDouble(oLista.GetValue("Col_4", i)) > 0)
                        //    {
                        //        oJE.Lines.Debit = 0;
                        //        oJE.Lines.Credit = Convert.ToDouble(oLista.GetValue("Col_4", i));
                        //    }

                        //    if (Convert.ToDouble(oLista.GetValue("Col_5", i)) > 0)
                        //    {
                        //        oJE.Lines.DebitSys = Convert.ToDouble(oLista.GetValue("Col_5", i));
                        //        oJE.Lines.CreditSys = 0;
                        //    }
                        //    if (Convert.ToDouble(oLista.GetValue("Col_6", i)) > 0)
                        //    {
                        //        oJE.Lines.DebitSys = 0;
                        //        oJE.Lines.CreditSys = Convert.ToDouble(oLista.GetValue("Col_6", i));
                        //    }

                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_7", i).ToString())) oJE.Lines.CostingCode = oLista.GetValue("Col_7", i).ToString();
                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_9", i).ToString())) oJE.Lines.CostingCode2 = oLista.GetValue("Col_9", i).ToString();
                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_11", i).ToString())) oJE.Lines.CostingCode3 = oLista.GetValue("Col_11", i).ToString();
                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_13", i).ToString())) oJE.Lines.CostingCode4 = oLista.GetValue("Col_13", i).ToString();
                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_15", i).ToString())) oJE.Lines.CostingCode5 = oLista.GetValue("Col_15", i).ToString();
                        //    if (!string.IsNullOrEmpty(oLista.GetValue("Col_17", i).ToString())) oJE.Lines.ProjectCode = oLista.GetValue("Col_17", i).ToString();
                        //}

                        //Globals.lRetCode = oJE.Add();
                        //if (Globals.lRetCode != 0)
                        //{
                        //    Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        //    throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                        //}
                        //else
                        //{
                        //    string TransId;
                        //    Globals.oCompany.GetNewObjectCode(out TransId);

                        //    SAPbobsCOM.UserTable oUserTable = Globals.oCompany.UserTables.Item("EXX_ADRG_HIST");
                        //    if (!oUserTable.GetByKey("0"))
                        //    {
                        //        oUserTable.UserFields.Fields.Item("U_EXX_ADRG_FECHAE").Value = DateTime.Now.ToString("dd/MM/yyyy");
                        //        oUserTable.UserFields.Fields.Item("U_EXX_ADRG_TRANSID").Value = TransId;
                        //        oUserTable.UserFields.Fields.Item("U_EXX_ADRG_FECHAC").Value = Globals.ConvertDate(((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Value).ToString("dd/MM/yyyy");
                        //        oUserTable.UserFields.Fields.Item("U_EXX_ADRG_GLOSA").Value = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Value;
                        //        oUserTable.UserFields.Fields.Item("U_EXX_ADRG_EST").Value = "G";

                        //        if (oUserTable.Add() != 0)
                        //        {
                        //            Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        //            throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                        //        }
                        //        else
                        //        {
                        //            Globals.CommitTransaction();
                        //            btnEjecutar.Item.Enabled = false;
                        //            btnAnterior.Item.Enabled = false;
                        //            btnSiguiente.Item.Enabled = true;
                        //            ((EditText)oForm.Items.Item("et4_1").Specific).Value = TransId;
                        //            ((StaticText)oForm.Items.Item("st4_1").Specific).Item.Visible = true;
                        //            ((EditText)oForm.Items.Item("et4_1").Specific).Item.Visible = true;
                        //            ((LinkedButton)oForm.Items.Item("lb4_1").Specific).Item.Visible = true;
                        //            Globals.MessageBox("El asiento de reclasificación se ha generado correctamente con el número de transacción " + TransId);
                        //        }
                        //        Globals.Release(oUserTable);
                        //    }
                        //}
                        #endregion

                        #region SL
                        oAsiento.LineMemo = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Value;
                        oAsiento.ReferenceDate = Globals.ConvertDate(((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Value).ToString("yyyy-MM-dd");
                        oAsiento.TransactionCode = "RCSF";
                        oAsiento.Details.ForEach(x =>
                        {
                            //x.BPLID = 2; //Comentar
                            if (string.IsNullOrEmpty(x.FCCurrency))
                                x.FCCurrency = null;
                            //{
                            //    x.FCDebit = null;
                            //    x.FCCredit = null;
                            //}
                        });


                        var rsp = SL.CrearAsiento(oAsiento);
                        dynamic ojdtResult = JObject.Parse(rsp.Content);
                        if (rsp.StatusCode != System.Net.HttpStatusCode.OK && rsp.StatusCode != System.Net.HttpStatusCode.Created && rsp.StatusCode != System.Net.HttpStatusCode.NoContent && rsp.StatusCode != 0)
                        {
                            btnEjecutar.Item.Enabled = true;
                            btnAnterior.Item.Enabled = true;
                            btnSiguiente.Item.Enabled = false;
                            Globals.ErrorMessage(SapError.get_message(rsp.Content));
                        }
                        else
                        {
                            string TransId = ojdtResult.JdtNum;

                            SAPbobsCOM.UserTable oUserTable = Globals.oCompany.UserTables.Item("EXX_ADRG_HIST");
                            if (!oUserTable.GetByKey("0"))
                            {
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_FECHAE").Value = DateTime.Now.ToString("dd/MM/yyyy");
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_TRANSID").Value = TransId;
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_FECHAC").Value = Globals.ConvertDate(((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Value).ToString("dd/MM/yyyy");
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_GLOSA").Value = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Value;
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_EST").Value = "G";

                                if (oUserTable.Add() != 0)
                                {
                                    btnEjecutar.Item.Enabled = true;
                                    btnAnterior.Item.Enabled = true;
                                    btnSiguiente.Item.Enabled = false;
                                    throw new SapError("Error SAP:" + rsp.Content);
                                }
                                else
                                {
                                    foreach (var item in oListaxDistribuir.GroupBy(x => x.TransId))
                                    {
                                        string Lines = string.Join(",", oListaxDistribuir.Where(x => x.TransId == item.Key && x.Line_ID != null).Select(y => y.Line_ID).ToList());
                                        Globals.Query = Properties.Resources.ActualizaAsiento;
                                        Globals.Query = string.Format(Globals.Query, TransId, item.Key, Lines);
                                        Globals.RunQuery(Globals.Query);
                                    }
                                    
                                    Globals.Query = Properties.Resources.ActualizaAsientoGenerado;
                                    Globals.Query = string.Format(Globals.Query, TransId);
                                    Globals.RunQuery(Globals.Query);

                                    btnEjecutar.Item.Enabled = false;
                                    btnAnterior.Item.Enabled = false;
                                    btnSiguiente.Item.Enabled = true;
                                    ((EditText)oForm.Items.Item("et4_1").Specific).Value = TransId;
                                    ((StaticText)oForm.Items.Item("st4_1").Specific).Item.Visible = true;
                                    ((EditText)oForm.Items.Item("et4_1").Specific).Item.Visible = true;
                                    ((LinkedButton)oForm.Items.Item("lb4_1").Specific).Item.Visible = true;
                                    Globals.MessageBox("El asiento de reclasificación se ha generado correctamente con el número de transacción " + TransId);
                                    Globals.SuccessMessage("Asiento creado correctamente");
                                }
                                Globals.Release(oUserTable);
                            }
                        }
                        #endregion


                    }
                }
            }
            catch (Exception ex)
            {
                if (Globals.InTransaction())
                    Globals.RollBackTransaction();

                btnEjecutar.Item.Enabled = true;
                btnAnterior.Item.Enabled = true;
                btnSiguiente.Item.Enabled = false;
                throw;
            }
        }

        public static void Anterior(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                StaticText stPaso = (StaticText)oForm.Items.Item("stPaso").Specific;
                StaticText stPasoDes = (StaticText)oForm.Items.Item("stPasoDes").Specific;
                Button btnEjecutar = (Button)oForm.Items.Item("btnEjec").Specific;
                Button btnAnterior = (Button)oForm.Items.Item("btnAnt").Specific;
                Button btnSiguiente = (Button)oForm.Items.Item("btnSig").Specific;
                Grid oGrid;
                DataTable oLista;

                oForm.Freeze(true);
                switch (stPaso.Caption.ToString())
                {
                    case "Paso 2 de 4":
                        oListaIndices = new List<int>();
                        stPaso.Caption = "Paso 1 de 4";
                        btnEjecutar.Item.Visible = false;
                        btnAnterior.Item.Enabled = false;
                        stPasoDes.Caption = "Selección de parámetros de búsqueda";
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Item.Click(BoCellClickType.ct_Right);
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_6").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_7").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_8").Specific).Item.Visible = true;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_6").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Button)oForm.Items.Item("btn1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = false;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");
                        if (oGrid.DataTable != null)
                        {
                            oGrid.DataTable.Rows.Clear();
                            oLista.Rows.Clear();
                            oGrid.DataTable = null;
                        }
                        break;
                    case "Paso 3 de 4":
                        oListaIndices = new List<int>();
                        stPaso.Caption = "Paso 2 de 4";
                        btnEjecutar.Item.Visible = false;
                        stPasoDes.Caption = "Selección de contabilizaciones";
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = false;
                        oProyectos = new List<string>();
                        //oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific;
                        //oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3_1");
                        //if (oGrid.DataTable != null)
                        //{
                        //    oGrid.DataTable.Rows.Clear();
                        //    oLista.Rows.Clear();
                        //    oGrid.DataTable = null;
                        //}
                        break;
                    case "Paso 4 de 4":
                        oListaIndices = new List<int>();
                        stPaso.Caption = "Paso 3 de 4";
                        btnEjecutar.Item.Visible = false;
                        btnSiguiente.Caption = "Siguiente";
                        btnSiguiente.Item.Enabled = true;
                        stPasoDes.Caption = "Datos para asiento contable";
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific).Item.Visible = true;
                        //((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = true;
                        if (oProyectos.Count > 0) ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = true;
                        else ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific).Item.Visible = false;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_4");
                        if (oGrid.DataTable != null)
                        {
                            oGrid.DataTable.Rows.Clear();
                            oLista.Rows.Clear();
                            oGrid.DataTable = null;
                        }
                        break;
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
        }

        public static void Siguiente(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                StaticText stPaso = (StaticText)oForm.Items.Item("stPaso").Specific;
                StaticText stPasoDes = (StaticText)oForm.Items.Item("stPasoDes").Specific;
                Button btnEjecutar = (Button)oForm.Items.Item("btnEjec").Specific;
                Button btnAnterior = (Button)oForm.Items.Item("btnAnt").Specific;
                Button btnSiguiente = (Button)oForm.Items.Item("btnSig").Specific;
                Grid oGrid;
                DataTable oLista;

                oForm.Freeze(true);
                switch (stPaso.Caption.ToString())
                {
                    case "Paso 1 de 4":
                        CargaLista("grid1_1", oForm);
                        CargaLista("grid1_2", oForm);
                        CargaLista("grid1_3", oForm);
                        Validar(oForm);
                        stPaso.Caption = "Paso 2 de 4";
                        stPasoDes.Caption = "Selección de contabilizaciones";
                        btnAnterior.Item.Enabled = true;
                        btnEjecutar.Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_6").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_7").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_8").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_6").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Button)oForm.Items.Item("btn1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = true;
                        CargaGrid("grid2", oForm);
                        break;
                    case "Paso 2 de 4":
                        CargaLista("grid2", oForm);
                        Validar(oForm);
                        stPaso.Caption = "Paso 3 de 4";
                        stPasoDes.Caption = "Datos para asiento contable";
                        btnEjecutar.Item.Visible = false;
                        btnSiguiente.Item.Enabled = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific).Item.Visible = true;
                        if (oProyectos.Count > 0) ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = true;
                        else ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = false;
                        CargaGrid("grid3_2", oForm);
                        break;
                    case "Paso 3 de 4":
                        CargaLista("grid1_1", oForm);
                        CargaLista("grid1_2", oForm);
                        CargaLista("grid2", oForm);
                        Validar(oForm);
                        stPaso.Caption = "Paso 4 de 4";
                        btnEjecutar.Item.Visible = true;
                        stPasoDes.Caption = "Simulación de distribucíón de gastos";
                        btnAnterior.Item.Enabled = true;
                        btnSiguiente.Item.Enabled = false;
                        btnSiguiente.Caption = "Finalizar";
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific).Item.Visible = true;
                        CargaGrid("grid4", oForm);
                        break;
                    case "Paso 4 de 4":
                        oForm.Close();
                        break;
                }
                //Validar(oForm);
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void Validar(Form oForm)
        {
            try
            {
                StaticText stPaso = (StaticText)oForm.Items.Item("stPaso").Specific;
                switch (stPaso.Caption.ToString())
                {
                    case "Paso 1 de 4":
                        SAPbouiCOM.ComboBox cb1_1 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific;
                        SAPbouiCOM.EditText et1_2 = (SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific;
                        SAPbouiCOM.EditText et1_3 = (SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific;
                        if (cb1_1.Selected == null) throw new Exception("Debe seleccionar una dimensión");
                        if (string.IsNullOrEmpty(et1_2.Value.Trim())) throw new Exception("Debe ingresar una fecha de inicio");
                        if (string.IsNullOrEmpty(et1_3.Value.Trim())) throw new Exception("Debe ingresar una fecha de fin");
                        if (oCecos.Count == 0) throw new Exception("Debe seleccionar al menos un centro de costo de la lista");
                        if (oCuentas.Count == 0) throw new Exception("Debe seleccionar al menos una cuenta de la lista");
                        if (oMonedas.Count == 0) throw new Exception("Debe seleccionar al menos una moneda de la lista");
                        else
                        {
                            if (string.IsNullOrEmpty(Globals.MonedaLocal)) Globals.ObtenerMonedaLocal();
                            if (oMonedas.Where(x => x != Globals.MonedaLocal).ToList().Count > 1) throw new Exception("Solo puede seleccionar una moneda extranjera.");
                        }
                        break;
                    case "Paso 2 de 4":
                        if (oListaIndices.Count == 0) throw new Exception("Debe seleccionar al menos un asiento de la lista");
                        break;
                    case "Paso 3 de 4":
                        SAPbouiCOM.EditText et3_2 = (SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific;
                        SAPbouiCOM.EditText et3_3 = (SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific;
                        SAPbouiCOM.ComboBox cb3_4 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific;
                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific;
                        SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3_1");
                        oDistribucion = new List<OPRC>();

                        if (string.IsNullOrEmpty(et3_2.Value.Trim())) throw new Exception("Debe ingresar una fecha de contabilización");
                        if (string.IsNullOrEmpty(et3_3.Value.Trim())) throw new Exception("Debe ingresar una glosa para el asiento");
                        if (cb3_4.Selected == null) throw new Exception("Debe seleccionar una reglas de distribución");
                        if (oLista == null || oLista.Rows.Count == 0) throw new Exception("No se encontró la lista para la distribución");
                        if (string.IsNullOrEmpty(oLista.GetValue("PrcCode", 0).ToString())) throw new Exception("No se encontró la lista para la distribución");
                        else
                        {
                            for (int i = 0; i < oLista.Rows.Count; i++)
                            {
                                if (Convert.ToInt32(oLista.GetValue("Peso", i).ToString()) > 0)
                                    oDistribucion.Add(new OPRC
                                    {
                                        PrcCode = oLista.GetValue("PrcCode", i).ToString(),
                                        PrcName = oLista.GetValue("PrcName", i).ToString(),
                                        Peso = Convert.ToInt32(oLista.GetValue("Peso", i).ToString())
                                    });
                            }

                            if (oDistribucion.Sum(x => x.Peso) == 0) throw new Exception("Los valores ingresados para la ditribución no puedan totalizar 0, por favor verifique");
                        }

                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3_2");
                        oProyectosD = new List<OPRJ>();
                        for (int i = 0; i < oLista.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(oLista.GetValue("U_EXX_ADRG_PESO", i).ToString()) > 0)
                                oProyectosD.Add(new OPRJ
                                {
                                    PrjCode = oLista.GetValue("Code", i).ToString(),
                                    PrjDestino = oLista.GetValue("U_EXX_ADRG_PRJD", i).ToString(),
                                    Peso = Convert.ToInt32(oLista.GetValue("U_EXX_ADRG_PESO", i).ToString())
                                });
                        }

                        List<string> ProyectoNoDist = new List<string>();
                        for (int i = 0; i < oProyectos.Count; i++)
                        {
                            if (oProyectosD.Where(x => x.PrjCode == oProyectos[i]).Count() == 0)
                                ProyectoNoDist.Add(oProyectos[i]);
                        }

                        if (ProyectoNoDist.Count > 0)
                            if ((Globals.SBO_Application.MessageBox("Para los proyectos: " + string.Join(", ", ProyectoNoDist) + ", no se encontró distribución, de continuar no se realizará distribución para estos proyectos. ¿Continuar?.", 1, "Si", "No") != 1))
                                throw new Exception("Verifique la distribución de los proyectos: " + string.Join(", ", ProyectoNoDist));
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ComboSelect(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                    switch (pVal.ItemUID)
                    {
                        case "cb1_1":
                            CargaGrid("grid1_1", oForm);
                            break;
                        case "cb3_4":
                            CargaGrid("grid3_1", oForm);
                            break;
                    }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void ChooseFromList(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    switch (pVal.ItemUID)
                    {
                        case "et1_5":

                            if (oDataTable != null)
                                ((SAPbouiCOM.EditText)oForm.Items.Item("et1_5").Specific).Value = oDataTable.GetValue("FormatCode", 0).ToString();
                            break;
                        case "et1_6":
                            if (oDataTable != null)
                                ((SAPbouiCOM.EditText)oForm.Items.Item("et1_6").Specific).Value = oDataTable.GetValue("FormatCode", 0).ToString();
                            break;
                        case "et3_5":

                            if (oDataTable != null)
                            {
                                ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Value = oDataTable.GetValue("Code", 0).ToString();
                                CargaGrid("grid3_1", oForm);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void CargaGrid(string ItemUID, Form oForm)
        {
            try
            {
                Grid oGrid;
                SAPbouiCOM.DataTable oLista;
                EditTextColumn etColumna;
                int Dimension;
                string DescDimension, fechaInicio, fechaFin;
                SAPbouiCOM.ComboBox cboDimension;

                oForm.Freeze(true);
                switch (ItemUID)
                {
                    case "grid1_1":
                        oCecos = new List<string>();
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1_1").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_1");
                        Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);

                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCCGasto;
                        Globals.Query = string.Format(Globals.Query, Dimension);
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
                        oGrid.Columns.Item("Select").TitleObject.Caption = "Seleccionar";
                        oGrid.Columns.Item("Select").Editable = true;
                        oGrid.Columns.Item("PrcCode").TitleObject.Caption = "Código";
                        oGrid.Columns.Item("PrcCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("PrcCode")));
                        etColumna.LinkedObjectType = "61";
                        oGrid.Columns.Item("PrcName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("PrcName").Editable = false;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid1_2":
                        oCuentas = new List<string>();
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_2");

                        string CuentaI = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_5").Specific).Value;
                        string CuentaF = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_6").Specific).Value;

                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCuentaGasto;
                        string filtro = string.Empty;
                        if (!string.IsNullOrEmpty(CuentaI) && !string.IsNullOrEmpty(CuentaF))
                            filtro = " AND \"FormatCode\" BETWEEN " + CuentaI + " AND " + CuentaF;

                        if (string.IsNullOrEmpty(Globals.Level)) Globals.ObtenerLevel();
                        Globals.Query = string.Format(Globals.Query, Globals.Level, filtro);
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
                        oGrid.Columns.Item("Select").TitleObject.Caption = "Seleccionar";
                        oGrid.Columns.Item("Select").Editable = true;
                        oGrid.Columns.Item("FormatCode").TitleObject.Caption = "N° Cuenta";
                        oGrid.Columns.Item("FormatCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("FormatCode")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("AcctCode").TitleObject.Caption = "Código";
                        oGrid.Columns.Item("AcctCode").Visible = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("AcctCode")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("AcctName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("AcctName").Editable = false;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid1_3":
                        oMonedas = new List<string>();
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1_3").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_3");

                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarMonedas;
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
                        oGrid.Columns.Item("Select").TitleObject.Caption = "Seleccionar";
                        oGrid.Columns.Item("Select").Editable = true;
                        oGrid.Columns.Item("CurrCode").TitleObject.Caption = "Código";
                        oGrid.Columns.Item("CurrCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("CurrCode")));
                        etColumna.LinkedObjectType = "37";
                        oGrid.Columns.Item("CurrName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("CurrName").Editable = false;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid2":
                        oListaIndices = new List<int>();
                        oListaxDistribuir = new List<JDT1>();
                        cboDimension = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific;
                        Dimension = Convert.ToInt32(cboDimension.Selected.Value);
                        DescDimension = Dimension == 1 ? "ProfitCode" : Dimension == 2 ? "OcrCode2" : Dimension == 3 ? "OcrCode3" : Dimension == 4 ? "OcrCode4" : "OcrCode5";
                        fechaInicio = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Value;
                        fechaFin = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Value;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");

                        string cuentas = "'" + string.Join("','", oCuentas) + "'";
                        string cecos = "'" + string.Join("','", oCecos) + "'";
                        string monedas = "'" + string.Join("','", oMonedas) + "'";
                        if (monedas.Contains(Globals.MonedaLocal))
                            monedas = monedas + ", ''";
                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarAsientos;
                        Globals.Query = string.Format(Globals.Query, DescDimension, fechaInicio, fechaFin, cuentas, cecos, monedas);
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
                        oGrid.Columns.Item("Select").TitleObject.Caption = "Seleccionar";
                        oGrid.Columns.Item("Select").Editable = true;
                        oGrid.Columns.Item("TransId").TitleObject.Caption = "Asiento";
                        oGrid.Columns.Item("TransId").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("TransId")));
                        etColumna.LinkedObjectType = "30";
                        oGrid.Columns.Item("Line_ID").TitleObject.Caption = "Línea";
                        oGrid.Columns.Item("Line_ID").Editable = false;
                        oGrid.Columns.Item("Account").TitleObject.Caption = "Cuenta";
                        oGrid.Columns.Item("Account").Visible = false;
                        oGrid.Columns.Item("FormatCode").TitleObject.Caption = "Cuenta";
                        oGrid.Columns.Item("FormatCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("FormatCode")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("AcctName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("AcctName").Editable = false;
                        oGrid.Columns.Item("AcctName").Visible = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Account")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("RefDate").TitleObject.Caption = "F.Contable";
                        oGrid.Columns.Item("RefDate").Editable = false;
                        if (cboDimension.ValidValues.Count > 0)
                        {
                            oGrid.Columns.Item("ProfitCode").TitleObject.Caption = cboDimension.ValidValues.Item(0).Description;
                            oGrid.Columns.Item("ProfitCode").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("ProfitCode")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("ProfitCodeName").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("ProfitCodeName").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("ProfitCode").Visible = false;
                            oGrid.Columns.Item("ProfitCodeName").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 1)
                        {
                            oGrid.Columns.Item("OcrCode2").TitleObject.Caption = cboDimension.ValidValues.Item(1).Description;
                            oGrid.Columns.Item("OcrCode2").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("OcrCode2")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("OcrCode2Name").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("OcrCode2Name").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("OcrCode2").Visible = false;
                            oGrid.Columns.Item("OcrCode2Name").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 2)
                        {
                            oGrid.Columns.Item("OcrCode3").TitleObject.Caption = cboDimension.ValidValues.Item(2).Description;
                            oGrid.Columns.Item("OcrCode3").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("OcrCode3")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("OcrCode3Name").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("OcrCode3Name").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("OcrCode3").Visible = false;
                            oGrid.Columns.Item("OcrCode3Name").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 3)
                        {
                            oGrid.Columns.Item("OcrCode4").TitleObject.Caption = cboDimension.ValidValues.Item(3).Description;
                            oGrid.Columns.Item("OcrCode4").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("OcrCode4")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("OcrCode4Name").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("OcrCode4Name").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("ProfitCode").Visible = false;
                            oGrid.Columns.Item("OcrCode4").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 4)
                        {
                            oGrid.Columns.Item("OcrCode5").TitleObject.Caption = cboDimension.ValidValues.Item(4).Description;
                            oGrid.Columns.Item("OcrCode5").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("OcrCode5")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("OcrCode5Name").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("OcrCode5Name").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("OcrCode5").Visible = false;
                            oGrid.Columns.Item("OcrCode5Name").Visible = false;
                        }
                        oGrid.Columns.Item("Total ML").Editable = false;
                        oGrid.Columns.Item("Total ML").RightJustified = true;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Total ML");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Total ME").Editable = false;
                        oGrid.Columns.Item("Total ME").RightJustified = true;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Total ME");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Total MS").Editable = false;
                        oGrid.Columns.Item("Total MS").RightJustified = true;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Total MS");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Project").TitleObject.Caption = "Proyecto";
                        oGrid.Columns.Item("Project").Editable = false;
                        oGrid.Columns.Item("Ref1").TitleObject.Caption = "Referencia 1";
                        oGrid.Columns.Item("Ref1").Editable = false;
                        oGrid.Columns.Item("Ref2").TitleObject.Caption = "Referencia 2";
                        oGrid.Columns.Item("Ref2").Editable = false;
                        oGrid.Columns.Item("Ref3Line").TitleObject.Caption = "Referencia 3";
                        oGrid.Columns.Item("Ref3Line").Editable = false;
                        oGrid.Columns.Item("FCCurrency").TitleObject.Caption = "Moneda";
                        oGrid.Columns.Item("FCCurrency").Editable = false;
                        oGrid.CommonSetting.FixedColumnsCount = 2;
                        for (int i = 0; i < oGrid.Columns.Count; i++)
                            oGrid.Columns.Item(i).TitleObject.Sortable = true;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid3_1":
                        oDistribucion = new List<OPRC>();
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3_1").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3_1");
                        Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);

                        if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Selected.Value == "1")
                        {
                            string DocEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Value;
                            ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Enabled = true;
                            Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCCProductivo_Opc1;
                            Globals.Query = string.Format(Globals.Query, Dimension, DocEntry);
                        }
                        else
                        {
                            ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Item.Enabled = false;
                            ((SAPbouiCOM.EditText)oForm.Items.Item("et3_5").Specific).Value = "";
                            Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCCProductivo_Opc2;
                            Globals.Query = string.Format(Globals.Query, Dimension);
                        }

                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("PrcCode").TitleObject.Caption = "Código";
                        oGrid.Columns.Item("PrcCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("PrcCode")));
                        etColumna.LinkedObjectType = "61";
                        oGrid.Columns.Item("PrcName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("PrcName").Editable = false;
                        oGrid.Columns.Item("Peso").TitleObject.Caption = "Peso";
                        oGrid.Columns.Item("Peso").RightJustified = true;
                        oGrid.Columns.Item("Peso").Editable = true;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Peso");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid3_2":
                        oProyectosD = new List<OPRJ>();
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3_2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3_2");
                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarDitribucionProyectos;
                        if (oProyectos.Count > 0)
                        {
                            string proyectos = "'" + string.Join("','", oProyectos) + "'";
                            Globals.Query = string.Format(Globals.Query, proyectos);
                            oGrid.DataTable = oLista;
                            oGrid.DataTable.Rows.Clear();
                            oLista.Rows.Clear();
                            oGrid.DataTable.ExecuteQuery(Globals.Query);
                            oGrid.Columns.Item("RowsHeader").Visible = false;
                            oGrid.Columns.Item("Code").TitleObject.Caption = "Proyecto Origen";
                            oGrid.Columns.Item("Code").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Code")));
                            etColumna.LinkedObjectType = "63";
                            oGrid.Columns.Item("U_EXX_ADRG_PRJD").TitleObject.Caption = "Proyecto Destino";
                            oGrid.Columns.Item("U_EXX_ADRG_PRJD").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("U_EXX_ADRG_PRJD")));
                            etColumna.LinkedObjectType = "63";
                            oGrid.Columns.Item("U_EXX_ADRG_PESO").TitleObject.Caption = "Peso";
                            oGrid.Columns.Item("U_EXX_ADRG_PESO").RightJustified = true;
                            oGrid.Columns.Item("U_EXX_ADRG_PESO").Editable = true;
                            etColumna = (EditTextColumn)oGrid.Columns.Item("U_EXX_ADRG_PESO");
                            etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                            oGrid.AutoResizeColumns();
                        }
                        else
                        {
                            oGrid.Item.Visible = false;
                        }
                        break;
                    case "grid4":
                        Globals.InformationMessage("Cargando sumilación de la distribución de asientos...");
                        cboDimension = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");
                        Dimension = Convert.ToInt32(cboDimension.Selected.Value);
                        DescDimension = Dimension == 1 ? "ProfitCode" : Dimension == 2 ? "OcrCode2" : Dimension == 3 ? "OcrCode3" : Dimension == 4 ? "OcrCode4" : "OcrCode5";

                        oAsiento = new OJDT();
                        oAsiento.Details = new List<JDT1>();
                        oListaxDistribuir = new List<JDT1>();

                        foreach (int index in oListaIndices)
                        {
                            JDT1 detalle = new JDT1();
                            detalle.TransId = oLista.GetValue("TransId", index).ToString();
                            detalle.Line_ID = oLista.GetValue("Line_ID", index).ToString();
                            detalle.AccountCode = oLista.GetValue("Account", index).ToString();
                            detalle.FormatCode = oLista.GetValue("FormatCode", index).ToString();
                            detalle.AccountName = oLista.GetValue("AcctName", index).ToString();
                            detalle.CostingCode = oLista.GetValue("ProfitCode", index).ToString();
                            detalle.CostingCodeName = oLista.GetValue("ProfitCodeName", index).ToString();
                            detalle.CostingCode2 = oLista.GetValue("OcrCode2", index).ToString();
                            detalle.CostingCode2Name = oLista.GetValue("OcrCode2Name", index).ToString();
                            detalle.CostingCode3 = oLista.GetValue("OcrCode3", index).ToString();
                            detalle.CostingCode3Name = oLista.GetValue("OcrCode3Name", index).ToString();
                            detalle.CostingCode4 = oLista.GetValue("OcrCode4", index).ToString();
                            detalle.CostingCode4Name = oLista.GetValue("OcrCode4Name", index).ToString();
                            detalle.CostingCode5 = oLista.GetValue("OcrCode5", index).ToString();
                            detalle.CostingCode5Name = oLista.GetValue("OcrCode5Name", index).ToString();
                            detalle.ProjectCode = oLista.GetValue("Project", index).ToString();
                            detalle.TotalML = Convert.ToDouble(oLista.GetValue("Total ML", index).ToString());
                            detalle.TotalME = Convert.ToDouble(oLista.GetValue("Total ME", index).ToString());
                            detalle.TotalMS = Convert.ToDouble(oLista.GetValue("Total MS", index).ToString());
                            detalle.FCCurrency = oLista.GetValue("FCCurrency", index).ToString();
                            oListaxDistribuir.Add(detalle);
                        }

                        int intcf = -1;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_4");
                        oLista.Rows.Clear();
                        int factor = oDistribucion.Sum(x => x.Peso);
                        var asientoAux = oListaxDistribuir.GroupBy(x => new
                        {
                            x.AccountCode,
                            x.FormatCode,
                            x.AccountName,
                            x.CostingCode,
                            x.CostingCodeName,
                            x.CostingCode2,
                            x.CostingCode2Name,
                            x.CostingCode3,
                            x.CostingCode3Name,
                            x.CostingCode4,
                            x.CostingCode4Name,
                            x.CostingCode5,
                            x.CostingCode5Name,
                            x.ProjectCode,
                            x.FCCurrency
                        }).Select(x => new JDT1
                        {
                            AccountCode = x.Key.AccountCode,
                            FormatCode = x.Key.FormatCode,
                            AccountName = x.Key.AccountName,
                            CostingCode = x.Key.CostingCode,
                            CostingCodeName = x.Key.CostingCodeName,
                            CostingCode2 = x.Key.CostingCode2,
                            CostingCode2Name = x.Key.CostingCode2Name,
                            CostingCode3 = x.Key.CostingCode3,
                            CostingCode3Name = x.Key.CostingCode3Name,
                            CostingCode4 = x.Key.CostingCode4,
                            CostingCode4Name = x.Key.CostingCode4Name,
                            CostingCode5 = x.Key.CostingCode5,
                            CostingCode5Name = x.Key.CostingCode5Name,
                            ProjectCode = x.Key.ProjectCode,
                            FCCurrency = x.Key.FCCurrency,
                            TotalML = Math.Round(x.Sum(y => Convert.ToDouble(y.TotalML)), 2, MidpointRounding.AwayFromZero),
                            TotalME = Math.Round(x.Sum(y => Convert.ToDouble(y.TotalME)), 2, MidpointRounding.AwayFromZero),
                            TotalMS = Math.Round(x.Sum(y => Convert.ToDouble(y.TotalMS)), 2, MidpointRounding.AwayFromZero)
                        }).Where(z => z.TotalML + z.TotalME + z.TotalMS != 0).ToList();

                        var lista = asientoAux.OrderBy(x => x.AccountCode).ThenBy(x => x.CostingCode).ThenBy(x => x.CostingCode2).ThenBy(x => x.CostingCode3).ThenBy(x => x.CostingCode4).ThenBy(x => x.CostingCode5).ThenBy(x => x.ProjectCode).ThenBy(x => x.FCCurrency).ToList();
                        foreach (JDT1 detalle in lista)
                        {
                            intcf++;
                            JDT1 detail = new JDT1();
                            double totalML = 0.0;
                            double totalME = 0.0;
                            double totalMS = 0.0;
                            double montoDistribucion = 0;
                            oLista.Rows.Add();
                            oLista.SetValue("Col_0", intcf, detalle.AccountCode);
                            oLista.SetValue("Col_1", intcf, detalle.FormatCode);
                            oLista.SetValue("Col_2", intcf, detalle.AccountName);
                            oLista.SetValue("Col_20", intcf, detalle.FCCurrency);

                            detail.AccountCode = detalle.AccountCode;

                            if (detalle.TotalML < 0)
                            {
                                montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalML)), 2, MidpointRounding.AwayFromZero);
                                totalML += Math.Abs(montoDistribucion);
                                oLista.SetValue("Col_3", intcf, montoDistribucion);
                                detail.Debit = montoDistribucion;
                                detail.Credit = 0;
                            }
                            else
                            {
                                montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalML)), 2, MidpointRounding.AwayFromZero);
                                totalML += Math.Abs(montoDistribucion);
                                oLista.SetValue("Col_4", intcf, montoDistribucion);
                                detail.Debit = 0;
                                detail.Credit = montoDistribucion;
                            }

                            if (detalle.FCCurrency != "" && detalle.TotalME != 0)
                            {
                                if (detalle.TotalME < 0)
                                {
                                    montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalME)), 2, MidpointRounding.AwayFromZero);
                                    totalME += montoDistribucion;
                                    oLista.SetValue("Col_18", intcf, montoDistribucion);
                                    detail.FCCurrency = detalle.FCCurrency;
                                    detail.FCDebit = montoDistribucion;
                                    detail.FCCredit = 0;
                                }
                                else
                                {
                                    montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalME)), 2, MidpointRounding.AwayFromZero);
                                    totalME += montoDistribucion;
                                    oLista.SetValue("Col_19", intcf, montoDistribucion);
                                    detail.FCCurrency = detalle.FCCurrency;
                                    detail.FCDebit = 0;
                                    detail.FCCredit = montoDistribucion;
                                }
                            }
                            else
                            {
                                detail.FCCurrency = "USD";
                                detail.FCDebit = 0;
                                detail.FCCredit = 0;
                            }

                            if (detalle.TotalMS < 0)
                            {
                                montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalMS)), 2, MidpointRounding.AwayFromZero);
                                totalMS += montoDistribucion;
                                oLista.SetValue("Col_5", intcf, montoDistribucion);
                                detail.DebitSys = montoDistribucion;
                                detail.CreditSys = 0;
                            }
                            else
                            {
                                montoDistribucion = Math.Round(Math.Abs(Convert.ToDouble(detalle.TotalMS)), 2, MidpointRounding.AwayFromZero);
                                totalMS += montoDistribucion;
                                oLista.SetValue("Col_6", intcf, montoDistribucion);
                                detail.DebitSys = 0;
                                detail.CreditSys = montoDistribucion;
                            }

                            if (cboDimension.ValidValues.Count > 0)
                            {
                                oLista.SetValue("Col_7", intcf, detalle.CostingCode);
                                oLista.SetValue("Col_8", intcf, detalle.CostingCodeName);
                                detail.CostingCode = detalle.CostingCode;
                            }
                            if (cboDimension.ValidValues.Count > 1)
                            {
                                oLista.SetValue("Col_9", intcf, detalle.CostingCode2);
                                oLista.SetValue("Col_10", intcf, detalle.CostingCode2Name);
                                detail.CostingCode2 = detalle.CostingCode2;
                            }
                            if (cboDimension.ValidValues.Count > 2)
                            {
                                oLista.SetValue("Col_11", intcf, detalle.CostingCode3);
                                oLista.SetValue("Col_12", intcf, detalle.CostingCode3Name);
                                detail.CostingCode3 = detalle.CostingCode3;
                            }
                            if (cboDimension.ValidValues.Count > 3)
                            {
                                oLista.SetValue("Col_13", intcf, detalle.CostingCode4);
                                oLista.SetValue("Col_14", intcf, detalle.CostingCode4Name);
                                detail.CostingCode4 = detalle.CostingCode4;
                            }
                            if (cboDimension.ValidValues.Count > 4)
                            {
                                oLista.SetValue("Col_15", intcf, detalle.CostingCode5);
                                oLista.SetValue("Col_16", intcf, detalle.CostingCode5Name);
                                detail.CostingCode5 = detalle.CostingCode5;
                            }
                            if (!string.IsNullOrEmpty(detalle.ProjectCode))
                            {
                                oLista.SetValue("Col_17", intcf, detalle.ProjectCode);
                                detail.ProjectCode = detalle.ProjectCode;
                            }
                            oAsiento.Details.Add(detail);

                            for (int i = 0; i < oDistribucion.Count; i++)
                            {
                                if (string.IsNullOrEmpty(detalle.ProjectCode))
                                {
                                    intcf++;
                                    detail = new JDT1();
                                    oLista.Rows.Add();
                                    oLista.SetValue("Col_0", intcf, detalle.AccountCode);
                                    oLista.SetValue("Col_1", intcf, detalle.FormatCode);
                                    oLista.SetValue("Col_2", intcf, detalle.AccountName);
                                    oLista.SetValue("Col_20", intcf, detalle.FCCurrency);
                                    detail.AccountCode = detalle.AccountCode;

                                    if (detalle.TotalML < 0)
                                    {
                                        montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalML)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        totalML -= montoDistribucion;
                                        oLista.SetValue("Col_4", intcf, montoDistribucion);
                                        detail.Debit = 0;
                                        detail.Credit = montoDistribucion;
                                    }
                                    else
                                    {
                                        montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalML)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        totalML -= montoDistribucion;
                                        oLista.SetValue("Col_3", intcf, montoDistribucion);
                                        detail.Debit = montoDistribucion;
                                        detail.Credit = 0;
                                    }

                                    if (detalle.FCCurrency != "" && detalle.TotalME != 0)
                                    {
                                        if (detalle.TotalME < 0)
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalME)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalME -= montoDistribucion;
                                            oLista.SetValue("Col_19", intcf, montoDistribucion);
                                            detail.FCCurrency = detalle.FCCurrency;
                                            detail.FCDebit = 0;
                                            detail.FCCredit = montoDistribucion;
                                        }
                                        else
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalME)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalME -= montoDistribucion;
                                            oLista.SetValue("Col_18", intcf, montoDistribucion);
                                            detail.FCCurrency = detalle.FCCurrency;
                                            detail.FCDebit = montoDistribucion;
                                            detail.FCCredit = 0;
                                        }
                                    }
                                    else
                                    {
                                        detail.FCCurrency = "USD";
                                        detail.FCDebit = 0;
                                        detail.FCCredit = 0;
                                    }

                                    if (detalle.TotalMS < 0)
                                    {
                                        montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalMS)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        totalMS -= montoDistribucion;
                                        oLista.SetValue("Col_6", intcf, montoDistribucion);
                                        detail.DebitSys = 0;
                                        detail.CreditSys = montoDistribucion;
                                    }
                                    else
                                    {
                                        montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalMS)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        totalMS -= montoDistribucion;
                                        oLista.SetValue("Col_5", intcf, montoDistribucion);
                                        detail.DebitSys = montoDistribucion;
                                        detail.CreditSys = 0;
                                    }

                                    if (cboDimension.ValidValues.Count > 0)
                                    {
                                        if (Dimension == 1)
                                        {
                                            oLista.SetValue("Col_7", intcf, oDistribucion[i].PrcCode);
                                            oLista.SetValue("Col_8", intcf, oDistribucion[i].PrcName);
                                            detail.CostingCode = oDistribucion[i].PrcCode;
                                        }
                                        else
                                        {
                                            oLista.SetValue("Col_7", intcf, detalle.CostingCode);
                                            oLista.SetValue("Col_8", intcf, detalle.CostingCodeName);
                                            detail.CostingCode = detalle.CostingCode;
                                        }
                                    }
                                    if (cboDimension.ValidValues.Count > 1)
                                    {
                                        if (Dimension == 2)
                                        {
                                            oLista.SetValue("Col_9", intcf, oDistribucion[i].PrcCode);
                                            oLista.SetValue("Col_10", intcf, oDistribucion[i].PrcName);
                                            detail.CostingCode2 = oDistribucion[i].PrcCode;
                                        }
                                        else
                                        {
                                            oLista.SetValue("Col_9", intcf, detalle.CostingCode2);
                                            oLista.SetValue("Col_10", intcf, detalle.CostingCode2Name);
                                            detail.CostingCode2 = detalle.CostingCode2;
                                        }
                                    }
                                    if (cboDimension.ValidValues.Count > 2)
                                    {
                                        if (Dimension == 3)
                                        {
                                            oLista.SetValue("Col_11", intcf, oDistribucion[i].PrcCode);
                                            oLista.SetValue("Col_12", intcf, oDistribucion[i].PrcName);
                                            detail.CostingCode3 = oDistribucion[i].PrcCode;
                                        }
                                        else
                                        {
                                            oLista.SetValue("Col_11", intcf, detalle.CostingCode3);
                                            oLista.SetValue("Col_12", intcf, detalle.CostingCode3Name);
                                            detail.CostingCode3 = detalle.CostingCode3;
                                        }
                                    }
                                    if (cboDimension.ValidValues.Count > 3)
                                    {
                                        if (Dimension == 4)
                                        {
                                            oLista.SetValue("Col_13", intcf, oDistribucion[i].PrcCode);
                                            oLista.SetValue("Col_14", intcf, oDistribucion[i].PrcName);
                                            detail.CostingCode4 = oDistribucion[i].PrcCode;
                                        }
                                        else
                                        {
                                            oLista.SetValue("Col_13", intcf, detalle.CostingCode4);
                                            oLista.SetValue("Col_14", intcf, detalle.CostingCode4Name);
                                            detail.CostingCode4 = detalle.CostingCode4;
                                        }
                                    }
                                    if (cboDimension.ValidValues.Count > 4)
                                    {
                                        if (Dimension == 5)
                                        {
                                            oLista.SetValue("Col_15", intcf, oDistribucion[i].PrcCode);
                                            oLista.SetValue("Col_16", intcf, oDistribucion[i].PrcName);
                                            detail.CostingCode5 = oDistribucion[i].PrcCode;
                                        }
                                        else
                                        {
                                            oLista.SetValue("Col_15", intcf, detalle.CostingCode5);
                                            oLista.SetValue("Col_16", intcf, detalle.CostingCode5Name);
                                            detail.CostingCode5 = detalle.CostingCode5;
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(detalle.ProjectCode))
                                    {
                                        oLista.SetValue("Col_17", intcf, detalle.ProjectCode);
                                        detail.ProjectCode = detalle.ProjectCode;
                                    }
                                    oAsiento.Details.Add(detail);
                                }
                                else
                                {
                                    var proyectoDist = oProyectosD.Where(x => x.PrjCode == detalle.ProjectCode).Select(y => y).ToList();
                                    double factorProyecto = proyectoDist.Sum(x => x.Peso);

                                    if (proyectoDist.Count == 0 || factorProyecto == 0)
                                    {
                                        intcf++;
                                        detail = new JDT1();
                                        oLista.Rows.Add();
                                        oLista.SetValue("Col_0", intcf, detalle.AccountCode);
                                        oLista.SetValue("Col_1", intcf, detalle.FormatCode);
                                        oLista.SetValue("Col_2", intcf, detalle.AccountName);
                                        oLista.SetValue("Col_20", intcf, detalle.FCCurrency);
                                        detail.AccountCode = detalle.AccountCode;

                                        if (detalle.TotalML < 0)
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalML)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalML -= montoDistribucion;
                                            oLista.SetValue("Col_4", intcf, montoDistribucion);
                                            detail.Debit = 0;
                                            detail.Credit = montoDistribucion;
                                        }
                                        else
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalML)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalML -= montoDistribucion;
                                            oLista.SetValue("Col_3", intcf, montoDistribucion);
                                            detail.Debit = montoDistribucion;
                                            detail.Credit = 0;
                                        }

                                        if (detalle.FCCurrency != "" && detalle.TotalME != 0)
                                        {
                                            if (detalle.TotalME < 0)
                                            {
                                                montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalME)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                                totalME -= montoDistribucion;
                                                oLista.SetValue("Col_19", intcf, montoDistribucion);
                                                detail.FCCurrency = detalle.FCCurrency;
                                                detail.FCDebit = 0;
                                                detail.FCCredit = montoDistribucion;
                                            }
                                            else
                                            {
                                                montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalME)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                                totalME -= montoDistribucion;
                                                oLista.SetValue("Col_18", intcf, montoDistribucion);
                                                detail.FCCurrency = detalle.FCCurrency;
                                                detail.FCDebit = montoDistribucion;
                                                detail.FCCredit = 0;
                                            }
                                        }
                                        else
                                        {
                                            detail.FCCurrency = "USD";
                                            detail.FCDebit = 0;
                                            detail.FCCredit = 0;
                                        }

                                        if (detalle.TotalMS < 0)
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalMS)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalMS -= montoDistribucion;
                                            oLista.SetValue("Col_6", intcf, montoDistribucion);
                                            detail.DebitSys = 0;
                                            detail.CreditSys = montoDistribucion;
                                        }
                                        else
                                        {
                                            montoDistribucion = Math.Round((Math.Abs(Convert.ToDouble(detalle.TotalMS)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                            totalMS -= montoDistribucion;
                                            oLista.SetValue("Col_5", intcf, montoDistribucion);
                                            detail.DebitSys = montoDistribucion;
                                            detail.CreditSys = 0;
                                        }

                                        if (cboDimension.ValidValues.Count > 0)
                                        {
                                            if (Dimension == 1)
                                            {
                                                oLista.SetValue("Col_7", intcf, oDistribucion[i].PrcCode);
                                                oLista.SetValue("Col_8", intcf, oDistribucion[i].PrcName);
                                                detail.CostingCode = oDistribucion[i].PrcCode;
                                            }
                                            else
                                            {
                                                oLista.SetValue("Col_7", intcf, detalle.CostingCode);
                                                oLista.SetValue("Col_8", intcf, detalle.CostingCodeName);
                                                detail.CostingCode = detalle.CostingCode;
                                            }
                                        }
                                        if (cboDimension.ValidValues.Count > 1)
                                        {
                                            if (Dimension == 2)
                                            {
                                                oLista.SetValue("Col_9", intcf, oDistribucion[i].PrcCode);
                                                oLista.SetValue("Col_10", intcf, oDistribucion[i].PrcName);
                                                detail.CostingCode2 = oDistribucion[i].PrcCode;
                                            }
                                            else
                                            {
                                                oLista.SetValue("Col_9", intcf, detalle.CostingCode2);
                                                oLista.SetValue("Col_10", intcf, detalle.CostingCode2Name);
                                                detail.CostingCode2 = detalle.CostingCode2;
                                            }
                                        }
                                        if (cboDimension.ValidValues.Count > 2)
                                        {
                                            if (Dimension == 3)
                                            {
                                                oLista.SetValue("Col_11", intcf, oDistribucion[i].PrcCode);
                                                oLista.SetValue("Col_12", intcf, oDistribucion[i].PrcName);
                                                detail.CostingCode3 = oDistribucion[i].PrcCode;
                                            }
                                            else
                                            {
                                                oLista.SetValue("Col_11", intcf, detalle.CostingCode3);
                                                oLista.SetValue("Col_12", intcf, detalle.CostingCode3Name);
                                                detail.CostingCode3 = detalle.CostingCode3;
                                            }
                                        }
                                        if (cboDimension.ValidValues.Count > 3)
                                        {
                                            if (Dimension == 4)
                                            {
                                                oLista.SetValue("Col_13", intcf, oDistribucion[i].PrcCode);
                                                oLista.SetValue("Col_14", intcf, oDistribucion[i].PrcName);
                                                detail.CostingCode4 = oDistribucion[i].PrcCode;
                                            }
                                            else
                                            {
                                                oLista.SetValue("Col_13", intcf, detalle.CostingCode4);
                                                oLista.SetValue("Col_14", intcf, detalle.CostingCode4Name);
                                                detail.CostingCode4 = detalle.CostingCode4;
                                            }
                                        }
                                        if (cboDimension.ValidValues.Count > 4)
                                        {
                                            if (Dimension == 5)
                                            {
                                                oLista.SetValue("Col_15", intcf, oDistribucion[i].PrcCode);
                                                oLista.SetValue("Col_16", intcf, oDistribucion[i].PrcName);
                                                detail.CostingCode5 = oDistribucion[i].PrcCode;
                                            }
                                            else
                                            {
                                                oLista.SetValue("Col_15", intcf, detalle.CostingCode5);
                                                oLista.SetValue("Col_16", intcf, detalle.CostingCode5Name);
                                                detail.CostingCode5 = detalle.CostingCode5;
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(detalle.ProjectCode))
                                        {
                                            oLista.SetValue("Col_17", intcf, detalle.ProjectCode);
                                            detail.ProjectCode = detalle.ProjectCode;
                                        }
                                        oAsiento.Details.Add(detail);
                                    }
                                    else
                                    {
                                        double totalDistribucionML = Math.Round(Math.Abs((Convert.ToDouble(detalle.TotalML)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        double totalDistribucionME = Math.Round(Math.Abs((Convert.ToDouble(detalle.TotalME)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        double totalDistribucionMS = Math.Round(Math.Abs((Convert.ToDouble(detalle.TotalMS)) * oDistribucion[i].Peso) / factor, 2, MidpointRounding.AwayFromZero);
                                        double montoDistribucionML = totalDistribucionML;
                                        double montoDistribucionME = totalDistribucionME;
                                        double montoDistribucionMS = totalDistribucionMS;

                                        for (int j = 0; j < proyectoDist.Count; j++)
                                        {
                                            if (proyectoDist[j].Peso > 0)
                                            {
                                                intcf++;
                                                detail = new JDT1();
                                                oLista.Rows.Add();
                                                oLista.SetValue("Col_0", intcf, detalle.AccountCode);
                                                oLista.SetValue("Col_1", intcf, detalle.FormatCode);
                                                oLista.SetValue("Col_2", intcf, detalle.AccountName);
                                                oLista.SetValue("Col_20", intcf, detalle.FCCurrency);
                                                detail.AccountCode = detalle.AccountCode;

                                                if (detalle.TotalML < 0)
                                                {
                                                    montoDistribucion = Math.Round((totalDistribucionML * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                    totalML -= montoDistribucion;
                                                    montoDistribucionML -= montoDistribucion;
                                                    oLista.SetValue("Col_4", intcf, montoDistribucion);
                                                    detail.Debit = 0;
                                                    detail.Credit = montoDistribucion;
                                                }
                                                else
                                                {
                                                    montoDistribucion = Math.Round((totalDistribucionML * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                    totalML -= montoDistribucion;
                                                    montoDistribucionML -= montoDistribucion;
                                                    oLista.SetValue("Col_3", intcf, montoDistribucion);
                                                    detail.Debit = montoDistribucion;
                                                    detail.Credit = 0;
                                                }

                                                if (detalle.FCCurrency != "" && detalle.TotalME != 0)
                                                {
                                                    if (detalle.TotalME < 0)
                                                    {
                                                        montoDistribucion = Math.Round((totalDistribucionME * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                        totalME -= montoDistribucion;
                                                        montoDistribucionME -= montoDistribucion;
                                                        oLista.SetValue("Col_19", intcf, montoDistribucion);
                                                        detail.FCCurrency = detalle.FCCurrency;
                                                        detail.FCDebit = 0;
                                                        detail.FCCredit = montoDistribucion;
                                                    }
                                                    else
                                                    {
                                                        montoDistribucion = Math.Round((totalDistribucionME * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                        totalME -= montoDistribucion;
                                                        montoDistribucionME -= montoDistribucion;
                                                        oLista.SetValue("Col_18", intcf, montoDistribucion);
                                                        detail.FCCurrency = detalle.FCCurrency;
                                                        detail.FCDebit = montoDistribucion;
                                                        detail.FCCredit = 0;
                                                    }
                                                }
                                                else
                                                {
                                                    detail.FCCurrency = "USD";
                                                    detail.FCDebit = 0;
                                                    detail.FCCredit = 0;
                                                }

                                                if (detalle.TotalMS < 0)
                                                {
                                                    montoDistribucion = Math.Round((totalDistribucionMS * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                    totalMS -= montoDistribucion;
                                                    montoDistribucionMS -= montoDistribucion;
                                                    oLista.SetValue("Col_6", intcf, montoDistribucion);
                                                    detail.DebitSys = 0;
                                                    detail.CreditSys = montoDistribucion;
                                                }
                                                else
                                                {
                                                    montoDistribucion = Math.Round((totalDistribucionMS * proyectoDist[j].Peso) / factorProyecto, 2, MidpointRounding.AwayFromZero);
                                                    totalMS -= montoDistribucion;
                                                    montoDistribucionMS -= montoDistribucion;
                                                    oLista.SetValue("Col_5", intcf, montoDistribucion);
                                                    detail.DebitSys = montoDistribucion;
                                                    detail.CreditSys = 0;
                                                }

                                                if (cboDimension.ValidValues.Count > 0)
                                                {
                                                    if (Dimension == 1)
                                                    {
                                                        oLista.SetValue("Col_7", intcf, oDistribucion[i].PrcCode);
                                                        oLista.SetValue("Col_8", intcf, oDistribucion[i].PrcName);
                                                        detail.CostingCode = oDistribucion[i].PrcCode;
                                                    }
                                                    else
                                                    {
                                                        oLista.SetValue("Col_7", intcf, detalle.CostingCode);
                                                        oLista.SetValue("Col_8", intcf, detalle.CostingCodeName);
                                                        detail.CostingCode = detalle.CostingCode;
                                                    }
                                                }
                                                if (cboDimension.ValidValues.Count > 1)
                                                {
                                                    if (Dimension == 2)
                                                    {
                                                        oLista.SetValue("Col_9", intcf, oDistribucion[i].PrcCode);
                                                        oLista.SetValue("Col_10", intcf, oDistribucion[i].PrcName);
                                                        detail.CostingCode2 = oDistribucion[i].PrcCode;
                                                    }
                                                    else
                                                    {
                                                        oLista.SetValue("Col_9", intcf, detalle.CostingCode2);
                                                        oLista.SetValue("Col_10", intcf, detalle.CostingCode2Name);
                                                        detail.CostingCode2 = detalle.CostingCode2;
                                                    }
                                                }
                                                if (cboDimension.ValidValues.Count > 2)
                                                {
                                                    if (Dimension == 3)
                                                    {
                                                        oLista.SetValue("Col_11", intcf, oDistribucion[i].PrcCode);
                                                        oLista.SetValue("Col_12", intcf, oDistribucion[i].PrcName);
                                                        detail.CostingCode3 = oDistribucion[i].PrcCode;
                                                    }
                                                    else
                                                    {
                                                        oLista.SetValue("Col_11", intcf, detalle.CostingCode3);
                                                        oLista.SetValue("Col_12", intcf, detalle.CostingCode3Name);
                                                        detail.CostingCode3 = detalle.CostingCode3;
                                                    }
                                                }
                                                if (cboDimension.ValidValues.Count > 3)
                                                {
                                                    if (Dimension == 4)
                                                    {
                                                        oLista.SetValue("Col_13", intcf, oDistribucion[i].PrcCode);
                                                        oLista.SetValue("Col_14", intcf, oDistribucion[i].PrcName);
                                                        detail.CostingCode4 = oDistribucion[i].PrcCode;
                                                    }
                                                    else
                                                    {
                                                        oLista.SetValue("Col_13", intcf, detalle.CostingCode4);
                                                        oLista.SetValue("Col_14", intcf, detalle.CostingCode4Name);
                                                        detail.CostingCode4 = detalle.CostingCode4;
                                                    }
                                                }
                                                if (cboDimension.ValidValues.Count > 4)
                                                {
                                                    if (Dimension == 5)
                                                    {
                                                        oLista.SetValue("Col_15", intcf, oDistribucion[i].PrcCode);
                                                        oLista.SetValue("Col_16", intcf, oDistribucion[i].PrcName);
                                                        detail.CostingCode5 = oDistribucion[i].PrcCode;
                                                    }
                                                    else
                                                    {
                                                        oLista.SetValue("Col_15", intcf, detalle.CostingCode5);
                                                        oLista.SetValue("Col_16", intcf, detalle.CostingCode5Name);
                                                        detail.CostingCode5 = detalle.CostingCode5;
                                                    }
                                                }

                                                if (!string.IsNullOrEmpty(detalle.ProjectCode))
                                                {
                                                    oLista.SetValue("Col_17", intcf, proyectoDist[j].PrjDestino);
                                                    detail.ProjectCode = proyectoDist[j].PrjDestino;
                                                }
                                                oAsiento.Details.Add(detail);
                                            }
                                        }

                                        montoDistribucionML = Math.Round(montoDistribucionML, 2, MidpointRounding.AwayFromZero);
                                        montoDistribucionME = Math.Round(montoDistribucionME, 2, MidpointRounding.AwayFromZero);
                                        montoDistribucionMS = Math.Round(montoDistribucionMS, 2, MidpointRounding.AwayFromZero);
                                        if (montoDistribucionML + montoDistribucionME + montoDistribucionMS != 0)
                                        {
                                            double monto = Convert.ToDouble(oLista.GetValue("Col_4", intcf).ToString());
                                            if (monto > 0)
                                            {
                                                totalML -= montoDistribucionML;
                                                oLista.SetValue("Col_4", intcf, monto + montoDistribucionML);
                                                oAsiento.Details[intcf].Credit = monto + montoDistribucionML;
                                            }
                                            else
                                            {
                                                monto = Convert.ToDouble(oLista.GetValue("Col_3", intcf).ToString());
                                                totalML -= montoDistribucionML;
                                                oLista.SetValue("Col_3", intcf, monto + montoDistribucionML);
                                                oAsiento.Details[intcf].Debit = monto + montoDistribucionML;
                                            }

                                            monto = Convert.ToDouble(oLista.GetValue("Col_19", intcf).ToString());
                                            if (monto > 0)
                                            {
                                                totalME -= montoDistribucionME;
                                                oLista.SetValue("Col_19", intcf, monto + montoDistribucionME);
                                                oAsiento.Details[intcf].Credit = monto + montoDistribucionME;
                                            }
                                            else
                                            {
                                                monto = Convert.ToDouble(oLista.GetValue("Col_18", intcf).ToString());
                                                totalME -= montoDistribucionME;
                                                oLista.SetValue("Col_18", intcf, monto + montoDistribucionME);
                                                oAsiento.Details[intcf].Debit = monto + montoDistribucionME;
                                            }

                                            monto = Convert.ToDouble(oLista.GetValue("Col_6", intcf).ToString());
                                            if (monto > 0)
                                            {
                                                totalMS -= montoDistribucionMS;
                                                oLista.SetValue("Col_6", intcf, monto + montoDistribucionMS);
                                                oAsiento.Details[intcf].CreditSys = monto + montoDistribucionMS;
                                            }
                                            else
                                            {
                                                monto = Convert.ToDouble(oLista.GetValue("Col_5", intcf).ToString());
                                                totalMS -= montoDistribucionMS;
                                                oLista.SetValue("Col_5", intcf, monto + montoDistribucionMS);
                                                oAsiento.Details[intcf].DebitSys = monto + montoDistribucionMS;
                                            }
                                        }
                                    }
                                }
                            }

                            totalML = Math.Round(totalML, 2, MidpointRounding.AwayFromZero);
                            totalME = Math.Round(totalME, 2, MidpointRounding.AwayFromZero);
                            totalMS = Math.Round(totalMS, 2, MidpointRounding.AwayFromZero);
                            if (totalML + totalME + totalMS != 0)
                            {
                                double monto = Convert.ToDouble(oLista.GetValue("Col_4", intcf).ToString());
                                if (monto > 0)
                                {
                                    oLista.SetValue("Col_4", intcf, monto + totalML);
                                    oAsiento.Details[intcf].Credit = monto + totalML;
                                }
                                else
                                {
                                    monto = Convert.ToDouble(oLista.GetValue("Col_3", intcf).ToString());
                                    oLista.SetValue("Col_3", intcf, monto + totalML);
                                    oAsiento.Details[intcf].Debit = monto + totalML;
                                }

                                monto = Convert.ToDouble(oLista.GetValue("Col_19", intcf).ToString());
                                if (monto > 0)
                                {
                                    oLista.SetValue("Col_19", intcf, monto + totalME);
                                    oAsiento.Details[intcf].FCCredit = monto + totalME;
                                }
                                else
                                {
                                    monto = Convert.ToDouble(oLista.GetValue("Col_18", intcf).ToString());
                                    oLista.SetValue("Col_18", intcf, monto + totalME);
                                    oAsiento.Details[intcf].FCDebit = monto + totalME;
                                }

                                monto = Convert.ToDouble(oLista.GetValue("Col_6", intcf).ToString());
                                if (monto > 0)
                                {
                                    oLista.SetValue("Col_6", intcf, monto + totalMS);
                                    oAsiento.Details[intcf].CreditSys = monto + totalMS;
                                }
                                else
                                {
                                    monto = Convert.ToDouble(oLista.GetValue("Col_5", intcf).ToString());
                                    oLista.SetValue("Col_5", intcf, monto + totalMS);
                                    oAsiento.Details[intcf].DebitSys = monto + totalMS;
                                }
                            }
                        }

                        oGrid.DataTable = oLista;

                        oGrid.Columns.Item("RowsHeader").Visible = false;
                        oGrid.Columns.Item("Col_0").TitleObject.Caption = "Cuenta";
                        oGrid.Columns.Item("Col_0").Visible = false;
                        oGrid.Columns.Item("Col_1").TitleObject.Caption = "Cod. Cuenta";
                        oGrid.Columns.Item("Col_1").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_1")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("Col_2").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("Col_2").Editable = false;
                        oGrid.Columns.Item("Col_20").TitleObject.Caption = "Moneda";
                        oGrid.Columns.Item("Col_20").Editable = false;
                        oGrid.Columns.Item("Col_3").TitleObject.Caption = "Debe ML";
                        oGrid.Columns.Item("Col_3").RightJustified = true;
                        oGrid.Columns.Item("Col_3").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_3");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_4").TitleObject.Caption = "Haber ML";
                        oGrid.Columns.Item("Col_4").RightJustified = true;
                        oGrid.Columns.Item("Col_4").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_4");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                        oGrid.Columns.Item("Col_18").TitleObject.Caption = "Debe ME";
                        oGrid.Columns.Item("Col_18").RightJustified = true;
                        oGrid.Columns.Item("Col_18").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_18");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_19").TitleObject.Caption = "Haber ME";
                        oGrid.Columns.Item("Col_19").RightJustified = true;
                        oGrid.Columns.Item("Col_19").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_19");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                        oGrid.Columns.Item("Col_5").TitleObject.Caption = "Debe MS";
                        oGrid.Columns.Item("Col_5").RightJustified = true;
                        oGrid.Columns.Item("Col_5").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_5");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_6").TitleObject.Caption = "Haber MS";
                        oGrid.Columns.Item("Col_6").RightJustified = true;
                        oGrid.Columns.Item("Col_6").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_6");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                        if (cboDimension.ValidValues.Count > 0)
                        {
                            oGrid.Columns.Item("Col_7").TitleObject.Caption = cboDimension.ValidValues.Item(0).Description;
                            oGrid.Columns.Item("Col_7").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_7")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("Col_8").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("Col_8").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("Col_7").Visible = false;
                            oGrid.Columns.Item("Col_8").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 1)
                        {
                            oGrid.Columns.Item("Col_9").TitleObject.Caption = cboDimension.ValidValues.Item(1).Description;
                            oGrid.Columns.Item("Col_9").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_9")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("Col_10").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("Col_10").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("Col_9").Visible = false;
                            oGrid.Columns.Item("Col_10").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 2)
                        {
                            oGrid.Columns.Item("Col_11").TitleObject.Caption = cboDimension.ValidValues.Item(2).Description;
                            oGrid.Columns.Item("Col_11").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_11")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("Col_12").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("Col_12").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("Col_11").Visible = false;
                            oGrid.Columns.Item("Col_12").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 3)
                        {
                            oGrid.Columns.Item("Col_13").TitleObject.Caption = cboDimension.ValidValues.Item(3).Description;
                            oGrid.Columns.Item("Col_13").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_13")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("Col_14").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("Col_14").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("Col_13").Visible = false;
                            oGrid.Columns.Item("Col_14").Visible = false;
                        }
                        if (cboDimension.ValidValues.Count > 4)
                        {
                            oGrid.Columns.Item("Col_15").TitleObject.Caption = cboDimension.ValidValues.Item(4).Description;
                            oGrid.Columns.Item("Col_15").Editable = false;
                            etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_15")));
                            etColumna.LinkedObjectType = "61";
                            oGrid.Columns.Item("Col_16").TitleObject.Caption = "Descripción";
                            oGrid.Columns.Item("Col_16").Editable = false;
                        }
                        else
                        {
                            oGrid.Columns.Item("Col_15").Visible = false;
                            oGrid.Columns.Item("Col_16").Visible = false;
                        }

                        oGrid.Columns.Item("Col_17").TitleObject.Caption = "Proyecto";
                        oGrid.Columns.Item("Col_17").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_17")));
                        etColumna.LinkedObjectType = "63";
                        oGrid.CommonSetting.FixedColumnsCount = 2;
                        oGrid.AutoResizeColumns();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void CargaLista(string ItemUID, Form oForm)
        {
            try
            {
                try
                {
                    SAPbouiCOM.DataTable oLista;
                    switch (ItemUID)
                    {
                        case "grid1_1":
                            oCecos = new List<string>();
                            oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_1");
                            for (int i = 0; i < oLista.Rows.Count; i++)
                            {
                                if (oLista.GetValue("Select", i).ToString() == "Y")
                                    oCecos.Add(oLista.GetValue("PrcCode", i).ToString());
                            }
                            break;
                        case "grid1_2":
                            oCuentas = new List<string>();
                            oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_2");
                            for (int i = 0; i < oLista.Rows.Count; i++)
                            {
                                if (oLista.GetValue("Select", i).ToString() == "Y")
                                    oCuentas.Add(oLista.GetValue("AcctCode", i).ToString());
                            }
                            break;
                        case "grid1_3":
                            oMonedas = new List<string>();
                            oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_3");
                            for (int i = 0; i < oLista.Rows.Count; i++)
                            {
                                if (oLista.GetValue("Select", i).ToString() == "Y")
                                    oMonedas.Add(oLista.GetValue("CurrCode", i).ToString());
                            }
                            break;
                        case "grid2":
                            oProyectos = new List<string>();
                            oListaIndices = new List<int>();
                            oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");
                            for (int i = 0; i < oLista.Rows.Count; i++)
                            {
                                if (oLista.GetValue("Select", i).ToString() == "Y")
                                {
                                    oProyectos.Add(oLista.GetValue("Project", i).ToString());
                                    oListaIndices.Add(i);
                                }
                            }
                            oProyectos = oProyectos.Where(x => !string.IsNullOrEmpty(x)).GroupBy(y => y).Select(z => z.Key).ToList();

                            break;
                    }
                }
                catch (Exception ex)
                {
                    oForm.Freeze(false);
                    throw ex;
                }
                finally
                {
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}