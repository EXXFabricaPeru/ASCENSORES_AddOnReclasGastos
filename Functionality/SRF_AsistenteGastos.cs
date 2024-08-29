using AddOnRclsGastos.App;
using AddOnRclsGastos.Entity;
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
        public static bool open = false;
        public static List<string> oCecos = new List<string>();
        public static List<string> oCuentas = new List<string>();
        public static List<int> oAsientos = new List<int>();
        public static List<OJDT> oAsientoDist = new List<OJDT>();
        public static List<OPRC> oDistribucion = new List<OPRC>();
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
                open = true;
                SAPbouiCOM.FormCreationParams fcp = default(SAPbouiCOM.FormCreationParams);
                fcp = (SAPbouiCOM.FormCreationParams)Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = Frm;
                fcp.UniqueID = Frm;
                string FormName = "\\Views\\frmAsistenteGastos.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);
                CargaGrid("grid1_2", oForm);
                oForm.Visible = true;
                open = false;
            }
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

        private static void Ejecutar(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            Button btnEjecutar = (Button)oForm.Items.Item("btnEjec").Specific;
            Button btnAnterior = (Button)oForm.Items.Item("btnAnt").Specific;
            Button btnSiguiente = (Button)oForm.Items.Item("btnSig").Specific;

            try
            {
                if ((Globals.SBO_Application.MessageBox("¿Esta seguro de crear el asiento distribuido?.", 1, "Si", "No") == 1))
                {
                    btnEjecutar.Item.Enabled = false;
                    btnAnterior.Item.Enabled = false;
                    btnSiguiente.Item.Enabled = false;

                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific;
                    SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_4");
                    int Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);

                    Globals.StartTransaction();
                    SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJE.Memo = ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Value;
                    oJE.ReferenceDate = Globals.ConvertDate(((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Value);
                    oJE.TransactionCode = "RCSF";

                    for (int i = 0; i < oLista.Rows.Count; i++)
                    {
                        if (i != 0) oJE.Lines.Add();

                        oJE.Lines.BPLID = 2;
                        oJE.Lines.AccountCode = oLista.GetValue("Col_0", i).ToString();
                        if (Convert.ToDouble(oLista.GetValue("Col_2", i)) > 0)
                        {
                            oJE.Lines.Debit = Convert.ToDouble(oLista.GetValue("Col_2", i));
                            oJE.Lines.Credit = 0;
                        }
                        if (Convert.ToDouble(oLista.GetValue("Col_3", i)) > 0)
                        {
                            oJE.Lines.Debit = 0;
                            oJE.Lines.Credit = Convert.ToDouble(oLista.GetValue("Col_3", i));
                        }
                        if (!string.IsNullOrEmpty(oLista.GetValue("Col_8", i).ToString())) oJE.Lines.ProjectCode = oLista.GetValue("Col_8", i).ToString();

                        switch (Dimension)
                        {
                            case 1:
                                oJE.Lines.CostingCode = oLista.GetValue("Col_6", i).ToString();
                                break;
                            case 2:
                                oJE.Lines.CostingCode2 = oLista.GetValue("Col_6", i).ToString();
                                break;
                            case 3:
                                oJE.Lines.CostingCode3 = oLista.GetValue("Col_6", i).ToString();
                                break;
                            case 4:
                                oJE.Lines.CostingCode4 = oLista.GetValue("Col_6", i).ToString();
                                break;
                            case 5:
                                oJE.Lines.CostingCode5 = oLista.GetValue("Col_6", i).ToString();
                                break;
                        }
                    }

                    Globals.lRetCode = oJE.Add();
                    if (Globals.lRetCode != 0)
                    {
                        Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                        throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                    }
                    else
                    {
                        string TransId;
                        Globals.oCompany.GetNewObjectCode(out TransId);

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
                                Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                                throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                            }
                            else
                            {
                                Globals.CommitTransaction();
                                btnEjecutar.Item.Enabled = false;
                                btnAnterior.Item.Enabled = false;
                                btnSiguiente.Item.Enabled = true;
                                ((EditText)oForm.Items.Item("et4_1").Specific).Value = TransId;
                                ((StaticText)oForm.Items.Item("st4_1").Specific).Item.Visible = true;
                                ((EditText)oForm.Items.Item("et4_1").Specific).Item.Visible = true;
                                ((LinkedButton)oForm.Items.Item("lb4_1").Specific).Item.Visible = true;
                                Globals.MessageBox("El asiento de reclasificación se ha generado correctamente con el número de transacción " + TransId);
                            }
                            Globals.Release(oUserTable);
                        }
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
                        oAsientos = new List<int>();
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
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_1").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific).Item.Visible = true;
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
                        stPaso.Caption = "Paso 2 de 4";
                        btnEjecutar.Item.Visible = false;
                        stPasoDes.Caption = "Selección de contabilizaciones";
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific).Item.Visible = false;
                        //oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific;
                        //oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3");
                        //if (oGrid.DataTable != null)
                        //{
                        //    oGrid.DataTable.Rows.Clear();
                        //    oLista.Rows.Clear();
                        //    oGrid.DataTable = null;
                        //}
                        break;
                    case "Paso 4 de 4":
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
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific).Item.Visible = true;
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
                Validar(oForm);

                oForm.Freeze(true);
                switch (stPaso.Caption.ToString())
                {
                    case "Paso 1 de 4":
                        stPaso.Caption = "Paso 2 de 4";
                        stPasoDes.Caption = "Selección de contabilizaciones";
                        btnAnterior.Item.Enabled = true;
                        btnEjecutar.Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("etAux").Specific).Item.Click();
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st1_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_1").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific).Item.Visible = true;
                        CargaGrid("grid2", oForm);
                        break;
                    case "Paso 2 de 4":
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
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = true;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = true;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = true;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific).Item.Visible = true;
                        break;
                    case "Paso 3 de 4":
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
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific).Item.Visible = false;
                        ((SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific).Item.Visible = false;
                        ((SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific).Item.Visible = true;
                        CargaGrid("grid4", oForm);
                        break;
                    case "Paso 4 de 4":
                        oForm.Close();
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
                        break;
                    case "Paso 2 de 4":
                        if (oAsientos.Count == 0) throw new Exception("Debe seleccionar al menos un asiento de la lista");
                        break;
                    case "Paso 3 de 4":
                        SAPbouiCOM.EditText et3_2 = (SAPbouiCOM.EditText)oForm.Items.Item("et3_2").Specific;
                        SAPbouiCOM.EditText et3_3 = (SAPbouiCOM.EditText)oForm.Items.Item("et3_3").Specific;
                        SAPbouiCOM.ComboBox cb3_4 = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific;
                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific;
                        SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3");

                        if (string.IsNullOrEmpty(et3_2.Value.Trim())) throw new Exception("Debe ingresar una fecha de contabilización");
                        if (string.IsNullOrEmpty(et3_3.Value.Trim())) throw new Exception("Debe ingresar una glosa para el asiento");
                        if (cb3_4.Selected == null) throw new Exception("Debe seleccionar una reglas de distribución");
                        if (oLista == null || oLista.Rows.Count == 0) throw new Exception("No se encontró la lista para la distribución");
                        if (string.IsNullOrEmpty(oLista.GetValue("PrcCode", 0).ToString())) throw new Exception("No se encontró la lista para la distribución");
                        else
                        {
                            oDistribucion = new List<OPRC>();
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
                            CargaGrid("grid3", oForm);
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

        private static void CargaGrid(string ItemUID, Form oForm)
        {
            try
            {
                Grid oGrid;
                SAPbouiCOM.DataTable oLista;
                EditTextColumn etColumna;
                int Dimension;
                string DescDimension, fechaInicio, fechaFin;

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
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1_2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_2");

                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCuentaGasto;
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
                        oGrid.Columns.Item("Select").Type = BoGridColumnType.gct_CheckBox;
                        oGrid.Columns.Item("Select").TitleObject.Caption = "Seleccionar";
                        oGrid.Columns.Item("Select").Editable = true;
                        oGrid.Columns.Item("AcctCode").TitleObject.Caption = "Código";
                        oGrid.Columns.Item("AcctCode").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("AcctCode")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("AcctName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("AcctName").Editable = false;
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid2":
                        ComboBox cboDimension = (SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific;
                        Dimension = Convert.ToInt32(cboDimension.Selected.Value);
                        DescDimension = Dimension == 1 ? "ProfitCode" : Dimension == 2 ? "OcrCode2" : Dimension == 3 ? "OcrCode3" : Dimension == 4 ? "OcrCode4" : "OcrCode5";
                        fechaInicio = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific).Value;
                        fechaFin = ((SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific).Value;
                        SAPbouiCOM.EditText FechaInicio = (SAPbouiCOM.EditText)oForm.Items.Item("et1_2").Specific;
                        SAPbouiCOM.EditText FechaFin = (SAPbouiCOM.EditText)oForm.Items.Item("et1_3").Specific;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid2").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");

                        string cuentas = "'" + string.Join("','", oCuentas) + "'";
                        string cecos = "'" + string.Join("','", oCecos) + "'";
                        Globals.Query = AddOnRclsGastos.Properties.Resources.ListarAsientos;
                        Globals.Query = string.Format(Globals.Query, DescDimension, fechaInicio, fechaFin, cuentas, cecos);
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
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
                        oGrid.Columns.Item("Account").Editable = false;
                        oGrid.Columns.Item("AcctName").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("AcctName").Editable = false;
                        oGrid.Columns.Item("AcctName").Visible = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Account")));
                        etColumna.LinkedObjectType = "1";
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
                        oGrid.AutoResizeColumns();
                        break;
                    case "grid3":
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid3").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_3");
                        Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);

                        if (((SAPbouiCOM.ComboBox)oForm.Items.Item("cb3_4").Specific).Selected.Value == "1")
                            //Se debe cambiar el query para obtener los valores de SISMAN
                            Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCCProductivo;
                        else
                            Globals.Query = AddOnRclsGastos.Properties.Resources.ListarCCProductivo;

                        Globals.Query = string.Format(Globals.Query, Dimension);
                        oGrid.DataTable = oLista;
                        oGrid.DataTable.Rows.Clear();
                        oLista.Rows.Clear();
                        oGrid.DataTable.ExecuteQuery(Globals.Query);
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
                    case "grid4":
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");
                        Dimension = Convert.ToInt32(((SAPbouiCOM.ComboBox)oForm.Items.Item("cb1_1").Specific).Selected.Value);
                        DescDimension = Dimension == 1 ? "ProfitCode" : Dimension == 2 ? "OcrCode2" : Dimension == 3 ? "OcrCode3" : Dimension == 4 ? "OcrCode4" : "OcrCode5";

                        oAsientoDist = new List<OJDT>();

                        foreach (int index in oAsientos)
                        {
                            OJDT detalle = new OJDT();
                            detalle.AcctCode = oLista.GetValue("Account", index).ToString();
                            detalle.AcctName = oLista.GetValue("AcctName", index).ToString();
                            detalle.PrcCode = oLista.GetValue(DescDimension, index).ToString();
                            detalle.PrcName = oLista.GetValue("PrcName", index).ToString();
                            detalle.Project = oLista.GetValue("Project", index).ToString();
                            detalle.TotalML = Convert.ToDouble(oLista.GetValue("Total ML", index).ToString());
                            detalle.TotalMS = Convert.ToDouble(oLista.GetValue("Total MS", index).ToString());
                            oAsientoDist.Add(detalle);
                        }

                        int intcf = -1;
                        oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid4").Specific;
                        oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_4");
                        oLista.Rows.Clear();
                        int factor = oDistribucion.Sum(x => x.Peso);
                        var ListaCuentas = oAsientoDist.GroupBy(x => x.AcctCode).Select(y => y.Key).ToList();
                        var asientoAux = oAsientoDist.GroupBy(x => new { x.AcctCode, x.AcctName, x.PrcCode, x.PrcName, x.Project }).Select(x => new OJDT
                        {
                            AcctCode = x.Key.AcctCode,
                            AcctName = x.Key.AcctName,
                            PrcCode = x.Key.PrcCode,
                            PrcName = x.Key.PrcName,
                            Project = x.Key.Project,
                            TotalML = x.Sum(y => y.TotalML),
                            TotalMS = x.Sum(y => y.TotalMS)
                        }).Where(z => z.TotalML + z.TotalMS > 0).ToList();
                        var asientoAux2 = oAsientoDist.GroupBy(x => new { x.AcctCode, x.AcctName, x.Project }).Select(x => new OJDT
                        {
                            AcctCode = x.Key.AcctCode,
                            AcctName = x.Key.AcctName,
                            Project = x.Key.Project,
                            TotalML = x.Sum(y => y.TotalML),
                            TotalMS = x.Sum(y => y.TotalMS)
                        }).Where(z => z.TotalML + z.TotalMS > 0).ToList();

                        foreach (string cta in ListaCuentas)
                        {
                            if (asientoAux.Where(x => x.AcctCode == cta).Count() > 0)
                            {
                                foreach (OJDT detalle in asientoAux.Where(x => x.AcctCode == cta))
                                {
                                    intcf++;
                                    oLista.Rows.Add();
                                    oLista.SetValue("Col_0", intcf, detalle.AcctCode);
                                    oLista.SetValue("Col_1", intcf, detalle.AcctName);
                                    if (detalle.TotalML < 0) oLista.SetValue("Col_2", intcf, detalle.TotalML * -1);
                                    else oLista.SetValue("Col_3", intcf, detalle.TotalML);
                                    if (detalle.TotalMS < 0) oLista.SetValue("Col_4", intcf, detalle.TotalML * -1);
                                    else oLista.SetValue("Col_5", intcf, detalle.TotalML);
                                    oLista.SetValue("Col_6", intcf, detalle.PrcCode);
                                    oLista.SetValue("Col_7", intcf, detalle.PrcName);
                                    if (!string.IsNullOrEmpty(detalle.Project)) oLista.SetValue("Col_8", intcf, detalle.Project);
                                }

                                foreach (OJDT detalle in asientoAux2.Where(x => x.AcctCode == cta))
                                {
                                    for (int j = 0; j < oDistribucion.Count; j++)
                                    {
                                        intcf++;
                                        oLista.Rows.Add();
                                        oLista.SetValue("Col_0", intcf, detalle.AcctCode);
                                        oLista.SetValue("Col_1", intcf, detalle.AcctName);
                                        if (detalle.TotalML < 0) oLista.SetValue("Col_3", intcf, Math.Round((detalle.TotalML * -1 * oDistribucion[j].Peso) / factor, 2, MidpointRounding.AwayFromZero));
                                        else oLista.SetValue("Col_2", intcf, Math.Round((detalle.TotalML * oDistribucion[j].Peso) / factor, 2, MidpointRounding.AwayFromZero));
                                        if (detalle.TotalMS < 0) oLista.SetValue("Col_5", intcf, Math.Round((detalle.TotalML * -1 * oDistribucion[j].Peso) / factor, 2, MidpointRounding.AwayFromZero));
                                        else oLista.SetValue("Col_4", intcf, Math.Round((detalle.TotalML * oDistribucion[j].Peso) / factor, 2, MidpointRounding.AwayFromZero));
                                        oLista.SetValue("Col_6", intcf, oDistribucion[j].PrcCode);
                                        oLista.SetValue("Col_7", intcf, oDistribucion[j].PrcName);
                                        if (!string.IsNullOrEmpty(detalle.Project)) oLista.SetValue("Col_8", intcf, detalle.Project);
                                    }
                                }
                            }
                        }

                        oGrid.DataTable = oLista;

                        oGrid.Columns.Item("Col_0").TitleObject.Caption = "Cuenta";
                        oGrid.Columns.Item("Col_0").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_0")));
                        etColumna.LinkedObjectType = "1";
                        oGrid.Columns.Item("Col_1").TitleObject.Caption = "Descripción";
                        oGrid.Columns.Item("Col_1").Editable = false;
                        oGrid.Columns.Item("Col_2").TitleObject.Caption = "Debe ML";
                        oGrid.Columns.Item("Col_2").RightJustified = true;
                        oGrid.Columns.Item("Col_2").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_2");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_3").TitleObject.Caption = "Haber ML";
                        oGrid.Columns.Item("Col_3").RightJustified = true;
                        oGrid.Columns.Item("Col_3").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_3");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_4").TitleObject.Caption = "Debe MS";
                        oGrid.Columns.Item("Col_4").RightJustified = true;
                        oGrid.Columns.Item("Col_4").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_4");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_5").TitleObject.Caption = "Haber MS";
                        oGrid.Columns.Item("Col_5").RightJustified = true;
                        oGrid.Columns.Item("Col_5").Editable = false;
                        etColumna = (EditTextColumn)oGrid.Columns.Item("Col_5");
                        etColumna.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                        oGrid.Columns.Item("Col_6").TitleObject.Caption = "Centro Costo";
                        oGrid.Columns.Item("Col_6").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_6")));
                        etColumna.LinkedObjectType = "61";
                        oGrid.Columns.Item("Col_7").TitleObject.Caption = "Descripción CC";
                        oGrid.Columns.Item("Col_7").Editable = false;
                        oGrid.Columns.Item("Col_8").TitleObject.Caption = "Proyecto";
                        oGrid.Columns.Item("Col_8").Editable = false;
                        etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_8")));
                        etColumna.LinkedObjectType = "63";

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

        public static void LostFocus(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            try
            {
                BubbleEvent = true;
                try
                {
                    if (pVal.ColUID == "Select")
                    {
                        SAPbouiCOM.DataTable oLista;
                        if (pVal.ActionSuccess)
                            switch (pVal.ItemUID)
                            {
                                case "grid1_1":
                                    oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_1");
                                    if (oLista.GetValue("Select", pVal.Row).ToString() == "Y")
                                        oCecos.Add(oLista.GetValue("PrcCode", pVal.Row).ToString());
                                    else
                                        oCecos.Remove(oLista.GetValue("PrcCode", pVal.Row).ToString());
                                    break;
                                case "grid1_2":
                                    oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1_2");
                                    if (oLista.GetValue("Select", pVal.Row).ToString() == "Y")
                                        oCuentas.Add(oLista.GetValue("AcctCode", pVal.Row).ToString());
                                    else
                                        oCuentas.Remove(oLista.GetValue("AcctCode", pVal.Row).ToString());
                                    break;
                                case "grid2":
                                    oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_2");
                                    if (oLista.GetValue("Select", pVal.Row).ToString() == "Y")
                                        oAsientos.Add(pVal.Row);
                                    else
                                        oAsientos.Remove(pVal.Row);
                                    break;
                                case "grid3":
                                    CargaGrid("grid3", oForm);
                                    break;
                                case "grid4":
                                    CargaGrid("grid4", oForm);
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
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
