using AddOnRclsGastos.App;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.Functionality
{
    public class SRF_HistoricoAsientos
    {
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
                string FormName = "\\Views\\frmHistoricoAsientos.srf";
                fcp.XmlData = Globals.LoadFromXML(ref FormName);
                oForm = Globals.SBO_Application.Forms.AddEx(fcp);
                oForm.Visible = true;

            }
        }

        public static void ItemPressed(ref ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess)
                    switch (pVal.ItemUID)
                    {
                        case "btnBuscar":
                            Buscar(pVal, oForm, out BubbleEvent);
                            break;
                        case "btnAnular":
                            Anular(pVal, oForm, out BubbleEvent);
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

        private static void Buscar(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1");
                EditTextColumn etColumna;
                ComboBoxColumn cbColumna;

                ComboBox Estado = (ComboBox)oForm.Items.Item("cbEstado").Specific;
                if (Estado.Selected == null) throw new Exception("Primero seleccione un estado.");
                string EstadoFiltro = "'" + (Estado.Selected.Value == "T" ? "G','A" : Estado.Selected.Value) + "'";

                Globals.Query = AddOnRclsGastos.Properties.Resources.ListarHistorico;
                Globals.Query = string.Format(Globals.Query, EstadoFiltro);
                oGrid.DataTable = oLista;
                oGrid.DataTable.Rows.Clear();
                oLista.Rows.Clear();
                oGrid.DataTable.ExecuteQuery(Globals.Query);
                oGrid.Columns.Item("Col_0").Type = BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("Col_0").TitleObject.Caption = "Seleccionar";
                oGrid.Columns.Item("Col_1").TitleObject.Caption = "Fecha Ejecución";
                oGrid.Columns.Item("Col_1").Editable = false;
                oGrid.Columns.Item("Col_2").TitleObject.Caption = "N° Asiento";
                oGrid.Columns.Item("Col_2").Editable = false;
                etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_2")));
                etColumna.LinkedObjectType = "30";
                oGrid.Columns.Item("Col_3").TitleObject.Caption = "Estado";
                oGrid.Columns.Item("Col_3").Editable = false;
                cbColumna = ((ComboBoxColumn)(oGrid.Columns.Item("Col_3")));
                cbColumna.ValidValues.Add("T", "Todo");
                cbColumna.ValidValues.Add("G", "Generado");
                cbColumna.ValidValues.Add("A", "Anulado");
                cbColumna.DisplayType = BoComboDisplayType.cdt_Description;
                oGrid.Columns.Item("Col_4").TitleObject.Caption = "Fecha contabilización";
                oGrid.Columns.Item("Col_4").Editable = false;
                oGrid.Columns.Item("Col_5").TitleObject.Caption = "Glosa";
                oGrid.Columns.Item("Col_5").Editable = false;
                oGrid.Columns.Item("Col_6").TitleObject.Caption = "Fecha anulación";
                oGrid.Columns.Item("Col_6").Editable = true;
                oGrid.Columns.Item("Col_7").TitleObject.Caption = "N° Asiento anulado";
                oGrid.Columns.Item("Col_7").Editable = true;
                etColumna = ((EditTextColumn)(oGrid.Columns.Item("Col_7")));
                etColumna.LinkedObjectType = "30";
                oGrid.Columns.Item("Code").Editable = false;
                oGrid.Columns.Item("Code").Visible = false;

                if (Estado.Selected.Value == "A")
                {
                    oGrid.Columns.Item("Col_0").Editable = true;
                    oGrid.Columns.Item("Col_0").Visible = true;
                    ((Button)oForm.Items.Item("btnAnular").Specific).Item.Visible = true;
                }
                else
                {
                    oGrid.Columns.Item("Col_0").Editable = false;
                    oGrid.Columns.Item("Col_0").Visible = false;
                    ((Button)oForm.Items.Item("btnAnular").Specific).Item.Visible = false;

                    if (Estado.Selected.Value == "G")
                    {
                        oGrid.Columns.Item("Col_6").TitleObject.Caption = "Fecha anulación";
                        oGrid.Columns.Item("Col_6").Editable = false;
                        oGrid.Columns.Item("Col_7").TitleObject.Caption = "N° Asiento anulado";
                        oGrid.Columns.Item("Col_7").Editable = false;
                    }
                }
                oGrid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
        }

        private static void Anular(ItemEvent pVal, Form oForm, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                SAPbouiCOM.DataTable oLista = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_1");
                DateTime fecha = DateTime.Now;
                int rpta = Globals.SBO_Application.MessageBox("EXX: Por favor elija una opción para la cancelación del asiento:\n\t1. Fecha actual.\n\t2.Fecha de documento.", 1, "Opción 1", "Opción 2", "Cancelar");
                if (rpta == 3) return;

                for (int i = 0; i < oLista.Rows.Count; i++)
                {
                    if (oLista.GetValue("Col_0", i).ToString() == "Y")
                    {
                        if (rpta == 2) fecha = Globals.ConvertDate(oLista.GetValue("Col_4", i).ToString());

                        Globals.StartTransaction();
                        SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        int TransId = Convert.ToInt32(oLista.GetValue("Col_2", i).ToString());
                        oJE.GetByKey(TransId);
                        oJE.StornoDate = fecha;
                        Globals.lRetCode = oJE.Cancel();
                        if (Globals.lRetCode != 0)
                        {
                            Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                            throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                        }
                        else
                        {
                            string TransId2;
                            Globals.oCompany.GetNewObjectCode(out TransId2);
                            SAPbobsCOM.UserTable oUserTable = Globals.oCompany.UserTables.Item("EXX_ADRG_HIST");
                            string Code = oLista.GetValue("Code", i).ToString();

                            if (oUserTable.GetByKey(Code))
                            {
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_EST").Value = "A";
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_FECHAA").Value = DateTime.Now.ToString("dd/MM/yyyy");
                                oUserTable.UserFields.Fields.Item("U_EXX_ADRG_TRANSIDA").Value = TransId2;
                                if (oUserTable.Update() != 0)
                                {
                                    Globals.oCompany.GetLastError(out Globals.sErrCode, out Globals.sErrMsg);
                                    throw new Exception("ErrorSAP: " + Convert.ToString(Globals.sErrCode) + " " + Globals.sErrMsg);
                                }
                                else
                                {
                                    Globals.CommitTransaction();
                                    Globals.MessageBox("El asiento " + TransId + " se anuló correctamente.");
                                }
                                Globals.Release(oUserTable);
                            }
                        }

                    }
                }
                ((Button)oForm.Items.Item("btnBuscar").Specific).Item.Click(BoCellClickType.ct_Regular);
            }
            catch (Exception ex)
            {
                if (Globals.InTransaction())
                    Globals.RollBackTransaction();
                oForm.Freeze(false);
                BubbleEvent = false;
                throw ex;
            }
        }
    }
}
