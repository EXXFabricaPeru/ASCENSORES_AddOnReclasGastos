using AddOnRclsGastos.App;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.DBStructure
{
    class CreateStructure
    {
        public static void CreateStruct()
        {
            List<List<String>> TrueFalseValues = new List<List<String>>();
            agregarValorValido("Y", "SI", ref TrueFalseValues);
            agregarValorValido("N", "NO", ref TrueFalseValues);

            List<List<String>> TipoCCValues = new List<List<String>>();
            agregarValorValido("1", "Gasto", ref TipoCCValues);
            agregarValorValido("2", "Productivo", ref TipoCCValues);

            List<List<String>> EstadoValues = new List<List<String>>();
            agregarValorValido("G", "Generado", ref EstadoValues);
            agregarValorValido("A", "Anulado", ref EstadoValues);

            #region UDO_NOMINAS
            //crearCampo("EXX_ANXN_NOMINA_CAB", "EXX_ANXN_FIERA", "FEC INI EERR ABONADO", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("OACT", "EXX_ADRG_CTAGASTO", "Cuenta de gasto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "N", "", false, TrueFalseValues);
            crearCampo("OPRC", "EXX_ADRG_TIPOCC", "Tipo centro de costo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", "", false, TipoCCValues);


            crearTabla("EXX_ADRG_HIST", "Histórico-Asiento Reclas Gasto", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAE", "Fecha Ejecución", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None,10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_TRANSID", "N° Asiento", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAC", "Fecha Contabilización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_GLOSA", "Glosa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_EST", "Estado Asiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "G", "", false, EstadoValues);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAA", "Fecha Anulación", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_TRANSIDA", "N° Asiento anulación", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            #endregion

        }

        private static void agregarValorValido(String Value, String Description, ref List<List<String>> ValidValues)
        {
            List<String> ValidValue = new List<String>();
            ValidValue.Add(Value);
            ValidValue.Add(Description);
            ValidValues.Add(ValidValue);
        }

        private static bool crearTabla(string tabla, string nombretabla, SAPbobsCOM.BoUTBTableType tipo = SAPbobsCOM.BoUTBTableType.bott_NoObject)//CHG
        {
            SAPbobsCOM.UserTablesMD oTablaUser = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                if (!oTablaUser.GetByKey(tabla))
                {
                    oTablaUser.TableName = tabla;
                    oTablaUser.TableDescription = nombretabla;
                    oTablaUser.TableType = tipo;

                    int RetVal = oTablaUser.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                    else
                        return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTablaUser);
            }
        }

        private static void crearCampo(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo,
            SAPbobsCOM.BoFldSubTypes subtipo, int tamaño, string ValorPorDefecto, string sLinkedTable,
            Boolean Mandatory, List<List<String>> ValidValues)//CHG
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string cadena = "select \"FieldID\" from CUFD where (\"TableID\"='" + tabla + "' or \"TableID\"='@" + tabla + "') and \"AliasID\"='" + campo + "'";
            rs.DoQuery(cadena);

            existeCampo = rs.RecordCount;
            int FieldID = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;

            SAPbobsCOM.UserFieldsMD oCampo = (SAPbobsCOM.UserFieldsMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (existeCampo == 0)//Crear
                {
                    oCampo.TableName = tabla;
                    oCampo.Name = campo;
                    oCampo.Description = descripcion;
                    oCampo.Type = tipo;
                    oCampo.SubType = subtipo;
                    oCampo.Mandatory = Mandatory ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    if (tamaño > 0)
                    {
                        oCampo.EditSize = tamaño;
                    }

                    if (sLinkedTable.ToString() != "")
                        oCampo.LinkedTable = sLinkedTable;

                    if (ValidValues != null)
                    {
                        foreach (List<String> ValidValue in ValidValues)
                        {
                            oCampo.ValidValues.Value = ValidValue[0];
                            oCampo.ValidValues.Description = ValidValue[1];
                            oCampo.ValidValues.Add();
                        }
                    }

                    if (ValorPorDefecto.ToString() != "")
                    {
                        oCampo.DefaultValue = ValorPorDefecto;
                    }

                    int RetVal = oCampo.Add();
                    if (RetVal != 0)
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                }
                //else//Actualizar
                //{
                //    oCampo.GetByKey("@" + tabla, FieldID);
                //    oCampo.Description = descripcion;
                //    if (ValidValues != null)
                //    {
                //        foreach (List<String> ValidValue in ValidValues)
                //        {
                //            Boolean Existe = false;
                //            for (int i = 0; i < oCampo.ValidValues.Count; i++)
                //            {
                //                oCampo.ValidValues.SetCurrentLine(i);
                //                if (oCampo.ValidValues.Value == ValidValue[0])
                //                    Existe = true;

                //            }

                //            if (!Existe)
                //            {
                //                oCampo.ValidValues.Value = ValidValue[0];
                //                oCampo.ValidValues.Description = ValidValue[1];
                //                oCampo.ValidValues.Add();
                //            }
                //        }
                //    }

                //    if (ValorPorDefecto.ToString() != "")
                //    {
                //        oCampo.DefaultValue = ValorPorDefecto;
                //    }

                //    int RetVal = oCampo.Update();
                //    if ((RetVal != 0))
                //    {
                //        String errMsg;
                //        int errCode;
                //        oCompany.GetLastError(out errCode, out errMsg);
                //        throw new Exception(errMsg);
                //    }
                //}
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCampo);
            }
        }
    }
}
