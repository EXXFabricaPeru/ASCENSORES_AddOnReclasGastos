using AddOnRclsGastos.App;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

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

            List<List<String>> TipoHoraValues = new List<List<String>>();
            agregarValorValido("01", "EXTRA_DOMINICAL", ref TipoHoraValues);
            agregarValorValido("02", "EXTRA_DIARIO", ref TipoHoraValues);
            agregarValorValido("03", "EXTRA_FERIADO", ref TipoHoraValues);

            crearTabla("EXX_ADRG_CONF", "Configuración - Reclas Gasto", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            crearCampo("EXX_ADRG_CONF", "EXX_CONF_VALOR", "Valor", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);

            crearCampo("OACT", "EXX_ADRG_CTAGASTO", "Cuenta de gasto", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "N", "", false, TrueFalseValues);
            crearCampo("OPRC", "EXX_ADRG_TIPOCC", "Tipo centro de costo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "", "", false, TipoCCValues);

            crearTabla("EXX_ADRG_HIST", "Histórico-Asiento Reclas Gasto", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAE", "Fecha Ejecución", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_TRANSID", "N° Asiento", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAC", "Fecha Contabilización", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_GLOSA", "Glosa", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_EST", "Estado Asiento", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 1, "G", "", false, EstadoValues);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_FECHAA", "Fecha Anulación", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);
            crearCampo("EXX_ADRG_HIST", "EXX_ADRG_TRANSIDA", "N° Asiento anulación", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 10, "", "", false, null);

            crearTabla("EXA_CMAC", "EXA - Grupo Maquinas", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            crearCampo("EXA_CMAC", "EXA_PERIODO", "Periodo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 12, "", "", false, null); 
            crearTabla("EXA_CMAD", "EXA - Detalle Grupo Maquinas", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            crearCampo("EXA_CMAD", "EXA_CONTRATO", "Contrato", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "", false, null);
            crearCampo("EXA_CMAD", "EXA_MAQUINA", "Maquina", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "", false, null);
            crearCampo("EXA_CMAD", "EXA_TIPO_HORA", "Tipo Hora", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 15, "", "", false, TipoHoraValues);
            crearCampo("EXA_CMAD", "EXA_CANTHH", "Cantidad", SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, 50, "", "", false, null);
            crearCampo("EXA_CMAD", "EXA_CECO", "Centro de Costo", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "", false, null);
            crearCampo("EXA_CMAD", "EXA_CODOBRERO", "Codigo Obrero", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 20, "", "", false, null);


            crearTabla("EXX_ADRG_OPRJ", "EXX - Grupo Proyectos", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            crearTabla("EXX_ADRG_PRJ1", "EXX - Detalle Grupo Proyectos", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            crearCampo("EXX_ADRG_PRJ1", "EXX_ADRG_PRJD", "Proyecto Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 50, "", "", false, null);
            crearCampo("EXX_ADRG_PRJ1", "EXX_ADRG_PRDD", "Desc.Proyecto Destino", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);
            crearCampo("EXX_ADRG_PRJ1", "EXX_ADRG_PESO", "Peso", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_Price, 10, "", "", false, null);
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

        private static void crearCampo(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo, SAPbobsCOM.BoFldSubTypes subtipo, int tamaño,
                                       string ValorPorDefecto, string sLinkedTable, Boolean Mandatory, List<List<String>> ValidValues, string LinkedType = null)
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
                    {
                        switch (LinkedType)
                        {
                            case Globals.LinkedSystemObject:
                                oCampo.LinkedSystemObject = (SAPbobsCOM.UDFLinkedSystemObjectTypesEnum)Convert.ToInt32(sLinkedTable);
                                break;
                            case Globals.LinkedUDO:
                                oCampo.LinkedUDO = sLinkedTable;
                                break;
                            case Globals.LinkedTable:
                                oCampo.LinkedTable = sLinkedTable;
                                break;
                        }
                    }

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
