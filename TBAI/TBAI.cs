using System;
using System.Windows.Forms;
using Common;
using System.Collections;
using System.Collections.Generic;
using AX_TBAICommon;

namespace TBAI
{
    public class TBAI
    {
        public static csTBAICommon TBAICommon = null;
        //private static csAX_PanelEnvioAConector frmAX_PEC;
        /// <summary>
        /// Entrada principal del Add-on
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // Establecer la conexión
                Common.Main.connect(null);

                TBAICommon = new csTBAICommon();
                // Añadir los gestores de formularios
                addHandles();

                // Crear los menues de usuario
                addUserMenus();

                // Inicializar componentes del Main
                Common.Main.init("TBAI");
                
                // Tablas de usuario con sus campos
                createUserTables();
                createUserFields();
                //EliminarUserFields();

                // Inicializar bucle de mensajes
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                // Solo logging del mensaje, ya se ha mostrado un error al usuario
                Common.Main.Log.write(ex.ToString());
            }
        }

        /// <summary>
        /// Añade los gestores de los formularios
        /// </summary>
        private static void addHandles()
        {
            try
            {
                //csAX_PanelEnvioAConector frmAX_PEC = new csAX_PanelEnvioAConector(LocalConstants.FormIds.frmAX_PEC);

                // Añadir gestores formularios de usuario con opción de menu
                /*
                FormsHandler.addMenuFormHandle(frmAX_PEC, LocalConstants.FormIds.frmAX_PEC);
                frmAX_PEC.FormType = LocalConstants.FormIds.frmAX_PEC;

                FormsHandler.addMenuFormHandle(frmAX_PEC, LocalConstants.FormIds.frmAX_PEC);
                frmAX_PEC.FormType = LocalConstants.FormIds.frmAX_PEC;
                */

                csAX_PanelEnvioAConector frmAX_PEC = new csAX_PanelEnvioAConector(LocalConstants.FormIds.frmAX_PEC);
                Common.FormsHandler.addFormHandle(frmAX_PEC, frmAX_PEC.FormType);

            }
            catch (Exception ex)
            {
                Functions.Messages.showError("Error al inicializar los gestores de formularios");
                throw ex;
            }
        }

        /// <summary> 
        /// Crea los menus necesarios 
        /// </summary>
        private static void addUserMenus()
        {
        }

        /// <summary>
        /// Crea las tablas de usuario
        /// </summary>
        private static void createUserTables()
        {
            UserTable oUserTable = new UserTable();
            try
            {
                oUserTable.create("AX_ESTTBAI_0", "Estado TBAI", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            }
            catch (Exception ex)
            {
                Functions.Messages.showError("Error creando tablas de usuario");
                //throw ex;
            }
            try
            {
                oUserTable.create("AX_ESTTBAI_1", "Estado TBAI - Env. a Conector", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            }
            catch (Exception ex)
            {
                Functions.Messages.showError("Error creando tablas de usuario");
                //throw ex;
            }
            try
            {
                oUserTable.create("AX_ESTTBAI_2", "Estado TBAI - Res. de Conector", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            }
            catch (Exception ex)
            {
                Functions.Messages.showError("Error creando tablas de usuario");
                //throw ex;
            }

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);

                oUserTable = null;

                GC.Collect();
            }
            catch (Exception ex) { }
        }
        /// <summary>
        /// Crea los campos de las tablas de usuario
        /// </summary>
        private static void createUserFields()
        {
            #region OVTG
            try
            {
                UserField oUserField = new UserField();
                
                oUserField.create("OVTG", "AX_TBAI_CodImp", "TBAI Código Impuesto", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FECasSuj", "TBAI Fac. Exp. Caso Sujeta", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FECasExe", "TBAI Fac. Exp. Caso Exenta", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FECauExe", "TBAI Fac. Exp. Causa Exenta", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FETipNExe", "TBAI Fac. Exp. Tipo No Exenta", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FECauNSuj", "TBAI Fac. Exp. Causa No Sujeta", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FEClaRegEsp", "TBAI Fac. Exp. Clave Reg. Esp.", SAPbobsCOM.BoFieldTypes.db_Alpha, 2);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OVTG", "AX_TBAI_FRClaRegEsp", "TBAI Fac. Rec. Clave Reg. Esp.", SAPbobsCOM.BoFieldTypes.db_Alpha, 2);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_Ser", "TBAI Servcio", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_InvSujPas", "TBAI Inversión Sujeto Pasivo", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_BieInv", "TBAI Bienes Inversión", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_DetOpeInt", "TBAI Det. Ope. Intracomunitaria", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_Imp", "TBAI Importación", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_FacSim", "TBAI Factura Simplificada", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_FacRegSim", "TBAI Fac. Reg. Simplificada", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_FacRegRecEqu", "TBAI Fac. Reg. Rec. Equiv.", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_CasREAV", "TBAI Caso REAV", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.UserFieldsMD.DefaultValue = "C";
                oUserField.create("OVTG", "AX_TBAI_FRCG240", "TBAI Fac. Rec. Caso Gasto 240", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "N";
                oUserField.create("OVTG", "AX_TBAI_Arr", "TBAI Arrendamiento", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.create("OVTG", "AX_TBAI_CasREGEAju", "TBAI Caso REGE Ajustar", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            #endregion
            #region OINV
            try
            {
                UserField oUserField = new UserField();

                oUserField.addValidValue("N", "No");
                oUserField.addValidValue("Y", "Sí");
                oUserField.UserFieldsMD.DefaultValue = "Y";
                oUserField.create("OINV", "AX_EnviarTBA", "Enviar a TBAI", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OINV", "AX_CodQR", "Código QR", SAPbobsCOM.BoFieldTypes.db_Memo, 1000);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("OINV", "AX_MotRec", "Motivo Rectificación", SAPbobsCOM.BoFieldTypes.db_Memo, 1000);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            #endregion
            #region AX_ESTTBAI_0
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "DocNum", "Nº Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "ObjType", "Tipo Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "CardCode", "Código Cliente", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "EnvCon", "Enviado a Conector", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "ResEnvCon", "Resultado Enviado a Conector", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "ResCon", "Respuesta de Conector", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "ResResCon", "Resultado Respuesta de Conector", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "GenQR", "Generación QR", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_0", "ResGenQR", "Resultado Respuesta Generación QR", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            #endregion
            #region AX_ESTTBAI_1
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_1", "Accion", "Acción", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_1", "NumInt", "Nº Intento", SAPbobsCOM.BoFieldTypes.db_Numeric, 4);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_1", "Info", "Información", SAPbobsCOM.BoFieldTypes.db_Memo, 1000);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_1", "FecHor", "Fecha Hora", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_1", "JSONDoc", "JSON Documento", SAPbobsCOM.BoFieldTypes.db_Memo, 1000);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            #endregion
            #region AX_ESTTBAI_2
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_2", "Accion", "Acción", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_2", "NumInt", "Nº Intento", SAPbobsCOM.BoFieldTypes.db_Numeric, 4);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_2", "Info", "Información", SAPbobsCOM.BoFieldTypes.db_Memo, 1000);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario");
                //throw ex;
            }
            try
            {
                UserField oUserField = new UserField();

                oUserField.create("AX_ESTTBAI_2", "FecHor", "Fecha Hora", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            }
            catch (Exception ex)
            {
                if (ex.Message != "Error UserField.add, Esta entrada ya existe en las tablas siguientes (ODBC -2035) (Code: -2035)") Functions.Messages.showError("Error creando Campos de usuario"); 
                //throw ex;
            }
            #endregion
        }

        private static void EliminarUserFields()
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Common.Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);

            //oUserFieldsMD = null;

            GC.Collect();

            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Common.Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            for (int i = 10; i <= 28; i++)
            {
                oUserFieldsMD.GetByKey("OVTG", i);

                if (oUserFieldsMD.Remove() != 0)
                {
                    string a = Common.Main.Company.GetLastErrorCode().ToString() + " " + Common.Main.Company.GetLastErrorDescription().ToString();
                }
            }
        }

    }
}
