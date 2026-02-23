using Common;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TBAI
{
    public class csAX_PanelEnvioAConector : DSFormEx
    {
        //SqlConnection sqlConnectionAHORA = null;
        //string EnterpriseAhoraERP = "1";
        //string ConsultaPreviaAHORA = "INSERT INTO Ahora_Sesion(SpId,IdEmpresa,IdDelegacion,IdEmpleado,IdDepartamento,IdGrupoSeguridad,IdAplic,Exclusivo,Equipo,Usuario) SELECT @@SPID,0,0,0,0,0,'AhoraInicio',0,HOST_NAME(),'ahora'";


        /// <summary>
        /// Clase con los identificadores de items del formulario
        /// </summary>
        private class FormItems
        {
            public const string GRID_ID = "8";
            public const string BTN_TEST = "B_Test";
        }

        /// <summary>
        /// Clase con los indices de las columnas del grid
        /// </summary>
        private class ColumnIndexes
        {
            public const int GRID_COL_CODE = 0;
            public const int GRID_COL_NAME = 1;
            public const int GRID_COL_1 = 2;
            public const int GRID_COL_2 = 3;
            public const int GRID_COL_3 = 4;
            public const int GRID_COL_4 = 5;
            public const int GRID_COL_5 = 6;
        }

        /// <summary>
        /// Campo de formulario "E_test1"
        /// </summary>
        private DSFieldEx m_oTestField1;

        /// <summary>
        /// Campo de formulario "E_test2"
        /// </summary>
        private DSFieldEx m_oTestField2;

        /// <summary>
        /// Campo de formulario "E_test3"
        /// </summary>
        private DSFieldEx m_oTestField3;


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="asFormType">tipo de formulario</param>
        public csAX_PanelEnvioAConector(string asFormType)
        {
            try
            {
                FormType = asFormType;

                MenuId = FormType;

                // Asignar flag de ordenacíón del grid (se creará una opción de menu!)
                //m_bSortGrid = true;

                //// Asignar flag de modalidad
                //Modal = true;

                //// Asignar el identificador del screenpainter del grid -> OJO: no hace falta NUNCA crear columnas dentro del screenpainter
                //m_sGridID = FormItems.GRID_ID;

                //// Asignar el query para la carga del grid
                //m_sQuery = @"SELECT code, name, u_col1, isnull(u_col2, 0) as COL1, 
                //	    isnull(u_col3, 0) as COL2, u_col4 as COL3, u_col5 as COL4 
                //	    from [@" + LocalConstants.UserTables.TEST_TABLE + "]";

                //// Asignar los tipos de columna
                //setColumnType(ColumnIndexes.GRID_COL_1, SAPbouiCOM.BoGridColumnType.gct_ComboBox);
                //setColumnType(ColumnIndexes.GRID_COL_5, SAPbouiCOM.BoGridColumnType.gct_ComboBox);

                //// Asignar columnas que aceptan valores null
                //setNullColumn(ColumnIndexes.GRID_COL_5);

                //// Inicializar los campos de la cabecera
                //m_oTestField1 = addField("E_Test1", SAPbouiCOM.BoDataType.dt_MEASURE, 20);
                //m_oTestField2 = addField("E_Test2", SAPbouiCOM.BoDataType.dt_PERCENT, 10);
                //m_oTestField3 = addField("E_Test3", SAPbouiCOM.BoDataType.dt_PRICE, 10);

                //// Prueba con un grid sin query
                ////addColumn("col1", SAPbouiCOM.BoFieldsType.ft_Text, 10);
                ////addColumn("col2", SAPbouiCOM.BoFieldsType.ft_Text, 10);
                ////addColumn("col3", SAPbouiCOM.BoFieldsType.ft_Text, 10);
                ////addColumn("col4", SAPbouiCOM.BoFieldsType.ft_Text, 10);
                setFilters();
            }
            catch (Exception ex)
            {
                Functions.Messages.showError(ex.ToString());
                Common.Main.Log.write(ex.ToString());
                throw ex;
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~csAX_PanelEnvioAConector()
        {
        }

        /// <summary>
        /// Asigna los filtros que se utiliza dentro de la clase
        /// </summary>
        private void setFilters()
        {
            try
            {
                Main.EventFilters.add(FormType, SAPbouiCOM.BoEventTypes.et_CLICK);
                Main.EventFilters.add(FormType, SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
                Main.EventFilters.add(FormType, SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
                Main.EventFilters.add(FormType, SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
                Main.EventFilters.add(FormType, SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
            }
            catch (Exception ex)
            {
                Common.Main.Log.write(ex.ToString());
                throw ex;
            }
        }

        /// <summary>
        /// Activacion formulario
        /// </summary>
        /// <returns></returns>
        public override void activate()
        {
            SAPbouiCOM.ComboBoxColumn ComboBoxColumn;
            string sQuery;

            try
            {
                base.activate();
                AccionesAlCargarFormulario(this.Form);
                TBAI.TBAICommon.EstablecerTablas();
                //ProgressBar.start("Cargando controles");

                //// Query de pruebas
                //sQuery = "SELECT TOP 100 itemcode, itemname FROM OITM";

                //// Cargar los valores del combo
                //m_oDSGrid.loadCombo(ColumnIndexes.GRID_COL_1, "Y", "Yes");
                //m_oDSGrid.loadCombo(ColumnIndexes.GRID_COL_1, "N", "No");

                //m_oDSGrid.loadCombo(ColumnIndexes.GRID_COL_5, "col1", "col1");
                //m_oDSGrid.loadCombo(ColumnIndexes.GRID_COL_5, "col2", "col2");
                //m_oDSGrid.loadCombo(ColumnIndexes.GRID_COL_5, sQuery);

                //// Cambiar el modo de visualización del combo
                //ComboBoxColumn = (SAPbouiCOM.ComboBoxColumn)m_oDSGrid.Grid.Columns.Item(ColumnIndexes.GRID_COL_5);
                //ComboBoxColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;


                //// Crear un grid con 3 niveles
                //m_oDSGrid.Grid.CollapseLevel = 3;
                //m_oDSGrid.addRow(2);
                //for (int i = 0; i < m_oDSGrid.RowCount; i++)
                //{
                //	m_oDSGrid.setValue(i, 0, "test0 - " + Convert.ToString(i));
                //	m_oDSGrid.setValue(i, 1, "test1 - " + Convert.ToString(i));
                //	m_oDSGrid.setValue(i, 2, "test2 - " + Convert.ToString(i));
                //	m_oDSGrid.setValue(i, 3, "test3 - " + Convert.ToString(i));
                //}
                //m_oDSGrid.Grid.Rows.ExpandAll();

            }
            catch (Exception ex)
            {
                Functions.Messages.showError(ex.ToString());
            }
            finally
            {
                //ProgressBar.stop();
            }
        }

        /// <summary>
        /// Crea una nueva instancía del gestor
        /// </summary>
        /// <returns>Instancía nueva del gestor</returns>
        public override FormHandle clone()
        {
            return new csAX_PanelEnvioAConector(FormType);
        }


        /// <summary>
        /// Función que recibirá los valores devuelto por el formulario de seleción
        /// </summary>
        /// <param name="aoValues">array de valores string</param>
        //private void setValuesSelectionForm(ArrayList aoValues)
        //{
        //	try
        //	{
        //		for (int i = 0; i < aoValues.Count; i++)
        //			Functions.Messages.showMsgBox((string)aoValues[i]);

        //		m_oTestField1.setValue((string)aoValues[0]);
        //	}
        //	catch (Exception ex)
        //	{
        //		Functions.Messages.showError("Error recuprando los valores");
        //		Main.Log.write("Error MyTestForm.setValuesSelectionForm, " + ex.ToString());
        //	}

        //}


        /// <summary>
        /// Metodo para lanzar tests varios
        /// </summary>
        private void startTesting()
        {
            //string sQuery;
            //double dTest;

            //try
            //{
            //	// Prueba selectionForm
            //	sQuery = @"SELECT DISTINCT OITM.itemcode as 'codigo de artículo', OITM.itemname as descripcion, ITM1.price as precio 
            //					FROM OITM 
            //					INNER JOIN ITM1 on ITM1.itemcode=OITM.itemcode
            //					WHERE ITM1.pricelist=1
            //					ORDER BY 'codigo de artículo'";

            //	Functions.Forms.loadSelectionForm(sQuery, true, new SelectionForm.SetValueDelegate(setValuesSelectionForm));

            //	// prueba de doubles
            //	dTest = Functions.Conversions.SAPtoDouble(m_oTestField1.getValue());
            //	dTest = Functions.Conversions.SAPtoDouble(m_oTestField2.getValue());
            //	dTest = Functions.Conversions.SAPtoDouble(m_oTestField3.getValue());

            //	m_oTestField1.setValue(Functions.Conversions.doubleToSAP(5.53));
            //	m_oTestField2.setValue(Functions.Conversions.doubleToSAP(10005.53));
            //	m_oTestField3.setValue(Functions.Conversions.doubleToSAP(5.53));


            //	dTest = Convert.ToDouble(m_oDSGrid.getValue(0, ColumnIndexes.GRID_COL_4));
            //	dTest = Convert.ToDouble(m_oDSGrid.getValue(1, ColumnIndexes.GRID_COL_4));
            //	dTest = Convert.ToDouble(m_oDSGrid.getValue(2, ColumnIndexes.GRID_COL_4));

            //	m_oDSGrid.setValue(0, ColumnIndexes.GRID_COL_4, (5.53).ToString());
            //	m_oDSGrid.setValue(1, ColumnIndexes.GRID_COL_4, (10005.53).ToString());
            //	m_oDSGrid.setValue(2, ColumnIndexes.GRID_COL_4, (5.53).ToString());
            //}
            //catch (Exception ex)
            //{
            //	throw ex;
            //}
            //finally
            //{
            //}
        }


        /// <summary>
        /// Gestión de eventos de formularios y controles 
        /// <param name="asFormUID">Identificador del formulario</param>
        /// <param name="aoItemEvent">Objeto con información sobre el evento</param>
        /// <param name="abBubbleEvent">Parametro de salida que decide si el SAP maneja el evento si (a true) o no (a false)</param>
        /// </summary>
        public override void itemEvent(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, out bool abBubbleEvent)
        {
            abBubbleEvent = true;

            try
            {
                base.itemEvent(asFormUID, ref aoItemEvent, out abBubbleEvent);
                if (!abBubbleEvent)
                    return;

                //if (aoItemEvent.BeforeAction)
                {
                    switch (aoItemEvent.EventType)
                    {
                        case BoEventTypes.et_ITEM_PRESSED:
                            //ItemEvent_et_ITEM_PRESSED(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                        //case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //	if (aoItemEvent.ItemUID == FormItems.BTN_TEST) // Boton de prueba
                        //	{
                        //		//FormsHandler.loadForm(B1Examples.ChildForm);
                        //		//startTesting();
                        //		return;
                        //	}
                        //	break;
                        case BoEventTypes.et_CLICK:
                            ItemEvent_et_CLICK(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                        case BoEventTypes.et_FORM_RESIZE:
                            ItemEvent_et_FORM_RESIZE(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                        case BoEventTypes.et_MATRIX_LINK_PRESSED:
                            ItemEvent_et_MATRIX_LINK_PRESSED(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                        case BoEventTypes.et_CHOOSE_FROM_LIST:
                            ItemEvent_et_CHOOSE_FROM_LIST(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                        case BoEventTypes.et_COMBO_SELECT:
                            ItemEvent_et_COMBO_SELECT(asFormUID, ref aoItemEvent, ref abBubbleEvent);
                            break;
                    }
                }
                //else
                {
                }
                //abBubbleEvent = true;
            }
            catch (Exception ex)
            {
                Main.Log.write(ex.ToString());
                Functions.Messages.showError("Error inesperado. Consulta log:" + ex.Message);
            }
        }

        /// <summary>
        /// Gestión de eventos de menu
        /// <param name="aoMenuEvent">Objeto con información sobre el evento</param>
        /// <param name="abBubbleEvent">Parametro de salida que decide si el SAP maneja el evento si (a true) o no (a false)</param>
        /// </summary>
        public override void menuEvent(ref SAPbouiCOM.MenuEvent aoMenuEvent, out bool abBubbleEvent)
        {
            abBubbleEvent = true;

            try
            {
                base.menuEvent(ref aoMenuEvent, out abBubbleEvent);
                if (!abBubbleEvent)
                    return;
            }
            catch (Exception ex)
            {
                Main.Log.write(ex.ToString());
                Functions.Messages.showError("Error inesperado. Consulta log:" + ex.Message);
            }
        }

        #region
        private void ItemEvent_et_CLICK(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, ref bool abBubbleEvent)
        {
            switch (aoItemEvent.ItemUID)
            {
                case "FolderTi0":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        this.Form.PaneLevel = 1;
                    }
                    break;
                case "FolderTi1":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        this.Form.PaneLevel = 2;
                    }
                    break;
                case "btnRec":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            ProcesoCargarEnvios(this.Form);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
                case "btnEnv":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            TBAI.TBAICommon.EnviarAConector(true);
                            ProcesoCargarEnvios(this.Form);
                        }
                        catch (Exception ex)
                        {
                            Main.Log.write(ex.ToString());
                            Functions.Messages.showError("Error inesperado. Consulta log:" + ex.Message);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
                case "btnGenQR":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            TBAI.TBAICommon.GenerarQR();
                            ProcesoCargarEnvios(this.Form);
                        }
                        catch (Exception ex)
                        {
                            Main.Log.write(ex.ToString());
                            Functions.Messages.showError("Error inesperado. Consulta log:" + ex.Message);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
                case "btnRes":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            TBAI.TBAICommon.RespuestaDeConector(TBAI.TBAICommon.ObtenerListaParaRespuestaDeConector());
                            ProcesoCargarEnvios(this.Form);
                        }
                        catch (Exception ex) { }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
                case "btnGenTBAI":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            TBAI.TBAICommon.EnviarAConector();
                            //TBAI.TBAICommon.RespuestaDeConector(TBAI.TBAICommon.ObtenerListaParaRespuestaDeConector());
                            ProcesoCargarEnvios(this.Form);
                        }
                        catch (Exception ex)
                        {
                            Main.Log.write(ex.ToString());
                            Functions.Messages.showError("Error inesperado. Consulta log:" + ex.Message);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
                case "grdDoc":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            this.Form.Freeze(true);

                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDoc").Specific;

                            int FilaSeleccionada = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder));
                            string Code = oGrid.DataTable.Columns.Item("Code").Cells.Item(FilaSeleccionada).Value.ToString();

                            CargarDetalleDocumentoEnvioAConector(this.Form, Code);
                            CargarDetalleDocumentoRespuestaDeConector(this.Form, Code);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                        }
                    }
                    break;
            }
        }

        private void ItemEvent_et_FORM_RESIZE(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, ref bool abBubbleEvent)
        {
            if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
            {
                try
                {
                    if (this.Form != null)
                    {
                        this.Form.Freeze(true);

                        ((SAPbouiCOM.Item)this.Form.Items.Item("lblDocDet")).Top = ((SAPbouiCOM.Item)this.Form.Items.Item("grdDoc")).Top + ((SAPbouiCOM.Item)this.Form.Items.Item("grdDoc")).Height + 10;
                        ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetEC")).Top = ((SAPbouiCOM.Item)this.Form.Items.Item("lblDocDet")).Top + ((SAPbouiCOM.Item)this.Form.Items.Item("lblDocDet")).Height + 10;
                        ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetEC")).Height = 130;
                        ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetRC")).Top = ((SAPbouiCOM.Item)this.Form.Items.Item("lblDocDet")).Top + ((SAPbouiCOM.Item)this.Form.Items.Item("lblDocDet")).Height + 10;
                        ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetRC")).Height = 130;

                        Main.Application.ActivateMenuItem("1300");
                    }
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (this.Form != null)
                    {
                        this.Form.Freeze(false);
                    }
                }
            }
        }

        private void ItemEvent_et_MATRIX_LINK_PRESSED(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, ref bool abBubbleEvent)
        {
            switch (aoItemEvent.ItemUID)
            {
                case "grdDoc":
                    switch (aoItemEvent.ColUID)
                    {
                        case "Nº Documento":
                            if (aoItemEvent.BeforeAction)
                            {
                                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDoc").Specific;
                                switch (oGrid.DataTable.Columns.Item("Tipo Documento").Cells.Item(aoItemEvent.Row).Value.ToString())
                                {
                                    case "Factura Venta":
                                        Main.Application.Menus.Item("2053").Activate();
                                        break;
                                    case "Abono Venta":
                                        Main.Application.Menus.Item("2055").Activate();
                                        break;
                                }
                                SAPbouiCOM.Form oForm = Main.Application.Forms.ActiveForm;
                                oForm.Mode = BoFormMode.fm_FIND_MODE;
                                oForm.Freeze(true);
                                string v = oGrid.DataTable.Columns.Item("Nº Documento").Cells.Item(aoItemEvent.Row).Value.ToString();
                                ((EditText)oForm.Items.Item("8").Specific).Value = v;
                                oForm.Items.Item("1").Click();
                                oForm.Freeze(false);

                                abBubbleEvent = false;
                            }
                            break;
                    }
                    break;
            }
        }

        private void ItemEvent_et_CHOOSE_FROM_LIST(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, ref bool abBubbleEvent)
        {
            switch (aoItemEvent.ItemUID)
            {
                case "txtDDocNum":
                case "txtHDocNum":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvent = (IChooseFromListEvent)aoItemEvent;

                            SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                            string Codigo = oDataTable.GetValue("DocNum", 0).ToString();

                            //((EditText)this.Form.Items.Item("txtDDocNum").Specific).Value = Codigo;
                            this.Form.DataSources.UserDataSources.Item(aoItemEvent.ItemUID.Replace("txt", "ds")).ValueEx = Codigo;
                        }
                        catch(Exception ex)
                        {

                        }
                    }
                    break;
                case "txtDIC":
                case "txtHIC":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvent = (IChooseFromListEvent)aoItemEvent;

                            SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                            string Codigo = oDataTable.GetValue("CardCode", 0).ToString();
                            string Nombre = oDataTable.GetValue("CardName", 0).ToString();

                            //((EditText)this.Form.Items.Item("txtDDocNum").Specific).Value = Codigo;
                            this.Form.DataSources.UserDataSources.Item(aoItemEvent.ItemUID.Replace("txt", "ds")).ValueEx = Codigo;
                            this.Form.DataSources.UserDataSources.Item($"{aoItemEvent.ItemUID.Replace("txt", "ds")}Nom").ValueEx = Nombre;
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    break;
            }
        }

        private void ItemEvent_et_COMBO_SELECT(string asFormUID, ref SAPbouiCOM.ItemEvent aoItemEvent, ref bool abBubbleEvent)
        {
            try
            {
                this.Form.Freeze(true);

                string TipoDocumento = ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbTipDoc").Specific).Value;

            ((EditText)this.Form.Items.Item("txtDDocNum").Specific).Value = "";
            ((EditText)this.Form.Items.Item("txtHDocNum").Specific).Value = "";

            ((SAPbouiCOM.Item)this.Form.Items.Item("txtNumReg")).Click();

            switch (aoItemEvent.ItemUID)
            {
                case "cmbTipDoc":
                    if (!aoItemEvent.BeforeAction && aoItemEvent.ActionSuccess)
                    {
                        switch (TipoDocumento)
                        {
                            case "Todos":
                                this.Form.Items.Item("txtDDocNum").Enabled = false;
                                this.Form.Items.Item("txtHDocNum").Enabled = false;
                                break;
                            case "Factura Venta":
                                this.Form.Items.Item("txtDDocNum").Enabled = true;
                                this.Form.Items.Item("txtHDocNum").Enabled = true;
                                ((EditText)this.Form.Items.Item("txtDDocNum").Specific).ChooseFromListUID = "CFLDFV";
                                ((EditText)this.Form.Items.Item("txtHDocNum").Specific).ChooseFromListUID = "CFLHFV";
                                break;
                            case "Abono Venta":
                                this.Form.Items.Item("txtDDocNum").Enabled = true;
                                this.Form.Items.Item("txtHDocNum").Enabled = true;
                                ((EditText)this.Form.Items.Item("txtDDocNum").Specific).ChooseFromListUID = "CFLDBV";
                                ((EditText)this.Form.Items.Item("txtHDocNum").Specific).ChooseFromListUID = "CFLHBV";
                                break;
                        }
                    }
                    break;
            }
            }
            catch (Exception ex) { }
            finally
            {
                this.Form.Freeze(false);
            }
        }
        #endregion



        #region Funcionalidad Específica
        public void AccionesAlCargarFormulario(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                ((SAPbouiCOM.StaticText)this.Form.Items.Item("lblVersion").Specific).Caption = $"v. {System.Diagnostics.FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion}";

                ((SAPbouiCOM.Item)this.Form.Items.Item("FolderTi0")).Click();

                EstablecerValoresPorDefecto(oForm);
                ProcesoCargarEnvios(oForm);

                this.Form.Items.Item("txtDNomIC").LinkTo = "cmbEst";
                this.Form.Items.Item("txtHNomIC").LinkTo = "cmbEst";
                this.Form.Items.Item("txtTotReg").LinkTo = "cmbEst";
                this.Form.Items.Item("btnRec").LinkTo = "cmbEst";
                this.Form.Items.Item("btnGenTBAI").LinkTo = "cmbEst";
                this.Form.Items.Item("btnRes").LinkTo = "cmbEst";

                ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetEC")).Height = 130;
                ((SAPbouiCOM.Item)this.Form.Items.Item("grdDDetRC")).Height = 130;

                ((SAPbouiCOM.Item)this.Form.Items.Item("txtNumReg")).Click();
                this.Form.Items.Item("txtDDocNum").Enabled = false;
                this.Form.Items.Item("txtHDocNum").Enabled = false;
            }
            catch (Exception ex) { }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void EstablecerValoresPorDefecto(SAPbouiCOM.Form oForm)
        {
            CargarComboEstados(oForm);
            CargarComboRespuesta(oForm);
            //CargarComboGeneradoQR(oForm);
            CargarComboTipoDocumento(oForm);

            ((SAPbouiCOM.EditText)this.Form.Items.Item("txtDesFec").Specific).Value = (new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-1)).ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)this.Form.Items.Item("txtHasFec").Specific).Value = (new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month))).ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)this.Form.Items.Item("txtNumReg").Specific).Value = "1000";
            ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbEst").Specific).Select("Enviado Conector - No", BoSearchKey.psk_ByValue);
            ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbRes").Specific).Select("Respuesta Conector - Todos", BoSearchKey.psk_ByValue);
            //((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbGenQR").Specific).Select("Generado QR - Todos", BoSearchKey.psk_ByValue);
            ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbTipDoc").Specific).Select("Todos", BoSearchKey.psk_ByValue);
        }

        public void CargarComboEstados(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oComboBox = null;

                oComboBox = (SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbEst").Specific;
                oComboBox.ValidValues.Add("Enviado Conector - Sí", "Enviado Conector - Sí");
                oComboBox.ValidValues.Add("Enviado Conector - No", "Enviado Conector - No");
                oComboBox.ValidValues.Add("Enviado Conector - Todos", "Enviado Conector - Todos");
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        public void CargarComboRespuesta(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oComboBox = null;

                oComboBox = (SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbRes").Specific;
                oComboBox.ValidValues.Add("Respuesta Conector - Sí", "Respuesta Conector - Sí");
                oComboBox.ValidValues.Add("Respuesta Conector - No", "Respuesta Conector - No");
                oComboBox.ValidValues.Add("Respuesta Conector - Todos", "Respuesta Conector - Todos");
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        public void CargarComboGeneradoQR(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oComboBox = null;

                oComboBox = (SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbGenQR").Specific;
                oComboBox.ValidValues.Add("Generado QR - Sí", "Generado QR - Sí");
                oComboBox.ValidValues.Add("Generado QR - No", "Generado QR - No");
                oComboBox.ValidValues.Add("Generado QR - Todos", "Generado QR - Todos");
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        public void CargarComboTipoDocumento(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.ComboBox oComboBox = null;

                oComboBox = (SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbTipDoc").Specific;
                oComboBox.ValidValues.Add("Factura Venta", "Factura Venta");
                oComboBox.ValidValues.Add("Abono Venta", "Abono Venta");
                oComboBox.ValidValues.Add("Todos", "Todos");
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        public void ProcesoCargarEnvios(SAPbouiCOM.Form oForm)
        {
            try
            {
                CargarEnvios(oForm);

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDoc").Specific;

                oGrid.Rows.SelectedRows.Add(0);
                int FilaSeleccionada = oGrid.GetDataTableRowIndex(oGrid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder));
                string Code = oGrid.DataTable.Columns.Item("Code").Cells.Item(FilaSeleccionada).Value.ToString();

                CargarDetalleDocumentoEnvioAConector(oForm, Code);
                CargarDetalleDocumentoRespuestaDeConector(oForm, Code);
            }
            catch (Exception ex)
            {
                Main.Log.write("CargarEnvios" + ex.ToString());
                throw ex;
            }
        }

        private string EstablecerCadenaWhere(SAPbouiCOM.Form oForm)
        {
            string CadenaWhere = "WHERE 1 = 1 ";

            string Estado = ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbEst").Specific).Value;
            string Respuesta = ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbRes").Specific).Value;
            //string GeneradoQR = ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbGenQR").Specific).Value;
            string TipoDocumento = ((SAPbouiCOM.ComboBox)this.Form.Items.Item("cmbTipDoc").Specific).Value;
            string DesdeFecha = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtDesFec").Specific).Value;
            string HastaFecha = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtHasFec").Specific).Value;
            string DesdeDocNum  = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtDDocNum").Specific).Value;
            string HastaDocNum  = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtHDocNum").Specific).Value;
            string DesdeIC = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtDIC").Specific).Value;
            string HastaIC = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtHIC").Specific).Value;

            switch (Estado)
            {
                case "Enviado Conector - Sí":
                    CadenaWhere += " AND K10_ET.\"U_EnvCon\" = 'Y' ";
                    break;
                case "Enviado Conector - No":
                    CadenaWhere += " AND K10_ET.\"U_EnvCon\" = 'N' ";
                    break;
            }
            switch (Respuesta)
            {
                case "Respuesta Conector - Sí":
                    CadenaWhere += " AND K10_ET.\"U_ResCon\" = 'Y' ";
                    break;
                case "Respuesta Conector - No":
                    CadenaWhere += " AND K10_ET.\"U_ResCon\" = 'N' ";
                    break;
            }
            //switch (GeneradoQR)
            //{
            //    case "Generado QR - Sí":
            //        CadenaWhere += " AND K10_ET.\"U_GenQR\" = 'Y' ";
            //        break;
            //    case "Generado QR - No":
            //        CadenaWhere += " AND K10_ET.\"U_GenQR\" = 'N' ";
            //        break;
            //}
            switch (TipoDocumento)
            {
                case "Factura Venta":
                    CadenaWhere += " AND T0.\"ObjType\" = '13' ";
                    break;
                case "Abono Venta":
                    CadenaWhere += " AND T0.\"ObjType\" = '14' ";
                    break;
            }
            if (!String.IsNullOrWhiteSpace(DesdeFecha))
            {
                //switch (csVG.oTipoBBDD)
                //{
                //	case BoDataServerTypes.dst_HANADB:
                CadenaWhere += $@" AND T0.""DocDate"" >= '{DesdeFecha}' ";
                //		break;
                //	default:
                //		CadenaWhere += $" AND T0.DocDate >= '{DesdeFecha}' ";
                //		break;
                //}
            }
            if (!String.IsNullOrWhiteSpace(HastaFecha))
            {
                //switch (csVG.oTipoBBDD)
                //{
                //	case BoDataServerTypes.dst_HANADB:
                CadenaWhere += $@" AND T0.""DocDate"" <= '{HastaFecha}' ";
                //		break;
                //	default:
                //		CadenaWhere += $" AND T0.DocDate <= '{HastaFecha}' ";
                //		break;
                //}
            }
            if (!String.IsNullOrWhiteSpace(DesdeDocNum))
            {
                CadenaWhere += $@" AND T0.""DocNum"" >= '{DesdeDocNum}' ";
            }
            if (!String.IsNullOrWhiteSpace(HastaDocNum))
            {
                CadenaWhere += $@" AND T0.""DocNum"" <= '{HastaDocNum}' ";
            }
            if (!String.IsNullOrWhiteSpace(DesdeIC))
            {
                CadenaWhere += $@" AND T0.""CardCode"" >= '{DesdeIC}' ";
            }
            if (!String.IsNullOrWhiteSpace(HastaIC))
            {
                CadenaWhere += $@" AND T0.""CardCode"" <= '{HastaIC}' ";
            }

            return CadenaWhere;
        }

        private void CargarEnvios(SAPbouiCOM.Form oForm)
        {
            try
            {
                string ColumnasConsulta = $@"
                    , K10_ET.""Code""
	                , ""NNM1"".""SeriesName"" AS ""Serie""
	                , K10_ET.""U_DocNum"" AS ""Nº Documento""
                    , T0.""NumAtCard"" AS ""Nº Referencia""
                    , T0.""DocDate"" AS ""Fecha Documento""
                    , T0.""DocTotal"" AS ""Total Documento""
	                , K10_ET.""U_CardCode"" AS ""Código IC""
	                , ""OCRD"".""CardName"" AS ""Nombre IC""
	                , K10_ET.""U_EnvCon"" AS ""Enviado Conector""
	                , K10_ET.""U_ResEnvCon"" AS ""Resultado Envío Conector""
                    --, K10_ET.""U_GenQR"" AS ""Generación QR""
	                --, K10_ET.""U_ResGenQR"" AS ""Resultado Respuesta Generación QR""
                    , K10_ET.""U_ResCon"" AS ""Respuesta Conector""
	                , K10_ET.""U_ResResCon"" AS ""Resultado Respuesta Conector""
                ";

                Main.Log.write("antes cadena");
                string CadenaWhere = EstablecerCadenaWhere(oForm);
                Main.Log.write(CadenaWhere);
                string Consulta = $@"
                    SELECT _NumeroLineas_ T.*
                    FROM (
                    SELECT 
                        'Factura Venta' AS ""Tipo Documento""
                    {ColumnasConsulta}
                    FROM ""@AX_ESTTBAI_0"" K10_ET
                    INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = K10_ET.""U_CardCode""
                    INNER JOIN ""OINV"" T0 ON Right(K10_ET.""Code"", 2) = '13' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#13')
                    INNER JOIN ""NNM1"" ""NNM1"" ON ""NNM1"".""Series"" = T0.""Series"" AND ""NNM1"".""ObjectCode"" = 13
                    {CadenaWhere}
                    UNION
                    SELECT 
                        'Abono Venta' AS ""Tipo Documento""
                    {ColumnasConsulta}
                    FROM ""@AX_ESTTBAI_0"" K10_ET
                    INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = K10_ET.""U_CardCode""
                    INNER JOIN ""ORIN"" T0 ON Right(K10_ET.""Code"", 2) = '14' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#14')
                    INNER JOIN ""NNM1"" ""NNM1"" ON ""NNM1"".""Series"" = T0.""Series"" AND ""NNM1"".""ObjectCode"" = 14
                    {CadenaWhere}
                    ) T
                    ORDER BY T.""Nº Documento"" DESC
                ";
                Main.Log.write(Consulta.Replace("_NumeroLineas_", ""));
                oForm.DataSources.DataTables.Item("dtDoc").ExecuteQuery(Consulta.Replace("_NumeroLineas_", ""));

                Main.Log.write("paso 1");
                if (oForm.DataSources.DataTables.Item("dtDoc").IsEmpty)
                {
                    ((SAPbouiCOM.EditText)this.Form.Items.Item("txtTotReg").Specific).Value = $"/0";
                }
                else
                {
                    ((SAPbouiCOM.EditText)this.Form.Items.Item("txtTotReg").Specific).Value = $"/{oForm.DataSources.DataTables.Item("dtDoc").Rows.Count}";
                }
                Main.Log.write("paso 2");
                oForm.DataSources.DataTables.Item("dtDoc").ExecuteQuery(Consulta.Replace("_NumeroLineas_", $"TOP {((SAPbouiCOM.EditText)this.Form.Items.Item("txtNumReg").Specific).Value}"));

                Main.Log.write(Consulta.Replace("_NumeroLineas_", $"TOP {((SAPbouiCOM.EditText)this.Form.Items.Item("txtNumReg").Specific).Value}"));
                Main.Log.write("paso 2");
                Main.Log.write("Filas encontradas " + oForm.DataSources.DataTables.Item("dtDoc").Rows.Count.ToString());
                //SAPbouiCOM.Grid oGrid = csK10_SBOGrid.EstablecerDataTableAGrid("grdDoc", oForm, "dtDoc", 0);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDoc").Specific;
                oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDoc");

                //oGrid.AñadirLinkedObjectAColumnaGrid(oGrid, "Código IC", csK10_UI.K10_SAP_ObjectType.BusinessPartners);
                //csK10_SBOGrid.AñadirLinkedObjectAColumnaGrid(oGrid, "Nº Documento", csK10_UI.K10_SAP_ObjectType.BusinessPartners);
                ((EditTextColumn)oGrid.Columns.Item("Código IC")).LinkedObjectType = "2";
                ((EditTextColumn)oGrid.Columns.Item("Nº Documento")).LinkedObjectType = "13";

                oGrid.Columns.Item("Enviado Conector").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("Respuesta Conector").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                //oGrid.Columns.Item("Generación QR").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                oGrid.Columns.Item("Total Documento").RightJustified = true;

                oGrid.Columns.Item("Code").Visible = false;

                oForm.Items.Item("grdDoc").Enabled = false;

                oGrid.AutoResizeColumns();

                //csK10_SBOGrid.NumerarFilasGrid(oGrid);
            }
            catch (Exception ex)
            {
                Main.Log.write(ex.ToString());
            }
            finally
            {

            }
        }

        private void CargarDetalleDocumentoEnvioAConector(SAPbouiCOM.Form oForm, string Code)
        {
            try
            {
                string Consulta = $@"
SELECT 
	T1.""Code""
    , T1.""U_Accion"" AS ""Acción""
	, T1.""U_NumInt"" AS ""Nº Intento Envío""
	, T1.""U_Info"" AS ""Información""
    , T1.""U_FecHor"" AS ""Fecha/Hora""
    , T1.""U_JSONDoc"" AS ""Documento JSON""
FROM ""@AX_ESTTBAI_1"" T1
WHERE T1.""Code"" Like '{Code}#%'
ORDER BY T1.""U_NumInt"" DESC
";

                oForm.DataSources.DataTables.Item("dtDDetEC").ExecuteQuery(Consulta);

                //SAPbouiCOM.Grid oGrid = csK10_SBOGrid.EstablecerDataTableAGrid("grdDDetEC", oForm, "dtDDetEC", 0);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDDetEC").Specific;
                oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDDetEC");

                oForm.Items.Item("grdDDetEC").Enabled = false;

                oGrid.Columns.Item("Code").Visible = false;
                oGrid.Columns.Item("Acción").Visible = false;

                oGrid.AutoResizeColumns();

                //csK10_SBOGrid.NumerarFilasGrid(oGrid);

                ColorearFilasGrid(oGrid);
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        private void CargarDetalleDocumentoRespuestaDeConector(SAPbouiCOM.Form oForm, string Code)
        {
            try
            {
                string Consulta = $@"
SELECT 
	T2.""Code""
    , T2.""U_Accion"" AS ""Acción""
	, T2.""U_NumInt"" AS ""Nº Intento Respuesta""
	, T2.""U_Info"" AS ""Información""
    , T2.""U_FecHor"" AS ""Fecha/Hora""
FROM ""@AX_ESTTBAI_2"" T2
WHERE T2.""Code"" Like '{Code}#%'
ORDER BY T2.""U_NumInt"" DESC
";

                oForm.DataSources.DataTables.Item("dtDDetRC").ExecuteQuery(Consulta);

                //SAPbouiCOM.Grid oGrid = csK10_SBOGrid.EstablecerDataTableAGrid("grdDDetRC", oForm, "dtDDetRC", 0);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)this.Form.Items.Item("grdDDetRC").Specific;
                oGrid.DataTable = oForm.DataSources.DataTables.Item("dtDDetRC");

                oForm.Items.Item("grdDDetRC").Enabled = false;

                oGrid.Columns.Item("Code").Visible = false;
                oGrid.Columns.Item("Acción").Visible = false;

                oGrid.AutoResizeColumns();

                //csK10_SBOGrid.NumerarFilasGrid(oGrid);

                ColorearFilasGrid(oGrid);
            }
            catch (Exception ex) { }
            finally
            {

            }
        }

        private void ColorearFilasGrid(SAPbouiCOM.Grid oGrid)
        {
            int ColorFila = 0;
            for (int Fila = 0; Fila <= oGrid.Rows.Count; Fila++)
            {
                switch (oGrid.DataTable.GetValue("Acción", oGrid.GetDataTableRowIndex(Fila)).ToString())
                {
                    case "E_C":
                    case "R_C":
                    case "Q_C":
                        ColorFila = 32768;
                        break;
                    case "E_F":
                    case "R_F":
                    case "Q_F":
                        ColorFila = 255;
                        break;
                    default:
                        ColorFila = 0; // csK10_SBOFormatear.ColorSAP_Gris_GridBackground;
                        break;
                }

                //oGrid.CommonSetting.SetRowBackColor(Fila + 1, ColorFila);//row numbering in commonsettings is different
                oGrid.CommonSetting.SetRowFontColor(Fila + 1, ColorFila);
            }
        }


        //        private void EnviarAConector()
        //        {
        //            SAPbobsCOM.Recordset oRecordset = null;
        //            string ColumnasConsulta = $@"
        //T0.""DocEntry""
        //	, T0.""ObjType""
        //	, T0.""DocNum""
        //	, T0.""CardCode""
        //	, T0.""CardName""
        //	, T0.""DocDate""
        //	, T0.""NumAtCard""
        //	, T0.""DocType""
        //--	, ""CRD1"".""Country""
        //--	, ""CRD1"".""County""
        //--	, ""CRD1"".""City""
        //--	, ""CRD1"".""Address""
        //--	, ""CRD1"".""ZipCode""
        //	, T12.""CountryB"" AS ""Country""
        //	, T12.""CountyB"" AS ""County""
        //	, T12.""CityB"" AS ""City""
        //	, T12.""StreetB"" AS ""Address""
        //	, T12.""ZipCodeB"" AS ""ZipCode""
        //	, T0.""U_B1SYS_INV_TYPE""
        //--	, T0.""DocTotal""
        //    , T0.""DocTotal"" - T0.""VatSum"" - T0.""DiscSum"" + T0.""WTSum"" - T0.""RoundDif"" - T0.""TotalExpns"" AS ""DocTotal""
        //	, T0.""DocCur""
        //	, Replace(T0.""LicTradNum"", 'ES', '') AS ""LicTradNum""
        //    , K10_ET.""Code"" AS ""CodigoEnvio""
        //	, TS.""SeriesName"" AS ""AX_SAPSerieName""
        //    , T0.""DocNum"" AS ""AX_DocNum""
        //    --, 'KD_10' AS ""AX_ClaveRegimenIvaOpTrascendencia""
        //";

        //            string Consulta = $@"
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""OINV"" T0
        //INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = T0.""CardCode""
        //LEFT OUTER JOIN ""CRD1"" ON ""CRD1"".""CardCode"" = ""OCRD"".""CardCode"" AND ""CRD1"".""AdresType"" = 'B' AND ""CRD1"".""Address"" = ""OCRD"".""BillToDef""
        //INNER JOIN [@AX_ESTTBAI_0] K10_ET ON Right(K10_ET.""Code"", 2) = '13' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#13') AND K10_ET.""U_EnvCon"" = 'N'
        //INNER JOIN ""INV12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN ""NNM1"" TS ON TS.""Series"" = T0.""Series""
        //UNION
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""ORIN"" T0
        //INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = T0.""CardCode""
        //LEFT OUTER JOIN ""CRD1"" ON ""CRD1"".""CardCode"" = ""OCRD"".""CardCode"" AND ""CRD1"".""AdresType"" = 'B' AND ""CRD1"".""Address"" = ""OCRD"".""BillToDef""
        //INNER JOIN [@AX_ESTTBAI_0] K10_ET ON Right(K10_ET.""Code"", 2) = '14' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#14') AND K10_ET.""U_EnvCon"" = 'N'
        //INNER JOIN ""RIN12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN ""NNM1"" TS ON TS.""Series"" = T0.""Series""
        //";

        //            oRecordset = (SAPbobsCOM.Recordset)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            oRecordset.DoQuery(Consulta);

        //            oRecordset.MoveFirst();

        //            List<csDocumentos> ListaDocumentos = new List<csDocumentos>();

        //            while (!oRecordset.EoF)
        //            {
        //                ListaDocumentos.Add(AñadirADocumento(oRecordset));

        //                oRecordset.MoveNext();
        //            }
        //            //csListaDocumentos ListaDocumentosJSON = new csListaDocumentos();
        //            //ListaDocumentosJSON.ListaDocumentos = ListaDocumentos;

        //            //string Documento_JSON = Newtonsoft.Json.JsonConvert.SerializeObject(ListaDocumentosJSON, Newtonsoft.Json.Formatting.None);
        //            foreach (csDocumentos Documento in ListaDocumentos)
        //            {
        //                List<string> ListaErrores = ValidarDocumentoEnvioConector(Documento);

        //                string FechaHora = "";
        //                int NumeroIntentoEnvio = 0;
        //                int cont = 1;

        //                ObtenerDatosIntentoEnvio(Documento.CodigoEnvio, ref FechaHora, ref NumeroIntentoEnvio);

        //                AñadirDetalleEnvioAConector(Documento.CodigoEnvio, FechaHora, NumeroIntentoEnvio, Documento, ref ListaErrores, ref cont);
        //                AñadirDetalleActualizarEnConector(Documento.CodigoEnvio, FechaHora, NumeroIntentoEnvio, Documento, ref ListaErrores, ref cont);
        //                //if (ListaErrores.Count == 0)
        //                //{

        //                //}
        //            }
        //        }

        //        private void GenerarQR()
        //        {
        //            SAPbobsCOM.Recordset oRecordset = null;
        //            string ColumnasConsulta = $@"
        //T0.""DocEntry""
        //	, T0.""ObjType""
        //	, T0.""DocNum""
        //	, T0.""CardCode""
        //	, T0.""CardName""
        //	, T0.""DocDate""
        //	, T0.""NumAtCard""
        //	, T0.""DocType""
        //--	, ""CRD1"".""Country""
        //--	, ""CRD1"".""County""
        //--	, ""CRD1"".""City""
        //--	, ""CRD1"".""Address""
        //--	, ""CRD1"".""ZipCode""
        //	, T12.""CountryB"" AS ""Country""
        //	, T12.""CountyB"" AS ""County""
        //	, T12.""CityB"" AS ""City""
        //	, T12.""StreetB"" AS ""Address""
        //	, T12.""ZipCodeB"" AS ""ZipCode""
        //	, T0.""U_B1SYS_INV_TYPE""
        //--	, T0.""DocTotal""
        //    , T0.""DocTotal"" - T0.""VatSum"" - T0.""DiscSum"" + T0.""WTSum"" - T0.""RoundDif"" - T0.""TotalExpns"" AS ""DocTotal""
        //	, T0.""DocCur""
        //	, Replace(T0.""LicTradNum"", 'ES', '') AS ""LicTradNum""
        //    , K10_ET.""Code"" AS ""CodigoEnvio""
        //	, TS.""SeriesName"" AS ""AX_SAPSerieName""
        //    , T0.""DocNum"" AS ""AX_DocNum"" 
        //    --, 'KD_10' AS ""AX_ClaveRegimenIvaOpTrascendencia""
        //";

        //            string Consulta = $@"
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""OINV"" T0
        //INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = T0.""CardCode""
        //LEFT OUTER JOIN ""CRD1"" ON ""CRD1"".""CardCode"" = ""OCRD"".""CardCode"" AND ""CRD1"".""AdresType"" = 'B' AND ""CRD1"".""Address"" = ""OCRD"".""BillToDef""
        //INNER JOIN [@AX_ESTTBAI_0] K10_ET ON Right(K10_ET.""Code"", 2) = '13' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#13') AND K10_ET.""U_EnvCon"" = 'Y' AND K10_ET.""U_GenQR"" = 'N'
        //INNER JOIN ""INV12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN ""NNM1"" TS ON TS.""Series"" = T0.""Series""
        //UNION
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""ORIN"" T0
        //INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = T0.""CardCode""
        //LEFT OUTER JOIN ""CRD1"" ON ""CRD1"".""CardCode"" = ""OCRD"".""CardCode"" AND ""CRD1"".""AdresType"" = 'B' AND ""CRD1"".""Address"" = ""OCRD"".""BillToDef""
        //INNER JOIN [@AX_ESTTBAI_0] K10_ET ON Right(K10_ET.""Code"", 2) = '14' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#14') AND K10_ET.""U_EnvCon"" = 'Y' AND K10_ET.""U_GenQR"" = 'N'
        //INNER JOIN ""RIN12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN ""NNM1"" TS ON TS.""Series"" = T0.""Series""
        //";

        //            oRecordset = (SAPbobsCOM.Recordset)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            oRecordset.DoQuery(Consulta);

        //            oRecordset.MoveFirst();

        //            List<csDocumentos> ListaDocumentos = new List<csDocumentos>();

        //            while (!oRecordset.EoF)
        //            {
        //                ListaDocumentos.Add(AñadirADocumento(oRecordset));

        //                oRecordset.MoveNext();
        //            }
        //            //csListaDocumentos ListaDocumentosJSON = new csListaDocumentos();
        //            //ListaDocumentosJSON.ListaDocumentos = ListaDocumentos;

        //            //string Documento_JSON = Newtonsoft.Json.JsonConvert.SerializeObject(ListaDocumentosJSON, Newtonsoft.Json.Formatting.None);
        //            foreach (csDocumentos Documento in ListaDocumentos)
        //            {
        //                List<string> ListaErrores = new List<string>();

        //                string FechaHora = "";
        //                int NumeroIntentoEnvio = 0;
        //                int cont = 1;

        //                ObtenerDatosIntentoEnvio(Documento.CodigoEnvio, ref FechaHora, ref NumeroIntentoEnvio);

        //                //AñadirDetalleEnvioAConector(Documento.CodigoEnvio, FechaHora, NumeroIntentoEnvio, Documento, ref ListaErrores, ref cont);
        //                AñadirDetalleActualizarEnConector(Documento.CodigoEnvio, FechaHora, NumeroIntentoEnvio, Documento, ref ListaErrores, ref cont);
        //            }
        //        }

        //private csDocumentos AñadirADocumento(SAPbobsCOM.Recordset oRecordset)
        //{
        //    string InvoiceType = "";

        //    switch (oRecordset.Fields.Item("ObjType").Value.ToString())
        //    {
        //        case "13":
        //        case "14":
        //            InvoiceType = "Sales";
        //            break;
        //    }

        //    List<csDocumentos_Lineas> DocumentoLinea = AñadirADocumentoLinea(oRecordset.Fields.Item("DocEntry").Value.ToString(), oRecordset.Fields.Item("ObjType").Value.ToString());

        //    csDocumentos Documento = new csDocumentos()
        //    {
        //        DocEntry = Convert.ToInt32(oRecordset.Fields.Item("DocEntry").Value.ToString()),
        //        ObjType = oRecordset.Fields.Item("ObjType").Value.ToString(),
        //        DocNum = oRecordset.Fields.Item("DocNum").Value.ToString(),//oRecordset.Fields.Item("DocEntry").Value.ToString(),
        //        CardCode = oRecordset.Fields.Item("CardCode").Value.ToString(),
        //        CardName = oRecordset.Fields.Item("CardName").Value.ToString(),
        //        DocDate = Convert.ToDateTime(oRecordset.Fields.Item("DocDate").Value.ToString()),
        //        NumAtCard = oRecordset.Fields.Item("NumAtCard").Value.ToString(),
        //        DocType = oRecordset.Fields.Item("DocType").Value.ToString(),
        //        Country = oRecordset.Fields.Item("Country").Value.ToString(),
        //        Province = oRecordset.Fields.Item("County").Value.ToString(),
        //        City = oRecordset.Fields.Item("City").Value.ToString(),
        //        Address = oRecordset.Fields.Item("Address").Value.ToString(),
        //        ZipCode = oRecordset.Fields.Item("ZipCode").Value.ToString(),
        //        U_B1SYS_INV_TYPE = oRecordset.Fields.Item("U_B1SYS_INV_TYPE").Value.ToString(),
        //        DocTotal = Convert.ToDouble(oRecordset.Fields.Item("DocTotal").Value.ToString()),
        //        DocCur = oRecordset.Fields.Item("DocCur").Value.ToString(),
        //        LicTradNum = oRecordset.Fields.Item("LicTradNum").Value.ToString(),
        //        CodigoEnvio = oRecordset.Fields.Item("CodigoEnvio").Value.ToString(),
        //        DocEntryR = "0", //Pending contemplar casuistica facturas de rectificacion. Aqui la factrua rectificada
        //        Enterprise = EnterpriseAhoraERP,
        //        InvoiceType = InvoiceType,
        //        AX_SAPSerieName = oRecordset.Fields.Item("AX_SAPSerieName").Value.ToString(),
        //        AX_DocNum = oRecordset.Fields.Item("AX_DocNum").Value.ToString(),
        //        // AX_ClaveRegimenIvaOpTrascendencia = oRecordset.Fields.Item("AX_ClaveRegimenIvaOpTrascendencia").Value.ToString(),
        //        DocumentLines = DocumentoLinea,
        //        AX_ClaveRegimenIvaOpTrascendencia = DocumentoLinea[0].U_AX_TBAI_CodImp,
        //        AX_NombreOperacion = DocumentoLinea[0].NombreOperacion,
        //        AX_TBAI_FECasSuj = DocumentoLinea[0].U_AX_TBAI_FECasSuj,
        //        AX_TBAI_FECasExe = DocumentoLinea[0].U_AX_TBAI_FECasExe,
        //        AX_TBAI_FECauExe = DocumentoLinea[0].U_AX_TBAI_FECauExe,
        //        AX_TBAI_FETipNExe = DocumentoLinea[0].U_AX_TBAI_FETipNExe,
        //        AX_TBAI_FECauNSuj = DocumentoLinea[0].U_AX_TBAI_FECauNSuj,
        //        AX_TBAI_FEClaRegEsp = DocumentoLinea[0].U_AX_TBAI_FEClaRegEsp,
        //        AX_TBAI_FRClaRegEsp = DocumentoLinea[0].U_AX_TBAI_FRClaRegEsp,
        //        AX_TBAI_Ser = DocumentoLinea[0].U_AX_TBAI_Ser,
        //        AX_TBAI_InvSujPas = DocumentoLinea[0].U_AX_TBAI_InvSujPas,
        //        AX_TBAI_BieInv = DocumentoLinea[0].U_AX_TBAI_BieInv,
        //        AX_TBAI_DetOpeInt = DocumentoLinea[0].U_AX_TBAI_DetOpeInt,
        //        AX_TBAI_Imp = DocumentoLinea[0].U_AX_TBAI_Imp,
        //        AX_TBAI_FacSim = DocumentoLinea[0].U_AX_TBAI_FacSim,
        //        AX_TBAI_FacRegSim = DocumentoLinea[0].U_AX_TBAI_FacRegSim,
        //        AX_TBAI_FacRegRecEqu = DocumentoLinea[0].U_AX_TBAI_FacRegRecEqu,
        //        AX_TBAI_CasREAV = DocumentoLinea[0].U_AX_TBAI_CasREAV,
        //        AX_TBAI_FRCG240 = DocumentoLinea[0].U_AX_TBAI_FRCG240,
        //        AX_TBAI_Arr = DocumentoLinea[0].U_AX_TBAI_Arr,
        //        AX_TBAI_CasREGEAju = DocumentoLinea[0].U_AX_TBAI_CasREGEAju
        //    };

        //    return Documento;
        //}

        //        private List<csDocumentos_Lineas> AñadirADocumentoLinea(string DocEntry, string ObjType)
        //        {
        //            SAPbobsCOM.Recordset oRecordset = null;
        //            string ColumnasConsulta = $@"
        //T0.""DocEntry""
        //    , T0.""ObjType""
        //	, T1.""LineNum""
        //    , T1.""ItemCode""
        //	, T1.""Dscription""
        //	, T1.""Quantity""	
        //	, T1.""VatSum""
        //	, T1.""LineTotal""
        //	, T0.""ResidenNum""
        //	, TAX1.""AbsEntry""
        //	, TAX1.""LineSeq""
        //	, TAX1.""TaxCode""
        //	, TAX1.""VatPercent""
        //	, TAX1.""NdPercent""
        //	, TAX1.""EqPercent""
        //	, TAX1.""BaseSum""
        //	, TAX1.""VatSum""
        //	, TAX1.""DeductSum""
        //	, TAX1.""EqSum""
        //	, TAX1.""CrditDebit""
        //	, OVTG.""U_AX_TBAI_CodImp""
        //	, AX1.""Name"" AS ""NombreOperacion""
        //    , OVTG.""U_AX_TBAI_FECasSuj""
        //    , OVTG.""U_AX_TBAI_FECasExe""
        //    , OVTG.""U_AX_TBAI_FECauExe""
        //    , OVTG.""U_AX_TBAI_FETipNExe""
        //    , OVTG.""U_AX_TBAI_FECauNSuj""
        //    , OVTG.""U_AX_TBAI_FEClaRegEsp""
        //    , OVTG.""U_AX_TBAI_FRClaRegEsp""
        //    , OVTG.""U_AX_TBAI_Ser""
        //    , OVTG.""U_AX_TBAI_InvSujPas""
        //    , OVTG.""U_AX_TBAI_BieInv""
        //    , OVTG.""U_AX_TBAI_DetOpeInt""
        //    , OVTG.""U_AX_TBAI_Imp""
        //    , OVTG.""U_AX_TBAI_FacSim""
        //    , OVTG.""U_AX_TBAI_FacRegSim""
        //    , OVTG.""U_AX_TBAI_FacRegRecEqu""
        //    , OVTG.""U_AX_TBAI_CasREAV""
        //    , OVTG.""U_AX_TBAI_FRCG240""
        //    , OVTG.""U_AX_TBAI_Arr""
        //    , OVTG.""U_AX_TBAI_CasREGEAju""
        //";

        //            string Consulta = $@"
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""OINV"" T0
        //INNER JOIN ""INV1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN OTAX OTAX ON OTAX.SrcObjAbs = T0.""DocEntry"" and OTAX.""SrcObjType"" = T0.""ObjType""
        //INNER JOIN TAX1 TAX1 ON TAX1.""AbsEntry"" = OTAX.""AbsEntry"" AND TAX1.""SrcLineNum"" = T1.""LineNum""
        //INNER JOIN OVTG OVTG ON OVTG.""Code"" = TAX1.""TaxCode""
        //INNER JOIN [@AX_TBAIOPERACIONES] AX1 ON AX1.""Code"" = OVTG.""U_AX_TBAI_CodImp""
        //WHERE T0.""DocEntry"" = {DocEntry} AND T0.""ObjType"" = {ObjType}
        //UNION
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""ORIN"" T0
        //INNER JOIN ""RIN1"" T1 ON T1.""DocEntry"" = T0.""DocEntry""
        //INNER JOIN OTAX OTAX ON OTAX.SrcObjAbs = T0.""DocEntry"" and OTAX.""SrcObjType"" = T0.""ObjType""
        //INNER JOIN TAX1 TAX1 ON TAX1.""AbsEntry"" = OTAX.""AbsEntry"" AND TAX1.""SrcLineNum"" = T1.""LineNum""
        //INNER JOIN OVTG OVTG ON OVTG.""Code"" = TAX1.""TaxCode""
        //INNER JOIN [@AX_TBAIOPERACIONES] AX1 ON AX1.""Code"" = OVTG.""U_AX_TBAI_CodImp""
        //WHERE T0.""DocEntry"" = {DocEntry} AND T0.""ObjType"" = {ObjType}
        //";

        //            oRecordset = (SAPbobsCOM.Recordset)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            oRecordset.DoQuery(Consulta);

        //            oRecordset.MoveFirst();

        //            List<csDocumentos_Lineas> ListaDocumentos_Lineas = new List<csDocumentos_Lineas>();

        //            while (!oRecordset.EoF)
        //            {
        //                string Description = "";

        //                switch (oRecordset.Fields.Item("ObjType").Value.ToString())
        //                {
        //                    case "13":
        //                    case "14":
        //                        Description = "VENTA";
        //                        break;
        //                }

        //                ListaDocumentos_Lineas.Add(new csDocumentos_Lineas
        //                {
        //                    DocEntry = Convert.ToInt32(DocEntry),
        //                    LineNum = Convert.ToInt32(oRecordset.Fields.Item("LineNum").Value.ToString()),
        //                    Dscription = oRecordset.Fields.Item("Dscription").Value.ToString(),
        //                    ItemCode = oRecordset.Fields.Item("ItemCode").Value.ToString(),
        //                    Quantity = Convert.ToDouble(oRecordset.Fields.Item("Quantity").Value.ToString()),
        //                    VatSum = Convert.ToDouble(oRecordset.Fields.Item("VatSum").Value.ToString()),
        //                    LineTotal = Convert.ToDouble(oRecordset.Fields.Item("LineTotal").Value.ToString()),
        //                    AbsEntry = oRecordset.Fields.Item("AbsEntry").Value.ToString(),
        //                    LineSeq = oRecordset.Fields.Item("LineSeq").Value.ToString(),
        //                    VatPercent = Convert.ToDouble(oRecordset.Fields.Item("VatPercent").Value.ToString()),
        //                    NdPercent = Convert.ToDouble(oRecordset.Fields.Item("NdPercent").Value.ToString()),
        //                    baseSum = Convert.ToDouble(oRecordset.Fields.Item("BaseSum").Value.ToString()),
        //                    DeductSum = Convert.ToDouble(oRecordset.Fields.Item("DeductSum").Value.ToString()),
        //                    EqSum = Convert.ToDouble(oRecordset.Fields.Item("EqSum").Value.ToString()),
        //                    CrditDebit = oRecordset.Fields.Item("CrditDebit").Value.ToString(),
        //                    TaxCode = oRecordset.Fields.Item("TaxCode").Value.ToString(),
        //                    Description = Description,
        //                    EqPercent = Convert.ToDouble(oRecordset.Fields.Item("EqPercent").Value.ToString()),
        //                    ResidenNum = oRecordset.Fields.Item("ResidenNum").Value.ToString(),
        //                    U_AX_TBAI_CodImp = oRecordset.Fields.Item("U_AX_TBAI_CodImp").Value.ToString(),
        //                    NombreOperacion = oRecordset.Fields.Item("NombreOperacion").Value.ToString(),
        //                    U_AX_TBAI_FECasSuj = oRecordset.Fields.Item("U_AX_TBAI_FECasSuj").Value.ToString(),
        //                    U_AX_TBAI_FECasExe = oRecordset.Fields.Item("U_AX_TBAI_FECasExe").Value.ToString(),
        //                    U_AX_TBAI_FECauExe = oRecordset.Fields.Item("U_AX_TBAI_FECauExe").Value.ToString(),
        //                    U_AX_TBAI_FETipNExe = oRecordset.Fields.Item("U_AX_TBAI_FETipNExe").Value.ToString(),
        //                    U_AX_TBAI_FECauNSuj = oRecordset.Fields.Item("U_AX_TBAI_FECauNSuj").Value.ToString(),
        //                    U_AX_TBAI_FEClaRegEsp = oRecordset.Fields.Item("U_AX_TBAI_FEClaRegEsp").Value.ToString(),
        //                    U_AX_TBAI_FRClaRegEsp = oRecordset.Fields.Item("U_AX_TBAI_FRClaRegEsp").Value.ToString(),
        //                    U_AX_TBAI_Ser = oRecordset.Fields.Item("U_AX_TBAI_Ser").Value.ToString(),
        //                    U_AX_TBAI_InvSujPas = oRecordset.Fields.Item("U_AX_TBAI_InvSujPas").Value.ToString(),
        //                    U_AX_TBAI_BieInv = oRecordset.Fields.Item("U_AX_TBAI_BieInv").Value.ToString(),
        //                    U_AX_TBAI_DetOpeInt = oRecordset.Fields.Item("U_AX_TBAI_DetOpeInt").Value.ToString(),
        //                    U_AX_TBAI_Imp = oRecordset.Fields.Item("U_AX_TBAI_Imp").Value.ToString(),
        //                    U_AX_TBAI_FacSim = oRecordset.Fields.Item("U_AX_TBAI_FacSim").Value.ToString(),
        //                    U_AX_TBAI_FacRegSim = oRecordset.Fields.Item("U_AX_TBAI_FacRegSim").Value.ToString(),
        //                    U_AX_TBAI_FacRegRecEqu = oRecordset.Fields.Item("U_AX_TBAI_FacRegRecEqu").Value.ToString(),
        //                    U_AX_TBAI_CasREAV = oRecordset.Fields.Item("U_AX_TBAI_CasREAV").Value.ToString(),
        //                    U_AX_TBAI_FRCG240 = oRecordset.Fields.Item("U_AX_TBAI_FRCG240").Value.ToString(),
        //                    U_AX_TBAI_Arr = oRecordset.Fields.Item("U_AX_TBAI_Arr").Value.ToString(),
        //                    U_AX_TBAI_CasREGEAju = oRecordset.Fields.Item("U_AX_TBAI_CasREGEAju").Value.ToString()
        //                });

        //                oRecordset.MoveNext();
        //            }

        //            return ListaDocumentos_Lineas;
        //        }

        //private List<string> ValidarDocumentoEnvioConector(csDocumentos Documento)
        //{
        //    List<string> ListaErrores = new List<string>();

        //    if (String.IsNullOrWhiteSpace(Documento.LicTradNum))
        //    {
        //        ListaErrores.Add($"C.I.F. no está informado en documento nº ({Documento.DocNum})");
        //    }

        //    if (String.IsNullOrWhiteSpace(Documento.Country))
        //    {
        //        ListaErrores.Add($"País no está informado en documento nº ({Documento.DocNum})");
        //    }

        //    if (String.IsNullOrWhiteSpace(Documento.Address))
        //    {
        //        ListaErrores.Add($"Dirección de Factura no está informado en documento nº ({Documento.DocNum})");
        //    }

        //    return ListaErrores;
        //}

        //        private void AñadirDetalleEnvioAConector_o(List<string> ListaErrores, string Code, csDocumentos Documento)
        //        {
        //            Main.Log.write("AñadirDetalleEnvioAConector_o" + Code);
        //            string FechaHora = $"{DateTime.Today.ToString("dd/MM/yyyy")} {DateTime.Now.ToString("hh:mm")}";
        //            //            List<List<csK10_DI.csGuardarDatoTablaUsuario>> ListasGuardarDatoTablaUsuario = new List<List<csK10_DI.csGuardarDatoTablaUsuario>>();

        //            string Consulta = $@"
        //SELECT IsNull(Max(K10_ET1.""U_NumInt""), 0) + 1
        //FROM [@AX_ESTTBAI_1] K10_ET1
        //WHERE K10_ET1.""Code"" Like '{Code}#%'
        //";
        //            //            int NumeroIntentoEnvio = Convert.ToInt32(csK10_Utilidades.DameValor(oCompany, Consulta));
        //            int NumeroIntentoEnvio = Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString());

        //            string Documento_JSON = $@"{{""Documents"": [{Newtonsoft.Json.JsonConvert.SerializeObject(Documento, Newtonsoft.Json.Formatting.Indented)}]}}";

        //            if (ListaErrores.Count == 0)
        //            {
        //                //                List<csK10_DI.csGuardarDatoTablaUsuario> ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_FecHor", Valor = FechaHora });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_JSONDoc", Valor = Documento_JSON });

        //                string ErrorEnvioAConector = "";
        //                bool EnviadoAAhora = EnviarAAHORA(Documento_JSON, Documento, ref ErrorEnvioAConector); // InsertarDocumentosAhora(Documento_JSON, connection);

        //                string ErrorEnvioAlActualizarEnAHORA = "";
        //                if (EnviadoAAhora)
        //                {
        //                    if (TBAI_CreaCuentasParaPoderActualizarAHORA(Documento, ref ErrorEnvioAlActualizarEnAHORA))
        //                    {
        //                        bool Actualizada = TBAI_ActualizaDocumentoAHORA(Documento, ref ErrorEnvioAlActualizarEnAHORA);
        //                    }
        //                }

        //                //                if (EnviadoAAhora)
        //                //                {
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Accion", Valor = "E_C" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Info", Valor = "Enviado a conector correctamente" });
        //                //                }
        //                //                else
        //                //                {
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Accion", Valor = "E_F" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Info", Valor = $"Envío a conector fallido. {ErrorEnvioAConector}" });

        //                //                    ListaErrores.Add($"C.I.F. no está informado en documento nº ({Documento.DocNum})");
        //                //                }

        //                //                ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);
        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //                oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //                oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = Documento_JSON;
        //                if (EnviadoAAhora)
        //                {
        //                    oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_C";
        //                    oUserTable.UserFields.Fields.Item("U_Info").Value = "Enviado a conector correctamente";
        //                }
        //                else
        //                {
        //                    oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_F";
        //                    oUserTable.UserFields.Fields.Item("U_Info").Value = $"Envío a conector fallido. {ErrorEnvioAConector}";

        //                    ListaErrores.Add($"C.I.F. no está informado en documento nº ({Documento.DocNum})");
        //                }

        //                if (oUserTable.Add() != 0)
        //                {
        //                    string Error = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);
        //            }
        //            else
        //            {
        //                //                List<csK10_DI.csGuardarDatoTablaUsuario> ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Accion", Valor = "E_F" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Info", Valor = "Envío a conector fallido" });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_FecHor", Valor = FechaHora });
        //                //                ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_JSONDoc", Valor = Documento_JSON });

        //                //                ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);

        //                //                int cont = 1;
        //                //                foreach (string Error in ListaErrores)
        //                //                {
        //                //                    ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#{cont}" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#{cont}" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Accion", Valor = "" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_Info", Valor = Error });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_FecHor", Valor = "" });
        //                //                    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_1", Clave = "U_JSONDoc", Valor = "" });

        //                //                    ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);

        //                //                    cont++;
        //                //                }

        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //                oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //                oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = Documento_JSON;
        //                oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_F";
        //                oUserTable.UserFields.Fields.Item("U_Info").Value = "Envío a conector fallido";

        //                if (oUserTable.Add() != 0)
        //                {
        //                    string Error2 = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);

        //                int cont = 1;
        //                foreach (string Error in ListaErrores)
        //                {
        //                    oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //                    oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //                    oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //                    oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                    oUserTable.UserFields.Fields.Item("U_FecHor").Value = "";
        //                    oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = "";
        //                    oUserTable.UserFields.Fields.Item("U_Accion").Value = "";
        //                    oUserTable.UserFields.Fields.Item("U_Info").Value = Error;

        //                    if (oUserTable.Add() != 0)
        //                    {
        //                        string Error2 = Main.Company.GetLastErrorDescription();
        //                    }

        //                    Common.Functions.ReleaseComObject(oUserTable);

        //                    cont++;
        //                }
        //            }

        //            //            csK10_DI.DI_GuardarDatosTablaUsuario(ListasGuardarDatoTablaUsuario, null);



        //            //            List<List<csK10_DI.csActualizarDatoTablaUsuario>> ListasActualizarDatoTablaUsuario = new List<List<csK10_DI.csActualizarDatoTablaUsuario>>();

        //            //            if (ListaErrores.Count == 0)
        //            //            {
        //            //                List<csK10_DI.csActualizarDatoTablaUsuario> ListaActualizarDatoTablaUsuario = new List<csK10_DI.csActualizarDatoTablaUsuario>();
        //            //                ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "Code", Valor = $"{Code}" });
        //            //                ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "Name", Valor = $"{Code}" });
        //            //                ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "U_EnvCon", Valor = "Y" });
        //            //                ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "U_ResEnvCon", Valor = "Enviado a conector correctamente" });

        //            //                ListasActualizarDatoTablaUsuario.Add(ListaActualizarDatoTablaUsuario);
        //            //            }

        //            //            csK10_DI.DI_ActualizarDatosTablaUsuario(ListasActualizarDatoTablaUsuario, null);

        //            if (ListaErrores.Count == 0)
        //            {
        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_0");

        //                oUserTable.GetByKey(Code);

        //                oUserTable.UserFields.Fields.Item("U_EnvCon").Value = "Y";
        //                oUserTable.UserFields.Fields.Item("U_ResEnvCon").Value = "Enviado a conector correctamente";

        //                if (oUserTable.Update() != 0)
        //                {
        //                    string Error2 = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);
        //            }
        //        }

        //        private void ObtenerDatosIntentoEnvio(string Code, ref string FechaHora, ref int NumeroIntentoEnvio)
        //        {
        //            FechaHora = $"{DateTime.Today.ToString("dd/MM/yyyy")} {DateTime.Now.ToString("hh:mm")}";

        //            string Consulta = $@"
        //SELECT IsNull(Max(K10_ET1.""U_NumInt""), 0) + 1
        //FROM [@AX_ESTTBAI_1] K10_ET1
        //WHERE K10_ET1.""Code"" Like '{Code}#%'
        //";
        //            //            int NumeroIntentoEnvio = Convert.ToInt32(csK10_Utilidades.DameValor(oCompany, Consulta));
        //            NumeroIntentoEnvio = Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString());
        //        }

        //private void AñadirDetalleEnvioAConector(string Code, string FechaHora, int NumeroIntentoEnvio, csDocumentos Documento, ref List<string> ListaErrores, ref int cont)
        //{
        //    string Documento_JSON = $@"{{""Documents"": [{Newtonsoft.Json.JsonConvert.SerializeObject(Documento, Newtonsoft.Json.Formatting.Indented)}]}}";

        //    if (ListaErrores.Count == 0)
        //    {
        //        string ErrorEnvioAConector = "";
        //        Main.Log.write("AñadirDetalleEnvioAConector" + Code);
        //        bool EnviadoAAhora = EnviarAAHORA(Documento_JSON, Documento, ref ErrorEnvioAConector); // InsertarDocumentosAhora(Documento_JSON, connection);

        //        //string ErrorEnvioAlActualizarEnAHORA = "";
        //        //if (EnviadoAAhora)
        //        //{
        //        //    bool Actualizada = TBAI_ActualizaDocumentoAHORA(Documento, ref ErrorEnvioAlActualizarEnAHORA);
        //        //}

        //        SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //        oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //        oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //        oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //        oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //        oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = Documento_JSON;
        //        if (EnviadoAAhora)
        //        {
        //            oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_C";
        //            oUserTable.UserFields.Fields.Item("U_Info").Value = "Enviado a conector correctamente";
        //        }
        //        else
        //        {
        //            oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_F";
        //            oUserTable.UserFields.Fields.Item("U_Info").Value = $"Envío a conector fallido. {ErrorEnvioAConector}";

        //            ListaErrores.Add($"Envío fallido ({Documento.DocNum})");
        //        }

        //        if (oUserTable.Add() != 0)
        //        {
        //            string Error = Main.Company.GetLastErrorDescription();
        //        }

        //        Common.Functions.ReleaseComObject(oUserTable);
        //    }
        //    else
        //    {
        //        SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //        oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //        oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //        oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //        oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //        oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = Documento_JSON;
        //        oUserTable.UserFields.Fields.Item("U_Accion").Value = "E_F";
        //        oUserTable.UserFields.Fields.Item("U_Info").Value = "Envío a conector fallido";

        //        if (oUserTable.Add() != 0)
        //        {
        //            string Error2 = Main.Company.GetLastErrorDescription();
        //        }

        //        Common.Functions.ReleaseComObject(oUserTable);

        //        //int cont = 1;
        //        foreach (string Error in ListaErrores)
        //        {
        //            oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //            oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //            oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //            oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //            oUserTable.UserFields.Fields.Item("U_FecHor").Value = "";
        //            oUserTable.UserFields.Fields.Item("U_JSONDoc").Value = "";
        //            oUserTable.UserFields.Fields.Item("U_Accion").Value = "";
        //            oUserTable.UserFields.Fields.Item("U_Info").Value = Error;

        //            if (oUserTable.Add() != 0)
        //            {
        //                string Error2 = Main.Company.GetLastErrorDescription();
        //            }

        //            Common.Functions.ReleaseComObject(oUserTable);

        //            cont++;
        //        }
        //    }


        //    if (ListaErrores.Count == 0)
        //    {
        //        SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_0");

        //        oUserTable.GetByKey(Code);

        //        oUserTable.UserFields.Fields.Item("U_EnvCon").Value = "Y";
        //        oUserTable.UserFields.Fields.Item("U_ResEnvCon").Value = "Enviado a conector correctamente";

        //        if (oUserTable.Update() != 0)
        //        {
        //            string Error2 = Main.Company.GetLastErrorDescription();
        //        }

        //        Common.Functions.ReleaseComObject(oUserTable);
        //    }
        //}

        //private void AñadirDetalleActualizarEnConector(string Code, string FechaHora, int NumeroIntentoEnvio, csDocumentos Documento, ref List<string> ListaErrores, ref int cont)
        //{
        //    //            string FechaHora = $"{DateTime.Today.ToString("dd/MM/yyyy")} {DateTime.Now.ToString("hh:mm")}";
        //    //            //            List<List<csK10_DI.csGuardarDatoTablaUsuario>> ListasGuardarDatoTablaUsuario = new List<List<csK10_DI.csGuardarDatoTablaUsuario>>();

        //    //            string Consulta = $@"
        //    //SELECT IsNull(Max(K10_ET1.""U_NumInt""), 0) + 1
        //    //FROM [@AX_ESTTBAI_1] K10_ET1
        //    //WHERE K10_ET1.""Code"" Like '{Code}#%'
        //    //";
        //    //            //            int NumeroIntentoEnvio = Convert.ToInt32(csK10_Utilidades.DameValor(oCompany, Consulta));
        //    //            int NumeroIntentoEnvio = Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString());

        //    //string Documento_JSON = $@"{{""Documents"": [{Newtonsoft.Json.JsonConvert.SerializeObject(Documento, Newtonsoft.Json.Formatting.Indented)}]}}";

        //    if (ListaErrores.Count == 0)
        //    {
        //        //string ErrorEnvioAConector = "";
        //        //bool EnviadoAAhora = EnviarAAHORA(Documento_JSON, Documento, ref ErrorEnvioAConector); // InsertarDocumentosAhora(Documento_JSON, connection);

        //        string ErrorEnvioAlActualizarEnAHORA = "";
        //        //bool Actualizada = false;
        //        //if (EnviadoAAhora)
        //        //{
        //        bool Actualizada = false;
        //        if (TBAI_CreaCuentasParaPoderActualizarAHORA(Documento, ref ErrorEnvioAlActualizarEnAHORA))
        //        {
        //            Actualizada = TBAI_ActualizaDocumentoAHORA(Documento, ref ErrorEnvioAlActualizarEnAHORA);
        //        }
        //        //}

        //        SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_1");

        //        oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //        oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //        oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //        oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //        if (Actualizada ||
        //            ErrorEnvioAlActualizarEnAHORA.PadRight(31, '0').Substring(0, 31) == "La factura ya está actualizada.")
        //        {
        //            oUserTable.UserFields.Fields.Item("U_Accion").Value = "Q_C";
        //            oUserTable.UserFields.Fields.Item("U_Info").Value = $"Generación QR correcta. {ErrorEnvioAlActualizarEnAHORA}";
        //        }
        //        else
        //        {
        //            oUserTable.UserFields.Fields.Item("U_Accion").Value = "Q_F";
        //            oUserTable.UserFields.Fields.Item("U_Info").Value = $"Generación QR fallida. {ErrorEnvioAlActualizarEnAHORA}";

        //            ListaErrores.Add($"Generación QR en documento nº ({Documento.DocNum})");
        //        }

        //        if (oUserTable.Add() != 0)
        //        {
        //            string Error = Main.Company.GetLastErrorDescription();
        //        }

        //        Common.Functions.ReleaseComObject(oUserTable);
        //    }

        //    if (ListaErrores.Count == 0)
        //    {
        //        SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_0");

        //        oUserTable.GetByKey(Code);

        //        oUserTable.UserFields.Fields.Item("U_GenQR").Value = "Y";
        //        oUserTable.UserFields.Fields.Item("U_ResGenQR").Value = "Generación QR correcta";

        //        if (oUserTable.Update() != 0)
        //        {
        //            string Error2 = Main.Company.GetLastErrorDescription();
        //        }

        //        Common.Functions.ReleaseComObject(oUserTable);
        //    }
        //}

        //        private List<string> ObtenerListaParaRespuestaDeConector()
        //        {
        //            SAPbobsCOM.Recordset oRecordset = null;
        //            string ColumnasConsulta = $@"
        //K10_ET_0.Code
        //	, T0.""DocEntry""
        //	, T0.""ObjType""
        //	, T0.""DocNum""
        //	, ""NNM1"".""SeriesName"" AS ""Serie""
        //";

        //            string Consulta = $@"
        //SELECT {ColumnasConsulta}
        //FROM [@AX_ESTTBAI_0] K10_ET_0
        //INNER JOIN ""OINV"" T0 ON Right(K10_ET_0.""Code"", 2) = '13' AND K10_ET_0.""Code"" Like CONCAT(T0.""DocEntry"", '#13') AND K10_ET_0.""U_EnvCon"" = 'Y'
        //INNER JOIN ""NNM1"" ""NNM1"" ON ""NNM1"".""Series"" = T0.""Series"" AND ""NNM1"".""ObjectCode"" = 13
        //WHERE (T0.""U_AX_CodQR"" Is null OR CAST(T0.""U_AX_CodQR"" AS nvarchar(MAX)) = '') AND (K10_ET_0.""U_ResResCon"" Is null OR CAST(K10_ET_0.""U_ResResCon"" AS nvarchar(MAX)) = '')
        //UNION
        //SELECT {ColumnasConsulta}
        //FROM [@AX_ESTTBAI_0] K10_ET_0
        //INNER JOIN ""ORIN"" T0 ON Right(K10_ET_0.""Code"", 2) = '14' AND K10_ET_0.""Code"" Like CONCAT(T0.""DocEntry"", '#14') AND K10_ET_0.""U_EnvCon"" = 'Y'
        //INNER JOIN ""NNM1"" ""NNM1"" ON ""NNM1"".""Series"" = T0.""Series"" AND ""NNM1"".""ObjectCode"" = 14
        //WHERE (T0.""U_AX_CodQR"" Is null OR CAST(T0.""U_AX_CodQR"" AS nvarchar(MAX)) = '') AND (K10_ET_0.""U_ResResCon"" Is null OR CAST(K10_ET_0.""U_ResResCon"" AS nvarchar(MAX)) = '')
        //";

        //            oRecordset = (SAPbobsCOM.Recordset)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            oRecordset.DoQuery(Consulta);

        //            oRecordset.MoveFirst();

        //            List<string> ListaDocumentos = new List<string>();

        //            while (!oRecordset.EoF)
        //            {
        //                //ListaDocumentos.Add(oRecordset.Fields.Item("DocEntry"/*"DocNum"*/).Value.ToString());
        //                ListaDocumentos.Add($"{oRecordset.Fields.Item("Serie").Value.ToString()}#{oRecordset.Fields.Item("DocNum").Value.ToString()}");

        //                oRecordset.MoveNext();
        //            }
        //            ////csListaDocumentos ListaDocumentosJSON = new csListaDocumentos();
        //            ////ListaDocumentosJSON.ListaDocumentos = ListaDocumentos;

        //            ////string Documento_JSON = Newtonsoft.Json.JsonConvert.SerializeObject(ListaDocumentosJSON, Newtonsoft.Json.Formatting.None);
        //            //foreach (string Documento in ListaDocumentos)
        //            //{
        //            //    List<string> ListaErrores = ValidarDocumentoEnvioConector(Documento);
        //            //    AñadirDetalleEnvioAConector(ListaErrores, Documento.CodigoEnvio, Documento);
        //            //}

        //            return ListaDocumentos;
        //        }

        //        private void RespuestaDeConector(List<string> ListaDocumentos)
        //        {
        //            ConectarAAHORA();
        //            string CadenaListaDocumentos = String.Join("', '", ListaDocumentos);

        //            string ColumnasConsulta = $@"
        //FCC.""IdFactura""
        //	, IsNull(TI.""urlQR"", '')
        //	/*, CASE FCC.""IdFactura""
        //		WHEN '435' THEN '439#13'
        //		WHEN '436' THEN '440#13'
        //		WHEN '437' THEN '441#13'
        //		WHEN '438' THEN '442#13'
        //		WHEN '439' THEN '443#13'
        //		WHEN '441' THEN '444#13'
        //		END AS ""Code""*/
        //    , CONCAT(CFC.""AX_SerieName"", '#', CFC.""AX_DocNum"", '#13')
        //    , TI.""QR"" AS ""QR_64""
        //    , TI.""Firmado"" AS ""Firmado""
        //";

        //            string Consulta = $@"
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""Facturas_Cli_Cab"" FCC
        //LEFT OUTER JOIN ""TBAI_Identificativo"" TI ON FCC.""IdFactura"" = TI.""IdFactura""
        //LEFT OUTER JOIN ""Conf_Facturas_Cli"" CFC ON FCC.""IdFactura"" = CFC.""IdFactura""
        //WHERE CONCAT(CFC.""AX_SerieName"", '#', CFC.""AX_DocNum"") IN ('{CadenaListaDocumentos}')--FCC.""IdFactura"" IN ('{CadenaListaDocumentos}')
        //AND (TI.""urlQR"" is not null AND TI.""urlQR"" <> '')
        ///*UNION
        //SELECT 
        //	{ColumnasConsulta}
        //FROM ""ORIN"" T0
        //INNER JOIN ""OCRD"" ""OCRD"" ON ""OCRD"".""CardCode"" = T0.""CardCode""
        //LEFT OUTER JOIN ""CRD1"" ON ""CRD1"".""CardCode"" = ""OCRD"".""CardCode"" AND ""CRD1"".""AdresType"" = 'B' AND ""CRD1"".""Address"" = ""OCRD"".""BillToDef""
        //INNER JOIN [@AX_ESTTBAI_0] K10_ET ON Right(K10_ET.""Code"", 2) = '14' AND K10_ET.""Code"" Like CONCAT(T0.""DocEntry"", '#14') AND K10_ET.""U_EnvCon"" = 'N'
        //INNER JOIN ""RIN12"" T12 ON T12.""DocEntry"" = T0.""DocEntry""*/
        //";

        //            System.Data.DataTable dataTable = new System.Data.DataTable();

        //            SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //            SqlDataAdapter da = new SqlDataAdapter(command);
        //            // this will query your database and return the result to your datatable
        //            da.Fill(dataTable);
        //            sqlConnectionAHORA.Close();
        //            da.Dispose();

        //            List<csDocumentosQR> ListaDocumentosQR = new List<csDocumentosQR>();

        //            foreach (DataRow DR in dataTable.Rows)
        //            {
        //                string Tabla = "";

        //                switch (DR[2].ToString().Split('#')[2])
        //                {
        //                    case "13":
        //                        Tabla = "OINV";
        //                        break;
        //                }
        //                Consulta = $@"
        //SELECT T0.""DocEntry""
        //FROM {Tabla} T0
        //INNER JOIN NNM1 TS ON T0.""Series"" = TS.""Series""
        //WHERE TS.""SeriesName"" = '{DR[2].ToString().Split('#')[0]}'
        //AND T0.""DocNum"" = {DR[2].ToString().Split('#')[1]}
        //";
        //                ListaDocumentosQR.Add(new csDocumentosQR()
        //                {
        //                    DocEntry = Convert.ToInt32(DR[0].ToString()),//Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString()),// Convert.ToInt32(DR[2].ToString().Split('#')[0]),
        //                    ObjType = DR[2].ToString().Split('#')[2],
        //                    DocNum = DR[2].ToString().Split('#')[1], //DR[0].ToString(),
        //                    Code = $"{Convert.ToInt32(DR[0].ToString())}#{DR[2].ToString().Split('#')[2]}",//$"{Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString())}#{DR[2].ToString().Split('#')[2]}",//DR[2].ToString(),
        //                    QR = DR[1].ToString(),
        //                    QR_64 = (byte[])DR[3],
        //                    DocumentoTBAIFirmado = DR[4].ToString(),
        //                    Serie = DR[2].ToString().Split('#')[0]
        //                });
        //            }

        //            //List<csDocumentos> ListaDocumentos = new List<csDocumentos>();

        //            //while (!oRecordset.EoF)
        //            //{
        //            //    ListaDocumentos.Add(AñadirADocumento(oRecordset));

        //            //    oRecordset.MoveNext();
        //            //}
        //            ////csListaDocumentos ListaDocumentosJSON = new csListaDocumentos();
        //            ////ListaDocumentosJSON.ListaDocumentos = ListaDocumentos;

        //            ////string Documento_JSON = Newtonsoft.Json.JsonConvert.SerializeObject(ListaDocumentosJSON, Newtonsoft.Json.Formatting.None);
        //            foreach (csDocumentosQR DocumentosQR in ListaDocumentosQR)
        //            {
        //                List<string> ListaErrores = ActualizarConQR(DocumentosQR);
        //                AñadirDetalleRespuestaDeConector(ListaErrores, DocumentosQR.Code, DocumentosQR.DocNum);
        //            }
        //        }

        //private List<string> ActualizarConQR(csDocumentosQR DocumentosQR)
        //{
        //    List<string> ListaErrores = new List<string>();

        //    if (String.IsNullOrWhiteSpace(DocumentosQR.QR))
        //    {
        //        ListaErrores.Add("Código QR no existe en Conector");
        //    }
        //    else
        //    {
        //        try
        //        {
        //            //oCompany.StartTransaction();

        //            SAPbobsCOM.Documents oDocuments = null;

        //            switch (DocumentosQR.ObjType)
        //            {
        //                case "13":
        //                    oDocuments = (SAPbobsCOM.Documents)Main.Company.GetBusinessObject(BoObjectTypes.oInvoices);
        //                    break;
        //                case "14":
        //                    oDocuments = (SAPbobsCOM.Documents)Main.Company.GetBusinessObject(BoObjectTypes.oCreditNotes);
        //                    break;
        //            }

        //            string PathFichero = "";

        //            oDocuments.GetByKey(DocumentosQR.DocEntry);

        //            oDocuments.UserFields.Fields.Item("U_AX_CodQR").Value = DocumentosQR.QR;
        //            //oDocuments.UserFields.Fields.Item("U_ImaQR").Value = PathFichero;
        //            //if (oDocuments.Update() != 0)
        //            //{
        //            //    ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //            //}


        //            List<string> ListaAdjuntos = new List<string>();
        //            ListaAdjuntos.Add(ConvertirBase64AImagen(DocumentosQR));
        //            ListaAdjuntos.Add(GenerarFichero(DocumentosQR));



        //            //oDocuments.GetByKey(DocumentosQR.DocEntry);
        //            //PathFichero = ConvertirBase64AImagen(DocumentosQR.DocEntry, DocumentosQR.QR_64);
        //            //AñadirAdjunto(PathFichero, false, ref oDocuments, ref ListaErrores);

        //            //oDocuments.GetByKey(DocumentosQR.DocEntry);
        //            //PathFichero = GenerarFichero(DocumentosQR.DocEntry, DocumentosQR.DocumentoTBAIFirmado);
        //            //AñadirAdjunto(PathFichero, true, ref oDocuments, ref ListaErrores);

        //            AñadirAdjuntos(ListaAdjuntos, true, ref oDocuments, ref ListaErrores);

        //            //oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        }
        //        catch (Exception ex)
        //        {
        //            //oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //    }

        //    return ListaErrores;
        //}

        //        private string ConvertirBase64AImagen(csDocumentosQR DocumentosQR)
        //        {
        //            //byte[] ImagenEn64 = Convert.FromBase64String(ImagenBase64);
        //            string Consulta = $@"
        //SELECT T0.BitmapPath
        //FROM OADP T0
        //";
        //            string PathImagenes = $"{Common.Functions.Connection.executeScalar(Consulta).ToString()}TBAI_QR\\";

        //            if (!Directory.Exists(PathImagenes)) Directory.CreateDirectory(PathImagenes);

        //            string NombreFichero = $"{PathImagenes}QR_{DocumentosQR.Serie}_{DocumentosQR.DocNum}.jpg";
        //            FileStream FicheroImagen = new FileStream(NombreFichero, FileMode.Create);

        //            FicheroImagen.Write(DocumentosQR.QR_64, 0, DocumentosQR.QR_64.Length);
        //            FicheroImagen.Flush();

        //            return NombreFichero;
        //        }

        //        private string GenerarFichero(csDocumentosQR DocumentosQR)
        //        {
        //            string Consulta = $@"
        //SELECT T0.""AttachPath""
        //FROM OADP T0
        //";
        //            string PathFichero = $"{Common.Functions.Connection.executeScalar(Consulta).ToString()}";

        //            if (!Directory.Exists(PathFichero)) Directory.CreateDirectory(PathFichero);

        //            string NombreFichero = $"{PathFichero}TBAI_Firmado_{DocumentosQR.Serie}_{DocumentosQR.DocNum}.xml";
        //            using (StreamWriter Fichero = new StreamWriter(NombreFichero))
        //            {
        //                Fichero.WriteLine(DocumentosQR.DocumentoTBAIFirmado);
        //            }

        //            return NombreFichero;
        //        }

        //private void AñadirAdjuntos(List<string> ListaAdjuntos, bool AñadirLinea, ref SAPbobsCOM.Documents oDocuments, ref List<string> ListaErrores)
        //{
        //    SAPbobsCOM.Attachments2 oAttachments2 = (SAPbobsCOM.Attachments2)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

        //    try
        //    {
        //        int AE = oDocuments.AttachmentEntry;
        //        oAttachments2.GetByKey(oDocuments.AttachmentEntry);

        //        int LineNo = oAttachments2.Lines.Count;

        //        foreach (var Adjunto in ListaAdjuntos)
        //        {
        //            oAttachments2.Lines.Add();
        //            oAttachments2.Lines.SetCurrentLine(LineNo);
        //            oAttachments2.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(Adjunto);
        //            oAttachments2.Lines.FileExtension = System.IO.Path.GetExtension(Adjunto).Substring(1);
        //            oAttachments2.Lines.SourcePath = System.IO.Path.GetDirectoryName(Adjunto);
        //            oAttachments2.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;

        //            LineNo++;
        //        }

        //        int Result = 0;

        //        if (AE == 0)
        //        {
        //            Result = oAttachments2.Add();
        //        }
        //        else
        //        {
        //            Result = oAttachments2.Update();
        //        }

        //        if (Result != 0)
        //        {
        //            ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //        }
        //        else
        //        {
        //            int NumeroAdjunto = Convert.ToInt32(Main.Company.GetNewObjectKey());

        //            oDocuments.AttachmentEntry = NumeroAdjunto;

        //            if (oDocuments.Update() != 0)
        //            {
        //                ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        ListaErrores.Add(ex.Message);
        //    }
        //    finally
        //    {
        //        Common.Functions.ReleaseComObject(oAttachments2);
        //    }
        //}

        //        private void AñadirAdjunto(string PathFichero, bool AñadirLinea, ref SAPbobsCOM.Documents oDocuments, ref List<string> ListaErrores)
        //        {
        //            SAPbobsCOM.Attachments2 oAttachments2 = (SAPbobsCOM.Attachments2)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);
        //            try
        //            {
        //                oAttachments2.Lines.Add();
        //                oAttachments2.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(PathFichero);
        //                oAttachments2.Lines.FileExtension = System.IO.Path.GetExtension(PathFichero).Substring(1);
        //                oAttachments2.Lines.SourcePath = System.IO.Path.GetDirectoryName(PathFichero);
        //                oAttachments2.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;

        //                if (oAttachments2.Add() != 0)
        //                {
        //                    ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //                }
        //                else
        //                {
        //                    int NumeroAdjunto = Convert.ToInt32(Main.Company.GetNewObjectKey());

        //                    oDocuments.AttachmentEntry = NumeroAdjunto;

        //                    if (oDocuments.Update() != 0)
        //                    {
        //                        ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //                    }
        //                }
        //            }
        //            catch (Exception ex)
        //            {

        //            }
        //            finally
        //            {
        //                Common.Functions.ReleaseComObject(oAttachments2);
        //            }
        //        }

        //        private void AñadirAdjunto_o(string PathFichero, bool AñadirLinea, ref SAPbobsCOM.Documents oDocuments, ref List<string> ListaErrores)
        //        {
        //            SAPbobsCOM.Attachments2 oAtt = (SAPbobsCOM.Attachments2)Main.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

        //            if (File.Exists(PathFichero))
        //            {
        //                //if (AñadirLinea)
        //                //{
        //                //    oAtt.Lines.Add();
        //                //}
        //                oAtt.Lines.SourcePath = System.IO.Path.GetDirectoryName(PathFichero);
        //                oAtt.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(PathFichero);
        //                oAtt.Lines.FileExtension = System.IO.Path.GetExtension(PathFichero).Substring(1);
        //                oAtt.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
        //                int attEntry = 0;
        //                if (oAtt.Add() == 0)
        //                {
        //                    attEntry = int.Parse(Main.Company.GetNewObjectKey());
        //                    //Documents oDoc = Main.Company.GetBusinessObject(BoObjectTypes.oInvoices);
        //                    //oDoc.CardCode = "C00000001";
        //                    //oDoc.DocDueDate = DateTime.Now;
        //                    //oDoc.DocType = BoDocumentTypes.dDocument_Service;
        //                    //oDoc.Lines.ItemDescription = "Test";
        //                    //oDoc.Lines.AccountCode = "5.01.02.05.39";
        //                    //oDoc.Lines.LineTotal = 100.00;
        //                    oDocuments.AttachmentEntry = attEntry;
        //                    if (oDocuments.Update() != 0)
        //                        //    //MessageBox.Show(oCompany.GetLastErrorDescription());
        //                        ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //                }
        //                else
        //                    //MessageBox.Show(oCompany.GetLastErrorDescription());
        //                    ListaErrores.Add($"{Main.Company.GetLastErrorCode().ToString()} - {Main.Company.GetLastErrorDescription()}");
        //            }
        //        }

        //        private void AñadirDetalleRespuestaDeConector(List<string> ListaErrores, string Code, string Documento)
        //        {
        //            string FechaHora = $"{DateTime.Today.ToString("dd/MM/yyyy")} {DateTime.Now.ToString("hh:mm")}";
        //            //List<List<csK10_DI.csGuardarDatoTablaUsuario>> ListasGuardarDatoTablaUsuario = new List<List<csK10_DI.csGuardarDatoTablaUsuario>>();

        //            string Consulta = $@"
        //SELECT IsNull(Max(K10_ET2.""U_NumInt""), 0) + 1
        //FROM [@AX_ESTTBAI_2] K10_ET2
        //WHERE K10_ET2.""Code"" Like '{Code}#%'
        //";
        //            int NumeroIntentoEnvio = Convert.ToInt32(Common.Functions.Connection.executeScalar(Consulta).ToString());

        //            //string Documento_JSON = $@"{{""Documents"": [{Newtonsoft.Json.JsonConvert.SerializeObject(Documento, Newtonsoft.Json.Formatting.Indented)}]}}";

        //            if (ListaErrores.Count == 0)
        //            {
        //                //List<csK10_DI.csGuardarDatoTablaUsuario> ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_FecHor", Valor = FechaHora });

        //                ////string ErrorEnvioAConector = "";
        //                ////bool EnviadoAAhora = EnviarAAHORA(Documento_JSON, ref ErrorEnvioAConector); // InsertarDocumentosAhora(Documento_JSON, connection);

        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Accion", Valor = "R_C" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Info", Valor = "Respuesta de conector recibida correctamente" });

        //                //ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);

        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_2");

        //                oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //                oUserTable.UserFields.Fields.Item("U_Accion").Value = "R_C";
        //                oUserTable.UserFields.Fields.Item("U_Info").Value = "Respuesta de conector recibida correctamente";

        //                if (oUserTable.Add() != 0)
        //                {
        //                    string Error = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);
        //            }
        //            else
        //            {
        //                //List<csK10_DI.csGuardarDatoTablaUsuario> ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#0" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Accion", Valor = "R_F" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Info", Valor = "Respuesta de conector fallida" });
        //                //ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_FecHor", Valor = FechaHora });

        //                //ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);

        //                //int cont = 1;
        //                //foreach (string Error in ListaErrores)
        //                //{
        //                //    ListaGuardarDatoTablaUsuario = new List<csK10_DI.csGuardarDatoTablaUsuario>();
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Code", Valor = $"{Code}#{NumeroIntentoEnvio}#{cont}" });
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "Name", Valor = $"{Code}#{NumeroIntentoEnvio}#{cont}" });
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Accion", Valor = "" });
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_NumInt", Valor = NumeroIntentoEnvio.ToString() });
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_Info", Valor = Error });
        //                //    ListaGuardarDatoTablaUsuario.Add(new csK10_DI.csGuardarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_2", Clave = "U_FecHor", Valor = "" });

        //                //    ListasGuardarDatoTablaUsuario.Add(ListaGuardarDatoTablaUsuario);

        //                //    cont++;
        //                //}

        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_2");

        //                oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#0";
        //                oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                oUserTable.UserFields.Fields.Item("U_FecHor").Value = FechaHora;
        //                oUserTable.UserFields.Fields.Item("U_Accion").Value = "R_F";
        //                oUserTable.UserFields.Fields.Item("U_Info").Value = "Respuesta de conector fallida";

        //                if (oUserTable.Add() != 0)
        //                {
        //                    string Error2 = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);

        //                int cont = 1;
        //                foreach (string Error in ListaErrores)
        //                {
        //                    oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_2");

        //                    oUserTable.Code = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //                    oUserTable.Name = $"{Code}#{NumeroIntentoEnvio}#{cont}";
        //                    oUserTable.UserFields.Fields.Item("U_NumInt").Value = NumeroIntentoEnvio.ToString();
        //                    oUserTable.UserFields.Fields.Item("U_FecHor").Value = "";
        //                    oUserTable.UserFields.Fields.Item("U_Accion").Value = "";
        //                    oUserTable.UserFields.Fields.Item("U_Info").Value = Error;

        //                    if (oUserTable.Add() != 0)
        //                    {
        //                        string Error2 = Main.Company.GetLastErrorDescription();
        //                    }

        //                    Common.Functions.ReleaseComObject(oUserTable);

        //                    cont++;
        //                }
        //            }

        //            //csK10_DI.DI_GuardarDatosTablaUsuario(ListasGuardarDatoTablaUsuario, null);



        //            //List<List<csK10_DI.csActualizarDatoTablaUsuario>> ListasActualizarDatoTablaUsuario = new List<List<csK10_DI.csActualizarDatoTablaUsuario>>();

        //            //if (ListaErrores.Count == 0)
        //            //{
        //            //    List<csK10_DI.csActualizarDatoTablaUsuario> ListaActualizarDatoTablaUsuario = new List<csK10_DI.csActualizarDatoTablaUsuario>();
        //            //    ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "Code", Valor = $"{Code}" });
        //            //    ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "Name", Valor = $"{Code}" });
        //            //    ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "U_ResCon", Valor = "Y" });
        //            //    ListaActualizarDatoTablaUsuario.Add(new csK10_DI.csActualizarDatoTablaUsuario() { TablaUsuario = "K10_ESTTBAI_0", Clave = "U_ResResCon", Valor = "Respuesta de conector recibida correctamente" });

        //            //    ListasActualizarDatoTablaUsuario.Add(ListaActualizarDatoTablaUsuario);
        //            //}

        //            //csK10_DI.DI_ActualizarDatosTablaUsuario(ListasActualizarDatoTablaUsuario, null);

        //            if (ListaErrores.Count == 0)
        //            {
        //                SAPbobsCOM.UserTable oUserTable = (SAPbobsCOM.UserTable)Main.Company.UserTables.Item("AX_ESTTBAI_0");

        //                oUserTable.GetByKey(Code);

        //                oUserTable.UserFields.Fields.Item("U_ResCon").Value = "Y";
        //                oUserTable.UserFields.Fields.Item("U_ResResCon").Value = "Respuesta de conector recibida correctamente";

        //                if (oUserTable.Update() != 0)
        //                {
        //                    string Error2 = Main.Company.GetLastErrorDescription();
        //                }

        //                Common.Functions.ReleaseComObject(oUserTable);
        //            }
        //        }

        //private void ConectarAAHORA()
        //{
        //    //string sCalve = "AxTicketBai";
        //    //string sConectionString = Common.crypto.decryptText(sCalve, @"nnWUCINHXuP16AGGMeUVelFZsBjfOZQkclW4O5UiR6XALUS4iulBkcGhJJU1ST116DGloT86ebEIinuj91yL5FZLl4OQveVxpJ9kgW77/WG9u8WQ46ufY1Jl+NYaKQfEVUluK4NPDSSsRwxoeq8Tl502tQdRIiYb205sbQn7TOURznc/AQAQqdiznk1x2aIeyWw9kUW6ZbNtZzVt01QzsEYxyDmjb08PnjBZ9HW67hI=");
        //    //string constrAHORA = "Server = ax-flexygo\\FLEXYGO; Database = AHORA_46; User Id = sa; Password = SB1Admin";
        //    string constrAHORA = "Server = 185.57.174.52,1444; Database = AHORA_46; User Id = sa; Password = SB1Admin;";
        //    //string sServer = Common.Functions.Xml.getXmlParamsValue("params/ahora/server", "", 0);
        //    //string sDataBase = Common.Functions.Xml.getXmlParamsValue("params/ahora/database", "", 0);
        //    //string sUserId = Common.Functions.Xml.getXmlParamsValue("params/ahora/userid", "", 0);
        //    //string sPassword = Common.Functions.Xml.getXmlParamsValue("params/ahora/password", "", 0);
        //    //string constrAHORA = $"Server = {sServer}; Database = {sDataBase}; User Id = {sUserId}; Password = {sPassword};";
        //    sqlConnectionAHORA = new SqlConnection(constrAHORA);

        //    if (sqlConnectionAHORA.State == System.Data.ConnectionState.Closed)
        //    {
        //        sqlConnectionAHORA.Open();
        //    }
        //}

        //        private bool EnviarAAHORA_o(string Documento_JSON, csDocumentos Documento, ref string ErrorEnvioAConector)
        //        {
        //            TBAI.TBAICommon.ConectarAAHORA();
        //            Main.Log.write("EnviarAAHORA_o" + Documento.AX_DocNum);
        //            //Cuando tengamos la conexion con AhoraERP cambiar esta query con la correcta
        //            string queryPre = "INSERT INTO Ahora_Sesion(SpId,IdEmpresa,IdDelegacion,IdEmpleado,IdDepartamento,IdGrupoSeguridad,IdAplic,Exclusivo,Equipo,Usuario) SELECT @@SPID,0,0,0,0,0,'AhoraInicio',0,HOST_NAME(),'ahora'";
        //            string Consulta = $"EXEC [AX_TBAI_AltaDocumento] '{Documento_JSON}' ";

        //            //CommonMini.Log.write("[AX_API TICKETBAI Offline][DataProcessHelper][InsertarDocumentosAhora] Json enviado: " + sJon);

        //            SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //            SqlCommand commandPre = new SqlCommand(queryPre, sqlConnectionAHORA);
        //            int cCount = 0;

        //            try
        //            {
        //                commandPre.ExecuteNonQuery();
        //                cCount = (int)command.ExecuteNonQuery();

        //                Consulta = $@"
        //UPDATE Clientes_Subcuentas
        //SET Subcuenta = '600000001'
        //WHERE IdCliente = '{Documento.CardCode}'
        //";

        //                command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //                commandPre = new SqlCommand(queryPre, sqlConnectionAHORA);
        //                cCount = 0;

        //                commandPre.ExecuteNonQuery();
        //                cCount = (int)command.ExecuteNonQuery();

        //                return true;
        //            }
        //            catch (Exception ex)
        //            {
        //                ErrorEnvioAConector = ex.Message;
        //                //Console.WriteLine($"[AX_API TICKETBAI] [InsertarDocumentosAhora] Error : " + ex.ToString());
        //                //CommonMini.Log.write("[AX_API TICKETBAI Offline][DataProcessHelper][InsertarDocumentosAhora] Error al llamar a Ahora ERP: " + ex.ToString());
        //                return false;
        //            }
        //        }

        //private bool EnviarAAHORA(string Documento_JSON, csDocumentos Documento, ref string ErrorEnvioAConector)
        //{
        //    ConectarAAHORA();
        //    Main.Log.write("EnviarAAHORA" + Documento.AX_DocNum);
        //    bool Resultado = true;

        //    try
        //    {
        //        Resultado = TBAI_AltaDocumento(Documento_JSON, ref ErrorEnvioAConector);

        //        if (Resultado)
        //        {
        //            Resultado = TBAI_ActualizaCuentasCliente(Documento, ref ErrorEnvioAConector);
        //        }

        //        return Resultado;
        //    }
        //    catch (Exception ex)
        //    {
        //        Main.Log.write(ex.ToString());
        //        ErrorEnvioAConector = ex.Message;

        //        return false;
        //    }
        //}

        //        public bool TBAI_AltaDocumento(string Documento_JSON, ref string ErrorEnvioAConector)
        //        {
        //            int cCount = 0;
        //            string Consulta = $"EXEC [AX_TBAI_AltaDocumento] '{Documento_JSON}' ";
        //            Main.Log.write(Consulta);
        //            SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //            SqlCommand commandPre = new SqlCommand(ConsultaPreviaAHORA, sqlConnectionAHORA);

        //            try
        //            {
        //                commandPre.ExecuteNonQuery();
        //                cCount = (int)command.ExecuteNonQuery();

        //                return true;
        //            }
        //            catch (Exception ex)
        //            {
        //                Main.Log.write(ex.ToString());
        //                ErrorEnvioAConector += $"{ex.Message}{Environment.NewLine}";
        //                //Console.WriteLine($"[AX_API TICKETBAI] [InsertarDocumentosAhora] Error : " + ex.ToString());
        //                //CommonMini.Log.write("[AX_API TICKETBAI Offline][DataProcessHelper][InsertarDocumentosAhora] Error al llamar a Ahora ERP: " + ex.ToString());
        //                return false;
        //            }
        //        }

        //        public bool TBAI_ActualizaCuentasCliente(csDocumentos Documento, ref string ErrorEnvioAConector)
        //        {
        //            int cCount = 0;
        //            string Consulta = $@"
        //UPDATE Clientes_Subcuentas
        //SET Subcuenta = '600000001'
        //WHERE IdCliente = '{Documento.CardCode}'
        //";

        //            SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //            SqlCommand commandPre = new SqlCommand(ConsultaPreviaAHORA, sqlConnectionAHORA);

        //            try
        //            {
        //                commandPre.ExecuteNonQuery();
        //                cCount = (int)command.ExecuteNonQuery();

        //                return true;
        //            }
        //            catch (Exception ex)
        //            {
        //                Main.Log.write(ex.ToString());
        //                ErrorEnvioAConector += $"{ex.Message}{Environment.NewLine}";

        //                return false;
        //            }
        //        }

        //public bool TBAI_ActualizaDocumentoAHORA(csDocumentos Documento, ref string ErrorEnvioAConector)
        //{
        //    ConectarAAHORA();

        //    int cCount = 0;
        //    string Fecha = $"{Documento.DocDate.ToString("yyyyMMdd")}";
        //    //            string Consulta = $@"
        //    //{ConsultaPreviaAHORA}
        //    //EXEC zsetContextInfo ahora
        //    //EXECUTE AS USER = 'ahora'
        //    //SET NOCOUNT ON

        //    //	Declare @vret int
        //    //		,@idfactura int = {Documento.DocEntry} -- indicamos la factura a actualizar
        //    //		, @TipoFactura int
        //    //		,@FechaActualizacion date
        //    //		,@IdEjercicio int

        //    //	-- obtenemos el tipo de factura
        //    //	SELECT @TipoFactura = CASE WHEN Deudor = 1 THEN 3 ELSE 0 END FROM Facturas_Cli_Cab WHERE IdFactura = @idfactura
        //    //	-- indicamos la fecha de actualizacion
        //    //	SET @FechaActualizacion = '{Fecha}'

        //    //	-- obtenemos el ejercicio de la factura
        //    //	SELECT @IdEjercicio = IdEjercicio FROM Conta_Ejercicios CE
        //    //									INNER JOIN Facturas_Cli_Cab F ON  F.IdEmpresa = CE.IdEmpresa AND F.IdFactura = @IdFactura
        //    //									WHERE @FechaActualizacion BETWEEN CE.FechaInicio AND CE.FechaFin
        //    //	-- actualizamos la factura
        //    //	EXEC  @vret = pfactura_cli_actualizar @idfactura,@FechaActualizacion,@IdEjercicio

        //    //	-- si se ha actualizado, generamos el TBAI
        //    //	IF @vret <> 0 BEGIN
        //    //		SET @vret = NULL
        //    //		EXEC @vret = pTBAI_generarTBAI_Factura  @IdFactura, @TipoFactura, 0

        //    //		-- si no se genera el tbai, desactualizamos la factura
        //    //		IF @vret = 0 BEGIN
        //    //			-- si falla, desactualizamos la factura
        //    //			EXEC Desactualiza_FacturaCliente @IdFactura, @TipoFactura, 1
        //    //		END
        //    //	END

        //    //REVERT
        //    //";
        //    string Consulta = $"EXEC [AX_TBAI_ActualizaDocumentoAHORA] '{Documento.DocEntry}', '{Fecha}' ";
        //    SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //    SqlCommand commandPre = new SqlCommand(ConsultaPreviaAHORA, sqlConnectionAHORA);

        //    try
        //    {
        //        commandPre.ExecuteNonQuery();
        //        cCount = (int)command.ExecuteNonQuery();

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorEnvioAConector += $"{ex.Message}{Environment.NewLine}";

        //        return false;
        //    }
        //}

        //public bool TBAI_CreaCuentasParaPoderActualizarAHORA(csDocumentos Documento, ref string ErrorEnvioAConector)
        //{
        //    ConectarAAHORA();

        //    int cCount = 0;
        //    string Fecha = $"{Documento.DocDate.ToString("yyyyMMdd")}";

        //    string Consulta = $"EXEC [pDameCuentasActualizarFacturaVenta] '{Documento.DocEntry}', '{Documento.DocEntry}', '{Fecha}', 0, null, 1";
        //    SqlCommand command = new SqlCommand(Consulta, sqlConnectionAHORA);
        //    //SqlCommand commandPre = new SqlCommand(ConsultaPreviaAHORA, sqlConnectionAHORA);

        //    try
        //    {
        //        //commandPre.ExecuteNonQuery();
        //        cCount = (int)command.ExecuteNonQuery();

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        ErrorEnvioAConector += $"{ex.Message}{Environment.NewLine}";

        //        return false;
        //    }
        //}
        #endregion
    }
}
