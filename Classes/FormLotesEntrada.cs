
using SAPbouiCOM;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using static Lotes_Hornitos.Classes.DataDocument;
using System.Xml.Serialization;
using Lotes_Hornitos.Modelos;
using System.Collections;
using System.Security.Cryptography;
using SAPbobsCOM;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace Lotes_Hornitos.Classes
{
    internal class FormLotesEntrada
    {
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application oApp;
        public static List<EntradaLotes> ListaEntrada;

        /// <summary>
        /// FormUID del documento de SAP al que pertenece la lista de lotes en memoria
        /// (la lista es estatica: una sola por sesion de SAP).
        /// </summary>
        public static string ListaOwnerFormUID = null;
        public bool DesdeLista;

        public FormLotesEntrada(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oApp)
        {
            this.oCompany = oCompany;
            this.oApp = oApp;
            ListaEntrada = new List<EntradaLotes>();
        }

        /// <summary>
        /// Mostrar el formulario
        /// </summary>
        internal bool showForm(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            List<bool> result = new List<bool>();
            bool res = true;
            try
            {

                //validar si ya se agregaron los lotes
                result = ValidasiexisteloteentablaEntrada(query);



                foreach (bool l in result)
                {
                    if (!l)
                    {
                        res = l;
                        break;
                    }
                }

                if (res)
                {
                    return true;
                }

                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTE.srf";


                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(oPath);
                StreamReader sr = new StreamReader(stream);
                stringForm = sr.ReadToEnd();

                oXmlDocument.LoadXml(stringForm);
                FormUID = oXmlDocument.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                // CORRECCION: si la ventana ya estaba abierta (por ejemplo el
                // usuario vuelve a pedir el formulario sin haberlo cerrado), se
                // cierra primero para poder abrirla de nuevo limpia. Antes esto
                // lanzaba "Form - already exists [66000-11]".
                try
                {
                    SAPbouiCOM.Form oFormExistente = oApp.Forms.Item(FormUID);
                    oFormExistente.Close();
                    Ini.ReleaseCOMObjetc(oFormExistente);
                }
                catch
                {
                    // No existe todavía, es el caso normal. No hay nada que hacer.
                }

                oApp.LoadBatchActions(ref stringForm);

                oForm = oApp.Forms.Item(FormUID);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //DT_Crear.Rows.Add();




                //query = "SELECT 'N' AS 'Sel',15 AS 'Numero Doc',T0.ItemCode,CAST((1.000000) AS INT) AS 'Cantidad','01' AS 'Almacen'FROM OITM T0 WHERE T0.ItemCode = 'TESTLOTES-H2' AND T0.QryGroup28 = 'Y'";

                DT_LineasDocumento.ExecuteQuery(query);



                if (!DT_LineasDocumento.IsEmpty)
                {

                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = true;
                    grid_LineasDoc.Item.AffectsFormMode = false;


                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                    ColumnGrid.LinkedObjectType = "4";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                    ColumnGrid.LinkedObjectType = "64";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                    {

                        grid_LineasDoc.Columns.Item(i).Editable = false;
                        grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                        grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                    }

                    grid_LineasDoc.Columns.Item("Sel").Editable = true;
                    grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                    grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";
                    try { grid_LineasDoc.Columns.Item("Descripcion").TitleObject.Caption = "Descripción del artículo"; } catch { }
                    //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                    grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                    grid_LineasDoc.AutoResizeColumns();
                    oForm.Visible = true;

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    oForm.Close();
                    res = true;
                }

                return res;

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(ColumnGrid);

            }
        }

        /// <summary>
        /// Mostrar formulario de entrada de lotes y crear lotes
        /// </summary>
        /// <param name="pVal"></param>
        public bool MostarItems(ItemEvent pVal, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Entrada = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;
            bool result = false;
            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes");

                Entrada = oApp.Forms.Item(pVal.FormUID);

                if (!Entrada.Mode.Equals(BoFormMode.fm_ADD_MODE)) return true;

                // CORRECCION (dos documentos abiertos a la vez): la lista de lotes en
                // memoria solo vale para UN documento. Si el usuario viene de otro
                // documento, se descarta la lista ajena y este documento empieza limpio.
                if (ListaOwnerFormUID != pVal.FormUID)
                {
                    Ini.RegistrarError($"[{Tipo}] Dueno de la lista de lotes cambia de '{ListaOwnerFormUID}' a '{pVal.FormUID}' (lineas en lista: {(ListaEntrada == null ? 0 : ListaEntrada.Count)}).");
                    if (ListaEntrada != null) ListaEntrada.Clear();
                    ListaOwnerFormUID = pVal.FormUID;
                }


                switch (Tipo)
                {
                    case "OIGN":
                        LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("13").Specific;
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("38").Specific;
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");

                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLote.sql")).ReadToEnd();

                // CORRECCION (cancelacion de factura de compra): el documento "Factura de proveedores - Cancelacion"
                // no mueve inventario por si mismo, no se piden lotes.
                if (Tipo == "OPCH" && EsCancelacionDeFactura(Entrada))
                {
                    // Factura copiada de una entrada de mercancia: no hay nada que registrar.
                    // Factura directa (movio inventario): la cancelacion es una SALIDA de lotes.
                    FormLotesSalida salidaCan = ObtenerSalidaCancelacion();
                    if (!salidaCan.CancelacionTieneLineasConMovimiento(Entrada)) return true;
                    Ini.RegistrarError("[OPCH] cancelacion de factura directa: se abre la seleccion de lotes de salida.");
                    return salidaCan.MostarItems(pVal, "OPCH");
                }

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;
                    //ajuste
                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        switch (Tipo)
                        {
                            case "OPCH":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("20"))
                                {
                                    return true;
                                }
                                break;
                            case "ORIN":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("16"))
                                {
                                    return true;
                                }
                                break;

                        }
                        switch (Tipo)
                        {

                            case "OIGN":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("9").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGN":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;

                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGN":
                                oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("7").Specific;

                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;

                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);
                        Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }



                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);

                Console.WriteLine("fin de consulta de articulos con lotes");

                // CORRECCION: si no se encontró ningún artículo con lote en las
                // líneas del documento, "Items" queda vacío y la consulta armada
                // también queda vacía — ejecutar eso contra SAP lanzaba
                // "Estructura de consulta no permitida". Ahora, si no hay nada
                // que mostrar, simplemente no se abre el formulario.
                if (Items.Count == 0 || string.IsNullOrWhiteSpace(query))
                {
                    Console.WriteLine($"No se encontraron artículos con lote en el documento (Tipo: {Tipo}). No se abre el formulario de lotes.");
                    return true;
                }

                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de entrada de lotes");
                result = showForm(query);
                System.Console.WriteLine("Fin carga de formulario de entrada de lotes");

                return result;

            }
            catch (Exception ex)
            {
                Entrada.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Entrada);
                Ini.ReleaseCOMObjetc(LineasEntrada);
                Ini.ReleaseCOMObjetc(oEditText);
                Ini.ReleaseCOMObjetc(oCantidad);
                Ini.ReleaseCOMObjetc(oAlmacen);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(RowNumber);

            }






        }

        /// <summary>
        /// Valida si los articulos con lote ya se encuentran agregados en la tabla de lotes entrada parta no abrir el formulario de ingreso de lotes y pueda pasar a crear el documento de sap.
        /// </summary>
        /// <param name="ItemCode"></param>
        /// <param name="oDocNum"></param>
        /// <param name="tipo"></param>
        /// <returns></returns>
        private List<bool> ValidasiexisteloteentablaEntrada(string query)
        {
            //string query = string.Empty;
            //SAPbobsCOM.Recordset oRs = null;
            //var Sb = new System.Text.StringBuilder();

            double cantidadLote = 0;
            bool existelote = false, existeenconsulta = false;
            SAPbobsCOM.Recordset oRs = null;

            List<bool> result = new List<bool>();

            try
            {

                if (ListaEntrada == null || !ListaEntrada.Any())
                {
                    result.Add(false);
                    return result;
                }



                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);
                oRs.MoveFirst();

                for (int i = 1; i <= oRs.RecordCount; i++)
                {

                    existelote = false;
                    cantidadLote = 0;

                    for (int j = 0; j < ListaEntrada.Count; j++)
                    {
                        var item = ListaEntrada[j];

                        System.Console.WriteLine($"desde lista {item.ItemCode}");
                        System.Console.WriteLine($"desde recordset {oRs.Fields.Item("ItemCode").Value.ToString()}");

                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {

                                if (!item.almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {

                                    ListaEntrada.RemoveAt(j);
                                    result.Add(false);
                                    break;
                                }

                                cantidadLote += item.CantLote;
                                existelote = true;
                            }
                        }




                    }
                    if (existelote)
                    {
                        // CORRECCION (traslado 590129): antes solo se volvia a pedir los lotes si la cantidad
                        // asignada era MENOR que la de la linea. Si el usuario bajaba la cantidad despues de
                        // asignar los lotes (asignado MAYOR), el documento se grababa con lotes de mas.
                        // Ahora se pide de nuevo cuando la cantidad asignada es DISTINTA a la de la linea.
                        if (Math.Abs(cantidadLote - Convert.ToDouble(oRs.Fields.Item("Cantidad").Value)) > 0.001)
                        {
                            result.Add(false);
                            break;
                        }
                    }
                    else
                    {
                        result.Add(false);
                    }
                    oRs.MoveNext();
                }



                //validar que los valores de la tabla tengan los mismos de la consulta.
                int o = 0;
                for (o = 0; o < ListaEntrada.Count; o++)
                {
                    existeenconsulta = false;
                    var item = ListaEntrada[o];
                    oRs.MoveFirst();
                    for (int i = 1; i <= oRs.RecordCount; i++)
                    {
                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {
                                if (item.almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    existeenconsulta = true;
                                }
                            }
                        }
                        oRs.MoveNext();
                    }
                    if (!existeenconsulta)
                    {
                        ListaEntrada.Remove(item);
                        o--;
                    }

                }





                //oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosenTablaEntrada.sql")).ReadToEnd();
                //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode, tipo, oDocNum).ToString();

                //oRs.DoQuery(query);
                //if (oRs == null)
                //{
                //    return false;
                //}
                //else
                //{
                //    oRs.MoveFirst();

                //    int valor = (int)oRs.Fields.Item(0).Value;
                //    if (valor.Equals(0)) return false;
                //}


                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en ValidasiexisteloteentablaEntrada {ex.Message}");
                return result;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Funcion para agregar al dt de crear la informacion de la linea
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="tipo"></param>
        internal void AddInfoLotes(ItemEvent pVal, string tipo)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            SAPbouiCOM.CellPosition oCellPosition = null;
            EditTextColumn ColumnGrid = null;
            Columns columns = null;
            Column column = null;

            double cantidad = 0;
            string articulo = string.Empty, query = string.Empty;
            string path = string.Empty;


            var Sb = new System.Text.StringBuilder();
            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");
                oCellPosition = grid_LineasDoc.GetCellFocus();


                cantidad = Convert.ToDouble(DT_LineasDocumento.GetValue("Cantidad", pVal.Row));

                // valida si hay informacion de lotes ya creados en la tabla de entradas, esto para mostrar la info de lotes en el gri de lotes creados

                /*query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteUDT.sql")).ReadToEnd();
                query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), tipo, DT_LineasDocumento.GetValue("Numero Doc", pVal.Row), cantidad, DT_LineasDocumento.GetValue("Almacen", pVal.Row)).ToString();
                DT_Crear.ExecuteQuery(query);
                */

                //Cambio por Lista y datatable
                CargarLista(ref DT_Crear, ListaEntrada, DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), DT_LineasDocumento.GetValue("Almacen", pVal.Row).ToString(), cantidad, Convert.ToInt32(DT_LineasDocumento.GetValue("NumLinea", pVal.Row)));


                //Si no hay lotes creados continua con la informacion de lotes a crear, desde la informacion del documento de entrada.
                if (DT_Crear.IsEmpty)
                {

                    Sb.Clear();
                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLote.sql")).ReadToEnd();
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), tipo, DT_LineasDocumento.GetValue("Numero Doc", pVal.Row), cantidad, DT_LineasDocumento.GetValue("Almacen", pVal.Row), DT_LineasDocumento.GetValue("NumLinea", pVal.Row)).ToString();
                    // carga la info en el grid de lotes a crear
                    DT_Crear.ExecuteQuery(query);
                }


                if (!DT_Crear.IsEmpty)
                {
                    grid_Crear.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Crear.Item.Enabled = true;
                    grid_Crear.Item.AffectsFormMode = true;
                    grid_Crear.Columns.Item("ItemCode").Visible = false;
                    grid_Crear.Columns.Item("Tipo doc").Visible = false;
                    grid_Crear.Columns.Item("Num Doc").Visible = false;
                    grid_Crear.Columns.Item("Index").Visible = false;
                    grid_Crear.Columns.Item("NumLinea").Visible = false;

                    ColumnGrid = (EditTextColumn)grid_Crear.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_Crear.Columns.Count; i++)
                    {

                        grid_Crear.Columns.Item(i).Editable = DesdeLista ? false : true;
                        grid_Crear.Columns.Item(i).AffectsFormMode = DesdeLista ? false : true;
                    }

                    grid_Crear.Columns.Item("Sel").Editable = true;

                    grid_Crear.AutoResizeColumns();

                    var s = DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 1);
                    if (DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 1).ToString().Contains("1900"))
                    {
                        DT_Crear.SetValue("Fecha vence", grid_Crear.Rows.Count - 1, "");
                    }
                    oForm.Items.Item("BtnAddLn").Enabled = true;
                }
                else
                {
                    grid_Crear.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Crear.Item.Enabled = false;
                }


                //columns = (SAPbouiCOM.Columns)grid_Crear.Columns;
                //column = columns.Item("Fecha vence");
                //column.DataBind.SetBound(true,"@CLC_LTS","U_FchVenLt");





            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(oCellPosition);
                Ini.ReleaseCOMObjetc(ColumnGrid);
                Ini.ReleaseCOMObjetc(columns);
                Ini.ReleaseCOMObjetc(column);
            }
        }

        /// <summary>
        /// Funcion para tomar info de la lista para mostrar en el datatable de crear
        /// </summary>
        /// <param name="dT_Crear"></param>
        /// <param name="listaEntrada"></param>
        /// <param name="_ItemCode"></param>
        private void CargarLista(ref DataTable dT_Crear, List<EntradaLotes> listaEntrada, string _ItemCode, string almacen, double cantidadlinea, int numLinea)
        {
            string query = string.Empty;
            int cont = 0;
            double cantidadEntrada = 0;
            List<string> ListLotes = new List<string>();
            var Sb = new System.Text.StringBuilder();

            dT_Crear.Rows.Clear();

            if (listaEntrada == null || !listaEntrada.Any())
            {
                DesdeLista = false;
                return;
            }


            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.LotesDatatable.sql")).ReadToEnd();


            for (int i = 0; i < listaEntrada.Count; i++)
            {
                var item = listaEntrada[i];
                // CORRECCION: antes solo comparaba ItemCode+Almacen, así que si el
                // documento tenía 2 líneas con el mismo artículo y almacén, al
                // seleccionar la segunda línea traía (y bloqueaba) los datos ya
                // capturados en la primera. Ahora también exige que sea la MISMA
                // línea del documento (NumLinea), para no mezclar líneas distintas.
                if (item.ItemCode.Equals(_ItemCode))
                {
                    if (item.almacen.Equals(almacen) && item.NumLinea == numLinea)
                    {
                        ListLotes.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, item.ItemCode, item.NumDoc, item.TipDoc, item.CantLote, item.NumLot, item.FchVenLt, item.FchAdmLt, item.almacen, i, item.NumLinea).ToString());
                        cantidadEntrada += item.CantLote;
                        Sb.Clear();
                    }


                }

            }

            if (ListLotes.Any())
            {
                query = string.Join(" UNION ALL ", ListLotes);
                dT_Crear.ExecuteQuery(query);
                if (cantidadlinea > cantidadEntrada)
                {
                    DesdeLista = false;

                }
                else
                {
                    DesdeLista = true;

                }

            }
            else
            {
                DesdeLista = false;
            }

        }

        internal void CleanData(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;

            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                oForm.Freeze(true);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                for (int i = 0; i < DT_LineasDocumento.Rows.Count; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).Equals("Y"))
                    {
                        DT_LineasDocumento.SetValue("Sel", i, "N");
                    }
                }

                DT_Crear.Clear();
                oForm.Freeze(false);


            }
            catch (Exception ex)
            {
                oForm.Freeze(false);

            }
            finally
            {
                oForm.Freeze(false);

                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);            
            }
        }


        /// <summary>
        /// Funcion que agrega la informacion del lote a la lista
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        internal bool AddLotesList(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_Lineas = null;
            DataTable DT_Crear = null;
            DataTable DT_Lineas = null;
            object fechaVence = null;
            Console.WriteLine("Inicio de proceso de creación lote");
            bool result = true;

            string loteAIngresar = string.Empty;
            double cantidadAingresar = 0, cantidadLineas = 0;
            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_Lineas = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_Lineas = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");



                if (oForm.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
                {

                    for (int i = 0; i <= DT_Crear.Rows.Count - 1; i++)
                    {
                        loteAIngresar = DT_Crear.GetValue("Numero lote", i).ToString();

                        if (string.IsNullOrEmpty(loteAIngresar))
                        {
                            oApp.MessageBox("Debe ingresar un número lote valido", 1, "Ok");
                            result = false;
                            break;
                        }
                        else if (!ValidaLoteCrear(loteAIngresar, DT_Crear.GetValue("Almacen", i).ToString(), DT_Crear.GetValue("ItemCode", i).ToString()))
                        {
                            oApp.MessageBox("El lote ingresado ya existe", 1, "Ok");
                            result = false;
                            break;
                        }

                        //valida que no haya lotes iguales dentro del grid
                        for (int j = i + 1; j <= DT_Crear.Rows.Count - 1; j++)
                        {
                            if (DT_Crear.GetValue("Numero lote", j).ToString().Equals(loteAIngresar))
                            {
                                oApp.MessageBox("Hay lotes duplicados", 1, "Ok");
                                result = false;
                                break;
                            }
                        }


                        fechaVence = DT_Crear.GetValue("Fecha vence", i);

                        if (fechaVence is null || string.IsNullOrEmpty(fechaVence.ToString()))
                        {
                            oApp.MessageBox("Debe ingresar una fecha vence valida", 1, "Ok");
                            result = false;
                            break;
                        }

                        if (!String.IsNullOrEmpty(DT_Crear.GetValue("Cantidad", i).ToString()))
                        {
                            cantidadAingresar += Convert.ToDouble(DT_Crear.GetValue("Cantidad", i));
                        }


                    }

                    if (!result) { return result; }




                    for (int i = 0; i <= grid_Lineas.Rows.Count - 1; i++)
                    {
                        if (DT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                        {
                            cantidadLineas = Convert.ToDouble(DT_Lineas.GetValue("Cantidad", i));
                            break;
                        }

                    }

                    if (cantidadAingresar > cantidadLineas)
                    {
                        oApp.MessageBox("La cantidad del lote no debe ser mayor a la cantidad del documento", 1, "Ok");
                        result = false;
                        return result;
                    }

                    // aca debe ir la asignacion de los lotes a el datatable global.

                    ConvertirALista(DT_Crear);


                }




                return result;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(DT_Lineas);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(grid_Lineas);
            }
        }

        private bool ValidaLoteCrear(string loteAIngresar, string almacen, string ItemCode)
        {
            SAPbobsCOM.Recordset oRs = null;
            string query = string.Empty;

            try
            {
                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ValidaLoteCrear.sql")).ReadToEnd();
                query = query.Replace("lote", loteAIngresar).Replace("bodega", almacen).Replace("item", ItemCode);

                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery(query);

                if (oRs == null)
                {
                    return false;
                }
                else
                {
                    oRs.MoveFirst();
                    if ((int)oRs.Fields.Item(0).Value > 0)
                    {
                        if (oApp.MessageBox($"Para el articulo {ItemCode} ya existe el lote {loteAIngresar} en el almacen {almacen}.\n ¿Desea continuar?", 1, "Si", "No").Equals(1))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }

                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }

        }

        internal void AddLine(ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            DataTable DT_Crear = null;
            Grid grid_Crear = null;


            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                if (oForm.Items.Item("BtnAddLn").Enabled.Equals(false)) return;

                oForm.Freeze(true);
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;



                DT_Crear.Rows.Add();
                DT_Crear.SetValue("ItemCode", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("ItemCode", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Cantidad", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Cantidad", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Fecha vence", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Fecha vence", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Fecha Admin", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Fecha Admin", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Tipo doc", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Tipo doc", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Num Doc", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Num Doc", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("Almacen", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("Almacen", grid_Crear.Rows.Count - 2));
                DT_Crear.SetValue("NumLinea", grid_Crear.Rows.Count - 1, DT_Crear.GetValue("NumLinea", grid_Crear.Rows.Count - 2));

            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Console.WriteLine($"Error en agregar línea {ex.Message}");

            }
            finally
            {
                oForm.Freeze(false);
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(grid_Crear);
            }
        }

        internal void RollBack(ItemEvent pVal, string Tipo)
        {
            SAPbouiCOM.Form Entrada = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            EditText ItemCode = null;
            EditText oDocNum = null;
            SAPbobsCOM.Recordset oRs = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de rollback de articulos con lote entrada");

                Entrada = oApp.Forms.Item(pVal.FormUID);
                LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("38").Specific;
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");


                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    ItemCode = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(ItemCode.String))
                    {
                        oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("8").Specific;

                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.DeleteArticulosconLote.sql")).ReadToEnd();

                        query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode.String, Tipo, oDocNum.String).ToString();

                        oRs.DoQuery(query);

                        Sb.Clear();

                        System.Console.WriteLine($"Query rollback entrada, {query}");
                    }
                }

                Console.WriteLine("Fin de rollback de articulos con lote entrada");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en rollback de articulos con lote entrada, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Entrada);
                Ini.ReleaseCOMObjetc(LineasEntrada);
                Ini.ReleaseCOMObjetc(ItemCode);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        public void ConvertirALista(SAPbouiCOM.DataTable dtSAP)
        {
            //var lista = new List<EntradaLotes>();            

            for (int i = 0; i < dtSAP.Rows.Count; i++)
            {

                //add
                var item = new EntradaLotes();

                item.ItemCode = dtSAP.GetValue("ItemCode", i)?.ToString();
                item.TipDoc = dtSAP.GetValue("Tipo doc", i)?.ToString();
                item.NumDoc = (int)dtSAP.GetValue("Num Doc", i);
                item.NumLot = dtSAP.GetValue("Numero lote", i)?.ToString();
                item.CantLote = Convert.ToDouble(dtSAP.GetValue("Cantidad", i));

                var fechaObj = dtSAP.GetValue("Fecha vence", i);
                DateTime fecha = Convert.ToDateTime(fechaObj);
                item.FchVenLt = fecha.ToString("dd/MM/yyyy");

                //if (
                //    DateTime.TryParse(fechaObj.ToString(), out DateTime fecha))
                //{
                //    item.FchVenLt = fecha;
                //}

                fechaObj = dtSAP.GetValue("Fecha Admin", i);
                fecha = Convert.ToDateTime(fechaObj);
                item.FchAdmLt = fecha.ToString("dd/MM/yyyy");

                //if (
                //    DateTime.TryParse(fechaObj.ToString(), out DateTime fecha1))
                //{
                //    item.FchAdmLt = fecha1;
                //}

                item.almacen = dtSAP.GetValue("Almacen", i)?.ToString();
                item.NumLinea = (int)dtSAP.GetValue("NumLinea", i);

                bool existe = ListaEntrada.Any(x =>
                x.ItemCode == item.ItemCode &&
                x.NumDoc == item.NumDoc &&
                x.NumLot == item.NumLot &&
                x.NumLinea == item.NumLinea
                );

                if (!existe)
                {
                    ListaEntrada.Add(item);
                }
            }

            //return lista;
        }

        /// <summary>
        /// Inserta en CLC_LTS las filas de ListaEntrada.
        /// </summary>
        /// <param name="estado">"P" = Pendiente (borrador en aprobación), "C" = Confirmado (documento definitivo).</param>
        public bool CrearLotes(string estado = "C")
        {
            bool res = true;
            SAPbobsCOM.UserTable Lotes = null;
            SAPbobsCOM.Recordset oRs = null;

            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                for (int i = 0; i < ListaEntrada.Count; i++)
                {
                    var item = ListaEntrada[i];

                    // CORRECCION: si el mismo número de lote se repite en varias
                    // líneas del mismo documento (mismo artículo+documento+lote+
                    // almacén+estado), se SUMA la cantidad a la fila que ya existe
                    // en vez de crear una fila duplicada. Evita que, al consumirlo
                    // después en una salida, se descuente por duplicado.
                    string queryBusca = $"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"Name\" = '{item.ItemCode}' " +
                                         $"AND \"U_TipDoc\" = '{item.TipDoc}' AND \"U_NumDoc\" = {item.NumDoc} AND \"U_NumLot\" = '{item.NumLot}' " +
                                         $"AND \"U_almacen\" = '{item.almacen}' AND \"U_Estado\" = '{estado}'";
                    oRs.DoQuery(queryBusca);

                    if (oRs.RecordCount > 0)
                    {
                        oRs.MoveFirst();
                        string codeExistente = oRs.Fields.Item("Code").Value.ToString();
                        string queryUpdate = $"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" + {item.CantLote}) WHERE \"Code\" = '{codeExistente}'";
                        oRs.DoQuery(queryUpdate);
                        Console.WriteLine($"Lote ya existía en el documento, se sumó la cantidad ({estado}) {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");
                        continue;
                    }

                    Lotes = oCompany.UserTables.Item("CLC_LTS");

                    Lotes.Name = item.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = item.TipDoc;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = item.NumDoc;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = item.NumLot;
                    Lotes.UserFields.Fields.Item("U_CantLote").Value = item.CantLote;
                    Lotes.UserFields.Fields.Item("U_FchVenLt").Value = item.FchVenLt;
                    Lotes.UserFields.Fields.Item("U_FchAdmLt").Value = item.FchAdmLt;
                    Lotes.UserFields.Fields.Item("U_almacen").Value = item.almacen;
                    Lotes.UserFields.Fields.Item("U_Estado").Value = estado;

                    if (Lotes.Add().Equals(0))
                    {
                        Console.WriteLine($"Datos agregados a tabla lotes exitosamente ({estado}) {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot}");

                    }
                    else
                    {
                        string msgErr = $"Error al agregar datos a tabla de lotes. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipDoc} - {item.NumDoc} - {item.NumLot} - Cant {item.CantLote} - Alm {item.almacen}";
                        Ini.RegistrarError(msgErr);
                        try { oApp.SetStatusBarMessage($"ATENCION: no se pudo guardar el lote {item.NumLot} del articulo {item.ItemCode} ({item.TipDoc} {item.NumDoc}). Avise a sistemas."); } catch { }
                        res = false;
                        continue; // sigue con las demas lineas en vez de perderlas
                    }
                }
                ListaEntrada.Clear();
                return res;

            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error en crear lotes de articulos con lote entrada, {ex.Message} - {ex.StackTrace}");
                try { oApp.SetStatusBarMessage("ATENCION: fallo el guardado de lotes de la entrada. Avise a sistemas."); } catch { }
                ListaEntrada.Clear();
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
                Ini.ReleaseCOMObjetc(oRs);
            }


        }

        /// <summary>
        /// Punto único de entrada para et_FORM_DATA_ADD de los documentos de entrada.
        /// - Si es Borrador (pendiente de aprobación): guarda el lote como Pendiente ("P").
        /// - Si es el documento definitivo: si ya existe un pendiente para ese Tipo+NumDoc
        ///   (venía de un borrador aprobado, tal vez desde otra sesión/usuario), lo confirma.
        ///   Si no existe pendiente (flujo normal, sin aprobación), crea el lote directo
        ///   como Confirmado ("C").
        /// </summary>
        public bool ProcesarAdd(bool esBorrador, string tipo, string objectKey = null, string formUID = null)
        {
            try
            {
                // CORRECCION (cancelacion de factura de compra): la cancelacion de una factura directa es una
                // SALIDA de lotes; si era de una entrada de mercancia no hay lotes y no se guarda nada.
                if (tipo == "OPCH" && !esBorrador && EsCancelacionPorClave(objectKey))
                {
                    return ObtenerSalidaCancelacion().ProcesarAdd(false, tipo, objectKey, formUID);
                }

                // La lista de lotes en memoria solo se usa si fue armada para ESTE documento.
                // Si la lista esta vacia no hay nada de otro documento que proteger: se sigue normal.
                Ini.RegistrarError($"[ProcesarAdd {tipo}] borrador={esBorrador} formUID del evento='{formUID}' dueno de la lista='{ListaOwnerFormUID}' lineas en lista={(ListaEntrada == null ? 0 : ListaEntrada.Count)}");
                bool listaEsDeEsteDocumento = string.IsNullOrEmpty(formUID) || ListaOwnerFormUID == formUID
                                              || ListaEntrada == null || ListaEntrada.Count == 0;

                if (esBorrador)
                {
                    if (!listaEsDeEsteDocumento)
                    {
                        Ini.RegistrarError($"ATENCION: el borrador de {tipo} no tiene lotes asignados desde este formulario ({formUID}); la lista en memoria era del formulario {ListaOwnerFormUID}. No se guardo nada.");
                        try { oApp.SetStatusBarMessage("ATENCION: este documento no tiene lotes asignados. Vuelva a abrirlo y asigne los lotes."); } catch { }
                        return true;
                    }
                    // El numero del documento definitivo todavia no existe. Los lotes
                    // pendientes se guardan con el numero del BORRADOR (llave estable);
                    // cuando el borrador se apruebe y se grabe el documento real, se
                    // vuelven a numerar con el DocNum real (ver mas abajo).
                    int draftEntry = Ini.ObtenerDocEntry(objectKey);
                    if (draftEntry > 0 && ListaEntrada != null)
                    {
                        foreach (var it in ListaEntrada) it.NumDoc = draftEntry;
                    }
                    Ini.RegistrarError($"Documento {tipo} enviado a aprobacion (borrador {draftEntry}). Lotes guardados como Pendiente.");
                    return CrearLotes("P");
                }

                // CORRECCION: usar el DocNum REAL del documento recien grabado (no el que
                // se veia en pantalla antes de guardar, que es solo una prediccion y
                // cambia si otro usuario guarda primero).
                int docReal = Ini.ObtenerDocNumReal(oCompany, tipo, objectKey);

                // CORRECCION (doble descuento, traslado 590074): a veces SAP entrega en el
                // evento una clave (DocEntry) que NO es de un documento real, por ejemplo la
                // del borrador. Antes se seguia con el "numero de pantalla" y se guardaban los
                // lotes como confirmados (C); luego, al llegar el documento real, se confirmaban
                // otra vez los del borrador y el lote quedaba descontado DOS veces.
                // Ahora, si el evento trae una clave pero no existe ese documento, no se guarda
                // nada: el documento real llega en su propio evento.
                if (docReal == 0 && System.Text.RegularExpressions.Regex.IsMatch(objectKey ?? string.Empty, "<DocEntry>\\d+</DocEntry>"))
                {
                    Ini.RegistrarError($"ATENCION: el evento de {tipo} trae una clave ({objectKey}) que no corresponde a un documento real. No se guardo nada; se espera el evento del documento real.");
                    return true;
                }

                string numDoc = docReal > 0 ? docReal.ToString() : ObtenerNumDocFormularioActivo();

                // Si el documento viene de un borrador aprobado, sus lotes pendientes
                // estan guardados con el numero del borrador: se renumeran con el
                // DocNum real y se confirman.
                int draftKey = Ini.ObtenerDraftKey(oCompany, tipo, objectKey);
                if (draftKey > 0 && docReal > 0 && ExistePendiente(tipo, draftKey.ToString()))
                {
                    // CORRECCION: solo se confirman las filas de ESTE borrador (por Code).
                    // CORRECCION (documento aprobado y luego modificado): si cambiaron lotes o
                    // cantidades despues de la aprobacion, prevalece la seleccion nueva.
                    if (listaEsDeEsteDocumento && ListaEntrada != null && ListaEntrada.Count > 0)
                    {
                        ReemplazarPendientesEntradaSiCambiaron(tipo, draftKey);
                    }
                    string codesBorrador = ObtenerCodesPendientesEntrada(tipo, draftKey);
                    if (string.IsNullOrEmpty(codesBorrador)) return true;
                    ReasignarNumeroPendientes(tipo, draftKey, docReal, codesBorrador);
                    if (listaEsDeEsteDocumento && ListaEntrada != null) ListaEntrada.Clear();
                    Ini.RegistrarError($"Borrador {draftKey} aprobado -> {tipo} {docReal}. Lotes confirmados.");
                    return ConfirmarPendientes(tipo, docReal.ToString(), codesBorrador);
                }

                // draftKey == 0: NO viene de borrador, no se busca ningun pendiente
                // (evita confundirlo con el borrador de otro usuario).
                // CORRECCION: ya no se confirman pendientes buscando por el numero real
                // (podian ser filas viejas de otros documentos con ese mismo numero).
                if (draftKey != 0 && !string.IsNullOrEmpty(numDoc) && ExistePendiente(tipo, numDoc))
                {
                    Ini.RegistrarError($"ATENCION: existen lotes pendientes viejos con el numero {tipo} {numDoc} que NO pertenecen a este documento; no se confirmaron. Revisar a mano.");
                }

                if (!listaEsDeEsteDocumento)
                {
                    // La lista en memoria es de OTRO documento: no se toca ni se usa aqui.
                    Ini.RegistrarError($"ATENCION: {tipo} {docReal} se grabo sin lotes asignados desde este formulario ({formUID}); la lista en memoria era del formulario {ListaOwnerFormUID}. No se guardo nada.");
                    try { oApp.SetStatusBarMessage($"ATENCION: el documento {docReal} no tiene lotes asignados. Avise a sistemas."); } catch { }
                    return true;
                }

                if (docReal > 0 && ListaEntrada != null)
                {
                    foreach (var it in ListaEntrada)
                    {
                        if (it.NumDoc != docReal)
                        {
                            Ini.RegistrarError($"Numero de documento corregido en {tipo}: en pantalla era {it.NumDoc}, el real es {docReal}.");
                            it.NumDoc = docReal;
                        }
                    }
                }
                else if (docReal == 0)
                {
                    Ini.RegistrarError($"ATENCION: no se pudo obtener el DocNum real de {tipo} (ObjectKey: {objectKey}). Se usara el numero de pantalla.");
                }

                return CrearLotes("C");
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error en ProcesarAdd de articulos con lote entrada, {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Lee el numero de documento (item "8") del formulario activo.
        /// </summary>
        private string ObtenerNumDocFormularioActivo()
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oDocNum = null;
            try
            {
                oForm = oApp.Forms.ActiveForm;
                oDocNum = (SAPbouiCOM.EditText)oForm.Items.Item("8").Specific;
                return oDocNum.String;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error obteniendo numero de documento activo, {ex.Message}");
                return string.Empty;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oForm);
            }
        }

        /// <summary>
        /// Verifica si hay filas Pendientes en CLC_LTS para un Tipo+NumDoc de documento.
        /// </summary>
        private bool ExistePendiente(string tipo, string numDoc)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $"SELECT COUNT(*) AS \"Total\" FROM \"@CLC_LTS\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'";
                oRs.DoQuery(query);
                if (oRs.RecordCount > 0)
                {
                    oRs.MoveFirst();
                    int total = Convert.ToInt32(oRs.Fields.Item("Total").Value);
                    return total > 0;
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validando pendientes en CLC_LTS, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Pasa de Pendiente ("P") a Confirmado ("C") las filas de CLC_LTS de un Tipo+NumDoc.
        /// </summary>
        /// <summary>
        /// Cambia el numero de los lotes pendientes: del numero del borrador al DocNum real.
        /// </summary>
        private void ReasignarNumeroPendientes(string tipo, int draftKey, int docReal, string codes)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_NumDoc\" = {docReal} WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P' AND \"Code\" IN ({codes})");
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error renumerando lotes pendientes {tipo} {draftKey}->{docReal}: {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Compara la seleccion de lotes en memoria con las filas pendientes (P) del borrador.
        /// Si son distintas (el documento se modifico despues de aprobarse), borra las filas
        /// pendientes viejas y guarda la seleccion nueva como pendiente con el numero del borrador.
        /// </summary>
        private void ReemplazarPendientesEntradaSiCambiaron(string tipo, int draftKey)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"SELECT \"Name\", \"U_NumLot\", \"U_almacen\", \"U_CantLote\" FROM \"@CLC_LTS\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
                var viejas = new System.Collections.Generic.List<string>();
                for (int i = 0; i < oRs.RecordCount; i++)
                {
                    viejas.Add(oRs.Fields.Item("Name").Value.ToString() + "|" + oRs.Fields.Item("U_NumLot").Value.ToString() + "|" +
                               oRs.Fields.Item("U_almacen").Value.ToString() + "|" +
                               Convert.ToDouble(oRs.Fields.Item("U_CantLote").Value).ToString("0.####", CultureInfo.InvariantCulture));
                    oRs.MoveNext();
                }
                var nuevas = ListaEntrada.Select(it => it.ItemCode + "|" + it.NumLot + "|" + it.almacen + "|" +
                                                       it.CantLote.ToString("0.####", CultureInfo.InvariantCulture)).ToList();
                viejas.Sort(StringComparer.Ordinal);
                nuevas.Sort(StringComparer.Ordinal);
                if (viejas.SequenceEqual(nuevas)) return;

                Ini.RegistrarError($"Borrador {draftKey} de {tipo} modificado despues de la aprobacion: se reemplazan {viejas.Count} fila(s) pendiente(s) por la seleccion nueva ({nuevas.Count}).");
                oRs.DoQuery($"DELETE FROM \"@CLC_LTS\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
                foreach (var it in ListaEntrada) it.NumDoc = draftKey;
                CrearLotes("P");
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error reemplazando pendientes del borrador {tipo} {draftKey}: {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Devuelve los Code (entre comillas y separados por coma) de las filas pendientes
        /// ('P') que pertenecen a ESTE borrador. Se usa para confirmar SOLO esas filas y no
        /// otras pendientes viejas que casualmente tengan el mismo numero.
        /// </summary>
        private string ObtenerCodesPendientesEntrada(string tipo, int draftKey)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
                var lista = new System.Collections.Generic.List<string>();
                for (int i = 0; i < oRs.RecordCount; i++)
                {
                    lista.Add("'" + oRs.Fields.Item("Code").Value.ToString().Replace("'", "''") + "'");
                    oRs.MoveNext();
                }
                return string.Join(",", lista);
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error leyendo pendientes del borrador {tipo} {draftKey}: {ex.Message}");
                return string.Empty;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        private bool ConfirmarPendientes(string tipo, string numDoc, string codes = null)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $"UPDATE \"@CLC_LTS\" SET \"U_Estado\" = 'C' WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'" + (string.IsNullOrEmpty(codes) ? "" : $" AND \"Code\" IN ({codes})");
                oRs.DoQuery(query);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error confirmando pendientes en CLC_LTS, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        public void formClose()
        {
            try
            {
                ListaEntrada.Clear();


            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en formClose lotes de articulos con lote entrada, {ex.Message}");
            }
            finally
            {

            }
        }

        internal void DeleteRow(ItemEvent pVal)
        {
            DataTable DTCrear = null;
            SAPbouiCOM.Form oForm = null;



            try
            {
                oForm = oApp.Forms.Item(pVal.FormUID);
                DTCrear = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DTCrea");

                for (int i = 0; i < DTCrear.Rows.Count; i++)
                {
                    if (DTCrear.GetValue("Sel", i).Equals("Y"))
                    {
                        DTCrear.Rows.Remove(i);
                        if (ListaEntrada.Any())
                        {
                            ListaEntrada.RemoveAt((int)DTCrear.GetValue("Index", i));

                        }

                    }
                }

                if (DTCrear.IsEmpty)
                {

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en  DeleteRow de articulos con lote entrada, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(DTCrear);
                Ini.ReleaseCOMObjetc(oForm);


            }

        }

        /// <summary>
        /// Cargar boton dentro de formularios entradas
        /// </summary>
        /// <param name="pVal">Variable de evento</param>
        /// <param name="tipo">Tipo de documento</param>
        internal void LoadControl(ItemEvent pVal, string tipo)
        {
            SAPbouiCOM.Form oForm = oApp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);            
            SAPbouiCOM.Button oButton = null;            
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oRefItem = null;

            try
            {
                oForm.Freeze(true);

                oRefItem = oForm.Items.Item("2");
                oItem = oForm.Items.Add("BtnVerLt", BoFormItemTypes.it_BUTTON);
                oItem.Top = oRefItem.Top;
                oItem.Left = oRefItem.Left+80;
                oItem.Width = oRefItem.Width;
                oItem.Height = oRefItem.Height;
                oItem.Enabled = true;
                oItem.FromPane = oRefItem.FromPane;
                oItem.ToPane = oRefItem.ToPane;
                oItem.LinkTo = "2";
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Ver lotes";
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Console.WriteLine($"Error en LoadControl de articulos con lote entrada, {ex.Message}");
            }
            finally
            {
                oForm.Freeze(false);
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(oButton);
                Ini.ReleaseCOMObjetc(oItem);
                Ini.ReleaseCOMObjetc(oRefItem);
            }

        }

        private static FormLotesSalida salidaCancelacion;

        /// <summary>
        /// Instancia de FormLotesSalida para las cancelaciones de factura directa (su lista de lotes es estatica y se comparte).
        /// </summary>
        private FormLotesSalida ObtenerSalidaCancelacion()
        {
            if (salidaCancelacion == null) salidaCancelacion = new FormLotesSalida(oCompany, oApp);
            return salidaCancelacion;
        }

        /// <summary>
        /// Al grabar: indica si el documento OPCH recien creado es una cancelacion (OPCH.CANCELED = 'C').
        /// </summary>
        private bool EsCancelacionPorClave(string objectKey)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                int docEntry = Ini.ObtenerDocEntry(objectKey);
                if (docEntry <= 0) return false;
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT \"CANCELED\" FROM \"OPCH\" WHERE \"DocEntry\" = " + docEntry);
                return rs.RecordCount > 0 && Convert.ToString(rs.Fields.Item("CANCELED").Value).Trim().Equals("C", StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("EsCancelacionPorClave fallo (se sigue como entrada): " + ex.Message);
                return false;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Factura de proveedores: indica si el documento que se esta creando es la cancelacion de otra factura
        /// (clic derecho > Cancelar). Se reconoce por el campo CANCELED = 'C' o por el titulo de la ventana.
        /// Si algo falla devuelve false (se piden lotes como siempre).
        /// </summary>
        private bool EsCancelacionDeFactura(SAPbouiCOM.Form oForm)
        {
            try
            {
                string cancelado = string.Empty;
                try { cancelado = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("CANCELED", 0).Trim(); } catch { }
                string titulo = string.Empty;
                try { titulo = oForm.Title ?? string.Empty; } catch { }

                bool porTitulo = titulo.IndexOf("ancelaci", StringComparison.OrdinalIgnoreCase) >= 0
                              || titulo.IndexOf("ancellation", StringComparison.OrdinalIgnoreCase) >= 0;
                string baseLinea1 = string.Empty;
                try { baseLinea1 = oForm.DataSources.DBDataSources.Item("PCH1").GetValue("BaseType", 0).Trim(); } catch { }
                bool esCancelacion = cancelado.Equals("C", StringComparison.OrdinalIgnoreCase) || porTitulo || baseLinea1 == "18";
                if (esCancelacion) Ini.RegistrarError($"[OPCH] cancelacion de factura detectada (CANCELED='{cancelado}', titulo='{titulo}', BaseType linea 1='{baseLinea1}').");
                return esCancelacion;
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("EsCancelacionDeFactura fallo (se piden lotes): " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Metodo de mostrar Items de lotes con variable de formulario
        /// </summary>
        /// <param name="oForm">Fomrulario sap</param>
        /// <param name="tipo">Nombre de tabla sap</param>
        internal void MostarItems(SAPbouiCOM.Form oForm, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Entrada = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;
            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes");

                Entrada = oForm;

                if (!Entrada.Mode.Equals(BoFormMode.fm_ADD_MODE)) return ;


                switch (Tipo)
                {
                    case "OIGN":
                        LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("13").Specific;
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Entrada.Items.Item("38").Specific;
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");

                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLote.sql")).ReadToEnd();

                // CORRECCION (cancelacion de factura de compra): el documento "Factura de proveedores - Cancelacion"
                // no mueve inventario por si mismo, no se piden lotes.
                if (Tipo == "OPCH" && EsCancelacionDeFactura(Entrada)) return;

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;
                    //ajuste
                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        switch (Tipo)
                        {
                            case "OPCH":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("20"))
                                {
                                    return;
                                }
                                break;
                            case "ORIN":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("16"))
                                {
                                    return;
                                }
                                break;

                        }
                        switch (Tipo)
                        {

                            case "OIGN":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("9").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGN":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;

                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGN":
                                oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("7").Specific;

                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Entrada.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;

                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);
                        Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }



                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);

                Console.WriteLine("fin de consulta de articulos con lotes");

                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de entrada de lotes");
                showFormMostarLotes(query);
                System.Console.WriteLine("Fin carga de formulario de entrada de lotes");

                return ;

            }
            catch (Exception ex)
            {
                Entrada.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Entrada);
                Ini.ReleaseCOMObjetc(LineasEntrada);
                Ini.ReleaseCOMObjetc(oEditText);
                Ini.ReleaseCOMObjetc(oCantidad);
                Ini.ReleaseCOMObjetc(oAlmacen);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(RowNumber);

            }
        }

        /// <summary>
        /// Mostrar los lotes seleccionados sin validaciones
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        private void showFormMostarLotes(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_Crear = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Crear = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

        
            try
            {
                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTE.srf";


                Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(oPath);
                StreamReader sr = new StreamReader(stream);
                stringForm = sr.ReadToEnd();

                oXmlDocument.LoadXml(stringForm);
                FormUID = oXmlDocument.SelectSingleNode("/application/forms/action/form | /Application/forms/action/form").Attributes["uid"].Value;
                // CORRECCION: si la ventana ya estaba abierta (por ejemplo el
                // usuario vuelve a pedir el formulario sin haberlo cerrado), se
                // cierra primero para poder abrirla de nuevo limpia. Antes esto
                // lanzaba "Form - already exists [66000-11]".
                try
                {
                    SAPbouiCOM.Form oFormExistente = oApp.Forms.Item(FormUID);
                    oFormExistente.Close();
                    Ini.ReleaseCOMObjetc(oFormExistente);
                }
                catch
                {
                    // No existe todavía, es el caso normal. No hay nada que hacer.
                }

                oApp.LoadBatchActions(ref stringForm);

                oForm = oApp.Forms.Item(FormUID);
                grid_Crear = (Grid)oForm.Items.Item("GrdCrear").Specific;
                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;
                DT_Crear = (DataTable)oForm.DataSources.DataTables.Item("DTCrea");
                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //DT_Crear.Rows.Add();


                //query = "SELECT 'N' AS 'Sel',15 AS 'Numero Doc',T0.ItemCode,CAST((1.000000) AS INT) AS 'Cantidad','01' AS 'Almacen'FROM OITM T0 WHERE T0.ItemCode = 'TESTLOTES-H2' AND T0.QryGroup28 = 'Y'";

                DT_LineasDocumento.ExecuteQuery(query);

                if (!DT_LineasDocumento.IsEmpty)
                {

                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = true;
                    grid_LineasDoc.Item.AffectsFormMode = false;


                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                    ColumnGrid.LinkedObjectType = "4";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                    ColumnGrid.LinkedObjectType = "64";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                    {

                        grid_LineasDoc.Columns.Item(i).Editable = false;
                        grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                        grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                    }

                    grid_LineasDoc.Columns.Item("Sel").Editable = true;
                    grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                    grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";
                    try { grid_LineasDoc.Columns.Item("Descripcion").TitleObject.Caption = "Descripción del artículo"; } catch { }
                    //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                    grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                    grid_LineasDoc.AutoResizeColumns();
                    oForm.Visible = true;

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    oForm.Close();
                   
                }

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oForm);
                Ini.ReleaseCOMObjetc(grid_Crear);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Crear);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(ColumnGrid);

            }

        }
    }
}