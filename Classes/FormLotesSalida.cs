using Lotes_Hornitos.Modelos;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.Security.Cryptography;

namespace Lotes_Hornitos.Classes
{
    internal class FormLotesSalida
    {
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application oApp;
        public static List<LoteSalida> listaLotesSalida;

        /// <summary>
        /// FormUID del documento de SAP al que pertenece la lista de lotes en memoria.
        /// La lista es estatica (una sola por sesion de SAP), asi que sin este dato dos
        /// documentos abiertos a la vez se mezclaban: el segundo borraba/reusaba los lotes
        /// del primero y al crear los documentos se grababan lotes cruzados.
        /// </summary>
        public static string ListaOwnerFormUID = null;
        public int SelectionManual = 0;

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        /// <param name="oCompany"></param>
        /// <param name="oApp"></param>
        public FormLotesSalida(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application oApp)
        {
            this.oCompany = oCompany;
            this.oApp = oApp;
            // CORRECCION: la lista es estatica y puede haber mas de una instancia de esta clase
            // (la de Ini y la que usa FormLotesEntrada para cancelar facturas): no se reinicia si ya existe.
            if (listaLotesSalida == null) listaLotesSalida = new List<LoteSalida>();
        }


        /// <summary>
        /// Funcion para agregar informacion de lote para salida
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddInfoLotes(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            int cantidad = 0;

            string query = string.Empty;

            bool EnableGridDisponible = true;

            var Sb = new System.Text.StringBuilder();

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                if (SelectionManual.Equals(1))
                {
                    //cantidad = (int)DT_LineasDocumento.GetValue("Cantidad", pVal.Row);

                    //Valida que las cantidades entre cantidad de linea de docuemnto sea igal a la cantidad de lineas en seelccionados
                    EnableGridDisponible = ValidarCantidadesDisponibleToSeleccionados(DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), Convert.ToDouble(DT_LineasDocumento.GetValue("Cantidad", pVal.Row)));

                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteUDT_Salida_Manual.sql")).ReadToEnd();
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", pVal.Row), DT_LineasDocumento.GetValue("Almacen", pVal.Row)).ToString();

                    DT_Disponible.ExecuteQuery(query);
                    DT_Seleccion.Clear();



                    if (!DT_Disponible.IsEmpty)
                    {
                        grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Disponible.Item.Enabled = true;
                        grid_Disponible.Item.AffectsFormMode = true;
                        grid_Disponible.Columns.Item("ItemCode").Visible = false;
                        grid_Disponible.Columns.Item("Tipo doc").Visible = false;
                        grid_Disponible.Columns.Item("Num Doc").Visible = false;
                        grid_Disponible.Columns.Item("Almacen").Visible = false;

                        ColumnGrid = (EditTextColumn)grid_Disponible.Columns.Item("Sel");
                        ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                        for (int i = 0; i < grid_Disponible.Columns.Count; i++)
                        {

                            grid_Disponible.Columns.Item(i).AffectsFormMode = false;
                            grid_Disponible.Columns.Item(i).TitleObject.Sortable = false;
                            grid_Disponible.Columns.Item(i).Editable = false;
                        }

                        grid_Disponible.Columns.Item("Cantidad").Editable = true;
                        grid_Disponible.Columns.Item("Sel").Editable = true;
                        grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;
                        grid_Disponible.Columns.Item("Fecha vence").Editable = false;
                        grid_Disponible.Columns.Item("Fecha Admin").Editable = false;

                        grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
                        grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                        //grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;

                        

                        grid_Disponible.AutoResizeColumns();

                        grid_Disponible.Item.Enabled = EnableGridDisponible;

                    }
                    else
                    {
                        grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Disponible.Item.Enabled = false;
                    }
                }


                infolotesentablasalida(ref DT_Seleccion, ref grid_Seleccion, DT_LineasDocumento.GetValue("ItemCode", pVal.Row).ToString(), DT_LineasDocumento.GetValue("Almacen", pVal.Row).ToString(), (int)DT_LineasDocumento.GetValue("NumLinea", pVal.Row));

                for (int i = 0; i < DT_Disponible.Rows.Count; i++)
                {
                    for (int j = 0; j < listaLotesSalida.Count; j++)
                    {
                        var item = listaLotesSalida[j];
                        var itemcode = DT_Disponible.GetValue("ItemCode", i).ToString();
                            if (item.ItemCode.Equals(DT_Disponible.GetValue("ItemCode", i).ToString()))
                            {
                                if (item.Almacen.Equals(DT_Disponible.GetValue("Almacen", i).ToString()))
                                {
                                    if (item.NumeroLote.Equals(DT_Disponible.GetValue("Numero lote", i).ToString()))
                                    {
                                        DT_Disponible.SetValue("Sel", i, "Y");
                                        DT_Disponible.SetValue("Cantidad", i, item.Cantidad);
                                    }                                        
                                }
                            }
                        

                    }
                }



                Salida.Freeze(false);
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());

            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (grid_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Disponible);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (grid_Seleccion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Seleccion);
                }
                if (DT_Seleccion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Seleccion);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (DT_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Disponible);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }

        /// <summary>
        /// Valida que las cantidades entre cantidad de linea de docuemnto sea igual a la cantidad de lineas en seelccionados
        /// </summary>
        /// <param name="oItem">Item a evaluar las cantidades</param>
        /// <returns></returns>
        private bool ValidarCantidadesDisponibleToSeleccionados(string oItem, double CantidadLinea)
        {
            double cantidadSeleccionada = 0;
            bool result = false;

            if (listaLotesSalida == null || !listaLotesSalida.Any())
            {
                //DesdeLista = false;
                result = true;
                return result;
            }



            for (int j = 0; j < listaLotesSalida.Count; j++)
            {
                var item = listaLotesSalida[j];

                if (item.ItemCode.Equals(oItem))
                {
                    cantidadSeleccionada += Convert.ToDouble(item.Cantidad);

                }
            }




            if (!cantidadSeleccionada.Equals(CantidadLinea))
            {
                result = true;
            }
            return result;
        }

        private void infolotesentablasalida(ref DataTable dT_Seleccion, ref Grid grid_Seleccion, string ItemCode, string almacen, int NumLinea)
        {

            CargarLista(ref dT_Seleccion, listaLotesSalida, ItemCode, almacen, NumLinea);


            //var Sb = new System.Text.StringBuilder();
            //string query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalida.sql")).ReadToEnd();
            //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode, NumDoc, TipoDoc).ToString();


            //dT_Seleccion.ExecuteQuery(query);


            if (!dT_Seleccion.IsEmpty)
            {
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                grid_Seleccion.Item.Enabled = false;
                grid_Seleccion.Item.AffectsFormMode = true;
                grid_Seleccion.Columns.Item("ItemCode").Visible = false;
                grid_Seleccion.Columns.Item("Sel").Visible = false;
                grid_Seleccion.Columns.Item("Tipo doc").Visible = false;
                grid_Seleccion.Columns.Item("Num Doc").Visible = false;
                grid_Seleccion.Columns.Item("NumLinea").Visible = false;



                for (int i = 0; i < grid_Seleccion.Columns.Count; i++)
                {

                    grid_Seleccion.Columns.Item(i).AffectsFormMode = false;
                    grid_Seleccion.Columns.Item(i).TitleObject.Sortable = false;
                    grid_Seleccion.Columns.Item(i).Editable = false;
                }
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_Single;
                grid_Seleccion.AutoResizeColumns();
            }
            else
            {
                grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                grid_Seleccion.Item.Enabled = false;
            }
        }


        private void CargarLista(ref DataTable dT_Seleccion, List<LoteSalida> listaLotesSalida, string _ItemCode, string almacen, int numLinea)
        {
            string query = string.Empty;
            int cont = 0;
            List<string> ListLotes = new List<string>();
            var Sb = new System.Text.StringBuilder();

            dT_Seleccion.Rows.Clear();

            if (listaLotesSalida == null || !listaLotesSalida.Any())
            {
                //DesdeLista = false;
                return;
            }


            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.LotesDatatable_salida.sql")).ReadToEnd();


            for (int i = 0; i < listaLotesSalida.Count; i++)
            {
                var item = listaLotesSalida[i];
                if (item.NumLinea.Equals(numLinea))
                {
                    if (item.ItemCode.Equals(_ItemCode))
                    {
                        if (item.Almacen.Equals(almacen))
                        {
                            ListLotes.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, item.ItemCode, item.NumeroDocumento, item.TipoDocumento, item.Cantidad, item.NumeroLote, item.FechaProceso, item.Almacen, i, item.NumLinea).ToString());
                            Sb.Clear();
                        }
                    }
                }

            }

            if (ListLotes.Any())
            {
                query = string.Join(" UNION ALL ", ListLotes);
                dT_Seleccion.ExecuteQuery(query);
                //DesdeLista = true;
            }
            //else
            //{
            //    DesdeLista = false;
            //}
        }


        /// <summary>
        /// Funciona que limpia los check selecionadso y el datable de lotes disponibles.
        /// </summary>
        /// <param name="pVal"></param>
        internal void CleanData(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;

            try
            {

                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");

                for (int i = 0; i < DT_LineasDocumento.Rows.Count; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).Equals("Y"))
                    {
                        DT_LineasDocumento.SetValue("Sel", i, "N");
                    }
                }

                DT_Disponible.Clear();

                Salida.Freeze(false);

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);

            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (grid_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_Disponible);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (DT_Disponible != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_Disponible);
                }
            }
        }


        /// <summary>
        /// Funcion para mostrar el formulario de salida de lotes
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        private bool showForm(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_LineasDoc = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            List<bool> result = new List<bool>();
            bool res = true, existeenconsulta = false;
            try
            {
                System.Console.WriteLine("Inico carga de formulario XML FRMLTS.srf de salida de lotes");
                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTS.srf";


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

                System.Console.WriteLine("Ajuste de campos de formulario de salida de lotes");

                oForm = oApp.Forms.Item(FormUID);

                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;

                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //validar si ya se agregaron los lotes
                result = ValidasiexisteloteentablaSalida(query);


                System.Console.WriteLine("Ejuctar query de formulario de salida de lotes");

                DT_LineasDocumento.ExecuteQuery(query);


                if (!DT_LineasDocumento.IsEmpty)
                {

                    foreach (bool l in result)
                    {
                        if (!l)
                        {
                            res = l;
                            break;
                        }
                    }
                    if (!res)
                    {
                        System.Console.WriteLine("Formateo de campos de formulario de salida de lotes");
                        grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                        grid_LineasDoc.Item.Enabled = true;
                        grid_LineasDoc.Item.AffectsFormMode = false;


                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                        ColumnGrid.LinkedObjectType = "4";
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                        ColumnGrid.LinkedObjectType = "64";
                        if (Ini.Tipo.Equals("OWTR"))
                        {
                            ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen destino");
                            ColumnGrid.LinkedObjectType = "64";
                        }
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                        ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                        for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                        {

                            grid_LineasDoc.Columns.Item(i).Editable = false;
                            grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                            grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                        }
                        if (SelectionManual.Equals(1))
                        {
                            grid_LineasDoc.Columns.Item("Sel").Editable = true;
                            oForm.Items.Item("BtnAutSel").Enabled = false;
                        }
                        else
                        {
                            grid_LineasDoc.Columns.Item("Sel").Editable = false;
                        }


                        grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";


                        try { grid_LineasDoc.Columns.Item("Descripcion").TitleObject.Caption = "Descripción del artículo"; } catch { }
                        //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                        grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                        grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                        grid_LineasDoc.AutoResizeColumns();
                        System.Console.WriteLine("Mostrar formulario de salida de lotes");
                        oForm.Visible = true;
                    }
                    else
                    {
                        oForm.Close();
                    }

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    System.Console.WriteLine("Cerrar formulario de salida de lotes");
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
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }


        /// <summary>
        /// Funcion que valida que el lote exista en la tabla temporal
        /// </summary>
        /// <param name="query">Consulta con la informacion de los lotes</param>
        /// <returns></returns>
        private List<bool> ValidasiexisteloteentablaSalida(string query)
        {
            double cantidadLote = 0;
            bool existelote = false, existeenconsulta = false;
            SAPbobsCOM.Recordset oRs = null;

            List<bool> result = new List<bool>();

            try
            {

                if (listaLotesSalida == null || !listaLotesSalida.Any())
                {
                    Ini.RegistrarError("[Validacion lotes salida] no hay lotes asignados en memoria: se pide la seleccion de lotes.");
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
                    for (int j = 0; j < listaLotesSalida.Count; j++)
                    {
                        var item = listaLotesSalida[j];

                        System.Console.WriteLine($"desde lista {item.ItemCode}");
                        System.Console.WriteLine($"desde recordset {oRs.Fields.Item("ItemCode").Value.ToString()}");

                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {

                                if (!item.Almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    Ini.RegistrarError($"[Validacion lotes salida] linea {item.NumLinea} {item.ItemCode}: el almacen de los lotes ({item.Almacen}) no coincide con el de la linea ({oRs.Fields.Item("Almacen").Value}). Se descartan y se pide seleccion.");
                                    listaLotesSalida.RemoveAt(j);
                                    result.Add(false);
                                    break;
                                }

                                cantidadLote += item.Cantidad;
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
                            Ini.RegistrarError($"[Validacion lotes salida] linea {oRs.Fields.Item("NumLinea").Value} {oRs.Fields.Item("ItemCode").Value}: cantidad de la linea {Convert.ToDouble(oRs.Fields.Item("Cantidad").Value).ToString(CultureInfo.InvariantCulture)} y lotes asignados {cantidadLote.ToString(CultureInfo.InvariantCulture)}. Se pide seleccion.");
                            result.Add(false);
                            break;
                        }
                    }
                    else
                    {
                        Ini.RegistrarError($"[Validacion lotes salida] linea {oRs.Fields.Item("NumLinea").Value} {oRs.Fields.Item("ItemCode").Value}: no tiene lotes asignados en la lista (lista: {string.Join(", ", listaLotesSalida.Select(x => x.NumLinea + ":" + x.ItemCode))}). Se pide seleccion.");
                        result.Add(false);
                    }
                    oRs.MoveNext();
                }



                //validar que los valores de la tabla tengan los mismos de la consulta.
                int o = 0;
                for (o = 0; o < listaLotesSalida.Count; o++)
                {
                    existeenconsulta = false;
                    var item = listaLotesSalida[o];
                    oRs.MoveFirst();
                    for (int i = 1; i <= oRs.RecordCount; i++)
                    {
                        if (item.NumLinea.Equals((int)oRs.Fields.Item("NumLinea").Value))
                        {
                            if (item.ItemCode.Equals(oRs.Fields.Item("ItemCode").Value.ToString()))
                            {
                                if (item.Almacen.Equals(oRs.Fields.Item("Almacen").Value.ToString()))
                                {
                                    existeenconsulta = true;
                                }
                            }
                        }


                        oRs.MoveNext();
                    }
                    if (!existeenconsulta)
                    {
                        Ini.RegistrarError($"[Validacion lotes salida] lotes de la linea {item.NumLinea} {item.ItemCode} ({item.Almacen}) ya no corresponden a ninguna linea del documento: se descartan.");
                        listaLotesSalida.Remove(item);
                        o--;
                    }

                }

                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en ValidasiexisteloteentablSalida {ex.Message}");
                return result;
            }
            finally
            {
                if (oRs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                }
            }
        }

        //private List<bool> ValidasiexisteloteentablaSalida(string query, string oDocNum, int oCantidad, string almacen)
        //{

        //    int cantidadLote = 0;
        //    bool existelote = false;
        //    try
        //    {

        //        if (listaLotesSalida == null || !listaLotesSalida.Any())
        //            return false;

        //        for (int i = 0; i < listaLotesSalida.Count; i++)
        //        {
        //            var item = listaLotesSalida[i];
        //            if (item.ItemCode.Equals(_ItemCode))
        //            {
        //                if (!item.Almacen.Equals(almacen))
        //                {
        //                    listaLotesSalida.RemoveAt(i);
        //                    return false;                            
        //                }

        //                cantidadLote += item.Cantidad;
        //                existelote = true;
        //            }


        //        }

        //        if (existelote)
        //        {
        //            if (cantidadLote < oCantidad)
        //            {
        //                return false;
        //            }
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //        //oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosenTablaSalida.sql")).ReadToEnd();
        //        //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode, tipo, oDocNum).ToString();
        //        //oRs.DoQuery(query);
        //        //if (oRs == null)
        //        //{
        //        //    return false;
        //        //}
        //        //else
        //        //{
        //        //    oRs.MoveFirst();

        //        //    int valor = (int)oRs.Fields.Item(0).Value;
        //        //    if (valor.Equals(0)) return false;
        //        //}


        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error en ValidasiexisteloteentablaSalida {ex.Message}");
        //        return false;
        //    }
        //    finally
        //    {
        //        //Ini.ReleaseCOMObjetc(oRs);
        //    }
        //}


        /// <summary>
        /// Mostrar los items de la salida de lotes
        /// </summary>
        /// <param name="pVal"></param>
        internal bool MostarItems(ItemEvent pVal, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oAlmacenDestino = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;
            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();
            bool result = false;
            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes salida");

                Salida = oApp.Forms.Item(pVal.FormUID);

                if (!Salida.Mode.Equals(BoFormMode.fm_ADD_MODE)) return true;

                // CORRECCION (dos documentos abiertos a la vez): la lista de lotes en
                // memoria solo vale para UN documento. Si el usuario viene de otro
                // documento, se descarta la lista ajena y este documento empieza limpio.
                if (ListaOwnerFormUID != pVal.FormUID)
                {
                    Ini.RegistrarError($"[{Tipo}] Dueno de la lista de lotes cambia de '{ListaOwnerFormUID}' a '{pVal.FormUID}' (lineas en lista: {(listaLotesSalida == null ? 0 : listaLotesSalida.Count)}).");
                    if (listaLotesSalida != null) listaLotesSalida.Clear();
                    ListaOwnerFormUID = pVal.FormUID;
                }

                switch (Tipo)
                {
                    case "OIGE":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("13").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;
                    case "OWTR":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("23").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalidaTransferencia.sql")).ReadToEnd();
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");



                Console.WriteLine("Armado de query de articulos con lotes salida");

                // CORRECCION (stock): total requerido por articulo+almacen en este documento
                var requeridoPorArticulo = new Dictionary<string, double>();

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        switch (Tipo)
                        {
                            case "OPCH":
                                // Cancelacion de una factura de compra: solo las lineas de facturas directas
                                // (que si movieron inventario) descuentan lotes.
                                if (!LineaCancelacionFacturaMueveInventario(Salida, i))
                                {
                                    continue;
                                }
                                break;
                            case "OINV":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("15"))
                                {
                                    return true;
                                }
                                break;
                            case "ORPC":
                                BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                                if (BaseType.Value.Equals("21"))
                                {
                                    return true;
                                }
                                // CORRECCION (nota de credito con "Sin contabilizacion de stock"): si la linea
                                // tiene marcada esa casilla (RPC1.NoInvtryMv = 'Y') no mueve inventario: no se piden lotes.
                                if (LineaSinContabilizacionStock(Salida, LineasEntrada, i))
                                {
                                    continue;
                                }
                                break;
                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001028").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001029").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470001039").Cells.Item(i).Specific;
                                oAlmacenDestino = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("5").Cells.Item(i).Specific;
                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("7").Specific;
                                break;
                            case "OWTR":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("11").Specific;
                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;
                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);

                        string claveStock = oEditText.String + "|" + oAlmacen.String;
                        double acumStock;
                        requeridoPorArticulo.TryGetValue(claveStock, out acumStock);
                        requeridoPorArticulo[claveStock] = acumStock + cantidad;

                        switch (Tipo)
                        {
                            case "OWTR":
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, oAlmacenDestino.String, RowNumber.Value).ToString());
                                break;
                            default:
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                                break;
                        }
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }

                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                // CORRECCION (cancelacion de factura de compra): si ninguna linea requiere lotes no se abre la ventana
                if (Items.Count == 0) return true;

                query = string.Join(" UNION ALL ", Items);

                // CORRECCION (stock): antes de abrir la ventana de lotes se revisa que SAP tenga
                // existencias en ese almacen. Si no hay, se avisa y no se abre la ventana.
                string faltaStock = VerificarStockSuficiente(requeridoPorArticulo);
                if (!string.IsNullOrEmpty(faltaStock))
                {
                    Ini.RegistrarError($"[{Tipo}] Stock insuficiente en SAP, no se abre la seleccion de lotes: {faltaStock.Replace(Environment.NewLine, " | ")}");
                    oApp.MessageBox("No hay stock suficiente en SAP para estos articulos:" + Environment.NewLine + Environment.NewLine + faltaStock + Environment.NewLine + "Corrija la cantidad o el almacen de la linea antes de continuar.", 1, "Ok");
                    return false;
                }


                Console.WriteLine("fin Armado de query de articulos con lotes salida");


                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de salida de lotes");
                result = showForm(query);
                System.Console.WriteLine("Fin carga de formulario de salida de lotes");

                return result;

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                return false;
            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (LineasEntrada != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(LineasEntrada);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oCantidad != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCantidad);
                }
                if (oAlmacen != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacen);
                }
                if (oAlmacenDestino != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacenDestino);
                }
                if (oDocNum != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocNum);
                }
                if (RowNumber != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RowNumber);
                }
            }
        }

        /// <summary>
        /// Cancelacion de factura de compra (documento "Factura de proveedores - Cancelacion", PCH1):
        /// indica si la linea cancela una linea de factura que SI movio inventario (la factura directa,
        /// no copiada de una entrada de mercancia: PCH1.BaseType original distinto de 20).
        /// Si no se puede determinar devuelve false (no se piden lotes) y deja los datos leidos en el log.
        /// </summary>
        internal bool LineaCancelacionFacturaMueveInventario(SAPbouiCOM.Form oForm, int fila)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("PCH1");
                string bt = ds.GetValue("BaseType", fila - 1).Trim();
                string be = ds.GetValue("BaseEntry", fila - 1).Trim();
                string bl = ds.GetValue("BaseLine", fila - 1).Trim();

                int baseEntry, baseLine;
                if (bt != "18" || !int.TryParse(be, out baseEntry) || !int.TryParse(bl, out baseLine))
                {
                    Ini.RegistrarError($"[OPCH cancelacion] linea {fila}: no se pudo ubicar la factura original (BaseType='{bt}', BaseEntry='{be}', BaseLine='{bl}'): no se piden lotes.");
                    return false;
                }

                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT \"BaseType\" FROM \"PCH1\" WHERE \"DocEntry\" = " + baseEntry + " AND \"LineNum\" = " + baseLine);
                if (rs.RecordCount == 0)
                {
                    Ini.RegistrarError($"[OPCH cancelacion] linea {fila}: la linea {baseLine} de la factura {baseEntry} no existe: no se piden lotes.");
                    return false;
                }

                bool mueve = Convert.ToInt32(rs.Fields.Item("BaseType").Value) != 20;
                Ini.RegistrarError($"[OPCH cancelacion] linea {fila}: factura {baseEntry} linea {baseLine} " + (mueve ? "es directa (movio inventario): se piden lotes de salida." : "viene de una entrada de mercancia: no se piden lotes."));
                return mueve;
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("LineaCancelacionFacturaMueveInventario fallo (no se piden lotes): " + ex.Message);
                return false;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// True si alguna linea de la cancelacion de factura de compra movio inventario.
        /// </summary>
        internal bool CancelacionTieneLineasConMovimiento(SAPbouiCOM.Form oForm)
        {
            try
            {
                int filas = oForm.DataSources.DBDataSources.Item("PCH1").Size;
                for (int i = 1; i <= filas; i++)
                {
                    if (LineaCancelacionFacturaMueveInventario(oForm, i)) return true;
                }
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("CancelacionTieneLineasConMovimiento fallo (no se piden lotes): " + ex.Message);
            }
            return false;
        }

        private static string uidColumnaSinStock = null;
        private static bool diagnosticoColumnasEscrito = false;

        /// <summary>
        /// Nota de credito de compras: indica si la linea tiene marcada la casilla "Sin contabilizacion de stock"
        /// (campo NoInvtryMv). Se lee directo de la matriz: se busca la columna cuyo campo enlazado es NoInvtryMv.
        /// Si algo falla devuelve false (se piden lotes) y deja el detalle en el log.
        /// </summary>
        private bool LineaSinContabilizacionStock(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix matriz, int fila)
        {
            try
            {
                if (uidColumnaSinStock == null)
                {
                    var detalle = new System.Text.StringBuilder();
                    for (int c = 0; c < matriz.Columns.Count; c++)
                    {
                        SAPbouiCOM.Column col = matriz.Columns.Item(c);
                        string alias = string.Empty;
                        try { alias = col.DataBind.Alias ?? string.Empty; } catch { }
                        if (alias.Equals("NoInvtryMv", StringComparison.OrdinalIgnoreCase))
                        {
                            uidColumnaSinStock = col.UniqueID;
                            break;
                        }
                        if (alias.Length > 0) detalle.Append(col.UniqueID + "=" + alias + "; ");
                    }
                    if (uidColumnaSinStock == null && !diagnosticoColumnasEscrito)
                    {
                        diagnosticoColumnasEscrito = true;
                        Ini.RegistrarError("[ORPC] no se encontro la columna 'Sin contabilizacion de stock' (alias NoInvtryMv). Columnas: " + detalle);
                    }
                }
                if (uidColumnaSinStock == null) return false;

                SAPbouiCOM.CheckBox chk = (SAPbouiCOM.CheckBox)matriz.Columns.Item(uidColumnaSinStock).Cells.Item(fila).Specific;
                bool sinStock = chk.Checked;
                if (sinStock) Ini.RegistrarError($"[ORPC] linea {fila}: marcada 'Sin contabilizacion de stock': no se piden lotes.");
                return sinStock;
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("LineaSinContabilizacionStock fallo (se piden lotes): " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Compara lo que pide el documento (por articulo y almacen) contra el stock de SAP (OITW.OnHand).
        /// Solo revisa articulos con QryGroup28 = 'Y' en almacenes con U_UtilLot = 'Y'.
        /// Devuelve el detalle de lo que no alcanza, o vacio si todo alcanza (o si la consulta falla: no bloquea).
        /// </summary>
        private string VerificarStockSuficiente(Dictionary<string, double> requerido)
        {
            SAPbobsCOM.Recordset rs = null;
            var faltan = new System.Text.StringBuilder();
            try
            {
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                foreach (var kv in requerido)
                {
                    string[] partes = kv.Key.Split('|');
                    string item = partes[0].Replace("'", "''");
                    string alm = partes[1].Replace("'", "''");

                    rs.DoQuery("SELECT T0.\"OnHand\" AS \"Stock\" FROM \"OITW\" T0 " +
                               "INNER JOIN \"OITM\" T1 ON T1.\"ItemCode\" = T0.\"ItemCode\" AND T1.\"QryGroup28\" = 'Y' " +
                               "INNER JOIN \"OWHS\" T2 ON T2.\"WhsCode\" = T0.\"WhsCode\" AND T2.\"U_UtilLot\" = 'Y' " +
                               "WHERE T0.\"ItemCode\" = '" + item + "' AND T0.\"WhsCode\" = '" + alm + "'");

                    if (rs.RecordCount == 0) continue; // no es articulo de lote en almacen de lote

                    double stock = Convert.ToDouble(rs.Fields.Item("Stock").Value);
                    if (stock < kv.Value - 0.001)
                    {
                        faltan.AppendLine(partes[0] + " en " + partes[1] + ": pide " +
                            kv.Value.ToString("0.####", CultureInfo.InvariantCulture) + " y hay " +
                            stock.ToString("0.####", CultureInfo.InvariantCulture));
                    }
                }
                return faltan.ToString();
            }
            catch (Exception ex)
            {
                Ini.RegistrarError("VerificarStockSuficiente fallo (no bloquea): " + ex.Message);
                return string.Empty;
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Funcion para mostar los lotes de salida
        /// </summary>
        /// <param name="oForm">Formulario activo</param>
        /// <param name="Tipo">Tipo de tabla de base de datos</param>
        /// <returns></returns>
        internal void MostarItems(SAPbouiCOM.Form oForm, string Tipo)
        {
            string query = string.Empty;
            string path = string.Empty;

            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasEntrada = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oCantidad = null;
            SAPbouiCOM.EditText oAlmacen = null;
            SAPbouiCOM.EditText oAlmacenDestino = null;
            SAPbouiCOM.EditText oDocNum = null;
            SAPbouiCOM.EditText RowNumber = null;
            SAPbouiCOM.EditText BaseType = null;

            double cantidad = 0;

            List<string> Items = new List<string>();
            var Sb = new System.Text.StringBuilder();
            bool result = false;
            try
            {
                Console.WriteLine("Inicio de consulta de articulos con lotes salida");

                Salida = oForm;

                if (!Salida.Mode.Equals(BoFormMode.fm_ADD_MODE)) return;

                switch (Tipo)
                {
                    case "OIGE":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("13").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;
                    case "OWTR":
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("23").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalidaTransferencia.sql")).ReadToEnd();
                        break;
                    default:
                        LineasEntrada = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.ArticulosconLoteSalida.sql")).ReadToEnd();
                        break;

                }
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");



                Console.WriteLine("Armado de query de articulos con lotes salida");

                for (int i = 1; i <= LineasEntrada.RowCount; i++)
                {
                    oEditText = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(oEditText.String))
                    {
                        //switch (Tipo)
                        //{
                        //case "OINV":
                        //    BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                        //    if (BaseType.Value.Equals("15"))
                        //    {
                        //        return true;
                        //    }
                        //    break;
                        //case "ORPC":
                        //    BaseType = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("43").Cells.Item(i).Specific;
                        //    if (BaseType.Value.Equals("21"))
                        //    {
                        //        return true;
                        //    }
                        //    break;
                        //}

                        switch (Tipo)
                        {
                            case "OIGE":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001028").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10001029").Cells.Item(i).Specific;
                                break;
                            default:
                                oCantidad = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("10002117").Cells.Item(i).Specific;
                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("15").Cells.Item(i).Specific;
                                break;
                            case "OWTR":
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("1470001039").Cells.Item(i).Specific;
                                oAlmacenDestino = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("5").Cells.Item(i).Specific;
                                break;
                            default:
                                oAlmacen = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("24").Cells.Item(i).Specific;

                                break;

                        }

                        switch (Tipo)
                        {
                            case "OIGE":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("7").Specific;
                                break;
                            case "OWTR":
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("11").Specific;
                                break;
                            default:
                                oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;
                                break;

                        }

                        RowNumber = (SAPbouiCOM.EditText)LineasEntrada.Columns.Item("0").Cells.Item(i).Specific;
                        cantidad = Ini.ConvertToSystemFormat(oCantidad.String);

                        switch (Tipo)
                        {
                            case "OWTR":
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, oAlmacenDestino.String, RowNumber.Value).ToString());
                                break;
                            default:
                                Items.Add(Sb.AppendFormat(CultureInfo.InvariantCulture, query, oEditText.String, oDocNum.String, oAlmacen.String, cantidad, Tipo, RowNumber.Value).ToString());
                                break;
                        }
                        Sb.Clear();
                        System.Console.WriteLine(oEditText.String);

                    }
                }

                //query = string.Join(",", Items.Select(x => $"'{x}'"));
                query = string.Join(" UNION ALL ", Items);


                Console.WriteLine("fin Armado de query de articulos con lotes salida");


                // Mostar el formulario de lotes.
                System.Console.WriteLine("Mostrar el formulario de salida de lotes");
                showFormMostrarLotes(query);
                System.Console.WriteLine("Fin carga de formulario de salida de lotes");

                //return result;

            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                //return false;
            }
            finally
            {
                if (Salida != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Salida);
                }
                if (LineasEntrada != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(LineasEntrada);
                }
                if (oEditText != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEditText);
                }
                if (oCantidad != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCantidad);
                }
                if (oAlmacen != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacen);
                }
                if (oAlmacenDestino != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oAlmacenDestino);
                }
                if (oDocNum != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocNum);
                }
                if (RowNumber != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RowNumber);
                }
            }
        }

        /// <summary>
        /// Mostar formulariode lotes sin condicionales
        /// </summary>
        /// <param name="query">Consulta con informacion de lotes a mostrar</param>
        /// <returns></returns>
        private void showFormMostrarLotes(string query)
        {
            System.Xml.XmlDocument oXmlDocument = null;

            string stringForm = null, oPath = null, FormUID = null;

            SAPbouiCOM.Form oForm = null;
            Grid grid_LineasDoc = null;
            DataTable DT_LineasDocumento = null;
            EditTextColumn ColumnGrid = null;

            //List<bool> result = new List<bool>();
            //bool res = true, existeenconsulta = false;
            try
            {
                System.Console.WriteLine("Inico carga de formulario XML FRMLTS.srf de salida de lotes");
                oXmlDocument = new System.Xml.XmlDocument();

                oPath = $"{Assembly.GetExecutingAssembly().GetName().Name}.Resources.XML.FRMLTS.srf";


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

                System.Console.WriteLine("Ajuste de campos de formulario de salida de lotes");

                oForm = oApp.Forms.Item(FormUID);

                grid_LineasDoc = (Grid)oForm.Items.Item("GrdLineas").Specific;

                DT_LineasDocumento = (DataTable)oForm.DataSources.DataTables.Item("DTLineas");

                //validar si ya se agregaron los lotes
                //result = ValidasiexisteloteentablaSalida(query);


                System.Console.WriteLine("Ejuctar query de formulario de salida de lotes");

                DT_LineasDocumento.ExecuteQuery(query);


                if (!DT_LineasDocumento.IsEmpty)
                {

                    //foreach (bool l in result)
                    //{
                    //    if (!l)
                    //    {
                    //        res = l;
                    //        break;
                    //    }
                    //}
                    //if (!res)
                    //{
                    System.Console.WriteLine("Formateo de campos de formulario de salida de lotes");
                    oForm.Items.Item("BtnAutSel").Enabled = false;
                    oForm.Items.Item("BtnMan").Enabled = false;
                    oForm.Items.Item("BtnAgre").Enabled = false;
                    oForm.Items.Item("BtnElim").Enabled = false;
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = true;
                    grid_LineasDoc.Item.AffectsFormMode = false;


                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("ItemCode");
                    ColumnGrid.LinkedObjectType = "4";
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen");
                    ColumnGrid.LinkedObjectType = "64";
                    if (Ini.Tipo.Equals("OWTR"))
                    {
                        ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Almacen destino");
                        ColumnGrid.LinkedObjectType = "64";
                    }
                    ColumnGrid = (EditTextColumn)grid_LineasDoc.Columns.Item("Sel");
                    ColumnGrid.Type = BoGridColumnType.gct_CheckBox;


                    for (int i = 0; i < grid_LineasDoc.Columns.Count; i++)
                    {

                        grid_LineasDoc.Columns.Item(i).Editable = false;
                        grid_LineasDoc.Columns.Item(i).TitleObject.Sortable = false;
                        grid_LineasDoc.Columns.Item(i).AffectsFormMode = false;
                    }

                    grid_LineasDoc.Columns.Item("Sel").Editable = true;

                    grid_LineasDoc.Columns.Item("ItemCode").TitleObject.Caption = "Número de articulo";

                    try { grid_LineasDoc.Columns.Item("Descripcion").TitleObject.Caption = "Descripción del artículo"; } catch { }
                    //grid_LineasDoc.Columns.Item("Cantidad solicitada").TitleObject.Caption = "Cantidad";
                    grid_LineasDoc.Columns.Item("Numero Doc").TitleObject.Caption = "Número documento";
                    grid_LineasDoc.Columns.Item("Tipo Doc").Visible = false;
                    grid_LineasDoc.AutoResizeColumns();
                    System.Console.WriteLine("Mostrar formulario de salida de lotes");
                    oForm.Visible = true;
                    //}
                    //else
                    //{
                    //    oForm.Close();
                    //}

                }
                else
                {
                    grid_LineasDoc.SelectionMode = BoMatrixSelect.ms_None;
                    grid_LineasDoc.Item.Enabled = false;
                    System.Console.WriteLine("Cerrar formulario de salida de lotes");
                    oForm.Close();
                    //res = true;

                }


                //return res;

            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
                //return false;

            }
            finally
            {
                if (oForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                if (grid_LineasDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(grid_LineasDoc);
                }
                if (DT_LineasDocumento != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DT_LineasDocumento);
                }
                if (ColumnGrid != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ColumnGrid);
                }
            }
        }

        /// <summary>
        /// Funcion que agrega los lotes de la salida a la tabla de usuario desde boton autoseleccion
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        internal void SalidaLotesAuto(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Sel = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            DataTable DT_Sel = null;
            EditTextColumn ColumnGrid = null;
            SAPbobsCOM.Recordset oRs = null;

            string query = string.Empty;
            bool result = true;
            var Sb = new System.Text.StringBuilder();

            double cantidadSalida = 0;
            var consumido = new Dictionary<string, double>(); // lo ya asignado por lote dentro de este documento

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);

                if (!Salida.Items.Item("BtnAutSel").Enabled)
                {
                    return;
                }

                Salida.Freeze(true);
                listaLotesSalida.Clear();
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Sel = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Sel = (DataTable)Salida.DataSources.DataTables.Item("DTSel");
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                {
                    Console.WriteLine(DT_LineasDocumento.GetValue("ItemCode", i));

                    //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalidaValidacion.sql")).ReadToEnd();
                    //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_LineasDocumento.GetValue("ItemCode", i), DT_LineasDocumento.GetValue("Numero Doc", i), DT_LineasDocumento.GetValue("Tipo Doc", i)).ToString();
                    //oRs.DoQuery(query);
                    //if (oRs == null)
                    //{
                    //    result = false;
                    //    break;
                    //}
                    //else
                    //{
                    //    val = (int)oRs.Fields.Item(0).Value;
                    //    if (!val.Equals(0))
                    //    {
                    //        result = false;
                    //        break;
                    //    }
                    //}

                    string itemLinea = DT_LineasDocumento.GetValue("ItemCode", i).ToString();
                    string almLinea = DT_LineasDocumento.GetValue("Almacen", i).ToString();

                    Sb.Clear();
                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosconLoteUDT_Salida.sql")).ReadToEnd(); // trae la informacion de los lotes desde la tabla de entrada 
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, itemLinea, almLinea).ToString();

                    DT_Disponible.ExecuteQuery(query);

                    if (DT_Disponible.IsEmpty)
                    {
                        Salida.Freeze(false);
                        listaLotesSalida.Clear();
                        oApp.MessageBox($"No hay lote disponible para {itemLinea} de los lotes ingresados en el almacen {almLinea}", 1, "Ok");
                        result = false;
                        break;
                    }

                    cantidadSalida = Convert.ToDouble(DT_LineasDocumento.GetValue("Cantidad", i)); //cantidad que se va a sacar en esta linea
                    double pendiente = cantidadSalida;

                    // Recorre los lotes (ya vienen ordenados por fecha de vencimiento).
                    // A cada lote se le resta lo que ya se asigno a lineas anteriores del mismo documento.
                    for (int j = 0; j < DT_Disponible.Rows.Count; j++)
                    {
                        if (pendiente <= 0.000001) break;

                        string claveLote = $"{itemLinea}|{almLinea}|{DT_Disponible.GetValue("Tipo doc", j)}|{DT_Disponible.GetValue("Num Doc", j)}|{DT_Disponible.GetValue("Numero lote", j)}";
                        double yaAsignado = 0;
                        consumido.TryGetValue(claveLote, out yaAsignado);

                        double libre = Convert.ToDouble(DT_Disponible.GetValue("Cantidad", j)) - yaAsignado;
                        if (libre <= 0.000001) continue; // este lote ya se agoto en lineas anteriores

                        double tomar = Math.Min(pendiente, libre);

                        var lote = new LoteSalida { };
                        var fechavence = DT_Disponible.GetValue("Fecha vence", j);
                        DateTime fecha = Convert.ToDateTime(fechavence);

                        switch (Ini.Tipo)
                        {
                            case "OWTR":
                                lote = new LoteSalida
                                {
                                    ItemCode = DT_Disponible.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = DT_Disponible.GetValue("Numero lote", j).ToString(),
                                    Almacen = almLinea,
                                    AlmacenDestino = DT_LineasDocumento.GetValue("Almacen destino", i).ToString(),
                                    NumeroDocumento = DT_LineasDocumento.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = tomar,
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)DT_LineasDocumento.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = DT_Disponible.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = DT_Disponible.GetValue("Tipo doc", j).ToString()
                                };
                                break;
                            default:
                                lote = new LoteSalida
                                {
                                    ItemCode = DT_Disponible.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = DT_Disponible.GetValue("Numero lote", j).ToString(),
                                    Almacen = almLinea,
                                    NumeroDocumento = DT_LineasDocumento.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = tomar,
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)DT_LineasDocumento.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = DT_Disponible.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = DT_Disponible.GetValue("Tipo doc", j).ToString()
                                };
                                break;
                        }

                        listaLotesSalida.Add(lote);
                        consumido[claveLote] = yaAsignado + tomar;
                        pendiente -= tomar;
                    }

                    if (pendiente > 0.000001)
                    {
                        // No alcanzo el saldo de lotes para esta linea (considerando lo ya asignado a lineas anteriores)
                        Salida.Freeze(false);
                        listaLotesSalida.Clear();
                        oApp.MessageBox($"No hay cantidad suficiente de lote para {itemLinea} en el almacen {almLinea}. Faltan {pendiente.ToString("0.######", CultureInfo.InvariantCulture)} unidades para completar la linea {i + 1} (considerando lo asignado en las lineas anteriores).", 1, "Ok");
                        result = false;
                        break;
                    }

                    //grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
                    //grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);


                    //cantidadSalida = (int)DT_Disponible.GetValue("Cantidad", 0);
                    //cantidadEntrada = (int)DT_LineasDocumento.GetValue("Cantidad", i);

                    //if (cantidadEntrada < cantidadSalida)
                    //{
                    //    Salida.Freeze(false);
                    //    oApp.MessageBox($"No hay cantidad disponible para lote disponible para {DT_LineasDocumento.GetValue("ItemCode", i).ToString()} de los lotes ingresados", 1, "Ok");
                    //    result = false;
                    //    break;
                    //}





                    //Sb.Clear();

                    ////Agregar a la lista de update los valores de cada linea
                    //query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateInfoEntradasLotes.sql")).ReadToEnd();
                    //query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, (cantidadEntrada - cantidadSalida), DT_Disponible.GetValue("ItemCode", 0), DT_Disponible.GetValue("Tipo doc", 0), DT_Disponible.GetValue("Num Doc", 0), DT_Disponible.GetValue("Numero lote", 0)).ToString();

                    //ListaUpdateEntrada.Add(query);
                    //Sb.Clear();

                }

                //if (!result)
                //{
                //    DT_Disponible.Clear();
                //    DT_Sel.Clear();
                //    Salida.Freeze(false);
                //    return;
                //}
                DT_Disponible.Clear();
                Salida.Freeze(false);
                //Salida.Close();
                if (result)
                {
                    SelectionManual = 2;
                    Salida.Items.Item("BtnMan").Enabled = false;
                    Salida.Items.Item("BtnElim").Enabled = false;
                    oApp.MessageBox($"Lotes agregados con exito", 1, "Ok");
                }
                grid_LineasDoc.Columns.Item("Sel").Editable = true;
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                Console.WriteLine($"Error SalidaLotes {ex.Message} - {ex.StackTrace}");
            }
            finally
            {
                Salida.Freeze(false);
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(grid_Sel);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(DT_Sel);
                Ini.ReleaseCOMObjetc(ColumnGrid);
                Ini.ReleaseCOMObjetc(oRs);
            }

        }


        private void setDTDisponible(ref DataTable dT_Disponible, ref Grid grid_Disponible, ref EditTextColumn columnGrid)
        {

            grid_Disponible.SelectionMode = BoMatrixSelect.ms_None;
            grid_Disponible.Item.Enabled = true;
            grid_Disponible.Item.AffectsFormMode = true;
            grid_Disponible.Columns.Item("ItemCode").Visible = false;
            grid_Disponible.Columns.Item("Tipo doc").Visible = false;
            grid_Disponible.Columns.Item("Num Doc").Visible = false;
            grid_Disponible.Columns.Item("Almacen").Visible = false;

            columnGrid = (EditTextColumn)grid_Disponible.Columns.Item("Sel");
            columnGrid.Type = BoGridColumnType.gct_CheckBox;


            for (int i = 0; i < grid_Disponible.Columns.Count; i++)
            {

                grid_Disponible.Columns.Item(i).AffectsFormMode = false;
                grid_Disponible.Columns.Item(i).TitleObject.Sortable = false;
                grid_Disponible.Columns.Item(i).Editable = false;
            }

            //grid_Disponible.Columns.Item("Cantidad").Editable = true;
            //grid_Disponible.Columns.Item("Sel").Editable = true;
            //grid_Disponible.Columns.Item("Sel").AffectsFormMode = true;
            grid_Disponible.Columns.Item("Fecha vence").Editable = false;
            grid_Disponible.Columns.Item("Fecha Admin").Editable = false;

            grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sortable = true;
            grid_Disponible.Columns.Item("Fecha vence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
            //grid_Disponible.Columns.Item("Cantidad").AffectsFormMode = true;


            grid_Disponible.AutoResizeColumns();
        }


        /// <summary>
        /// Funcion que actualiza la tabl CLC_LTS de entradas con las cantidades de las salidas
        /// </summary>
        /// <param name="listaUpdateEntrada"></param>
        private void UpdateTablaEntradaLotes(List<string> listaUpdateEntrada)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                foreach (var lista in listaUpdateEntrada)
                {
                    oRs.DoQuery(lista);
                    if (oRs == null)
                    {
                        Console.WriteLine("Error al ejecutar recordset de update entrada lotes");
                    }
                    else
                    {
                        Console.WriteLine("Exito al actualizar entrada lotes");
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }


        /// <summary>
        /// Funcion que agrega a la tabla CLS_LTSSA los datos de los lotes a sacar.
        /// </summary>
        /// <param name="listaLotesSalida"></param>
        private void AddLotesSalida(List<LoteSalida> listaLotesSalida)
        {
            SAPbobsCOM.UserTable Lotes = null;
            try
            {
                foreach (var l in listaLotesSalida)
                {
                    Console.WriteLine(
                        $"{l.ItemCode} | {l.Almacen} | {l.TipoDocumento} |{l.NumeroDocumento} | {l.NumeroLote} | {l.Cantidad} | {l.FechaProceso}"
                    );

                    Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                    Lotes.Name = l.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = l.TipoDocumento;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = l.NumeroDocumento;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = l.NumeroLote;
                    Lotes.UserFields.Fields.Item("U_Cantidad").Value = l.Cantidad;
                    Lotes.UserFields.Fields.Item("U_FchSal").Value = l.FechaProceso;
                    Lotes.UserFields.Fields.Item("U_Almacen").Value = l.Almacen;

                    if (Lotes.Add().Equals(0))
                    {
                        Console.WriteLine("Datos agregados a tabla lotes salida exitosamente");

                    }
                    else
                    {
                        Console.WriteLine($"Error al agregar datos a tabla de lotes salida. {oCompany.GetLastErrorDescription()} ");

                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
            }

        }


        /// <summary>
        /// Funcion para mostrar informacion de lotes para salida cuando se selecciona manualmente el check de lotes de las lienas del documento
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddInfoLotesSalida(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            Grid grid_Sel = null;
            DataTable DT_Disponible = null;
            DataTable DT_LineasDocumento = null;
            DataTable DT_Sel = null;
            //EditTextColumn ColumnGrid = null;

            string query = string.Empty, Tipo = string.Empty;
            var Sb = new System.Text.StringBuilder();



            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Sel = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                DT_LineasDocumento = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                DT_Sel = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                //if (DT_Disponible.GetValue("Sel", pVal.Row).Equals("N"))
                //{
                //    DT_Sel.Clear();
                //    Salida.Freeze(false);
                //    return;
                //}

                //Console.WriteLine(DT_Disponible.GetValue("ItemCode", pVal.Row));

                for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                {
                    if (DT_LineasDocumento.GetValue("Sel", i).ToString().Equals("Y"))
                    {
                        Tipo = DT_LineasDocumento.GetValue("Tipo Doc", i).ToString();
                    }
                }


                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosUDTSalida.sql")).ReadToEnd();
                query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", pVal.Row), DT_Disponible.GetValue("Numero lote", pVal.Row), Tipo).ToString();


                DT_Sel.ExecuteQuery(query);

                //if (DT_Sel.IsEmpty)
                //{
                //    Sb.Clear();
                //    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoArticulosLoteDisponible.sql")).ReadToEnd();
                //    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", pVal.Row), DT_Disponible.GetValue("Tipo doc", pVal.Row), DT_Disponible.GetValue("Num Doc", pVal.Row), DT_Disponible.GetValue("Numero lote", pVal.Row), DT_Disponible.GetValue("Cantidad", pVal.Row)).ToString();
                //    DT_Sel.ExecuteQuery(query);
                //}

                if (!DT_Sel.IsEmpty)
                {
                    grid_Sel.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Sel.Item.Enabled = false;
                    grid_Sel.Item.AffectsFormMode = true;
                    grid_Sel.Columns.Item("ItemCode").Visible = false;
                    grid_Sel.Columns.Item("Tipo doc").Visible = false;
                    grid_Sel.Columns.Item("Num Doc").Visible = false;



                    for (int i = 0; i < grid_Sel.Columns.Count; i++)
                    {

                        grid_Sel.Columns.Item(i).AffectsFormMode = false;
                        grid_Sel.Columns.Item(i).TitleObject.Sortable = false;
                        grid_Sel.Columns.Item(i).Editable = false;
                    }
                    grid_Sel.AutoResizeColumns();
                }
                else
                {
                    grid_Sel.SelectionMode = BoMatrixSelect.ms_None;
                    grid_Sel.Item.Enabled = false;
                }

                Salida.Freeze(false);
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                Console.WriteLine($"Error SalidaLotes {ex.Message} - {ex.StackTrace}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(grid_Sel);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(DT_LineasDocumento);
                Ini.ReleaseCOMObjetc(DT_Sel);
            }
        }

        /// <summary>
        /// Habilitar la opcion de seleccion manual 
        /// </summary>
        /// <param name="pVal">Variable con valores de evento</param>
        internal void EnableSelectionManual(ItemEvent pVal)
        {

            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);
                if (Salida.Items.Item("BtnMan").Enabled)
                {
                    grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                    DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                    grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;

                    grid_Disponible.Item.Enabled = true;
                    grid_LineasDoc.Columns.Item("Sel").Editable = true;
                    Salida.Items.Item("BtnAutSel").Enabled = false;

                    SelectionManual = 1;

                    query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.SinDatos.sql")).ReadToEnd();
                    query = Sb.AppendFormat(CultureInfo.InvariantCulture, query).ToString();

                    DT_Disponible.ExecuteQuery(query);
                    grid_Disponible.AutoResizeColumns();
                    grid_Disponible.Columns.Item("RowsHeader").Width = 20;
                }


            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Disponible);

            }
        }

        /// <summary>
        /// Adicionar manual los lotes en el grid de seleccionados
        /// </summary>
        /// <param name="pVal"></param>
        internal void AddLoteSel(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Disponible = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();
            List<string> ListaLotesAdd = new List<string>();

            int lineaseleccionada = 0;
            double cantidadDispo = 0, cantidadSel = 0;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                if (!Salida.Items.Item("BtnMan").Enabled)
                {
                    return;
                }
                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                lineaseleccionada = listaLotesSalida.Count;


                for (int i = 0; i < DT_Disponible.Rows.Count; i++)
                {
                    if (DT_Disponible.GetValue("Sel", i).Equals("Y"))
                    {
                        //Validar que la cantidad ingresada no sea mayor a la cantidad disponible
                        var catidaDisponible = DT_Disponible.GetValue("Cantidad disponible", i);
                        var catidaIngresada = DT_Disponible.GetValue("Cantidad", i);

                        cantidadDispo = Convert.ToDouble(catidaDisponible);
                        cantidadSel = Convert.ToDouble(catidaIngresada);

                        if (cantidadDispo < cantidadSel)
                        {
                            oApp.MessageBox($"La cantidad ingresada es mayor a la cantidad disponible.", 1, "Ok");
                            return;
                        }

                        // Descontar lo que ya se asigno de este mismo lote a OTRAS lineas del documento
                        if (cantidadSel > 0 && listaLotesSalida != null)
                        {
                            double usadoOtrasLineas = 0;
                            string itemSel = DT_Disponible.GetValue("ItemCode", i).ToString();
                            string almSel = DT_Disponible.GetValue("Almacen", i).ToString();
                            string tipoSel = DT_Disponible.GetValue("Tipo doc", i).ToString();
                            string numSel = DT_Disponible.GetValue("Num Doc", i).ToString();
                            string loteSel = DT_Disponible.GetValue("Numero lote", i).ToString();

                            int numLineaActual = -1;
                            DataTable DT_LineasTmp = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                            for (int k = 0; k < DT_LineasTmp.Rows.Count; k++)
                            {
                                if (DT_LineasTmp.GetValue("Sel", k).ToString().Equals("Y"))
                                {
                                    numLineaActual = Convert.ToInt32(DT_LineasTmp.GetValue("NumLinea", k));
                                    break;
                                }
                            }
                            Ini.ReleaseCOMObjetc(DT_LineasTmp);

                            foreach (var x in listaLotesSalida)
                            {
                                if (x.NumLinea != numLineaActual && x.ItemCode == itemSel && x.Almacen == almSel
                                    && x.TipoDocumentoEntrada == tipoSel && x.NumeroDocumentoEntrada == numSel && x.NumeroLote == loteSel)
                                {
                                    usadoOtrasLineas += x.Cantidad;
                                }
                            }

                            if (cantidadSel > cantidadDispo - usadoOtrasLineas + 0.000001)
                            {
                                oApp.MessageBox($"El lote {loteSel} solo tiene {(cantidadDispo - usadoOtrasLineas).ToString("0.######", CultureInfo.InvariantCulture)} disponibles, porque {usadoOtrasLineas.ToString("0.######", CultureInfo.InvariantCulture)} ya se asignaron a otras lineas de este documento.", 1, "Ok");
                                return;
                            }
                        }

                        if (cantidadSel > 0)
                        {
                            Sb.Clear();
                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.InfoLotesAddmanual.sql")).ReadToEnd();
                            var fechaVence = DT_Disponible.GetValue("Fecha vence", i);
                            DateTime ofechaVence = Convert.ToDateTime(fechaVence);

                            var fechaAdmin = DT_Disponible.GetValue("Fecha Admin", i);
                            DateTime ofechaAdim = Convert.ToDateTime(fechaVence);

                            query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, DT_Disponible.GetValue("ItemCode", i), DT_Disponible.GetValue("Tipo doc", i), DT_Disponible.GetValue("Num Doc", i), DT_Disponible.GetValue("Numero lote", i), DT_Disponible.GetValue("Cantidad", i), ofechaVence.ToString("dd/MM/yyyy"), ofechaAdim.ToString("dd/MM/yyyy"), DT_Disponible.GetValue("Almacen", i), lineaseleccionada, i).ToString();
                            ListaLotesAdd.Add(query);
                            lineaseleccionada++;

                        }
                    }
                }


                if (ListaLotesAdd.Any())
                {
                    query = string.Join(" UNION ALL ", ListaLotesAdd);
                    DT_Seleccion.ExecuteQuery(query);

                    if (!DT_Seleccion.IsEmpty)
                    {
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Seleccion.Item.Enabled = false;
                        grid_Seleccion.Item.AffectsFormMode = true;
                        grid_Seleccion.Columns.Item("ItemCode").Visible = false;
                        grid_Seleccion.Columns.Item("Tipo doc").Visible = false;
                        grid_Seleccion.Columns.Item("Num Doc").Visible = false;
                        grid_Seleccion.Columns.Item("NumLinea").Visible = false;



                        for (int i = 0; i < grid_Seleccion.Columns.Count; i++)
                        {

                            grid_Seleccion.Columns.Item(i).AffectsFormMode = false;
                            grid_Seleccion.Columns.Item(i).TitleObject.Sortable = false;
                            grid_Seleccion.Columns.Item(i).Editable = false;
                        }
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_Single;
                        grid_Seleccion.AutoResizeColumns();
                    }
                    else
                    {
                        grid_Seleccion.SelectionMode = BoMatrixSelect.ms_None;
                        grid_Seleccion.Item.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());
            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);

            }
        }

        /// <summary>
        /// Metodo para eliminar fila de lote seleccionado
        /// </summary>
        /// <param name="pVal">Variable con valores del evento</param>
        internal void DelLoteSel(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_Disponible = null;
            DataTable DT_Disponible = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);
                Salida.Freeze(true);

                if (!Salida.Items.Item("BtnMan").Enabled)
                {
                    return;
                }

                grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");
                grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");

                for (int i = 0; i < grid_Seleccion.Rows.Count; i++)
                {
                    if (grid_Seleccion.Rows.IsSelected(i))
                    {

                        if (listaLotesSalida.Any())
                        {
                            listaLotesSalida.RemoveAll(x =>
                                                    x.ItemCode == DT_Seleccion.GetValue("ItemCode", i).ToString() &&
                                                    x.NumeroLote == DT_Seleccion.GetValue("Numero lote", i).ToString() &&
                                                    x.Almacen == DT_Seleccion.GetValue("Almacen", i).ToString() &&
                                                    x.NumLinea == (int)DT_Seleccion.GetValue("NumLinea", i));

                            //listaLotesSalida.RemoveAt((int)DT_Seleccion.GetValue("Index", i));

                        }
                        DT_Seleccion.Rows.Remove(i);                        

                        if (!grid_Disponible.Item.Enabled)
                        {
                            grid_Disponible.Item.Enabled = true;

                        }

                        Salida.Mode = BoFormMode.fm_UPDATE_MODE;

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Salida.Freeze(false);
                oApp.SetStatusBarMessage("Error Add-On " + Ini.asm.GetName().Name + " : " + ex.Message.ToString() + " " + ex.StackTrace.ToString());

            }
            finally
            {
                Salida.Freeze(false);

                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(DT_Disponible);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);
            }

            
        }

        /// <summary>
        /// Metodo para validar cantidades y lotes seleccionados pra luego agregarlos a la lista de lotes salida
        /// </summary>
        /// <param name="pVal">Variable con valores de evento</param>
        /// <returns></returns>
        internal bool AddLoteSelToList(ItemEvent pVal)
        {
            SAPbouiCOM.Form Salida = null;
            Grid grid_LineasDoc = null;
            DataTable DT_Lineas = null;
            Grid grid_Seleccion = null;
            DataTable DT_Seleccion = null;
            Grid grid_Disponible = null;
            DataTable DT_Disponible = null;

            double cantidadLineas = 0, cantidadDisponible = 0, cantidadDisponibleSelecionada = 0, cantidadDisponibleLinea = 0, cantidadDisponibleSelecionadaLinea = 0;

            bool result = true;

            string LoteDisponibleLinea = string.Empty, LoteSeleccionadoLinea = string.Empty;

            try
            {
                Salida = oApp.Forms.Item(pVal.FormUID);

                if (Salida.Mode.Equals(BoFormMode.fm_UPDATE_MODE))
                {
                    if (SelectionManual.Equals(1))
                    {
                        DT_Lineas = (DataTable)Salida.DataSources.DataTables.Item("DTLineas");
                        grid_LineasDoc = (Grid)Salida.Items.Item("GrdLineas").Specific;
                        grid_Seleccion = (Grid)Salida.Items.Item("GrdSel").Specific;
                        DT_Seleccion = (DataTable)Salida.DataSources.DataTables.Item("DTSel");
                        grid_Disponible = (Grid)Salida.Items.Item("GrdDisp").Specific;
                        DT_Disponible = (DataTable)Salida.DataSources.DataTables.Item("DTDispo");

                        //Capturar el valor principal de cantidad de la salida
                        for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
                        {
                            if (DT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadLineas = Convert.ToDouble(DT_Lineas.GetValue("Cantidad", i));
                                break;
                            }

                        }

                        //Capturar el valor de cantidad de lote ingresado como disponible en seleccionados
                        for (int i = 0; i <= grid_Seleccion.Rows.Count - 1; i++)
                        {

                            cantidadDisponibleSelecionada += Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", i));

                        }

                        //Capturar el valor de cantidad de lote ingresado como disponible en disponibles
                        for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        {
                            if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadDisponible += Convert.ToDouble(DT_Disponible.GetValue("Cantidad", i));
                            }
                        }

                        //Validar que las cantidades entre principal y disponble ingresada sean iguales
                        if (!cantidadLineas.Equals(cantidadDisponible))
                        {
                            oApp.MessageBox("La cantidad del lote ingresado debe ser igual a la cantidad del documento.", 1, "Ok");
                            result = false;
                            return result;
                        }

                        //Validar que las cantidades entre Disponible y Selecionadas ingresada sean iguales
                        if (!cantidadDisponibleSelecionada.Equals(cantidadDisponible))
                        {
                            oApp.MessageBox("La cantidad del lote ingresado en disponibles debe ser igual a la cantidad del lote ingresado en seleccionados.", 1, "Ok");
                            result = false;
                            return result;
                        }


                        //validar que las cantidades entre disponible y seleccionados sean iguales por linea de cada lote
                        for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        {
                            if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                            {
                                cantidadDisponibleLinea = Convert.ToDouble(DT_Disponible.GetValue("Cantidad", i));
                                LoteDisponibleLinea = DT_Disponible.GetValue("Numero lote", i).ToString();

                                for (int j = 0; j <= grid_Seleccion.Rows.Count - 1; j++)
                                {
                                    var numlinea = (int)DT_Seleccion.GetValue("NumLinea", j);
                                    if (numlinea.Equals(i))
                                    {
                                        cantidadDisponibleSelecionadaLinea = Convert.ToDouble(DT_Seleccion.GetValue("Cantidad", j));
                                        LoteSeleccionadoLinea = DT_Seleccion.GetValue("Numero lote", j).ToString();
                                        break;
                                    }

                                }
                                if (!cantidadDisponibleLinea.Equals(cantidadDisponibleSelecionadaLinea))
                                {
                                    oApp.MessageBox("La cantidad del lote ingresado en disponibles por líneas difiere de lo ingresado en seleccionados.", 1, "Ok");
                                    result = false;
                                    return result;
                                }

                                //Validar que los lotes sean iguales por linea Disponible y seleccionada
                                if (!LoteDisponibleLinea.Equals(LoteSeleccionadoLinea))
                                {
                                    oApp.MessageBox("Existe diferencia entre los lotes seleccionados en disponible y los agregados en seleccionados.", 1, "Ok");
                                    result = false;
                                    return result;
                                }
                            }
                        }

                        ////validar que los lotes entre disponible y seleccionados sean iguales por linea
                        //for (int i = 0; i <= grid_Disponible.Rows.Count - 1; i++)
                        //{
                        //    if (DT_Disponible.GetValue("Sel", i).ToString().Equals("Y"))
                        //    {
                        //        LoteDisponibleLinea = DT_Disponible.GetValue("Numero lote", i).ToString();

                        //        existelote = false;
                        //        for (int j = 0; j <= grid_Seleccion.Rows.Count - 1; j++)
                        //        {

                        //            LoteSeleccionadoLinea = DT_Seleccion.GetValue("Numero lote", j).ToString();
                        //            //Validar que los lotes sean iguales por linea Disponible y seleccionada

                        //            if (LoteDisponibleLinea.Equals(LoteSeleccionadoLinea))
                        //            {
                        //                existelote = true;
                        //                break;
                        //            }                                    
                        //        }
                        //    }
                        //}

                        //if (!existelote)
                        //{
                        //    oApp.MessageBox("Existe diferencia entre los lotes seleccionados en disponible y los agregados en seleccionados.", 1, "Ok");
                        //    result = false;
                        //    return result;
                        //}


                        AddLotesSalidaManualToLits(DT_Seleccion, grid_LineasDoc, DT_Lineas);
                    }
                }


                return result;
            }
            catch (Exception ex)
            {

                oApp.MessageBox($"Error: Add-On {Ini.asm.GetName().Name} - {ex.Message}", 1, "Ok");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(grid_LineasDoc);
                Ini.ReleaseCOMObjetc(DT_Lineas);
                Ini.ReleaseCOMObjetc(grid_Seleccion);
                Ini.ReleaseCOMObjetc(DT_Seleccion);
                Ini.ReleaseCOMObjetc(grid_Disponible);
                Ini.ReleaseCOMObjetc(DT_Disponible);
            }
        }

        /// <summary>
        /// Metodo para agregar los lotes selecionados a la lista de salida
        /// </summary>
        /// <param name="dT_Seleccion">Data table SAP con información de lotes de salida</param>
        private void AddLotesSalidaManualToLits(DataTable dT_Seleccion, Grid grid_LineasDoc, DataTable dT_Lineas)
        {
            var tipo = Ini.Tipo;


            for (int i = 0; i <= grid_LineasDoc.Rows.Count - 1; i++)
            {
                if (dT_Lineas.GetValue("Sel", i).ToString().Equals("Y"))
                {
                    for (int j = 0; j < dT_Seleccion.Rows.Count; j++)
                    {

                        var lote = new LoteSalida { };
                        var fechavence = dT_Seleccion.GetValue("Fecha vence", j);
                        DateTime fecha = Convert.ToDateTime(fechavence);

                        switch (Ini.Tipo)
                        {
                            case "OWTR":
                                lote = new LoteSalida
                                {
                                    ItemCode = dT_Seleccion.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = dT_Seleccion.GetValue("Numero lote", j).ToString(),
                                    Almacen = dT_Lineas.GetValue("Almacen", i).ToString(),
                                    AlmacenDestino = dT_Lineas.GetValue("Almacen destino", i).ToString(),
                                    NumeroDocumento = dT_Lineas.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = dT_Lineas.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = Convert.ToDouble(dT_Seleccion.GetValue("Cantidad", j).ToString()),
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)dT_Lineas.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = dT_Seleccion.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = dT_Seleccion.GetValue("Tipo doc", j).ToString()


                                };
                                break;
                            default:
                                lote = new LoteSalida
                                {
                                    ItemCode = dT_Seleccion.GetValue("ItemCode", j).ToString(),
                                    NumeroLote = dT_Seleccion.GetValue("Numero lote", j).ToString(),
                                    Almacen = dT_Lineas.GetValue("Almacen", i).ToString(),
                                    NumeroDocumento = dT_Lineas.GetValue("Numero Doc", i).ToString(),
                                    TipoDocumento = dT_Lineas.GetValue("Tipo Doc", i).ToString(),
                                    Cantidad = Convert.ToDouble(dT_Seleccion.GetValue("Cantidad", j).ToString()),
                                    FechaProceso = fecha.ToString("dd/MM/yyyy"),
                                    NumLinea = (int)dT_Lineas.GetValue("NumLinea", i),
                                    NumeroDocumentoEntrada = dT_Seleccion.GetValue("Num Doc", j).ToString(),
                                    TipoDocumentoEntrada = dT_Seleccion.GetValue("Tipo doc", j).ToString()

                                };
                                break;
                        }


                        listaLotesSalida.Add(lote);

                    }
                }
            }
        }

        internal void RollBack(ItemEvent pVal, string Tipo)
        {
            SAPbouiCOM.Form Salida = null;
            SAPbouiCOM.Matrix LineasSalida = null;
            EditText ItemCode = null;
            EditText oDocNum = null;
            SAPbobsCOM.Recordset oRs = null, oRsUpdate = null;

            string query = string.Empty;
            var Sb = new System.Text.StringBuilder();

            try
            {
                Console.WriteLine("Inicio de rollback de articulos con lote entrada");

                Salida = oApp.Forms.Item(pVal.FormUID);
                LineasSalida = (SAPbouiCOM.Matrix)Salida.Items.Item("38").Specific;
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRsUpdate = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //dBDataSourceLineasDocumento = Entrada.DataSources.DBDataSources.Item("PDN1");


                for (int i = 1; i <= LineasSalida.RowCount; i++)
                {
                    ItemCode = (SAPbouiCOM.EditText)LineasSalida.Columns.Item("1").Cells.Item(i).Specific;

                    if (!string.IsNullOrEmpty(ItemCode.String))
                    {
                        oDocNum = (SAPbouiCOM.EditText)Salida.Items.Item("8").Specific;

                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.SelectLoteSalidasRollback.sql")).ReadToEnd();
                        query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode.String, oDocNum.String, Tipo).ToString();
                        oRs.DoQuery(query);
                        System.Console.WriteLine($"Query select lote rollback salida {query}");
                        if (oRs == null)
                        {
                            System.Console.WriteLine($"Error con recordset query select lote rollback salida {query}");
                        }
                        else
                        {
                            oRs.MoveFirst();
                            while (!oRs.EoF)
                            {
                                Sb.Clear();
                                query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.UpdateEntradaSalidasRollback.sql")).ReadToEnd();
                                query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, oRs.Fields.Item("U_NumLot"), oRs.Fields.Item("U_Cantidad")).ToString();
                                oRsUpdate.DoQuery(query);
                                if (oRsUpdate == null)
                                {
                                    System.Console.WriteLine($"Error con recordset update query rollback salida, {query}");
                                }
                                else
                                {
                                    System.Console.WriteLine($"Update de entrada exitoso rollback Salida, {query}");

                                }
                            }
                            Sb.Clear();
                        }

                        query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.DeleteArticulosconLoteSalida.sql")).ReadToEnd();
                        query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, ItemCode.String, Tipo, oDocNum.String).ToString();
                        oRs.DoQuery(query);

                        Sb.Clear();
                        System.Console.WriteLine($"Query rollback Salida delete, {query}");


                    }
                }

                Console.WriteLine("Fin de rollback de articulos con lote Salida");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en rollback de articulos con lote Salida, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Salida);
                Ini.ReleaseCOMObjetc(LineasSalida);
                Ini.ReleaseCOMObjetc(ItemCode);
                Ini.ReleaseCOMObjetc(oDocNum);
                Ini.ReleaseCOMObjetc(oRs);
                Ini.ReleaseCOMObjetc(oRsUpdate);
            }
        }

        internal void formClose()
        {
            try
            {
                listaLotesSalida.Clear();
                SelectionManual = 0;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en formClose lotes de articulos con lote salida, {ex.Message}");
            }
            finally
            {

            }
        }

        /// <summary>
        /// Inserta en CLC_LTSSA las filas de listaLotesSalida.
        /// </summary>
        /// <param name="estado">
        /// "P" = Pendiente (borrador en aprobación): SOLO se guarda el registro,
        ///       NO se descuenta cantidad del lote de origen todavía (el documento
        ///       puede ser rechazado, no hay que tocar el saldo disponible).
        /// "C" = Confirmado (documento definitivo, flujo normal sin aprobación):
        ///       se guarda Y se descuenta la cantidad del lote de origen, igual
        ///       que siempre se hizo.
        /// </param>
        internal bool SalidaLotes(string estado = "C")
        {
            bool res = true;
            SAPbobsCOM.UserTable Lotes = null;
            SAPbobsCOM.UserTable LotesEntrada = null;
            List<string> ListaUpdateEntrada = new List<string>();
            SAPbobsCOM.Recordset oRs = null;
            var Sb = new System.Text.StringBuilder();
            var query = string.Empty;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                for (int i = 0; i < listaLotesSalida.Count; i++)
                {
                    var item = listaLotesSalida[i];

                    Lotes = oCompany.UserTables.Item("CLC_LTSSA");

                    Lotes.Name = item.ItemCode;
                    Lotes.UserFields.Fields.Item("U_TipDoc").Value = item.TipoDocumento;
                    Lotes.UserFields.Fields.Item("U_NumDoc").Value = item.NumeroDocumento;
                    Lotes.UserFields.Fields.Item("U_NumLot").Value = item.NumeroLote;
                    Lotes.UserFields.Fields.Item("U_Cantidad").Value = item.Cantidad;
                    Lotes.UserFields.Fields.Item("U_FchSal").Value = DateTime.Now;
                    Lotes.UserFields.Fields.Item("U_almacen").Value = item.Almacen;
                    Lotes.UserFields.Fields.Item("U_Estado").Value = estado;
                    Lotes.UserFields.Fields.Item("U_TDocEnt").Value = item.TipoDocumentoEntrada ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_NDocEnt").Value = item.NumeroDocumentoEntrada ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_AlmDest").Value = item.AlmacenDestino ?? string.Empty;
                    Lotes.UserFields.Fields.Item("U_FchProc").Value = item.FechaProceso ?? string.Empty;

                    string codeFila = string.Empty;
                    if (Lotes.Add().Equals(0))
                    {
                        // CORRECCION (filas en F): se guarda el Code de ESTA fila para marcar como Faltante
                        // solo esa fila y no todas las del mismo documento/articulo/lote.
                        try { codeFila = Convert.ToString(oCompany.GetNewObjectKey()); } catch { codeFila = string.Empty; }
                        Console.WriteLine($"Datos agregados a tabla lotes exitosamente ({estado}) {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");

                        if (estado.Equals("P"))
                        {
                            // Pendiente de aprobación: no se descuenta cantidad ni se
                            // crea la entrada de traslado todavía. Eso se hace en
                            // ConfirmarPendientesSalida() cuando el documento se
                            // confirme como definitivo.
                            continue;
                        }

                        if (Ini.Tipo.Equals("OWTR"))
                        {
                            Sb.Clear();

                            query = new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream($"{Assembly.GetExecutingAssembly().GetName().Name}.Querys.AlmacenManejaLote.sql")).ReadToEnd();
                            query = Sb.AppendFormat(CultureInfo.InvariantCulture, query, item.AlmacenDestino).ToString();
                            oRs.DoQuery(query);
                            if (oRs != null)
                            {
                                oRs.MoveFirst();

                                string almacenDes = oRs.Fields.Item(0).Value.ToString();
                                if (!String.IsNullOrEmpty(almacenDes))
                                {
                                    // CORRECCION: si el traslado tiene 2 líneas con el mismo
                                    // artículo+lote hacia el mismo almacén destino, se SUMA la
                                    // cantidad a la fila que ya existe en vez de duplicarla.
                                    string queryBuscaDestino = $"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"Name\" = '{item.ItemCode}' " +
                                                                $"AND \"U_TipDoc\" = '{item.TipoDocumento}' AND \"U_NumDoc\" = {item.NumeroDocumento} AND \"U_NumLot\" = '{item.NumeroLote}' " +
                                                                $"AND \"U_almacen\" = '{item.AlmacenDestino}'";
                                    oRs.DoQuery(queryBuscaDestino);

                                    if (oRs.RecordCount > 0)
                                    {
                                        oRs.MoveFirst();
                                        string codeExistenteDestino = oRs.Fields.Item("Code").Value.ToString();
                                        oRs.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" + {item.Cantidad}) WHERE \"Code\" = '{codeExistenteDestino}'");
                                        Console.WriteLine($"Lote destino ya existía en el traslado, se sumó cantidad {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                    }
                                    else
                                    {
                                        LotesEntrada = oCompany.UserTables.Item("CLC_LTS");
                                        LotesEntrada.Name = item.ItemCode;
                                        LotesEntrada.UserFields.Fields.Item("U_TipDoc").Value = item.TipoDocumento;
                                        LotesEntrada.UserFields.Fields.Item("U_NumDoc").Value = item.NumeroDocumento;
                                        LotesEntrada.UserFields.Fields.Item("U_NumLot").Value = item.NumeroLote;
                                        LotesEntrada.UserFields.Fields.Item("U_CantLote").Value = item.Cantidad;
                                        LotesEntrada.UserFields.Fields.Item("U_FchVenLt").Value = item.FechaProceso;
                                        LotesEntrada.UserFields.Fields.Item("U_FchAdmLt").Value = DateTime.Now;
                                        LotesEntrada.UserFields.Fields.Item("U_almacen").Value = item.AlmacenDestino;
                                        if (LotesEntrada.Add().Equals(0))
                                        {
                                            Console.WriteLine($"Datos agregados a tabla lotes entrada exitosamente con traslado {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                        }
                                        else
                                        {
                                            Ini.RegistrarError($"Error al agregar datos a tabla de lotes entrada con traslado. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote}");
                                            try { oApp.SetStatusBarMessage($"ATENCION: no se pudo crear el lote destino {item.NumeroLote} del articulo {item.ItemCode}. Avise a sistemas."); } catch { }
                                            res = false;
                                            continue;
                                        }
                                    }
                                }

                            }

                        }
                    }
                    else
                    {
                        Ini.RegistrarError($"Error al agregar datos a tabla de lotes. {oCompany.GetLastErrorDescription()}. {item.ItemCode} - {item.TipoDocumento} - {item.NumeroDocumento} - {item.NumeroLote} - Cant {item.Cantidad} - Alm {item.Almacen}");
                        try { oApp.SetStatusBarMessage($"ATENCION: no se pudo guardar la salida del lote {item.NumeroLote} del articulo {item.ItemCode} ({item.TipoDocumento} {item.NumeroDocumento}). Avise a sistemas."); } catch { }
                        res = false;
                        continue;
                    }


                    //Descuento seguro: la resta SOLO se aplica si en ese instante
                    //todavía alcanza la cantidad en el lote de origen. Si dos
                    //usuarios chocan por el mismo lote, al segundo no le alcanza
                    //y se marca la fila como Faltante ("F") en vez de descontar
                    //de más y dejar el lote en negativo.
                    bool descontado = DescontarCantidadLoteSeguro(item.ItemCode, item.NumeroLote, item.Almacen, item.TipoDocumentoEntrada, item.NumeroDocumentoEntrada, item.Cantidad);
                    if (!descontado)
                    {
                        MarcarSalidaComoFaltante(codeFila, item.TipoDocumento, item.NumeroDocumento, item.ItemCode, item.NumeroLote, item.Cantidad);
                        Console.WriteLine($"ATENCION: no alcanzó cantidad del lote {item.NumeroLote} para {item.ItemCode} (posible choque con otro usuario). Revisar CLC_LTSSA, quedó como Faltante (F).");
                        try { oApp.SetStatusBarMessage($"Atención: no alcanzó cantidad del lote {item.NumeroLote} para {item.ItemCode}. Revisar."); } catch { }
                    }

                }
                listaLotesSalida.Clear();
                return res;

            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error en crear lotes de articulos con lote salida, {ex.Message} - {ex.StackTrace}");
                try { oApp.SetStatusBarMessage("ATENCION: fallo el guardado de lotes de la salida. Avise a sistemas."); } catch { }
                listaLotesSalida.Clear();
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(Lotes);
                Ini.ReleaseCOMObjetc(LotesEntrada);
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Punto único de entrada para et_FORM_DATA_ADD de los documentos de salida.
        /// - Si es Borrador (pendiente de aprobación): guarda el registro como
        ///   Pendiente ("P"), SIN descontar cantidad del lote de origen todavía.
        /// - Si es el documento definitivo: si ya existe un pendiente para ese
        ///   Tipo+NumDoc (venía de un borrador aprobado, tal vez desde otra
        ///   sesión/usuario), descuenta la cantidad, recrea la entrada en el
        ///   almacen destino si es traslado (OWTR), y lo confirma. Si no existe
        ///   pendiente (flujo normal, sin aprobación), procesa directo como
        ///   Confirmado ("C") igual que siempre.
        /// </summary>
        internal bool ProcesarAdd(bool esBorrador, string tipo, string objectKey = null, string formUID = null)
        {
            try
            {
                // La lista de lotes en memoria solo se usa si fue armada para ESTE documento.
                // Si la lista esta vacia no hay nada de otro documento que proteger: se sigue normal.
                Ini.RegistrarError($"[ProcesarAdd {tipo}] borrador={esBorrador} formUID del evento='{formUID}' dueno de la lista='{ListaOwnerFormUID}' lineas en lista={(listaLotesSalida == null ? 0 : listaLotesSalida.Count)}");
                bool listaEsDeEsteDocumento = string.IsNullOrEmpty(formUID) || ListaOwnerFormUID == formUID
                                              || listaLotesSalida == null || listaLotesSalida.Count == 0;

                if (esBorrador)
                {
                    if (!listaEsDeEsteDocumento)
                    {
                        Ini.RegistrarError($"ATENCION: el borrador de {tipo} no tiene lotes asignados desde este formulario ({formUID}); la lista en memoria era del formulario {ListaOwnerFormUID}. No se guardo nada.");
                        try { oApp.SetStatusBarMessage("ATENCION: este documento no tiene lotes asignados. Vuelva a abrirlo y asigne los lotes."); } catch { }
                        return true;
                    }
                    // El numero del documento definitivo todavia no existe. Los registros
                    // pendientes se guardan con el numero del BORRADOR (llave estable);
                    // cuando el borrador se apruebe y se grabe el documento real, se
                    // renumeran con el DocNum real.
                    int draftEntry = Ini.ObtenerDocEntry(objectKey);
                    if (draftEntry > 0 && listaLotesSalida != null)
                    {
                        foreach (var it in listaLotesSalida) it.NumeroDocumento = draftEntry.ToString();
                    }
                    Ini.RegistrarError($"Documento {tipo} enviado a aprobacion (borrador {draftEntry}). Salida de lote guardada como Pendiente.");
                    return SalidaLotes("P");
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

                // Si el documento viene de un borrador aprobado, sus registros pendientes
                // estan guardados con el numero del borrador: se renumeran con el
                // DocNum real y se confirman.
                int draftKey = Ini.ObtenerDraftKey(oCompany, tipo, objectKey);
                if (draftKey > 0 && docReal > 0 && ExistePendienteSalida(tipo, draftKey.ToString()))
                {
                    // CORRECCION: se toman los Code de las filas de ESTE borrador ANTES de
                    // renumerar, y solo esas se confirman. Antes se confirmaban todas las
                    // pendientes con el mismo numero (incluidas viejas de otros documentos).
                    // CORRECCION (documento aprobado y luego modificado): si el usuario cambio
                    // las cantidades/lotes DESPUES de la aprobacion, la ventana de lotes se abrio
                    // de nuevo y la lista en memoria tiene la seleccion NUEVA, pero las filas
                    // pendientes del borrador conservaban la seleccion vieja. Si son distintas,
                    // prevalece la seleccion nueva (la que el usuario ve en el documento).
                    if (listaEsDeEsteDocumento && listaLotesSalida != null && listaLotesSalida.Count > 0)
                    {
                        ReemplazarPendientesSalidaSiCambiaron(tipo, draftKey);
                    }
                    string codesBorrador = ObtenerCodesPendientesSalida(tipo, draftKey);
                    if (string.IsNullOrEmpty(codesBorrador)) return true;
                    ReasignarNumeroPendientesSalida(tipo, draftKey, docReal, codesBorrador);
                    if (listaEsDeEsteDocumento) listaLotesSalida.Clear();
                    Ini.RegistrarError($"Borrador {draftKey} aprobado -> {tipo} {docReal}. Salidas de lote confirmadas.");
                    return ConfirmarPendientesSalida(tipo, docReal.ToString(), codesBorrador);
                }

                // draftKey == 0: NO viene de borrador, no se busca ningun pendiente
                // (evita confundirlo con el borrador de otro usuario).
                // CORRECCION: ya no se confirman pendientes buscando por el numero real
                // (podian ser filas viejas de otros documentos con ese mismo numero).
                if (draftKey != 0 && !string.IsNullOrEmpty(numDoc) && ExistePendienteSalida(tipo, numDoc))
                {
                    Ini.RegistrarError($"ATENCION: existen salidas pendientes viejas con el numero {tipo} {numDoc} que NO pertenecen a este documento; no se confirmaron. Revisar a mano.");
                }

                if (!listaEsDeEsteDocumento)
                {
                    // La lista en memoria es de OTRO documento: no se toca ni se usa aqui.
                    Ini.RegistrarError($"ATENCION: {tipo} {docReal} se grabo sin lotes asignados desde este formulario ({formUID}); la lista en memoria era del formulario {ListaOwnerFormUID}. No se guardo nada.");
                    try { oApp.SetStatusBarMessage($"ATENCION: el documento {docReal} no tiene lotes asignados. Avise a sistemas."); } catch { }
                    return true;
                }

                if (docReal > 0 && listaLotesSalida != null)
                {
                    foreach (var it in listaLotesSalida)
                    {
                        if (it.NumeroDocumento != docReal.ToString())
                        {
                            Ini.RegistrarError($"Numero de documento corregido en {tipo}: en pantalla era {it.NumeroDocumento}, el real es {docReal}.");
                            it.NumeroDocumento = docReal.ToString();
                        }
                    }
                }
                else if (docReal == 0)
                {
                    Ini.RegistrarError($"ATENCION: no se pudo obtener el DocNum real de {tipo} (ObjectKey: {objectKey}). Se usara el numero de pantalla.");
                }

                return SalidaLotes("C");
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error en ProcesarAdd de articulos con lote salida, {ex.Message}");
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
        /// Verifica si hay filas Pendientes en CLC_LTSSA para un Tipo+NumDoc de documento.
        /// </summary>
        private bool ExistePendienteSalida(string tipo, string numDoc)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $"SELECT COUNT(*) AS \"Total\" FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'";
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
                Console.WriteLine($"Error validando pendientes en CLC_LTSSA, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Descuenta la cantidad del lote de origen para cada fila Pendiente de
        /// CLC_LTSSA que coincida con Tipo+NumDoc, y las pasa a Confirmado ("C").
        /// Si el tipo es OWTR (traslado) y el almacen destino maneja lotes,
        /// además recrea la entrada del lote en el almacen destino (igual que
        /// hace SalidaLotes() en el flujo normal sin aprobación).
        /// </summary>
        /// <summary>
        /// Cambia el numero de las salidas pendientes: del numero del borrador al DocNum real.
        /// </summary>
        private void ReasignarNumeroPendientesSalida(string tipo, int draftKey, int docReal, string codes)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"UPDATE \"@CLC_LTSSA\" SET \"U_NumDoc\" = {docReal} WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P' AND \"Code\" IN ({codes})");
            }
            catch (Exception ex)
            {
                Ini.RegistrarError($"Error renumerando salidas pendientes {tipo} {draftKey}->{docReal}: {ex.Message}");
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
        private void ReemplazarPendientesSalidaSiCambiaron(string tipo, int draftKey)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"SELECT \"Name\", \"U_NumLot\", \"U_Almacen\", \"U_Cantidad\" FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
                var viejas = new System.Collections.Generic.List<string>();
                for (int i = 0; i < oRs.RecordCount; i++)
                {
                    viejas.Add(oRs.Fields.Item("Name").Value.ToString() + "|" + oRs.Fields.Item("U_NumLot").Value.ToString() + "|" +
                               oRs.Fields.Item("U_Almacen").Value.ToString() + "|" +
                               Convert.ToDouble(oRs.Fields.Item("U_Cantidad").Value).ToString("0.####", CultureInfo.InvariantCulture));
                    oRs.MoveNext();
                }
                var nuevas = listaLotesSalida.Select(it => it.ItemCode + "|" + it.NumeroLote + "|" + it.Almacen + "|" +
                                                           it.Cantidad.ToString("0.####", CultureInfo.InvariantCulture)).ToList();
                viejas.Sort(StringComparer.Ordinal);
                nuevas.Sort(StringComparer.Ordinal);
                if (viejas.SequenceEqual(nuevas)) return;

                Ini.RegistrarError($"Borrador {draftKey} de {tipo} modificado despues de la aprobacion: se reemplazan {viejas.Count} fila(s) pendiente(s) por la seleccion nueva ({nuevas.Count}).");
                oRs.DoQuery($"DELETE FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
                foreach (var it in listaLotesSalida) it.NumeroDocumento = draftKey.ToString();
                SalidaLotes("P");
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
        private string ObtenerCodesPendientesSalida(string tipo, int draftKey)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs.DoQuery($"SELECT \"Code\" FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {draftKey} AND \"U_Estado\" = 'P'");
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

        private bool ConfirmarPendientesSalida(string tipo, string numDoc, string codes = null)
        {
            SAPbobsCOM.Recordset oRs = null;
            SAPbobsCOM.Recordset oRsAux = null;
            SAPbobsCOM.UserTable LotesEntrada = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRsAux = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string filtroCodes = string.IsNullOrEmpty(codes) ? "" : $" AND \"Code\" IN ({codes})";
                string query = $"SELECT \"Code\", \"Name\" AS \"ItemCode\", \"U_NumLot\", \"U_Cantidad\", \"U_Almacen\", \"U_TDocEnt\", \"U_NDocEnt\", \"U_AlmDest\", \"U_FchProc\" " +
                               $"FROM \"@CLC_LTSSA\" WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'" + filtroCodes;
                oRs.DoQuery(query);

                if (oRs.RecordCount == 0)
                {
                    return true; // nada pendiente que confirmar
                }

                for (int i = 0; i < oRs.RecordCount; i++)
                {
                    string codeFilaPend = oRs.Fields.Item("Code").Value.ToString();
                    string itemCode = oRs.Fields.Item("ItemCode").Value.ToString();
                    string numLot = oRs.Fields.Item("U_NumLot").Value.ToString();
                    double cantidad = Convert.ToDouble(oRs.Fields.Item("U_Cantidad").Value);
                    string almacen = oRs.Fields.Item("U_Almacen").Value.ToString();
                    string tDocEnt = oRs.Fields.Item("U_TDocEnt").Value.ToString();
                    string nDocEnt = oRs.Fields.Item("U_NDocEnt").Value.ToString();
                    string almacenDest = oRs.Fields.Item("U_AlmDest").Value.ToString();
                    string fechaProceso = oRs.Fields.Item("U_FchProc").Value.ToString();

                    // 1) Descontar la cantidad del lote de origen.
                    // Descuento seguro: la resta SOLO se aplica si en ese instante
                    // todavía alcanza (protegido por la base de datos, evita choques
                    // entre usuarios concurrentes).
                    bool descontado = DescontarCantidadLoteSeguro(itemCode, numLot, almacen, tDocEnt, nDocEnt, cantidad);
                    if (!descontado)
                    {
                        MarcarSalidaComoFaltante(codeFilaPend, tipo, numDoc, itemCode, numLot, cantidad);
                        Console.WriteLine($"ATENCION: al confirmar la aprobación, no alcanzó cantidad del lote {numLot} para {itemCode}. Quedó como Faltante (F).");
                    }

                    // 2) Si es traslado (OWTR) y el almacen destino maneja lotes,
                    //    recrear la entrada del lote en el almacen destino.
                    if (tipo.Equals("OWTR") && !string.IsNullOrEmpty(almacenDest))
                    {
                        string queryAlmacen = $"SELECT T1.\"WhsCode\" FROM OWHS T1 WHERE T1.\"WhsCode\" = '{almacenDest}' AND T1.\"U_UtilLot\" = 'Y'";
                        oRsAux.DoQuery(queryAlmacen);

                        if (oRsAux.RecordCount > 0)
                        {
                            oRsAux.MoveFirst();
                            string whsDestValido = oRsAux.Fields.Item(0).Value.ToString();

                            if (!string.IsNullOrEmpty(whsDestValido))
                            {
                                // CORRECCION: si el traslado tiene 2 líneas con el mismo
                                // artículo+lote hacia el mismo almacén destino, se SUMA la
                                // cantidad a la fila que ya existe en vez de duplicarla.
                                string queryBuscaDestino = $"SELECT \"Code\" FROM \"@CLC_LTS\" WHERE \"Name\" = '{itemCode}' " +
                                                            $"AND \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_NumLot\" = '{numLot}' " +
                                                            $"AND \"U_almacen\" = '{almacenDest}' AND \"U_Estado\" = 'C'";
                                oRsAux.DoQuery(queryBuscaDestino);

                                if (oRsAux.RecordCount > 0)
                                {
                                    oRsAux.MoveFirst();
                                    string codeExistenteDestino = oRsAux.Fields.Item("Code").Value.ToString();
                                    oRsAux.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" + {cantidad}) WHERE \"Code\" = '{codeExistenteDestino}'");
                                    Console.WriteLine($"Lote destino ya existía en el traslado aprobado, se sumó cantidad {itemCode} - {tipo} - {numDoc} - {numLot}");
                                }
                                else
                                {
                                    LotesEntrada = oCompany.UserTables.Item("CLC_LTS");
                                    LotesEntrada.Name = itemCode;
                                    LotesEntrada.UserFields.Fields.Item("U_TipDoc").Value = tipo;
                                    LotesEntrada.UserFields.Fields.Item("U_NumDoc").Value = numDoc;
                                    LotesEntrada.UserFields.Fields.Item("U_NumLot").Value = numLot;
                                    LotesEntrada.UserFields.Fields.Item("U_CantLote").Value = cantidad;
                                    LotesEntrada.UserFields.Fields.Item("U_FchVenLt").Value = fechaProceso;
                                    LotesEntrada.UserFields.Fields.Item("U_FchAdmLt").Value = DateTime.Now;
                                    LotesEntrada.UserFields.Fields.Item("U_almacen").Value = almacenDest;
                                    LotesEntrada.UserFields.Fields.Item("U_Estado").Value = "C";

                                    if (!LotesEntrada.Add().Equals(0))
                                    {
                                        Console.WriteLine($"Error al confirmar entrada de traslado con aprobación. {oCompany.GetLastErrorDescription()}. {itemCode} - {tipo} - {numDoc} - {numLot}");
                                        return false;
                                    }
                                    Ini.ReleaseCOMObjetc(LotesEntrada);
                                    LotesEntrada = null;
                                }
                            }
                        }
                    }

                    oRs.MoveNext();
                }

                string queryConfirma = $"UPDATE \"@CLC_LTSSA\" SET \"U_Estado\" = 'C' WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"U_Estado\" = 'P'" + filtroCodes;
                oRsAux.DoQuery(queryConfirma);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error confirmando pendientes de salida en CLC_LTSSA, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
                Ini.ReleaseCOMObjetc(oRsAux);
                Ini.ReleaseCOMObjetc(LotesEntrada);
            }
        }

        /// <summary>
        /// Descuenta la cantidad del lote de origen de forma SEGURA: la resta solo
        /// se aplica si en ese instante exacto todavía alcanza (la condición
        /// "U_CantLote >= cantidad" va dentro del mismo UPDATE, así que queda
        /// protegida por la base de datos). Si dos usuarios chocan por el mismo
        /// lote casi al mismo tiempo, al que llegue después esta función le
        /// devuelve false en vez de dejar el lote en negativo.
        /// </summary>
        private bool DescontarCantidadLoteSeguro(string itemCode, string numLot, string almacen, string tipDocEntrada, string numDocEntrada, double cantidad)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string filtro = $"\"Name\" = '{itemCode}' AND \"U_NumLot\" = '{numLot}' AND \"U_almacen\" = '{almacen}' " +
                                 $"AND \"U_TipDoc\" = '{tipDocEntrada}' AND \"U_NumDoc\" = {numDocEntrada}";

                // Cantidad ANTES de intentar el descuento (para poder comparar después).
                oRs.DoQuery($"SELECT \"U_CantLote\" FROM \"@CLC_LTS\" WHERE {filtro}");
                if (oRs.RecordCount == 0)
                {
                    Console.WriteLine($"No se encontró el lote de origen {numLot} de {itemCode} en {almacen} ({tipDocEntrada} {numDocEntrada}).");
                    return false;
                }
                oRs.MoveFirst();
                double cantidadAntes = Convert.ToDouble(oRs.Fields.Item("U_CantLote").Value);

                // Descuento protegido: SOLO se aplica si en este instante alcanza.
                oRs.DoQuery($"UPDATE \"@CLC_LTS\" SET \"U_CantLote\" = (\"U_CantLote\" - {cantidad}) WHERE {filtro} AND \"U_CantLote\" >= {cantidad}");

                // Cantidad DESPUES: si de verdad se aplicó, bajó exactamente "cantidad".
                oRs.DoQuery($"SELECT \"U_CantLote\" FROM \"@CLC_LTS\" WHERE {filtro}");
                oRs.MoveFirst();
                double cantidadDespues = Convert.ToDouble(oRs.Fields.Item("U_CantLote").Value);

                bool seAplico = Math.Abs((cantidadAntes - cantidadDespues) - cantidad) < 0.0001;
                return seAplico;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en descuento seguro de lote, {ex.Message}");
                return false;
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

        /// <summary>
        /// Marca una fila de CLC_LTSSA como Faltante ("F") cuando, al momento de
        /// descontar, ya no alcanzaba la cantidad en el lote de origen (choque
        /// entre usuarios concurrentes). Queda ahí para revisión manual, en vez
        /// de aparecer como Confirmada ("C") sin haberse podido descontar de verdad.
        /// </summary>
        private void MarcarSalidaComoFaltante(string code, string tipo, string numDoc, string itemCode, string numLot, double cantidad)
        {
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                oRs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query;
                if (!string.IsNullOrWhiteSpace(code) && !code.Contains("<"))
                {
                    // CORRECCION: solo la fila que no se pudo descontar. Antes se marcaban TODAS las filas
                    // del mismo documento, articulo y lote (incluidas las que si se habian descontado bien).
                    query = $"UPDATE \"@CLC_LTSSA\" SET \"U_Estado\" = 'F' WHERE \"Code\" = '{code.Trim()}'";
                }
                else
                {
                    // Sin Code: se limita al menos a las filas con la misma cantidad que aun no estan en F.
                    query = $"UPDATE \"@CLC_LTSSA\" SET \"U_Estado\" = 'F' " +
                            $"WHERE \"U_TipDoc\" = '{tipo}' AND \"U_NumDoc\" = {numDoc} AND \"Name\" = '{itemCode}' AND \"U_NumLot\" = '{numLot}' " +
                            $"AND \"U_Cantidad\" = {cantidad.ToString(CultureInfo.InvariantCulture)} AND \"U_Estado\" <> 'F'";
                }
                oRs.DoQuery(query);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error marcando salida como Faltante, {ex.Message}");
            }
            finally
            {
                Ini.ReleaseCOMObjetc(oRs);
            }
        }

    }
}