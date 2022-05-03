using OfficeOpenXml;
using SimpleRESTServiceCRUD.Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleRESTServiceCRUD.Business
{
    public class BoService
    {

        public BoService() { }


        public void ReadDataAndCopy()
        {

        }

        public void MoveDirectory()
        {
            string InitialPath = @"C:\Users\RAUL\Downloads\move1";
            string FinalPath = @"C:\Users\RAUL\Downloads\move2";

            Directory.Move(InitialPath, FinalPath);

        }

        #region Reporte de Errores

        public void ReporteErrores()
        {
            string etapa = "";
            byte[] excel;

            string servidor = string.Empty;
            string asuntoCorreo = string.Empty;
            string emisorCorreo = string.Empty;
            int nroEmpresa = 7;
            DataTable dtServidor = new DataTable();
            DataTable dtPiePagina = new DataTable();
            string cuerpoMensaje = string.Empty;
            string pieCorreo = "PD: ";



            etapa = "SETEO TÍTULOS EXCEL COMPRAS";



            excel = ReporteErrorExcel();



            //Obtenemos datos del servidor
           // dtServidor = _doPagoPresupuesto.ObtenerServidor(7);
            if (dtServidor.Rows.Count > 0)
            {
                servidor = dtServidor.Rows[0]["desc"].ToString().Trim();
                if (servidor != "PLATAFORMA")
                {
                    asuntoCorreo = asuntoCorreo + "Correo de Prueba: ";
                }
            }


           // cuerpoMensaje = "Este es un correo de prueba";

           // //Seteamos datos del correo
           // emisorCorreo = "Pago de Presupuesto Web Luz del Sur";
           // asuntoCorreo = asuntoCorreo + " " + " " + "Reporte de errores en pagos  ";
           // //cuerpoCorreo = cuerpoCorreo + "Se generó contrato para la solicitud N° " + nroSolicitud + ". Se adjunta en PDF.";

           // //Obtenemos datos de pie de página
           //// dtPiePagina = _doPagoPresupuesto.ObtenerPiePagina(nroEmpresa);
           // if (dtPiePagina.Rows.Count != 0)
           // {
           //     pieCorreo = pieCorreo + dtPiePagina.Rows[0]["desc"].ToString().Trim() + " (" + ConstantesPptoWeb.APP_ORIGEN + ")";  //MS021424_20201211_EVE;
           // }

           // ///Seteamos parámetros de correo
           // CorreoParametros correoParametro = new CorreoParametros();
           // correoParametro.Asunto = asuntoCorreo;
           // correoParametro.NombreRemite = emisorCorreo;
           // correoParametro.CorreoRemite = ConstantesPptoWeb.NTF_SPOMAIL;  //MS021424_20201118_EVE
           // correoParametro.Cuerpo = cuerpoMensaje;
           // correoParametro.Pie1 = pieCorreo;
           // correoParametro.Cabecera = "";
           // correoParametro.Pie2 = "";
           // correoParametro.NombrePlantilla = "generico.html";
           // correoParametro.SistemaPlantilla = CorreoParametros.Sistema.Plataforma;
           // correoParametro.SistemaGrupo = ConstantesPptoWeb.NTF_SISTEMA;  //MS021424_20201118_EVE
           // //if (!CopiaA.Equals(""))
           // //     correoParametro.ConCopia = CopiaA;//MS021424_20201123_RMAB
           // Dictionary<string, byte[]> adjuntos = new Dictionary<string, byte[]>();
           // adjuntos.Add("REPORTE_ERRORES" + ".xlsx", excel);
           // correoParametro.Adjunto = adjuntos;//.Add//.Add("",archivofisic);


           // List<string> correosEnvio = new List<string>();
           // correosEnvio.Add("rarmas@stefanini-it.com");
           // //correosEnvio.Add("mlinares@luzdelsur.com.pe");

           // auditorialog("[Ejecutando EnviarCorreoElectronico]");
           // //EnviarCorreoElectronico(correoParametro, correosEnvio);

        }

        public byte[] ReporteErrorExcel()
        {
            string etapa = "";
            int nroFila = 0;
            // string ruta = @"C:\Users\RAUL\Documents\rmab\report";
            string ruta = "";
            byte[] archivoExcelSalida = null;
            OfficeOpenXml.ExcelPackage paqueteExcel = null;
            OfficeOpenXml.ExcelWorksheet hojaExcel = null;
            string nombreArchivo = "ReporteErrores2.xlsx";
            ExcelRange celda;
            System.IO.FileInfo archivoExcel = null;

            ruta = System.IO.Path.GetTempPath();
            if (!ruta.EndsWith("\\"))
                ruta += "\\";


            ruta += nombreArchivo;
            //archivoExcel = new System.IO.FileStream(ruta, System.IO.FileMode.OpenOrCreate);

            archivoExcel = new System.IO.FileInfo(ruta);
            if (archivoExcel.Exists)
                archivoExcel.Delete();
            etapa = "CARGÓ FILESTREAM";

            paqueteExcel = new OfficeOpenXml.ExcelPackage(archivoExcel);
            etapa = "ABRIÓ PAQUETE EXCEL";

            //Hoja de solicitudes
            hojaExcel = paqueteExcel.Workbook.Worksheets.Add("SOLICITUDES");
            etapa = "GENERÓ HOJA EXCEL";


            List<RequerimientoDTO> ListaRequerimientos = new List<RequerimientoDTO>();
            RequerimientoDTO spo = new RequerimientoDTO();
            //spo.nro_solicitud = "3120002";
            //spo.nro_empresa = "7";
            ListaRequerimientos.Add(spo);


            ////cabecera
            //nroFila = 1;
            //SeteaCabeceraExcel(ref hojaExcel, 1, 1, 5, "Reporte de errores en pagos presupuestos web", "LUZ DEL SUR");
            //etapa = "SETEÓ CABECERA EXCEL COMPRAS";
            ////titulos
            //nroFila += 5;

            ////Fecha atención al cliente 
            //hojaExcel.Cells[nroFila + 1, 1].Value = "Fecha ofrecida al cliente";
            //hojaExcel.Cells[nroFila + 1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //hojaExcel.Cells[nroFila + 1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
            //hojaExcel.Cells[nroFila + 1, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //hojaExcel.Cells[nroFila + 1, 1].Style.Font.Bold = true;

            //nroFila += 2;

            //hojaExcel.Cells[nroFila, 1].Value = "Nro_Solicitud";
            //hojaExcel.Cells[nroFila, 2].Value = "Fecha_pagador";
            //hojaExcel.Cells[nroFila, 3].Value = "Hora_pagador";
            //hojaExcel.Cells[nroFila, 4].Value = "Proceso";
            //hojaExcel.Cells[nroFila, 5].Value = "Descripción del error";

            //hojaExcel.Cells[nroFila, 1, nroFila, 5].Style.Font.Size = 12;
            //hojaExcel.Cells[nroFila, 1, nroFila, 5].Style.Font.Bold = true;
            ////hojaExcel.Cells[nroFila, 1, nroFila, 5].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
            ////hojaExcel.Cells[nroFila, 1, nroFila, 17].AutoFilter = true;

            //nroFila += 1;

            //foreach (RequerimientoDTO spo2 in cuerpoMensaje = "Este es un correo de prueba";

            ////Seteamos datos del correo
            //emisorCorreo = "Pago de Presupuesto Web Luz del Sur";
            //asuntoCorreo = asuntoCorreo + " " + " " + "Reporte de errores en pagos  ";
            ////cuerpoCorreo = cuerpoCorreo + "Se generó contrato para la solicitud N° " + nroSolicitud + ". Se adjunta en PDF.";

            ////Obtenemos datos de pie de página
            //// dtPiePagina = _doPagoPresupuesto.ObtenerPiePagina(nroEmpresa);
            //if (dtPiePagina.Rows.Count != 0)
            //{
            //    pieCorreo = pieCorreo + dtPiePagina.Rows[0]["desc"].ToString().Trim() + " (" + ConstantesPptoWeb.APP_ORIGEN + ")";  //MS021424_20201211_EVE;
            //}

            /////Seteamos parámetros de correo
            //CorreoParametros correoParametro = new CorreoParametros();
            //correoParametro.Asunto = asuntoCorreo;
            //correoParametro.NombreRemite = emisorCorreo;
            //correoParametro.CorreoRemite = ConstantesPptoWeb.NTF_SPOMAIL;  //MS021424_20201118_EVE
            //correoParametro.Cuerpo = cuerpoMensaje;
            //correoParametro.Pie1 = pieCorreo;
            //correoParametro.Cabecera = "";
            //correoParametro.Pie2 = "";
            //correoParametro.NombrePlantilla = "generico.html";
            //correoParametro.SistemaPlantilla = CorreoParametros.Sistema.Plataforma;
            //correoParametro.SistemaGrupo = ConstantesPptoWeb.NTF_SISTEMA;  //MS021424_20201118_EVE
            ////if (!CopiaA.Equals(""))
            ////     correoParametro.ConCopia = CopiaA;//MS021424_20201123_RMAB
            //Dictionary<string, byte[]> adjuntos = new Dictionary<string, byte[]>();
            //adjuntos.Add("REPORTE_ERRORES" + ".xlsx", excel);
            //correoParametro.Adjunto = adjuntos;//.Add//.Add("",archivofisic);


            //List<string> correosEnvio = new List<string>();
            //correosEnvio.Add("rarmas@stefanini-it.com");
            ////correosEnvio.Add("mlinares@luzdelsur.com.pe");

            //auditorialog("[Ejecutando EnviarCorreoElectronico]");
            //EnviarCorreoElectronico(correoParametro, correosEnvio);)
            //{
            //    nroFila += 1;
            //    hojaExcel.Cells[nroFila, 1].Value = spo2.nro_solicitud;
            //    hojaExcel.Cells[nroFila, 2].Value = "18/06/2021"; //spo.fec_pagador;
            //    hojaExcel.Cells[nroFila, 3].Value = "17:53"; // spo.hor_pagador;
            //    hojaExcel.Cells[nroFila, 4].Value = "OBRA";
            //    hojaExcel.Cells[nroFila, 5].Value = "No se registró la pr_orden_obra";
            //}

            for (int nroCol = 1; nroCol <= 5; nroCol++)
            {
                hojaExcel.Column(nroCol).AutoFit();
            }

            ListaRequerimientos.Clear();
            spo = new RequerimientoDTO();
            //spo.nro_solicitud = "3120003";
            //spo.nro_empresa = "7";
            ListaRequerimientos.Add(spo);
            nroFila += 1;

            //Solicitudes con error en la pr_orden_obra 
            hojaExcel.Cells[nroFila, 1].Value = "Solicitudes con error en la pr_orden_obra";
            hojaExcel.Cells[nroFila, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            hojaExcel.Cells[nroFila, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
            hojaExcel.Cells[nroFila, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            hojaExcel.Cells[nroFila, 1].Style.Font.Bold = true;


            nroFila += 1;
            hojaExcel.Cells[nroFila, 1].Value = "Nro_Solicitud";
            hojaExcel.Cells[nroFila, 2].Value = "Fecha_pagador";
            hojaExcel.Cells[nroFila, 3].Value = "Hora_pagador";
            hojaExcel.Cells[nroFila, 4].Value = "Proceso";
            hojaExcel.Cells[nroFila, 5].Value = "Descripción del error";

            foreach (RequerimientoDTO spo2 in ListaRequerimientos)
            {
                nroFila += 1;
                hojaExcel.Cells[nroFila, 1].Value = "";
                hojaExcel.Cells[nroFila, 2].Value = "18/06/2021"; //spo.fec_pagador;
                hojaExcel.Cells[nroFila, 3].Value = "17:53"; // spo.hor_pagador;
                hojaExcel.Cells[nroFila, 4].Value = "NUCLI";
                hojaExcel.Cells[nroFila, 5].Value = "Error detallado...";
            }

            for (int nroCol = 1; nroCol <= 5; nroCol++)
            {
                hojaExcel.Column(nroCol).AutoFit();
            }



            etapa = "SETEÓ DATOS EXCEL COMPRAS";

            paqueteExcel.Save();
            etapa = "GUARDANDO EXCEL";

            archivoExcelSalida = System.IO.File.ReadAllBytes(ruta);
            etapa = "RETORNANDO EXCEL";

            return archivoExcelSalida;
        }

        private int SeteaCabeceraExcel(ref OfficeOpenXml.ExcelWorksheet hojaExcel, int nroFilaInicial, int nroColumnaInicial, int nroColumnaFinal, string titulo, string nombreEmpresa)
        {
            try
            {
                int nroFilaFinal;
                string cadenaEstados = "TODOS";
                string cadenaFechas = "TODOS";
                string cadenaFechasRegistro = "TODOS";
                string cadenaFolios = "TODOS";
                string cadenaCR = "TODOS";
                string cadenaMoneda = "TODOS";
                string cadenaTipoDocumento = "TODOS";
                string cadenaNumero = "TODOS";
                string cadenaMontoNeto = "TODOS";
                string cadenaRUC = "TODOS";


                //textos
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial + 7, nroColumnaFinal].Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial + 7, nroColumnaFinal].Style.Font.Bold = false;

                //Fila 1
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial].Value = "Empresa";
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial].Style.Font.Color.SetColor(System.Drawing.Color.White);
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial].Style.Font.Bold = true;

                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial + 1].Value = nombreEmpresa;
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial + 1, nroFilaInicial + 1, nroColumnaFinal - 2].Merge = true;
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial + 1, nroFilaInicial + 1, nroColumnaFinal - 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaInicial + 1, nroFilaInicial + 1, nroColumnaFinal - 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;


                //Fecha y Hora
                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal - 1].Value = "Fecha";
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaFinal - 1].Value = "Hora";
                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal - 1, nroFilaInicial + 1, nroColumnaFinal - 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal - 1, nroFilaInicial + 1, nroColumnaFinal - 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal - 1, nroFilaInicial + 1, nroColumnaFinal - 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal - 1, nroFilaInicial + 1, nroColumnaFinal - 1].Style.Font.Bold = true;

                hojaExcel.Cells[nroFilaInicial, nroColumnaFinal].Value = DateTime.Now.ToString("dd/MM/yyyy");
                hojaExcel.Cells[nroFilaInicial + 1, nroColumnaFinal].Value = DateTime.Now.ToString("HH:mm:ss");

                nroFilaInicial += 3;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial].Value = titulo;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Merge = true;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.Font.Color.SetColor(System.Drawing.Color.White);
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.Font.Bold = true;
                hojaExcel.Cells[nroFilaInicial, nroColumnaInicial, nroFilaInicial, nroColumnaFinal].Style.Font.Size = 14;

                nroFilaFinal = nroFilaInicial + 6;

                return nroFilaFinal;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        //public List<DetalleSPOCarga> gf_RecuperaSolicitudesErrores(string nro_empresa, string proceso)
        //{
        //    List<DetalleSPOCarga> detalle = new List<DetalleSPOCarga>();
        //    DataTable dtListaSPO = null;
        //    DataTable dtMaximo = null;
        //    string ultimook = string.Empty;
        //    string ultimafecha = string.Empty;
        //    try
        //    {
        //        //selecciona el último ok MS021424_20201117_RMAB
        //        dtListaSPO = new DataTable();
        //        dtMaximo = new DataTable();
        //        // dtMaximo = ObtenerUltimoOK();


        //        dtListaSPO = _doPagoPresupuesto.gf_RecuperaSolicitudesErrores("FOC");
        //        if (dtListaSPO != null && dtListaSPO.Rows.Count > 0)
        //        {
        //            foreach (DataRow dr in dtListaSPO.Rows)
        //            {
        //                detalle.Add(new DetalleSPOCarga()
        //                {
        //                    nro_empresa = Helper.ToString(dr["nro_empresa"], String.Empty),
        //                    nro_solicitud = Helper.ToString(dr["nro_solicitud"], String.Empty),
        //                    estado_pago_web = Helper.ToString(dr["estado_pago_web"], String.Empty),
        //                    hor_pagador = Helper.ToString(dr["estado_pago_web"], String.Empty),
        //                    fec_pagador = Helper.ToString(dr["fec_pagador"], String.Empty),
        //                    etapa_proceso = Helper.ToString(dr["etapa_proceso"], String.Empty),
        //                    estado_solicitud = Helper.ToString(dr["estado_solicitud"], String.Empty),
        //                    motivo = Helper.ToString(dr["motivo"], String.Empty),
        //                    estado_web_batch = Helper.ToString(dr["estado_web_batch"], String.Empty),
        //                    estado_proceso_log = Helper.ToString(dr["estado_ejecuc_log"], String.Empty),
        //                    metodo_ejecucion = Helper.ToString(dr["metodo_ejecucion"], String.Empty),
        //                    ind_nuevo_pago = Helper.ToString(dr["ind_nuevo_pago"], String.Empty),
        //                    TipoDocumento = Helper.ToString(dr["tip_documento"], String.Empty)
        //                });
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw LogPagoPresupuesto.LogErrorEx(ex, "Ocurrió un error al obtener la lista de solicitudes a procesar.");
        //    }
        //    finally
        //    {
        //        dtListaSPO = null;
        //    }
        //    return detalle;
        //}
        #endregion
    }
}
