using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using System.Xml.Linq;
//using Microsoft.Build.Framework;
namespace inElectra.Ine024
{
   public class Claseine24
   {
      [CommandMethod("INE24")]
      public void Main()
      {
         Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
         Editor editor2 = doc.Editor;
         string outputName;
         string path3 = Directory.GetCurrentDirectory();
         outputName = string.Concat(path3, "/LoopConfig.xml");
         editor2.WriteMessage("\n OUTPUT {0}", outputName);
         if (File.Exists(outputName))
         {
            int i = 0;
            string pathtSignaList = "";
            string pathtTypical = "";
            string pathLoop = "";
            string pathFormato = "";
            XElement xelement = XElement.Load(outputName);
            IEnumerable<XElement> configuracion = xelement.Elements();
            editor2.WriteMessage("\nConfiguracion :");
            foreach (var employee in configuracion)
            {
               if (i == 0) { pathtSignaList = employee.Element("ArchDir").Value; };
               if (i == 1) { pathtTypical = employee.Element("ArchDir").Value; };
               if (i == 2) { pathLoop = employee.Element("ArchDir").Value; };
               if (i == 3) { pathFormato = employee.Element("ArchDir").Value; };
               i++;
            }
            editor2.WriteMessage("\nArchivo con la lista de Planos            :{0} ", pathtSignaList);
            editor2.WriteMessage("\npDirectorio donde estan los Tipicos       :{0}", pathtTypical);
            editor2.WriteMessage("\nDirectorio donde  se copiaran los archivos:{0}", pathLoop);
            editor2.WriteMessage("\nNombre del Bloque que servira de Formato  :{0}", pathFormato);
            PromptStringOptions blkFor = new PromptStringOptions("\nAre you sure? Y/N :");
            PromptResult promptResult = doc.Editor.GetString(blkFor);
            if (promptResult.StringResult.ToUpper() == "Y")
            {
               //doc.Editor.WriteMessage("\nEscriba el comando INE24");
               Clase28(editor2, pathtSignaList, pathtTypical, pathLoop, pathFormato);
            }
            else
            {
               doc.Editor.WriteMessage("\nOperacion cancelada por el usuario...");
            }
         }
         else
         { editor2.WriteMessage("\nERROR: El archivo SIDC.XLSX debe estar en el mismo directorio del DLL"); }
      }
      public void Clase28(Editor ed, string signaList, string archTipico, string archInecad, string forMato)
      {
         Planos plano = new Planos();      //objeto que van a contener la data de la tabla Indice
         BuscaCelda buscarV = new BuscaCelda();
         //Atributos AtribExel = new Atributos();
         //AtribForm atribForm = new AtribForm();
         Excel.Workbook wbkObj;    //workBook -archivo excel
         Excel.Worksheet wshtObj;  //workSheet INDICE
         Excel.Range rngObj;       //rango en INDICE
         Excel.Worksheet wshtObj2; //worhSheet LISTA
         Excel.Range rngObj2;      //rango LISTA
         Excel.Application ExcelServer = new Excel.Application();
         wbkObj = ExcelServer.Workbooks.Open(@signaList, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         wshtObj = (Excel.Worksheet)wbkObj.Worksheets.get_Item(2);  //indice
         rngObj = wshtObj.UsedRange;
         wshtObj2 = (Excel.Worksheet)wbkObj.Worksheets.get_Item(1); //lista
         rngObj2 = wshtObj2.UsedRange;
         int rCnt1;          // contador de filas workSheet Indice
         int rTot1 = 0;      // total de filas workSheet Indice
         int cTot1 = 0;      // total de columnas workSheet Indice
         int cCnt1 = 0;      // contador de columnas workSheet Indice
         int rCnt2;          // contador de filas workSheet Lista
         int rTot2 = 0;      // total filas workSheet Lista
         int cCnt2;          // contador de columnas worksheet Lista
         int cTot2 = 0;      // total de columnas worksheet Lista
         int cSw0;           // switch, compara Nros de Planos en Indice contra Lista // luego se usap ara buscar el bloqeu de FORMATO
         int cSw1;           // switch, si Wplano no tiene bloque asignado en Lista
         int cSw2;           // switch, compara nombre de Bloque de Autocad contra el de Excel
         int colPos;         // *** nro de columna en LISTA, que contiene lo que se va a escribir en el Atributo ***
         string outputName;  // nombre del archivo a Generar (Inecad)
         string fileName;    // nombre el Tipico
         string celdAux;      // c/u de las filas de Lista, tal que Wplano = Nplano // re-utilizada en 87
         if ((archTipico != null) && (archInecad != null))
         {
            rTot1 = rngObj.Rows.Count;
            cTot1 = rngObj.Columns.Count;
            rTot2 = rngObj2.Rows.Count;
            cTot2 = rngObj2.Columns.Count;
            List<string> lista2 = new List<string>();
            for (cCnt1 = 1; cCnt1 <= cTot1; cCnt1++) // Hacer una lista con todos los nombres de los Atributos en INDICE
            {
               celdAux = (string)(rngObj.Cells[1, cCnt1] as Excel.Range).Value2;
               lista2.Add(celdAux);
            }
            List<string> lista3 = new List<string>();
            for (cCnt2 = 1; cCnt2 <= cTot2; cCnt2++) // Hacer una lista con todos los nombres de los Atributos en LISTA
            {
               //AtribExel.WAtributo = (string)(rngObj2.Cells[1, cCnt2] as Excel.Range).Value2;
               //lista3.Add(AtribExel.WAtributo);
               celdAux = (string)(rngObj2.Cells[1, cCnt2] as Excel.Range).Value2;
               lista3.Add(celdAux);
               // { lista3.Add(new Atributos() { WAtributo = celdAux }); }
            }
            for (int i = 0; i < lista3.Count; i++) { ed.WriteMessage("\ncolumna {0} :{1}", i, lista3[i]); }
            List<BuscaCelda> lista1 = new List<BuscaCelda>();
            for (rCnt1 = 2; rCnt1 <= rTot1; rCnt1++) // a partir de la fila 2 porque la 1 tiene los titulos, leer INDICE y buscar 'NPlano' en PLANOS
            {
               cSw0 = 1;
               plano.Nplano = (string)(rngObj.Cells[rCnt1, 1] as Excel.Range).Value2;   //plano
               plano.Ntipico = (string)(rngObj.Cells[rCnt1, 2] as Excel.Range).Value2;  //tipico
               plano.NindeCad = (string)(rngObj.Cells[rCnt1, 3] as Excel.Range).Value2; //inecad
               plano.Nfila = rCnt1;                                                     //# de fila 
               outputName = string.Concat(archInecad, plano.NindeCad, ".dwg");
               fileName = string.Concat(archTipico, plano.Ntipico, ".dwg");
               ed.WriteMessage("\n-----------------------------------------------------------------------------\n");
               ed.WriteMessage("\nTipico {0} inecad {1}", fileName, outputName);
               for (rCnt2 = 2; rCnt2 <= rTot2; rCnt2++) // a partir de la fila 2 porque la 1 tiene los titulos, hacer como un arreglo dinamico de los # de filas que contienen el nro de plano leido en Nplano
               {
                  cSw0 = 1;
                  celdAux = (string)(rngObj2.Cells[rCnt2, 1] as Excel.Range).Value2;
                  cSw0 = string.Compare(plano.Nplano, celdAux);
                  if (cSw0 == 0)   //si es el Tipico correcto
                  { lista1.Add(new BuscaCelda() { WPlano = celdAux, WfiLa = rCnt2 }); }
               }
               if (lista1.Count > 0)
               {
                  Database db = new Database(false, false);
                  using (db)
                  {
                     try
                     {
                        //ed.WriteMessage("\nAbriendo el archivo: " + fileName);
                        db.ReadDwgFile(fileName, FileShare.ReadWrite, false, "");
                        Transaction tr = db.TransactionManager.StartTransaction();
                        ObjectId msId;
                        using (tr)
                        {
                           cSw1 = 0;
                           cSw0 = 1; //re-utilizar cSw0
                           BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                           msId = bt[BlockTableRecord.ModelSpace];
                           BlockTableRecord btr = (BlockTableRecord)tr.GetObject(msId, OpenMode.ForRead);
                           foreach (ObjectId entId in btr)// revisa todos obketos en btr...
                           {
                              Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                              if (ent != null)
                              {
                                 BlockReference br = ent as BlockReference;
                                 if (br != null)
                                 {
                                    cSw2 = 1;
                                    BlockTableRecord bd = (BlockTableRecord)tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                    string blkName = bd.Name.ToUpper();   // nombre del Bloque Autocad
                                    cSw0 = string.Compare(blkName, forMato); // si es el bloque de Formato
                                    //ed.WriteMessage("\n BLKNAME: {0},  FORMATO: {1}", blkName, forMato);
                                    if (cSw0 == 0)
                                    {
                                       foreach (ObjectId arId in br.AttributeCollection)
                                       {
                                          DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                          AttributeReference ar = obj as AttributeReference;
                                          if (ar != null)
                                          {
                                             string celdaAtr = ar.Tag.ToUpper();
                                             if (celdaAtr != null)
                                             {
                                                colPos = 0;
                                                colPos = lista2.FindIndex(b => b == celdaAtr);
                                                if (colPos > 0) // si el nombre del Atributo coincide
                                                {
                                                   colPos++;
                                                   dynamic celDaTrib = rngObj.Cells[plano.Nfila, colPos].Value;
                                                   string myResult = Convert.ToString(celDaTrib);
                                                   if (celDaTrib != null)
                                                   {
                                                      //  ed.WriteMessage("\nrowPos: {0}, blkName: {1},  celdaAtr: {2}, colPos:{3},  myResult:{4}", plano.Nfila, blkName, celdaAtr, colPos, myResult);
                                                      //cSw1 = 1;
                                                      ar.UpgradeOpen();
                                                      ar.TextString = myResult;
                                                      ar.DowngradeOpen();
                                                   }
                                                }
                                             }
                                          }
                                       }

                                    }
                                    else
                                    {
                                       foreach (BuscaCelda c in lista1)
                                       {
                                          string celDaFila = (string)(rngObj2.Cells[c.WfiLa, 2] as Excel.Range).Value2;
                                          //string celDaFila = c.WPlano;
                                          if (celDaFila != null)
                                          {
                                             cSw2 = string.Compare(blkName, celDaFila);
                                             if (cSw2 == 0)
                                             {
                                                foreach (ObjectId arId in br.AttributeCollection)
                                                {
                                                   DBObject obj = tr.GetObject(arId, OpenMode.ForRead);
                                                   AttributeReference ar = obj as AttributeReference;
                                                   if (ar != null)
                                                   {
                                                      string celdaAtr = ar.Tag.ToUpper();
                                                      if (celdaAtr != null)
                                                      {
                                                         colPos = 0;
                                                         colPos = lista3.FindIndex(a => a == celdaAtr);
                                                         if (colPos > 0) // si el nombre del Atributo coincide
                                                         {
                                                            colPos++;
                                                            dynamic celDaTrib = rngObj2.Cells[c.WfiLa, colPos].Value;
                                                            string myResult = Convert.ToString(celDaTrib);
                                                            if (celDaTrib != null)
                                                            {
                                                               ed.WriteMessage("\nrowPos: {0}, blkName: {1}, celDaFila:{2}, celdaAtr: {3}, colPos:{4},  myResult:{5}", c.WfiLa, blkName, celDaFila, celdaAtr, colPos, myResult);
                                                               cSw1 = 1;
                                                               ar.UpgradeOpen();
                                                               ar.TextString = myResult;
                                                               ar.DowngradeOpen();
                                                            }
                                                         }
                                                      }
                                                   }
                                                }
                                             }
                                          }
                                       }
                                    }

                                 }
                              }
                           }
                           tr.Commit();
                        }
                        if (cSw1 == 1)
                        {
                           ed.WriteMessage("\nOK, Procesado archivo: {0}", outputName);
                           db.SaveAs(outputName, DwgVersion.Current);
                        }
                        else { ed.WriteMessage("\nProblema: el PLANO '{0}' NO TIENE BLOQUE asociado en el worksheet Lista ", outputName); }
                     }
                     catch (System.Exception ex)
                     {
                        ed.WriteMessage("\nProblema procesando archivo: {0}", fileName, ex.Message);
                     }
                  }
               }
               else { ed.WriteMessage("\nProblema: el PLANO '{0}' NO existe en el worksheet Lista ", outputName); }
               lista1.Clear();
            }
            wbkObj.Close(true, null, null);
            ExcelServer.Quit();
            _ = Marshal.ReleaseComObject(wshtObj2);
            _ = Marshal.ReleaseComObject(wshtObj);
            _ = Marshal.ReleaseComObject(wbkObj);
            _ = Marshal.ReleaseComObject(ExcelServer);
         }
         else
         {
            ed.WriteMessage("\n ** Error: Revise la tabla Directorios **");
         }
      }
      public class Planos   //2 miembreos de esta clase
      {
         public string Nplano { get; set; }
         public string Ntipico { get; set; }
         public string NindeCad { get; set; }
         public int Nfila { get; set; }
         public Planos()
         {
            Nplano = "";
            Ntipico = "";
            NindeCad = "";
            Nfila = 0;
         }
      }
      //public class Atributos
      // {
      //    public string WAtributo { get; set; }
      //   public Atributos()
      //   {
      //      WAtributo = "";
      // }
      //}

      //public class AtribForm
      // {
      //     public string Wformato { get; set; }
      //     public int Wfila2 { get; set; }
      //     public AtribForm()
      //    {
      //         Wformato = "";
      //         Wfila2 = 0;
      //     }
      // }

      public class BuscaCelda //3 miembreos de esta clase (propiedad, constructor, metodo)
      {
         public string WPlano { get; set; } // Propiedades
         public int WfiLa { get; set; }
         //[Required]
         //public string celdAux { get; set; }
         //[Required]
         //public string celDaFila { get; set; }
         public BuscaCelda()              // Constructor.
         {
            WPlano = "";
            WfiLa = 0;  // *** nro de fila en LISTA, que contiene lo que se va a escribir en el Atributo ***
                        //    celdAux = "";
                        //    celDaFila = "";
         }
         //public void CeldaDetails()      //Metodo
         //{
         //}
      }
   }
}
