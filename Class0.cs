using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using Autodesk.Windows;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace InelectrApp.Ine023
{
   public class Clase023
   {
      [CommandMethod("INE23")]
      public void RibbonSplitButton()
      {
         RibbonControl ribbonControl = ComponentManager.Ribbon;
         // creae el Ribbon Inelectra
         RibbonTab Tab = new RibbonTab
         {
            Title = "Inelectra",
            Id = "EDGARIBBON_TAB_ID"
         };
         ribbonControl.Tabs.Add(Tab);
         // agrgar panel 1 al Ribbon
         RibbonPanelSource srcPanel = new RibbonPanelSource
         { Title = "  CONFIGURAR  " };
         RibbonPanel Panel = new RibbonPanel
         { Source = srcPanel };
         Tab.Panels.Add(Panel);
         // agrgar panel 2 al Ribbon
         RibbonPanelSource srcPanel2 = new RibbonPanelSource         
         { Title = "  bbbb  " };
         RibbonPanel Panel2 = new RibbonPanel
         { Source = srcPanel2 };
         Tab.Panels.Add(Panel2);

         RibbonPanelSource srcPanel3 = new RibbonPanelSource
         { Title = "  ccccc  " };
         RibbonPanel Panel3 = new RibbonPanel
         { Source = srcPanel3 };
         Tab.Panels.Add(Panel3);
         RibbonPanelSource srcPanel4 = new RibbonPanelSource
         { Title = "  dddd  " };
         RibbonPanel Panel4 = new RibbonPanel
         { Source = srcPanel4 };
         Tab.Panels.Add(Panel4);


         // crear TODOS los botones (5 botones)
         RibbonButton button1 = new RibbonButton
         {
            Id = "1",
            Text = "Lista de Planos",
            ShowText = true,
            ShowImage = true,
            CommandHandler = new MyCmdHandler()
         };
         RibbonButton button2 = new RibbonButton
         {
            Id = "2",
            Text = "Directorio Tipicos",
            ShowText = true,
            ShowImage = true,
            CommandHandler = new MyCmdHandler()
         };
         RibbonButton button3 = new RibbonButton
         {
            Id = "3",
            Text = "Directorio Destino",
            ShowText = true,
            ShowImage = true,
            CommandHandler = new MyCmdHandler()
         };
         RibbonButton button4 = new RibbonButton
         {
            Id = "4",
            Text = "Nombre del Formato",
            ShowText = true,
            ShowImage = true,
            CommandHandler = new MyCmdHandler()
         };

         // crea el split boton (para separar los botones dentro del Panel 1)
         RibbonSplitButton ribSplitButton = new RibbonSplitButton
         {
            // Requerido para evitar error de AutoCAD cuando se usa el localizador de cmd
            Text = "RibbonSplitButton",
            ShowText = true
         };
         // crea el split boton, aunque es un solo boton (requerido)
         RibbonSplitButton ribSplitButton2 = new RibbonSplitButton
         {
            // Requerido para evitar error de AutoCAD cuando se usa el localizador de cmd
            Text = "RibbonSplitButton2",
            ShowText = true
         };

         RibbonSplitButton ribSplitButton3 = new RibbonSplitButton
         {
            // Requerido para evitar error de AutoCAD cuando se usa el localizador de cmd
            Text = "RibbonSplitButton3",
            ShowText = true
         };
         RibbonSplitButton ribSplitButton4 = new RibbonSplitButton
         {
            // Requerido para evitar error de AutoCAD cuando se usa el localizador de cmd
            Text = "RibbonSplitButton4",
            ShowText = true
         };

         // agregar los botones a los paneles a donde pertenecen
         ribSplitButton.Items.Add(button1);
         srcPanel.Items.Add(ribSplitButton);
         ribSplitButton2.Items.Add(button2);
         srcPanel2.Items.Add(ribSplitButton2);
         ribSplitButton3.Items.Add(button3);
         srcPanel3.Items.Add(ribSplitButton3);
         ribSplitButton4.Items.Add(button4);
         srcPanel4.Items.Add(ribSplitButton4);


         Tab.IsActive = true;
      }
      /// <summary>
      /// dependiendo del Boton seleccionado 'Execute' enviara a determinads Metodos
      /// </summary>
      public class MyCmdHandler : System.Windows.Input.ICommand
      {
         public bool CanExecute(object parameter)
         { return true; }
         public event EventHandler CanExecuteChanged
         {
            // agregados para evitar CS0535
            add { }
            remove { }
         }
         public void Execute(object parameter)
         {
            MyDirectory myDirectory = new MyDirectory();
            string path1 = myDirectory.Path3;
            if (parameter is RibbonButton)
            {
               RibbonButton button = parameter as RibbonButton;
               switch (button.Id)
               {
                  case "1":
                     BuscArchivo(myDirectory.Path3, path1);
                     // ArchivoTxt("1", myDirectory.Path3, path1);
                     break;
                  case "2":
                     BuscaDirectory(myDirectory.Path3);
                     break;
                  case "3":
                     BuscaDirectory2(myDirectory.Path3);
                     break;
                  case "4":
                     BuscaBloque(myDirectory.Path3);
                     break;
               }
            }
         }
      }
      /// <summary>
      /// aqui se configura el archivo .XLSX que contiene 2 Woksheet: Planos e Indice. 
      /// </summary>
      /// <param name="path2"></param>
      /// <param name="path1"></param>
      public static void BuscArchivo(string path2, string path1)
      {
         Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Excel con la Lista de Lazos", null, "xls; xlsx", "ExcelFileToLink", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
         DialogResult dr = ofd.ShowDialog();
         if (dr != DialogResult.OK)
            return;
         Directory.SetCurrentDirectory(path2);
         ArchivoTxt("1", ofd.Filename, path1);
      }
      public static void BuscaDirectory(string path2)
      {
         Directory.SetCurrentDirectory(path2);
         using (var fbd = new FolderBrowserDialog())
         {
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
               _ = Directory.GetFiles(fbd.SelectedPath);
               ArchivoTxt("2", fbd.SelectedPath, path2);
            }
         }
      }
      public static void BuscaDirectory2(string path2)
      {
         Directory.SetCurrentDirectory(path2);
         using (var fbd = new FolderBrowserDialog())
         {
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
               _ = Directory.GetFiles(fbd.SelectedPath);
               ArchivoTxt("3", fbd.SelectedPath, path2);
            }
         }
      }
      public static void BuscaBloque(string path2)
      {
         //_ = MessageBox.Show("path2: " + path2);
         Directory.SetCurrentDirectory(path2);
         Document doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
         PromptStringOptions blkFor = new PromptStringOptions("Nombre del Formato:");
         PromptResult promptResult = doc.Editor.GetString(blkFor);
         if (promptResult.Status == PromptStatus.OK)
         {
            string blOqResult = promptResult.StringResult.Trim();
            ArchivoTxt("4", blOqResult, path2);
         }
         else
         {
            doc.Editor.WriteMessage("\nERRROR");
         }
      }
      /// <summary>
      /// crear el archivo de Configuracio 
      /// </summary>
      /// <param name="botonTxt"></param> //la opcion seleccionada del menu
      /// <param name="valAux"></param>   //directorio/archivo/bloque seleccionado
      /// <param name="path3"></param>    //directorio donde se encuentra la Aplicacion
      public static void ArchivoTxt(string botonTxt, string valAux, string path3)
      {
         int i = 0;
         string cfgName;
         string pathtSignaList = "C:\\SignaList.xlsx\\";
         string pathtTypical = "C:\\";
         string pathLoop = "C:\\" ;
         string pathFormato = "BLOQUE";
         cfgName = string.Concat(path3, "\\LoopConfig.xml");
         if (File.Exists(cfgName))
         {
            XElement xelement = XElement.Load(cfgName);
            IEnumerable<XElement> configuracion = xelement.Elements();
            foreach (var archCfg in configuracion)
            {
               if (i == 0) { pathtSignaList = archCfg.Element("ArchDir").Value; }
               if (i == 1) { pathtTypical = archCfg.Element("ArchDir").Value; }
               if (i == 2) { pathLoop = archCfg.Element("ArchDir").Value; }
               if (i == 3) { pathFormato = archCfg.Element("ArchDir").Value; }
               i++;
            }
         }
         switch (botonTxt)
         {
            case "1":
               pathtSignaList = valAux;
               break;
            case "2":
               valAux += "\\";
               pathtTypical = valAux;
               break;
            case "3":
               valAux += "\\";
               pathLoop = valAux;
               break;
            case "4":
               pathFormato = valAux;
               break;
         }
         XmlWriterSettings xlmSeteo = new XmlWriterSettings
         {
            Indent = true,
            IndentChars = "    ",
            CloseOutput = true,
            OmitXmlDeclaration = false,
            Encoding = Encoding.UTF8
         };
         //_ = MessageBox.Show("DIRECTORIO DE LA APLICACION path3: " + path3, "*****DIRECTORIO*** ");
         using (XmlWriter ArchivoXlm = XmlWriter.Create("LoopConfig.xml", xlmSeteo))
         {
            ArchivoXlm.WriteStartDocument();
            ArchivoXlm.WriteStartElement("Configurar"); // escribe <Configurar>
            ArchivoXlm.WriteStartElement("Archivo");
            ArchivoXlm.WriteElementString("ArchDir", pathtSignaList);
            ArchivoXlm.WriteElementString("Description", "PATH DATABASE, debe ser un archivo excel .xlsx");
            ArchivoXlm.WriteElementString("Tipo", "1");
            ArchivoXlm.WriteEndElement();
            ArchivoXlm.WriteStartElement("Archivo");
            ArchivoXlm.WriteElementString("ArchDir", pathtTypical);
            ArchivoXlm.WriteElementString("Description", "PATH TYPICAL,dbe ser el nombre de un Directorio");
            ArchivoXlm.WriteElementString("Tipo", "2");
            ArchivoXlm.WriteEndElement();
            ArchivoXlm.WriteStartElement("Archivo");
            ArchivoXlm.WriteElementString("ArchDir", pathLoop);
            ArchivoXlm.WriteElementString("Description", "PATH LOOP, debe ser el nombre de un Directorio");
            ArchivoXlm.WriteElementString("Tipo", "3");
            ArchivoXlm.WriteEndElement();
            ArchivoXlm.WriteStartElement("Archivo");
            ArchivoXlm.WriteElementString("ArchDir", pathFormato);
            ArchivoXlm.WriteElementString("Description", "FORMAT BLOCK, debe ser el nombre de un BLOCK AutocaD");
            ArchivoXlm.WriteElementString("Tipo", "4");
            ArchivoXlm.WriteEndElement();
            ArchivoXlm.Flush();
         }
      }


      /// <summary>
      /// Directorio donde se encuentra la App
      /// y adonde se creara el archivo de Configuracion (si no existe)
      /// </summary>
      public class MyDirectory
      {
         public string Path3 { get; set; }
         public MyDirectory()
         {
            Path3 = Directory.GetCurrentDirectory();
         }
      }
   }
}
