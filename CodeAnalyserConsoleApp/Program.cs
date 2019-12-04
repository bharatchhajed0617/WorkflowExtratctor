using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Threading.Tasks;

namespace CodeAnalyserConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            LoadXAML xaml = new LoadXAML();

            string outputPath = @"C:\test\";
                outputPath = outputPath.Substring(0,outputPath.Length - 1);

            Console.WriteLine("Please enter UiPath Project Path!");

            string folderPath = @"C:\Users\Bharat\Documents\UiPath\BlankProcess27";

            string excelFolderPath =  Path.Combine( folderPath, "CodeAnalyser.xlsx");

            


            Excel excel = new Excel();
            Console.WriteLine("Processing...");
            #region Datatable

           
            DataTable dtSelector = new DataTable("Selectors");
            dtSelector.Columns.Add("Source");
            dtSelector.Columns.Add("Target");
            dtSelector.Columns.Add("DisplayName");
            dtSelector.Columns.Add("Timeoutms");
            dtSelector.Columns.Add("WaitForReady");
            dtSelector.Columns.Add("Selector");

            DataTable dtCommentedCode = new DataTable("Commented Code");
            dtCommentedCode.Columns.Add("Source");
            dtCommentedCode.Columns.Add("Detected");
            dtCommentedCode.Columns.Add("Count");


            DataTable dtMessageBox = new DataTable("MessageBox");
            dtMessageBox.Columns.Add("Source");
            dtMessageBox.Columns.Add("Detected");
            dtMessageBox.Columns.Add("Count");

            DataTable dtKillProcess = new DataTable("KillProcess");
            dtKillProcess.Columns.Add("Source");
            dtKillProcess.Columns.Add("Detected");
            dtKillProcess.Columns.Add("Count");


            DataTable dtTryCatch = new DataTable("TryCatch");
            dtTryCatch.Columns.Add("Source");
            dtTryCatch.Columns.Add("Detected");
            dtTryCatch.Columns.Add("Count");

            DataTable dtDelay = new DataTable("Delay");
            dtDelay.Columns.Add("Source");
            dtDelay.Columns.Add("Type");
            dtDelay.Columns.Add("Duration");
            dtDelay.Columns.Add("Parent XML");


            DataTable dtLogMessage = new DataTable("LogMessage");
            dtLogMessage.Columns.Add("Source");
            dtLogMessage.Columns.Add("Level");
            dtLogMessage.Columns.Add("Message");

            DataTable dtCredential = new DataTable("Credential");
            dtCredential.Columns.Add("Source");
            dtCredential.Columns.Add("Target");
            dtCredential.Columns.Add("Password");
            dtCredential.Columns.Add("Parent XML");

            DataTable dtSendHotKey = new DataTable("SendHotkey");
            dtSendHotKey.Columns.Add("Source");
            dtSendHotKey.Columns.Add("Name");
            dtSendHotKey.Columns.Add("Key");
            dtSendHotKey.Columns.Add("KeyModifier");
            dtSendHotKey.Columns.Add("Selector");

            DataTable dtVariables = new DataTable("Variables");
            dtVariables.Columns.Add("Source");
            dtVariables.Columns.Add("Name");
            dtVariables.Columns.Add("Type");

            DataTable dtArgument = new DataTable("Arguments");
            dtArgument.Columns.Add("Source");
            dtArgument.Columns.Add("Name");
            dtArgument.Columns.Add("Type");

            DataTable dtIF = new DataTable("Nested If");
            dtIF.Columns.Add("Source");
            dtIF.Columns.Add("Detected");
            dtIF.Columns.Add("Count");


            DataTable dtTerminateWF = new DataTable("Terminate Workflow");
            dtTerminateWF.Columns.Add("Source");
            dtTerminateWF.Columns.Add("Detected");
            dtTerminateWF.Columns.Add("Count");

            DataTable dtIDXSelector = new DataTable("IDX Selector");
            dtIDXSelector.Columns.Add("Source");
            dtIDXSelector.Columns.Add("Target");
            dtIDXSelector.Columns.Add("IDX Value");
            dtIDXSelector.Columns.Add("Selector");

            DataTable dtWorkflow = new DataTable("Workflow");
            dtWorkflow.Columns.Add("Workflow");
            dtWorkflow.Columns.Add("No. of Arguments");
            dtWorkflow.Columns.Add("No. of Variables");

            DataTable dtSoftware = new DataTable("SoftwareEvent");
            dtSoftware.Columns.Add("Source");
            dtSoftware.Columns.Add("Target");
            dtSoftware.Columns.Add("Simulate");
            dtSoftware.Columns.Add("SendWindowMessage");
            dtSoftware.Columns.Add("DisplayName");
            dtSoftware.Columns.Add("Selector");

            DataTable dtSendWindowMessage = new DataTable("SendWindowMessage");
            dtSendWindowMessage.Columns.Add("Source");
            dtSendWindowMessage.Columns.Add("Target");
            dtSendWindowMessage.Columns.Add("SendWindowMessage");
            dtSendWindowMessage.Columns.Add("DisplayName");
            dtSendWindowMessage.Columns.Add("Selector");

            DataTable dtActivities = new DataTable("Activities");
            dtActivities.Columns.Add("Source");
            dtActivities.Columns.Add("Target");
            dtActivities.Columns.Add("SendWindowMessage");
            dtActivities.Columns.Add("DisplayName");
            dtActivities.Columns.Add("Depth");

            DataTable dtSimulate = new DataTable("Simulate");
            dtSimulate.Columns.Add("Source");
            dtSimulate.Columns.Add("Target");
            dtSimulate.Columns.Add("Simulate");
            dtSimulate.Columns.Add("DisplayName");
            dtSimulate.Columns.Add("Selector");


            DataTable dtImage = new DataTable("Image");
            dtImage.Columns.Add("Source");
            dtImage.Columns.Add("Target");
            dtImage.Columns.Add("DisplayName");
            dtImage.Columns.Add("Selector");

            #endregion





            foreach (string xamlFilePath in Directory.GetFiles(folderPath, "*xaml", SearchOption.AllDirectories))
            {
                //string xamlFile = @"C:\Users\Bharat\Desktop\FA_Monitoring\Excel.xaml";
                string source = xamlFilePath.Trim().Replace(folderPath.Trim(), "");
                XmlDocument xmlDocument = xaml.GetDocument(xamlFilePath);

                Dictionary<string,int> diccUnique = new Dictionary<string, int>();
                diccUnique = xaml.UniqueActivityCount(xmlDocument, diccUnique);


                Dictionary < int, int> dicNested = new Dictionary<int, int>();
                xaml.AttributeExist(xmlDocument, "Sequence", dicNested);
           
                if (dicNested.Max(x => x.Value)>5)
                {
                    Console.WriteLine("Too Nested sequence "+dicNested.Max(x => x.Value));
                }

                dicNested = new Dictionary<int, int>();
                xaml.AttributeExist(xmlDocument, "If", dicNested);

                if (dicNested.Max(x => x.Value) > 5)
                {
                    Console.WriteLine("Too Nested sequence " + dicNested.Max(x => x.Value));
                }
                xaml.GetNodesWithAttributes(xmlDocument, "sap2010:WorkflowViewState.IdRef", "DisplayName", dtActivities, source, 0);

                if (xmlDocument.GetElementsByTagName("If") != null && xmlDocument.GetElementsByTagName("If").Count > 0)
                {
                    XmlNode msgNode = xmlDocument.GetElementsByTagName("If")[0];
                    foreach (XmlNode node in xmlDocument.GetElementsByTagName("If"))
                    {
                        List<string> lstNodes = new List<string>();
                        xaml.GetChildNodes(node, lstNodes, "If");
                        if (lstNodes.Count > 3)
                        {
                            dtIF.Rows.Add(source, "Yes", lstNodes.Count);
                            break;
                        }
                    }
                }

                foreach (XmlNode node in xmlDocument.GetElementsByTagName("ui:ImageTarget"))
                {
                    dtImage.Rows.Add(source, "ImageTarget", node.ParentNode.ParentNode.Attributes["DisplayName"].Value, ((XmlElement)(node.ParentNode.ParentNode)).GetElementsByTagName("ui:Target")[0].Attributes["Selector"].Value);
                }
                DataTable distinctValues = dtActivities.DefaultView.ToTable(true, "Depth");
                for (int i = 0; i < distinctValues.Rows.Count; i++)
                {
                    Console.WriteLine(distinctValues.Rows[i][0].ToString());
                    if (Convert.ToInt32(distinctValues.Rows[i][0])!=i)
                    {
                        for (int j = 0; j < dtActivities.Rows.Count ; j++)
                        {
                            if(dtActivities.Rows[j]["Depth"].ToString().Trim().Equals(distinctValues.Rows[i][0].ToString().Trim()))
                            dtActivities.Rows[j]["Depth"] = dtActivities.Rows[j]["Depth"].ToString().Replace(distinctValues.Rows[i][0].ToString(), (i ).ToString());
                        }
                    }
               
                }

                xaml.GetNodesWithAttributes(xmlDocument, "SendWindowMessages", "DisplayName", dtSendWindowMessage, source);


                //xaml.GetNodesWithAttributes(xmlDocument, "SimulateClick", "SendWindowMessages", "DisplayName", dtSoftware, source);
                //xaml.GetNodesWithAttributes(xmlDocument, "SimulateType", "SendWindowMessages", "DisplayName", dtSoftware, source);

                xaml.GetNodesWithAttributes(xmlDocument, "Password", dtCredential, source);

                xaml.GetNodesWithAttributes(xmlDocument, "TimeoutMS", "WaitForReady", "Selector", dtSelector, source);
                //xaml.GetNodesWithAttributes(xmlDocument, "Selector", dtSelector, source);
                foreach (XmlNode item in xmlDocument.GetElementsByTagName("x:Property"))
                {
                    dtArgument.Rows.Add(source, item.Attributes["Name"].Value, item.Attributes["Type"].Value);
                }


                foreach (XmlNode item in xmlDocument.GetElementsByTagName("Variable"))
                {
                    dtVariables.Rows.Add(source, item.Attributes["Name"].Value, item.Attributes["x:TypeArguments"].Value);
                }

                dtWorkflow.Rows.Add(source, xmlDocument.GetElementsByTagName("x:Property").Count.ToString(), xmlDocument.GetElementsByTagName("Variable").Count.ToString());


                foreach (XmlNode item in xmlDocument.GetElementsByTagName("ui:SendHotkey"))
                {
                    dtSendHotKey.Rows.Add(source, item.Attributes["DisplayName"].Value, item.Attributes["Key"].Value, item.Attributes["KeyModifiers"].Value, ((XmlElement)item).GetElementsByTagName("ui:Target")[0].Attributes["Selector"].Value);
                }

                foreach (XmlNode item in xmlDocument.GetElementsByTagName("ui:LogMessage"))
                {
                    dtLogMessage.Rows.Add(source, item.Attributes["Level"].Value, item.Attributes["Message"] == null?"": item.Attributes["Message"].Value);
                }

                if (xmlDocument.GetElementsByTagName("TerminateWorkflow") != null && xmlDocument.GetElementsByTagName("TerminateWorkflow").Count > 0)
                {
                    dtTerminateWF.Rows.Add(source, "Yes", xmlDocument.GetElementsByTagName("TerminateWorkflow").Count);
                }

                if (xmlDocument.GetElementsByTagName("TryCatch") != null && xmlDocument.GetElementsByTagName("TryCatch").Count > 0)
                {
                    dtTryCatch.Rows.Add(source, "Yes", xmlDocument.GetElementsByTagName("TryCatch").Count);
                }

                if (xmlDocument.GetElementsByTagName("ui:CommentOut") != null && xmlDocument.GetElementsByTagName("ui:CommentOut").Count > 0)
                {
                    dtCommentedCode.Rows.Add(source, "Yes", xmlDocument.GetElementsByTagName("ui:CommentOut").Count);
                }

                if (xmlDocument.GetElementsByTagName("ui:MessageBox") != null && xmlDocument.GetElementsByTagName("ui:MessageBox").Count > 0)
                {
                    XmlNode node = xmlDocument.GetElementsByTagName("ui:MessageBox")[0].ParentNode;

                    foreach (XmlNode item in xmlDocument.GetElementsByTagName("ui:MessageBox"))
                    {
                        List<string> lstNodes = new List<string>();
                        xaml.GetParentNodes(item, lstNodes);
                        if (!lstNodes.Contains("ui:CommentOut"))
                        {
                            dtMessageBox.Rows.Add(source, "Yes", xmlDocument.GetElementsByTagName("ui:MessageBox").Count);
                        }
                    }

                }


                if (xmlDocument.GetElementsByTagName("ui:KillProcess") != null && xmlDocument.GetElementsByTagName("ui:KillProcess").Count > 0)
                {
                    XmlNode node = xmlDocument.GetElementsByTagName("ui:KillProcess")[0].ParentNode;

                    foreach (XmlNode item in xmlDocument.GetElementsByTagName("ui:KillProcess"))
                    {
                        List<string> lstNodes = new List<string>();
                        xaml.GetParentNodes(item, lstNodes);
                        if (!lstNodes.Contains("ui:CommentOut"))
                        {
                            dtKillProcess.Rows.Add(source, "Yes", xmlDocument.GetElementsByTagName("ui:KillProcess").Count);
                        }
                    }

                }


                #region Delay
                xaml.GetNodesWithAttributes(xmlDocument, "Duration", dtDelay, source);
                xaml.GetNodesWithAttributes(xmlDocument, "DelayMS", dtDelay, source);
                xaml.GetNodesWithAttributes(xmlDocument, "DelayBefore", dtDelay, source);


                #endregion

            }




            List<string> listIDXGreater = new List<string>();
            List<string> listIDXLesser = new List<string>();

            foreach (DataRow item in dtSelector.Rows)
            {
                XmlDocument xmlSelector = new XmlDocument();
                xmlSelector.LoadXml("<Selector>" + item["Selector"].ToString().Replace("omit:","") + " </Selector>");
                List<string> lstIDX = new List<string>();
                xaml.GetNodesWithAttributes(xmlSelector, "idx", dtIDXSelector, item["Source"].ToString());

                if (lstIDX.Count > 0)
                {
                    foreach (var idx in lstIDX)
                    {
                        int intIDX = 0;
                        if (Int32.TryParse(idx, out intIDX))
                        {
                            if (intIDX > 3)
                                listIDXGreater.Add(item["Selector"].ToString());
                            else
                                listIDXLesser.Add(item["Selector"].ToString());
                        }
                    }
                }
            }


            for (int i = 0; i < dtDelay.Rows.Count; i++)
            {
                TimeSpan outSec = new TimeSpan();
                int second = 0;
                if (int.TryParse(dtDelay.Rows[i][2].ToString(), out second))
                {

                }
                else
                {
                    TimeSpan.TryParse(dtDelay.Rows[i][2].ToString(), out outSec);
                    second = outSec.Seconds * 1000;

                }
                dtDelay.Rows[i][2] = second;
                if (second == 0)
                {
                    dtDelay.Rows.Remove(dtDelay.Rows[i]);
                    i--;
                }

            }

            DataSet dataSet = new DataSet("CodeAnalyserDataSet");

            //dataSet.Tables.Add(dtSimulate);
            dataSet.Tables.Add(dtSendWindowMessage);
            dataSet.Tables.Add(dtSoftware);
            dataSet.Tables.Add(dtImage);
            dataSet.Tables.Add(dtCredential.DefaultView.ToTable(false, "Source", "Target", "Password"));
            dataSet.Tables.Add(dtMessageBox.DefaultView.ToTable(true));
            dataSet.Tables.Add(dtCommentedCode.DefaultView.ToTable(true));
            dataSet.Tables.Add(dtLogMessage);
            dataSet.Tables.Add(dtVariables);
            dataSet.Tables.Add(dtArgument);
            dataSet.Tables.Add(dtWorkflow);
            dataSet.Tables.Add(dtTerminateWF);
            dataSet.Tables.Add(dtKillProcess.DefaultView.ToTable(true));
            dataSet.Tables.Add(dtSendHotKey);
            dataSet.Tables.Add(dtDelay.DefaultView.ToTable(true, "Source", "Type", "Duration"));
            dataSet.Tables.Add(dtSelector);
            dataSet.Tables.Add(dtIDXSelector);
            dataSet.Tables.Add(dtIF);
            dataSet.Tables.Add(dtTryCatch);

            excel.ExportDataTableToExcel(dataSet, excelFolderPath);
            Console.WriteLine("Completed..");

        }


        private static List<String> DirSearch(string sDir)
        {
            List<String> files = new List<String>();
            try
            {
                foreach (string f in Directory.GetFiles(sDir, "*xaml"))
                {
                    files.Add(f);
                }
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files.AddRange(DirSearch(d));
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
            return files;
        }
    }

}
