using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace CodeAnalyserConsoleApp
{
    class LoadXAML
    {
        public XmlDocument GetDocument(string XAMLFilePath)
        {
            string xml = File.ReadAllText(XAMLFilePath);

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(xml);
            XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
            xmlNamespaceManager.AddNamespace("ui", "http://schemas.uipath.com/workflow/activities");

            return xmlDocument;

        }

        public void GetParentNodes(XmlNode node, List<string> lstNodes)
        {
            if ((node != null && node.Name != null))
            {
                lstNodes.Add(node.Name);
                GetParentNodes(node.ParentNode, lstNodes);
            }
        }

        public void GetChildNodes(XmlNode node, List<string> lstNodes, string NodeName)
        {
            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    if ((xmlNode.Name != null) && xmlNode.Name.Equals(NodeName))
                        lstNodes.Add(xmlNode.Name);
                    GetChildNodes(xmlNode, lstNodes, NodeName);
                }
            }
        }

        public void GetNodesWithAttributes(XmlNode node, string attributeName, DataTable dt, string source)
        {
            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    if ((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes[attributeName] != null)
                    {
                        dt.Rows.Add(source, xmlNode.Name, xmlNode.Attributes[attributeName].Value, xmlNode.ParentNode.OuterXml);

                    }
                    GetNodesWithAttributes(xmlNode, attributeName, dt, source);
                }
            }
        }

        public void GetNodesWithAttributes(XmlNode node, string attributeName, string attributeName1, DataTable dt, string source)
        {

            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    if ((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes[attributeName] != null)
                    {

                        string attr1 = "", attr2 = "";

                        if (xmlNode.Attributes[attributeName] != null)
                        {
                            attr1 = xmlNode.Attributes[attributeName].Value;
                        }
                        if (xmlNode.Attributes[attributeName1] != null)
                        {
                            attr2 = xmlNode.Attributes[attributeName1].Value;
                        }

                        if ((attributeName.Equals("SimulateClick") || attributeName.Equals("SimulateType") || attributeName.Equals("SendWindowMessages")) && (((XmlElement)xmlNode).GetElementsByTagName("ui:Target")).Count > 0)
                            dt.Rows.Add(source, xmlNode.Name, attr1, attr2, (((XmlElement)xmlNode).GetElementsByTagName("ui:Target")[0].Attributes["Selector"] == null) ? "" : ((XmlElement)xmlNode).GetElementsByTagName("ui:Target")[0].Attributes["Selector"].Value);
                        else
                        {
                            if (!attr2.Equals("Do"))
                            {

                                dt.Rows.Add(source, xmlNode.Name, attr1, attr2);
                            }
                        }
                    }

                    GetNodesWithAttributes(xmlNode, attributeName, attributeName1, dt, source);
                }
            }
        }

        public void GetNodesWithAttributes(XmlNode node, string attributeName, string attributeName1, DataTable dt, string source, int count = 1)
        {

            if (node.ChildNodes.Count > 0)
            {
                count++;
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    if ((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes[attributeName] != null)
                    {
                        if (xmlNode.Name.Equals("TryCatch"))
                        {
                            LoadXAML xaml = new LoadXAML();

                            DataTable dtActivities = new DataTable("Activities");
                            dtActivities.Columns.Add("Source");
                            dtActivities.Columns.Add("Target");
                            dtActivities.Columns.Add("SendWindowMessage");
                            dtActivities.Columns.Add("DisplayName");
                            dtActivities.Columns.Add("Depth");
                            xaml.GetNodesWithAttributes(xmlNode.ChildNodes[1], "sap2010:WorkflowViewState.IdRef", "DisplayName", dtActivities, source, 0);
                        }
                        string attr1 = "", attr2 = "";

                        if (xmlNode.Attributes[attributeName] != null)
                        {
                            attr1 = xmlNode.Attributes[attributeName].Value;
                        }
                        if (xmlNode.Attributes[attributeName1] != null)
                        {
                            attr2 = xmlNode.Attributes[attributeName1].Value;
                        }

                        if ((attributeName.Equals("SimulateClick") || attributeName.Equals("SimulateType") || attributeName.Equals("SendWindowMessages")) && (((XmlElement)xmlNode).GetElementsByTagName("ui:Target")).Count > 0)
                            dt.Rows.Add(source, xmlNode.Name, attr1, attr2, (((XmlElement)xmlNode).GetElementsByTagName("ui:Target")[0].Attributes["Selector"] == null) ? "" : ((XmlElement)xmlNode).GetElementsByTagName("ui:Target")[0].Attributes["Selector"].Value);
                        else
                        {
                            if (!attr2.Equals("Do"))
                            {

                                dt.Rows.Add(source, xmlNode.Name, attr1, attr2, count.ToString());
                            }
                        }
                    }

                    GetNodesWithAttributes(xmlNode, attributeName, attributeName1, dt, source, count);
                }
            }
        }

        public Dictionary<string, int> UniqueActivityCount(XmlNode node,  Dictionary<string, int> dicNodes)
        {
            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    
                    if (((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes["sap2010:WorkflowViewState.IdRef"] != null))
                    {
                        string displayName = "";
                        if (xmlNode.Attributes["DisplayName"] != null)
                        {
                            displayName = xmlNode.Attributes["DisplayName"].Value;
                        }
                        if (!dicNodes.ContainsKey(xmlNode.Name))
                            dicNodes[xmlNode.Name] = 1;
                        else
                            dicNodes[xmlNode.Name]++;

                        Console.WriteLine("DisplayName: " + displayName + "   NodeName: " + xmlNode.Name);
                        Console.WriteLine("----------");
                    }
                    UniqueActivityCount(xmlNode, dicNodes);
                }
            }
            return dicNodes;
        }

        public Dictionary<int, int> AttributeExist(XmlNode node, string attributeName, Dictionary<int, int> dicNodes, int counter = 0)
        {

            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {

                    if (((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes["sap2010:WorkflowViewState.IdRef"] != null) || xmlNode.Name.Contains(attributeName))
                    {

                        AttributeExist(xmlNode, attributeName, dicNodes, counter);
                        if (xmlNode.Name.Equals(attributeName))
                        {
                            string displayName = "";
                            if (xmlNode.Attributes["DisplayName"] != null)
                            {
                                displayName = xmlNode.Attributes["DisplayName"].Value;
                            }




                            if (!xmlNode.Name.Equals("Activity"))
                            {
                                counter = counter + 1;
                                if (!dicNodes.Keys.Contains(counter))
                                {
                                    dicNodes[counter] = 1;
                                }
                                else
                                    dicNodes[counter]++;
                                Console.WriteLine(counter + "   DisplayName: " + displayName + "   NodeName: " + xmlNode.Name);
                                Console.WriteLine("----------");
                            }
                        }
                    }


                }

            }
            return dicNodes;
        }

        public void GetNodesWithAttributes(XmlNode node, string attributeName, string attributeName1, string attributeName2, DataTable dt, string source)
        {
            if (node.ChildNodes.Count > 0)
            {
                foreach (XmlNode xmlNode in node.ChildNodes)
                {
                    if ((xmlNode.Name != null) && xmlNode.Attributes != null && xmlNode.Attributes[attributeName] != null)
                    {



                        string attr1 = "";
                        string attr2 = "";
                        string attr3 = "";
                        if (xmlNode.Attributes[attributeName] != null)
                        {
                            attr1 = xmlNode.Attributes[attributeName].Value;
                        }
                        if (xmlNode.Attributes[attributeName1] != null)
                        {
                            attr2 = xmlNode.Attributes[attributeName1].Value;
                            if (xmlNode.Attributes["SimulateType"] != null)
                            {
                                attr1 = xmlNode.Attributes["SimulateType"].Value;
                            }
                        }
                        if (xmlNode.Attributes[attributeName2] != null)
                        {
                            attr3 = xmlNode.Attributes[attributeName2].Value;
                        }

                        if (attributeName.Equals("SimulateClick") || attributeName.Equals("SimulateType"))
                            dt.Rows.Add(source, xmlNode.Name, attr1, attr2, attr3, ((XmlElement)xmlNode).GetElementsByTagName("ui:Target")[0].Attributes["Selector"].Value);
                        else if (attributeName.Equals("TimeoutMS"))
                        {
                            if (xmlNode.Name.Equals("ui:Target"))
                                dt.Rows.Add(source, xmlNode.Name, (xmlNode.ParentNode.ParentNode.Attributes["DisplayName"] == null) ? "" : xmlNode.ParentNode.ParentNode.Attributes["DisplayName"].Value, attr1, attr2, attr3);
                            else
                                dt.Rows.Add(source, xmlNode.Name, "", attr1, attr2, attr3);

                        }
                        else
                            dt.Rows.Add(source, xmlNode.Name, attr1, attr2, attr3);

                    }
                    GetNodesWithAttributes(xmlNode, attributeName, attributeName1, attributeName2, dt, source);
                }
            }
        }


    }
}
