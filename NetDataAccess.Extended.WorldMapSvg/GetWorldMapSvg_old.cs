using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Config;
using System.Threading;
using System.Windows.Forms;
using mshtml;
using NetDataAccess.Base.Definition;
using System.IO;
using NetDataAccess.Base.Common;
using NPOI.SS.UserModel;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using Newtonsoft.Json.Linq;
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.Reader;
using NetDataAccess.Base.CsvHelper;
using NetDataAccess.Base.DB;
using HtmlAgilityPack;
using System.Web;
using System.Runtime.Remoting;
using System.Reflection;
using System.Collections;
using NetDataAccess.Base.MathProcessor;
using System.Xml;

namespace NetDataAccess.Extended.WorldMapSvg
{
    public class GetWorldMapSvg_old : ExternalRunWebPage
    {
        public override bool AfterAllGrab(IListSheet listSheet)
        {
            this.GetSvgFromFile(listSheet);
            return true;
        }

        private void GetSvgFromFile(IListSheet listSheet)
        { 

            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            string detailPageUrlColumnName = SysConfig.DetailPageUrlFieldName;
            JObject mainJson = new JObject();
            mainJson.Add("code", "world");
            mainJson.Add("name", "World");

            Dictionary<string, string> row = listSheet.GetRow(0);
            bool giveUp = "Y".Equals(row[SysConfig.GiveUpGrabFieldName]);
            if (!giveUp)
            {
                string url = row[detailPageUrlColumnName];

                string filePath = this.RunPage.GetFilePath(url, pageSourceDir);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(filePath);

                XmlNamespaceManager m = new XmlNamespaceManager(xmlDoc.NameTable);
                m.AddNamespace("s", "http://www.w3.org/2000/svg");  

                XmlNodeList pathNodes = xmlDoc.DocumentElement.SelectNodes("s:path", m);
                XmlNode oceanPathNode = pathNodes[0];
                Dictionary<string, object> oceanSvgInfo = this.GetSvgInfoByPath(oceanPathNode, m);
                List<double[]> mainSvgPoints = (List<double[]>)oceanSvgInfo["path"];
                string mainBoundary = this.GetSvgPathStr(mainSvgPoints);

                JArray mainBoundaryArray = new JArray();
                mainBoundaryArray.Add(mainBoundary);
                mainJson.Add("boundaryArray", mainBoundaryArray);

                double[] minMax = this.GetMinMax(mainSvgPoints);
                mainJson.Add("minX", minMax[0]);
                mainJson.Add("minY", minMax[1]);
                mainJson.Add("maxX", minMax[2]);
                mainJson.Add("maxY", minMax[3]);


                JArray nextLevelArray = new JArray(); 

                for (int j = 1; j < pathNodes.Count; j++)
                {
                    XmlNode pathNode = pathNodes[j];
                    Dictionary<string, object> svgInfo = this.GetSvgInfoByPath(pathNode, m);
                    string code = (string)svgInfo["id"];
                    string name = (string)svgInfo["name"];
                    List<double[]> svgPoints = (List<double[]>)svgInfo["path"];
                    string boundary = this.GetSvgPathStr(svgPoints);

                    JObject nextLevelJson = new JObject();
                    nextLevelJson.Add("code", code);
                    nextLevelJson.Add("name", name);

                    JArray boundaryArray = new JArray();
                    boundaryArray.Add(boundary);
                    nextLevelJson.Add("boundaryArray", boundaryArray);

                    nextLevelArray.Add(nextLevelJson);
                }

                XmlNodeList groupNodes = xmlDoc.DocumentElement.SelectNodes("s:g", m);
                foreach (XmlNode groupNode in groupNodes)
                {
                    Dictionary<string, object> svgInfo = this.GetSvgInfoByGroup(groupNode, m);

                    string code = (string)svgInfo["id"];
                    string name = (string)svgInfo["name"];

                    JObject nextLevelJson = new JObject();
                    nextLevelJson.Add("code", code);
                    nextLevelJson.Add("name", name);

                    JArray boundaryArray = new JArray();
                    List<List<double[]>> svgPointsList = (List<List<double[]>>)svgInfo["path"];
                    foreach (List<double[]> svgPoints in svgPointsList)
                    {
                        string boundary = this.GetSvgPathStr(svgPoints);
                        boundaryArray.Add(boundary);
                    }
                    nextLevelJson.Add("boundaryArray", boundaryArray);

                    nextLevelArray.Add(nextLevelJson);
                }
                mainJson.Add("nextLevelArray", nextLevelArray);
            }

            String exportDir = this.RunPage.GetExportDir();
            string resultFilePath = Path.Combine(exportDir, "world.js");
            string worldJs = "boundaryList[\"world\"] = " + mainJson.ToString() +";";
            FileHelper.SaveTextToFile(worldJs, resultFilePath);
        }

        private Dictionary<string, object> GetSvgInfoByGroup(XmlNode groupNode, XmlNamespaceManager m)
        {
            string id = groupNode.Attributes["id"].Value;
            XmlNode titleNode = groupNode.SelectSingleNode("s:title", m);
            string name = titleNode == null ? "" : titleNode.InnerText.Trim();

            List<List<double[]>> svgPointsList = new List<List<double[]>>();

            XmlNodeList pathNodes = groupNode.SelectNodes("s:path", m);
            foreach (XmlNode pathNode in pathNodes)
            {
                List<double[]> pathValues = this.GetPathInfo(pathNode.Attributes["d"].Value);
                svgPointsList.Add(pathValues);
            }

            XmlNodeList subPathNodes = groupNode.SelectNodes("s:g/s:path", m);
            foreach (XmlNode subPathNode in subPathNodes)
            {
                List<double[]> pathValues = this.GetPathInfo(subPathNode.Attributes["d"].Value);
                svgPointsList.Add(pathValues);
            }
             
            Dictionary<string, object> svgInfo = new Dictionary<string, object>();
            svgInfo.Add("id", id);
            svgInfo.Add("name", name);
            svgInfo.Add("path", svgPointsList);
            return svgInfo;
        }

        private string GetSvgPathStr(List<double[]> svgPoints)
        {
            List<string> svgPointStr = new List<string>();
            foreach (double[] p in svgPoints)
            {
                svgPointStr.Add(p[0] + "," + p[1]);
            }
            return CommonUtil.StringArrayToString(svgPointStr.ToArray(), " ");
        }

        private Dictionary<string, object> GetSvgInfoByPath(XmlNode pathNode, XmlNamespaceManager m)
        {
            string id = pathNode.Attributes["id"].Value;
            XmlNode titleNode = pathNode.SelectSingleNode("s:title", m);
            string name = titleNode ==null ?"" :titleNode.InnerText.Trim();
            List<double[]> pathValues = this.GetPathInfo(pathNode.Attributes["d"].Value);
            Dictionary<string, object> svgInfo = new Dictionary<string, object>();
            svgInfo.Add("id", id);
            svgInfo.Add("name", name);
            svgInfo.Add("path", pathValues);
            return svgInfo;
        }

        private double[] GetMinMax(List<double[]> pathValues)
        {
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            foreach (double[] pathValue in pathValues)
            {
                if (pathValue[0] > maxX)
                {
                    maxX = pathValue[0];
                }
                if (pathValue[1] > maxY)
                {
                    maxY = pathValue[1];
                }
                if (pathValue[0] < minX)
                {
                    minX = pathValue[0];
                }
                if (pathValue[1] < minX)
                {
                    minY = pathValue[1];
                }
            }
            return new double[] { minX, minY, maxX, maxY };
        }

        private List<double[]> GetPathInfo(string sourcePath)
        {
            List<double[]> pathValues = new List<double[]>();
            string[] pathParts = sourcePath.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            bool goTo = true;
            double lastX = 0;
            double lastY = 0;
            for (int i = 0; i < pathParts.Length; i++)
            {
                string pathPart = pathParts[i];
                try
                {
                    switch (pathPart)
                    {
                        case "m":
                        case "z":
                        case "l":
                            goTo = true;
                            break;
                        case "c":
                            goTo = false;
                            break;
                        default:
                            {
                                string[] xy = pathPart.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                double x = double.Parse(xy[0].Trim());
                                double y = double.Parse(xy[1].Trim());
                                if (!goTo)
                                {
                                    x = lastX + x;
                                    y = lastY + y;
                                }

                                pathValues.Add(new double[] { x, y });
                                lastX = x;
                                lastY = y;
                            }

                            break;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return pathValues;
        } 

    }
}