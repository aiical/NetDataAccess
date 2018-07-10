using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class HtmlDocumentHelper
    {
        public static HtmlNode GetNextNode(HtmlNode currentNode, string nextNodeName)
        {
            HtmlNode nextNode = currentNode.NextSibling;
            while (nextNode != null)
            {
                if (nextNode.Name == nextNodeName)
                {
                    break;
                }
                else
                {
                    nextNode = nextNode.NextSibling;
                }
            }
            return nextNode;
        }

        public static HtmlAgilityPack.HtmlDocument Load(string localFilePath)
        {
            return HtmlDocumentHelper.Load(localFilePath, Encoding.UTF8);
        }

        public static HtmlAgilityPack.HtmlDocument Load(string localFilePath, Encoding encoding)
        {
            StreamReader tr = new StreamReader(localFilePath, encoding);
            string webPageHtml = tr.ReadToEnd();

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(webPageHtml);
            tr.Close();
            tr.Dispose();
            return htmlDoc;
        }

        public static string JointNodeAttributeValue(HtmlNode parentNode, string xPath, string attributeName, string splitString, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            HtmlNodeCollection nodes = parentNode.SelectNodes(xPath);
            return JointNodeAttributeValue(nodes, attributeName, splitString, ignoreBlank, needTrim, prefix, postfix);
        }

        public static string JointNodeAttributeValue(IList<HtmlNode> nodes, string attributeName, string splitString, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            return JointNodeAttributeValue(nodes == null ? null : nodes.ToArray(), attributeName, splitString, ignoreBlank, needTrim, prefix, postfix);
        }

        public static string JointNodeAttributeValue(HtmlNode[] nodes, string attributeName, string splitString, bool ignoreBlank, bool needTrim)
        {
            return JointNodeAttributeValue(nodes, attributeName, splitString, ignoreBlank, needTrim, null, null);
        }

        public static string JointNodeAttributeValue(HtmlNode[] nodes, string attributeName, string splitString, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            string s = null;
            if (nodes != null && nodes.Length > 0)
            {
                StringBuilder sBuilder = new StringBuilder();
                foreach (HtmlNode node in nodes)
                {
                    string value = node.GetAttributeValue(attributeName, "");
                    if (needTrim)
                    {
                        value = value.Trim();
                    }
                    if (value.Length > 0 || !ignoreBlank)
                    {
                        sBuilder.Append(node.InnerText.Trim() + splitString);
                    }
                }
                s =  sBuilder.ToString();
            }
            return ((s == null || s.Length == 0) && ignoreBlank) ? null : (prefix + s + postfix);
        }

        public static string JointNodeInnerText(IList<HtmlNode> nodes, string splitString, bool ignoreBlank, bool needTrim)
        {
            return JointNodeInnerText(nodes, splitString, ignoreBlank, needTrim, null, null);
        }

        public static string JointNodeInnerText(IList<HtmlNode> nodes, string splitString, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            return JointNodeInnerText(nodes == null ? null : nodes.ToArray(), splitString, ignoreBlank, needTrim, prefix, postfix);
        }

        public static string JointNodeInnerText(HtmlNode[] nodes, string splitString, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            string s = null;
            if (nodes != null && nodes.Length > 0)
            {
                StringBuilder sBuilder = new StringBuilder();
                foreach (HtmlNode node in nodes)
                {
                    string value = CommonUtil.HtmlDecode(node.InnerText);

                    if (needTrim)
                    {
                        value = value.Trim();
                    }
                    if (value.Length > 0 || !ignoreBlank)
                    {
                        sBuilder.Append(value + splitString);
                    }
                }
                s =  sBuilder.ToString();
            } 
            return ((s == null || s.Length == 0) && ignoreBlank) ? null : (prefix + s + postfix);
        }

        public static string JointNodeInnerText(HtmlNode parentNode, string xPath, string splitString, bool ignoreBlank, bool needTrim, string  prefix, string  postfix)
        {
            HtmlNodeCollection nodes = parentNode.SelectNodes(xPath);
            return JointNodeInnerText(nodes, splitString, ignoreBlank, needTrim, prefix, postfix);
        }

        public static string TryGetNodeInnerText(HtmlNode parentNode, string xPath, bool needTrim)
        {
            HtmlNode node = parentNode.SelectSingleNode(xPath);
            return node == null ? null : CommonUtil.HtmlDecode(needTrim ? node.InnerText.Trim() : node.InnerText);
        }

        /// <summary>
        /// 判断xpath对应的节点中包含checkText文字
        /// </summary>
        /// <param name="parentNode"></param>
        /// <param name="xPath"></param>
        /// <param name="checkText"></param>
        /// <returns></returns>
        public static bool CheckNodeContainsText(HtmlNode parentNode, string xPath, string checkText, bool caseSensitive)
        {
            HtmlNodeCollection nodes = parentNode.SelectNodes(xPath);
            if (nodes == null || nodes.Count == 0)
            {
                return false;
            }
            else
            {
                foreach (HtmlNode node in nodes)
                {
                    if (caseSensitive)
                    {
                        if (CommonUtil.HtmlDecode(node.InnerText).Contains(checkText))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        if (CommonUtil.HtmlDecode(node.InnerText).ToLower().Contains(checkText.ToLower()))
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
        } 

        public static string TryGetNodeInnerText(HtmlNode parentNode, string xPath, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            HtmlNode node = parentNode.SelectSingleNode(xPath);
            string value = node == null ? null : CommonUtil.HtmlDecode(needTrim ? node.InnerText.Trim() : node.InnerText);
            return ((value == null || value.Length == 0) && ignoreBlank) ? null : (prefix + value + postfix);
        }

        public static string TryGetNodeInnerText(HtmlNode node, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        { 
            string value = node == null ? null : CommonUtil.HtmlDecode(needTrim ? node.InnerText.Trim() : node.InnerText);
            return ((value == null || value.Length == 0) && ignoreBlank) ? null : (prefix + value + postfix);
        }

        public static string TryGetNodeAttributeValue(HtmlNode parentNode, string xPath, string attributeName, bool needTrim)
        {
            HtmlNode node = parentNode.SelectSingleNode(xPath);
            return node == null ? null : needTrim ? node.GetAttributeValue(attributeName, "").Trim() : node.GetAttributeValue(attributeName, "");
        }

        public static string TryGetNodeAttributeValue(HtmlNode parentNode, string xPath, string attributeName, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        {
            HtmlNode node = parentNode.SelectSingleNode(xPath);
            string value = node == null ? null : needTrim ? node.GetAttributeValue(attributeName, "").Trim() : node.GetAttributeValue(attributeName, "");
            return ((value == null || value.Length == 0) && ignoreBlank) ? null : (prefix + value + postfix);
        }

        public static string TryGetNodeAttributeValue(HtmlNode node, string attributeName, bool ignoreBlank, bool needTrim, string prefix, string postfix)
        { 
            string value = node == null ? null : needTrim ? node.GetAttributeValue(attributeName, "").Trim() : node.GetAttributeValue(attributeName, "");
            return ((value == null || value.Length == 0) && ignoreBlank) ? null : (prefix + value + postfix);
        }
    }
}
