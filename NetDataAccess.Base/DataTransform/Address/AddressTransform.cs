using NetDataAccess.Base.Reader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.DataTransform.Address
{
    public class AddressTransform
    {
        private XZQHMap _DefaultXZQHMap = null;
        public XZQHMap DefaultXZQHMap
        {
            get
            {
                return _DefaultXZQHMap;
            }
        }

        public void InitDeaultXZQHMap(string xzqhExcelFilePath)
        {
            XZQHMap map = new XZQHMap();

            ExcelReader er = new ExcelReader(xzqhExcelFilePath);
            int rowCount = er.GetRowCount();
            List<string> codeList = new List<string>();
            for (int i = 0; i < rowCount; i++)
            {
                Dictionary<string, string> row = er.GetFieldValues(i);
                string code = row["code"];
                string name = row["name"];
                string isCity = row["isCity"];
                string isProvince = row["isProvince"];

                XZQHArea area = new XZQHArea();
                area.Code = code;
                area.Name = name;
                area.IsCity = isCity == "是";
                area.IsProvince = isProvince == "是";

                string alias = row["alias"];
                
                if (alias != null && alias.Length>0)
                {
                    string[] aliasNames = alias.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    area.AliasNames = new List<string>(aliasNames);
                }

                map.AreaMap.Add(area.Code, area);
                codeList.Add(area.Code);
            }
            
            List<string> rootAreaCodes = new List<string>();
            foreach (string code in codeList)
            {
                XZQHArea area = map.GetArea(code);
                if (code.EndsWith("0000"))
                {
                    rootAreaCodes.Add(code);
                }
                else if (code.EndsWith("00"))
                {
                    string parentAreaCode = code.Substring(0, 2) + "0000";
                    XZQHArea parentArea = map.GetArea(parentAreaCode);
                    area.ParentAreaCode = parentAreaCode;
                    if (parentArea.ChildAreaCodes == null)
                    {
                        parentArea.ChildAreaCodes = new List<string>();
                    }
                    parentArea.ChildAreaCodes.Add(code);
                }
                else
                {
                    string parentAreaCode = code.Substring(0, 4) + "00";
                    XZQHArea parentArea = map.GetArea(parentAreaCode);
                    if (parentArea == null)
                    {
                        parentAreaCode = code.Substring(0, 2) + "0000";
                        parentArea = map.GetArea(parentAreaCode);
                    }

                    area.ParentAreaCode = parentAreaCode;
                    if (parentArea.ChildAreaCodes == null)
                    {
                        parentArea.ChildAreaCodes = new List<string>();
                    }
                    parentArea.ChildAreaCodes.Add(code);
                }
            }
            map.RootAreaCodes = rootAreaCodes;
            this._DefaultXZQHMap = map;
        }

        public List<string> GetAddressParts(string address, bool returnWithCodeAndName)
        {
            foreach (string rootAreaCode in this.DefaultXZQHMap.RootAreaCodes)
            {
                XZQHArea rootArea = this.DefaultXZQHMap.GetArea(rootAreaCode);
                List<string> parts = this.GetAddressParts(address, rootArea, returnWithCodeAndName);
                if (parts != null)
                {
                    return parts;
                }
            }
            foreach (string rootAreaCode in this.DefaultXZQHMap.RootAreaCodes)
            {
                List<String> childAreaCodes = this.DefaultXZQHMap.GetArea(rootAreaCode).ChildAreaCodes;
                if(childAreaCodes!=null ){
                    foreach (string childAreaCode in childAreaCodes)
                    {
                        XZQHArea area = this.DefaultXZQHMap.GetArea(childAreaCode);
                        List<string> parts = this.GetAddressParts(address, area, returnWithCodeAndName);
                        if (parts != null)
                        {
                            return parts;
                        }
                    }
                }
            } 
            return null;
        }


        private List<string> GetAddressParts(string address, XZQHArea area, bool returnWithCodeAndName)
        {
            List<string> parts = area.CheckInArea(address, returnWithCodeAndName);
            if (parts != null)
            {
                List<string> childAreaCodes = area.ChildAreaCodes;
                if (childAreaCodes != null)
                {
                    string partAddress = parts[1];
                    foreach (string childAreaCode in childAreaCodes)
                    {
                        XZQHArea childArea = this.DefaultXZQHMap.GetArea(childAreaCode);
                        List<string> childParts = this.GetAddressParts(partAddress, childArea, returnWithCodeAndName);
                        if (childParts != null)
                        {
                            parts.RemoveAt(parts.Count - 1);
                            parts.AddRange(childParts.ToArray());
                            return parts;
                        }
                    }
                }
                return parts;
            }
            return null;
        }
        

        public List<string> GetAddressParts(string parentAreaCode, string address, bool returnWithCodeAndName)
        {
            XZQHArea parentArea = this.DefaultXZQHMap.GetArea(parentAreaCode);
            if (parentArea != null)
            {
                foreach (string areaCode in this.DefaultXZQHMap.AreaMap.Keys)
                {
                    XZQHArea area = this.DefaultXZQHMap.GetArea(areaCode);
                    List<string> parts = this.GetAddressParts(address, area, returnWithCodeAndName);
                    if (parts != null)
                    {
                        string lastLevelAreaInfo = parts[parts.Count - 2];
                        string checkAreaCode = lastLevelAreaInfo.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries)[0].Substring(5);
                        if (this.IsParent(parentAreaCode, checkAreaCode))
                        {
                            return parts;
                        }
                    }
                }

                if (parentArea.IsCity)
                {
                    return new List<string>() { address };
                }
            }
            return null;
        }

        private bool IsParent(string parentAreaCode, string checkAreaCode)
        {
            if (parentAreaCode == checkAreaCode)
            {
                return true;
            }
            else
            {
                if (checkAreaCode == null || checkAreaCode.Length == 0)
                {
                    return false;
                }
                else
                {
                    XZQHArea checkArea = this.DefaultXZQHMap.GetArea(checkAreaCode);
                    return this.IsParent(parentAreaCode, checkArea.ParentAreaCode);
                }
            }
        }

        public List<string> GetAreaParts(string areaFullName)
        {
            foreach (string rootAreaCode in this.DefaultXZQHMap.RootAreaCodes)
            {
                XZQHArea rootArea = this.DefaultXZQHMap.GetArea(rootAreaCode);
                List<string> parts = this.GetAreaParts(areaFullName, rootArea);
                if (parts != null)
                {
                    return parts;
                }
            }
            return null;
        }

        private List<string> GetAreaParts(string areaFullName, XZQHArea area)
        {
            List<string> areaParts = area.CheckIsArea(areaFullName);
            if (areaParts != null)
            {
                if (areaParts[areaParts.Count - 1].Length == 0)
                {
                    areaParts.RemoveAt(areaParts.Count - 1);
                    return areaParts;
                }
                else
                {
                    List<string> childAreaCodes = area.ChildAreaCodes;
                    if (childAreaCodes != null)
                    {
                        string partAreaName = areaParts[1];
                        if (partAreaName != null && partAreaName.Length > 0)
                        {
                            foreach (string childAreaCode in childAreaCodes)
                            {
                                XZQHArea childArea = this.DefaultXZQHMap.GetArea(childAreaCode);
                                List<string> childParts = this.GetAreaParts(partAreaName, childArea);
                                if (childParts != null)
                                {
                                    areaParts.RemoveAt(areaParts.Count - 1);
                                    areaParts.AddRange(childParts.ToArray());
                                    return areaParts;
                                }
                            }
                        }
                        else
                        {
                            return areaParts;
                        }
                    }
                }
                return areaParts;
            }
            return null;
        }
    }
}
