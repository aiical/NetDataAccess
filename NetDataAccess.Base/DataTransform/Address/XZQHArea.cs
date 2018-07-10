using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.DataTransform.Address
{
    public class XZQHArea
    {
        private string _Code = "";
        public string Code
        {
            get
            {
                return this._Code;
            }
            set
            {
                this._Code = value;
            }
        }

        private string _Name = "";
        public string Name
        {
            get
            {
                return this._Name;
            }
            set
            {
                this._Name = value;
            }
        }

        private bool _IsCity = false;
        public bool IsCity
        {
            get
            {
                return this._IsCity;
            }
            set
            {
                this._IsCity = value;
            }
        }

        private bool _IsProvince = false;
        public bool IsProvince
        {
            get
            {
                return this._IsProvince;
            }
            set
            {
                this._IsProvince = value;
            }
        }

        private List<string> _AliasNames = null;
        public List<string> AliasNames
        {
            get
            {
                return this._AliasNames;
            }
            set
            {
                this._AliasNames = value;
            }
        }


        private List<string> _ChildAreaCodes = null;
        public List<string> ChildAreaCodes
        {
            get
            {
                return this._ChildAreaCodes;
            }
            set
            {
                this._ChildAreaCodes = value;
            }
        }

        private string _ParentAreaCode = null;
        public string ParentAreaCode
        {
            get
            {
                return this._ParentAreaCode;
            }
            set
            {
                this._ParentAreaCode = value;
            }
        }

        public bool CheckIsSame(string checkName)
        {
            if (checkName == this.Name)
            {
                return true;
            }
            else
            {
                if (this.AliasNames.Contains(checkName))
                {
                    return true;
                }
            }
            return false;
        }

        public List<string> CheckInArea(string address, bool returnWithCodeAndName)
        {
            if (address.StartsWith(this.Name))
            {
                List<string> parts = new List<string>();
                parts.Add(!returnWithCodeAndName ? this.Code : ("code:" + this.Code + ",name:" + this.Name));
                parts.Add(address.Substring(this.Name.Length));
                return parts;
            }
            else
            {
                foreach (string name in this.AliasNames)
                {
                    if (address.StartsWith(name))
                    {
                        List<string> parts = new List<string>();
                        parts.Add(!returnWithCodeAndName ? this.Code : ("code:" + this.Code + ",name:" + name));
                        parts.Add(address.Substring(name.Length));
                        return parts;
                    }
                }
            }
            return null;
        }

        public List<string> CheckIsArea(string areaFullName)
        {
            if (areaFullName.StartsWith(this.Name))
            {
                List<string> parts = new List<string>();
                parts.Add(this.Code);
                parts.Add(areaFullName.Substring(this.Name.Length));
                return parts;
            }
            else
            {
                foreach (string name in this.AliasNames)
                {
                    if (areaFullName.StartsWith(name))
                    {
                        List<string> parts = new List<string>();
                        parts.Add(this.Code);
                        parts.Add(areaFullName.Substring(name.Length));
                        return parts;
                    }
                }
            }
            return null;
        }
    }
}
