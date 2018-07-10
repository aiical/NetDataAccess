using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NetDataAccess.AppAccessBase
{
    /// <summary>
    /// NcpAppElement
    /// </summary>
    public class NcpAppElement
    {
        #region Id
        private string _Id = "";
        public string Id
        {
            get
            {
                return _Id;
            }
            set
            {
                _Id= value;
            }
        }
        #endregion

        #region Name
        private string _Name = "";
        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                _Name = value;
            }
        }
        #endregion

        #region TypeName
        private string _TypeName = "";
        public string TypeName
        {
            get
            {
                return _TypeName;
            }
            set
            {
                _TypeName = value;
            }
        }
        #endregion

        #region Location
        private Point _Location = new Point();
        public Point Location
        {
            get
            {
                return _Location;
            }
            set
            {
                _Location = value;
            }
        }
        #endregion

        #region Size
        private Size _Size = new Size();
        public Size Size
        {
            get
            {
                return _Size;
            }
            set
            {
                _Size = value;
            }
        }
        #endregion

        #region Attributes
        private Dictionary<string, string> _Attributes = new Dictionary<string, string>();
        public Dictionary<string, string> Attributes
        {
            get
            {
                return _Attributes;
            }
            set
            {
                _Attributes = value;
            }
        }
        #endregion

        #region Children
        private List<NcpAppElement> _Children = new List<NcpAppElement>();
        public List<NcpAppElement> Children
        {
            get
            {
                return _Children;
            }
        }
        #endregion
    }
}
