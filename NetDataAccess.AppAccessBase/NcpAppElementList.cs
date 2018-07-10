using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NetDataAccess.AppAccessBase
{
    /// <summary>
    /// NcpAppElementList
    /// </summary>
    public class NcpAppElementList : List<NcpAppElement>
    {
        #region SortByPosition
        public NcpAppElementList SortByPosition()
        {
            NcpAppElementList newList = new NcpAppElementList();
            foreach (NcpAppElement element in this)
            {
                Point location = element.Location;
                int posNum = location.Y * 1000 + location.X;
                int toIndex = newList.Count;
                for (int i = 0; i < newList.Count; i++)
                {
                    NcpAppElement tempElement = newList[i];
                    Point tempLocation = tempElement.Location;
                    int tempPosNum = tempLocation.Y * 1000 + tempLocation.X;
                    if (tempPosNum > posNum)
                    {
                        toIndex = i;
                        break;
                    }
                }
                newList.Insert(toIndex, element);
            }
            return newList;
        }
        #endregion

        #region Add
        public NcpAppElement Add(string id, string name, string typeName, Point location, Size size)
        {
            NcpAppElement element = new NcpAppElement();
            element.Id = id;
            element.Name = name;
            element.Location = location;
            element.TypeName = typeName;
            element.Size = size;
            this.Add(element);
            return element;
        }
        #endregion

        #region Add
        public NcpAppElement Add(string id, string name, string typeName)
        {
            NcpAppElement element = new NcpAppElement();
            element.Id = id;
            element.Name = name; 
            element.TypeName = typeName; 
            this.Add(element);
            return element;
        }
        #endregion

        #region Exist
        public bool Exist(string id)
        {
            foreach (NcpAppElement element in this)
            {
                if (element.Id == id)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion
    }
}
