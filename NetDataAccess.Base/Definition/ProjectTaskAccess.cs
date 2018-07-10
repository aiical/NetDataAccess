using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.EnumTypes;
using System.Data;
using NetDataAccess.Base.Common;
using NetDataAccess.Base.Proxy; 

namespace NetDataAccess.Base.Definition
{
    /// <summary>
    /// 对象操作本地数据库
    /// </summary>
    public class ProjectTaskAccess
    {
        #region 添加新组
        /// <summary>
        /// 添加新组
        /// </summary>
        /// <param name="newGroup"></param>
        /// <returns></returns>
        public static string AddNewGroup(Proj_Group newGroup)
        {
            string id = Guid.NewGuid().ToString();
            string addSql = "insert into Proj_Group(id, name, description) "
                + "values(:id, :name, :description)";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", id);
            p2vs.Add("name", newGroup.Name);
            p2vs.Add("description", newGroup.Description);
            if (SqliteHelper.MainDbHelper.ExecuteSql(addSql, p2vs))
            {
                return id;
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 保存更新组
        /// <summary>
        /// 保存更新组
        /// </summary>
        /// <param name="newGroup"></param>
        /// <returns></returns>
        public static bool UpdateGroup(Proj_Group group)
        { 
            string addSql = "update Proj_Group set name=:name, description=:description "
                + "where id=:id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", group.Id);
            p2vs.Add("name", group.Name);
            p2vs.Add("description", group.Description);
            return SqliteHelper.MainDbHelper.ExecuteSql(addSql, p2vs);
        }
        #endregion

        #region 从数据库中删除此分组（打标记，不实际删除记录）
        /// <summary>
        /// 从数据库中删除此分组（打标记，不实际删除记录）
        /// </summary>
        /// <param name="id"></param>
        public static bool DeleteGroup(string id)
        {
            string sql = "update Proj_Group set isdeleted = :isdeleted where id = :id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("isdeleted", "Y");
            p2vs.Add("id", id);
            try
            {
                return SqliteHelper.MainDbHelper.ExecuteSql(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法删除此分组. Id = " + id + "\r\n" + ex.Message);
                return false;
            }
        }
        #endregion

        #region 获取所有分组
        /// <summary>
        /// 获取所有分组
        /// </summary> 
        /// <returns></returns>
        public static List<Proj_Group> GetAllGroupsFromDB()
        {
            string sql = "select g.id as id, g.name as name, g.description as description "
                + "from Proj_Group g where g.isdeleted ='N' order by g.name"; 
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, null);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取分组信息. \r\n" + ex.Message);
            }
            if (dt != null)
            {
                List<Proj_Group> allGroups = new List<Proj_Group>();
                foreach (DataRow row in dt.Rows)
                {
                    Proj_Group g = new Proj_Group(); 
                    g.Id = (string)row["id"];
                    g.Name = (string)row["name"];
                    g.Description = (string)row["description"];
                    allGroups.Add(g);
                }
                return allGroups;
            }
            return null;
        }
        #endregion

        #region 根据分组名称获取分组
        /// <summary>
        /// 根据分组名称获取分组
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public static Proj_Group GetGroupInfoByNameFromDB(string groupName)
        {
            string sql = "select g.id as id, g.name as name, g.description as description "
                + "from Proj_Group g where g.isdeleted ='N' and g.name = :name";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("name", groupName);
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取分组信息. GroupName = " + groupName + "\r\n" + ex.Message);
            }
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Proj_Group g = new Proj_Group();
                    g.Id = (string)row["id"];
                    g.Name = (string)row["name"];
                    g.Description = (string)row["description"];
                    return g;
                }
            }
            return null;
        }
        #endregion

        #region 根据Id获取分组
        /// <summary>
        /// 根据Id获取分组
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static Proj_Group GetGroupInfoByIdFromDB(string id)
        {
            string sql = "select g.id as id, g.name as name, g.description as description "
                + "from Proj_Group g where g.isdeleted ='N' and g.id = :id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", id);
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取分组信息. Id = " + id + "\r\n" + ex.Message);
            }
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Proj_Group g = new Proj_Group();
                    g.Id = (string)row["id"];
                    g.Name = (string)row["name"];
                    g.Description = (string)row["description"];
                    return g;
                }
            }
            return null;
        }
        #endregion
        
        #region 添加项目
        /// <summary>
        /// 添加项目
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static string AddNewProject(Proj_Main newProject)
        {
            string id = Guid.NewGuid().ToString();
            string addSql = @"insert into Proj_Main(id, 
                                name, 
                                description ,
                                group_id, 
                                logintype,
                                loginpageinfo,  
                                detailgrabtype, 
                                detailgrabinfo, 
                                programaftergraball, 
                                programexternalrun) 
                                values(:id, 
                                :name, 
                                :description ,
                                :group_id, 
                                :logintype,
                                :loginpageinfo,  
                                :detailgrabtype, 
                                :detailgrabinfo, 
                                :programaftergraball, 
                                :programexternalrun)";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", id);
            p2vs.Add("name", newProject.Name);
            p2vs.Add("description", newProject.Description); 
            p2vs.Add("group_id", newProject.Group_Id);
            p2vs.Add("logintype", newProject.LoginType);
            p2vs.Add("loginpageinfo", newProject.LoginPageInfo); 
            p2vs.Add("detailgrabtype", newProject.DetailGrabType);
            p2vs.Add("detailgrabinfo", newProject.DetailGrabInfo);
            p2vs.Add("programaftergraball", newProject.ProgramAfterGrabAll);
            p2vs.Add("programexternalrun", newProject.ProgramExternalRun);  
            if (SqliteHelper.MainDbHelper.ExecuteSql(addSql, p2vs))
            {
                return id;
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 保存更新项目
        /// <summary>
        /// 保存更新项目
        /// </summary>
        /// <param name="project"></param>
        /// <returns></returns>
        public static bool UpdateProject(Proj_Main project)
        {
            string addSql = @"update Proj_Main set name=:name,
                                description=:description ,
                                group_id=:group_id , 
                                logintype=:logintype,
                                loginpageinfo=:loginpageinfo,  
                                detailgrabtype=:detailgrabtype, 
                                detailgrabinfo=:detailgrabinfo, 
                                programaftergraball=:programaftergraball,
                                programexternalrun=:programexternalrun 
                                where id=:id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", project.Id);
            p2vs.Add("name", project.Name);
            p2vs.Add("description", project.Description); 
            p2vs.Add("group_id", project.Group_Id);  
            p2vs.Add("logintype", project.LoginType);
            p2vs.Add("loginpageinfo", project.LoginPageInfo); 
            p2vs.Add("detailgrabtype", project.DetailGrabType);
            p2vs.Add("detailgrabinfo", project.DetailGrabInfo);
            p2vs.Add("programaftergraball", project.ProgramAfterGrabAll);
            p2vs.Add("programexternalrun", project.ProgramExternalRun);  
            return SqliteHelper.MainDbHelper.ExecuteSql(addSql, p2vs);
        }
        #endregion

        #region 从数据库中删除此项目（打标记，不实际删除记录）
        /// <summary>
        /// 从数据库中删除此项目（打标记，不实际删除记录）
        /// </summary>
        /// <param name="id"></param>
        public static bool DeleteProject(string id)
        {
            string sql = "update Proj_Main set isdeleted = :isdeleted where id = :id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("isdeleted", "Y");
            p2vs.Add("id", id);
            try
            {
                return SqliteHelper.MainDbHelper.ExecuteSql(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法删除此分组. Id = " + id + "\r\n" + ex.Message);
                return false;
            }
        }
        #endregion

        #region 获取所有项目
        /// <summary>
        /// 获取所有项目
        /// </summary> 
        /// <returns></returns>
        public static List<Proj_Main> GetAllProjectsFromDB()
        {
            string sql = @"select p.id as id,
                            p.name as name, 
                            p.description as description, 
                            p.group_id as group_id,  
                            p.logintype as logintype,
                            p.loginpageinfo as loginpageinfo,
                            p.detailgrabtype as detailgrabtype,
                            p.detailgrabinfo as detailgrabinfo,
                            p.programaftergraball as programaftergraball,
                            p.programexternalrun as programexternalrun from Proj_Main p where p.isdeleted = 'N' order by p.name";
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, null);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取项目信息. \r\n" + ex.Message);
            }
            if (dt != null)
            {
                List<Proj_Main> allProjects = new List<Proj_Main>();
                foreach (DataRow row in dt.Rows)
                {
                    Proj_Main p = new Proj_Main();
                    p.Id = (string)row["id"];
                    p.Name = (string)row["name"];
                    p.Description = (string)row["description"]; 
                    p.Group_Id = (string)row["group_id"];  
                    p.LoginType = (LoginLevelType)Enum.Parse(typeof(LoginLevelType), (string)row["logintype"]);
                    p.LoginPageInfo = (string)row["loginpageinfo"];
                    p.DetailGrabType = (DetailGrabType)Enum.Parse(typeof(DetailGrabType), (string)row["detailgrabtype"]);
                    p.DetailGrabInfo = (string)row["detailgrabinfo"];
                    p.ProgramAfterGrabAll = (string)row["programaftergraball"];
                    p.ProgramExternalRun = (string)row["programexternalrun"];
                    allProjects.Add(p);
                }
                return allProjects;
            }
            return null;
        }
        #endregion

        #region 根据名称获取项目
        /// <summary>
        /// 根据名称获取项目
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Proj_Main GetProjectInfoByNameFromDB(string name)
        {
            string sql = @"select p.id as id,
                            p.name as name, 
                            p.description as description, 
                            p.group_id as group_id,  
                            p.logintype as logintype,
                            p.loginpageinfo as loginpageinfo, 
                            p.detailgrabtype as detailgrabtype,
                            p.detailgrabinfo as detailgrabinfo,
                            p.programaftergraball as programaftergraball,
                            p.programexternalrun as programexternalrun from Proj_Main p where p.isdeleted = 'N' and p.name = :name";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("name", name);
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取项目信息. ProjectName = " + name + "\r\n" + ex.Message);
            }
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Proj_Main p = new Proj_Main();
                    p.Id = (string)row["id"];
                    p.Name = (string)row["name"];
                    p.Description = (string)row["description"]; 
                    p.Group_Id = (string)row["group_id"];  
                    p.LoginType = (LoginLevelType)Enum.Parse(typeof(LoginLevelType), (string)row["logintype"]);
                    p.LoginPageInfo = (string)row["loginpageinfo"];
                    p.DetailGrabType = (DetailGrabType)Enum.Parse(typeof(DetailGrabType), (string)row["detailgrabtype"]);
                    p.DetailGrabInfo = (string)row["detailgrabinfo"];
                    p.ProgramAfterGrabAll = (string)row["programaftergraball"];
                    p.ProgramExternalRun = (string)row["programexternalrun"];
                    return p;
                }
            }
            return null;
        }
        #endregion

        #region 根据Id获取项目
        /// <summary>
        /// 根据Id获取项目
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static Proj_Main GetProjectInfoByIdFromDB(string id)
        {
            string sql = @"select p.id as id,
                            p.name as name,  
                            p.description as description,  
                            p.group_id as group_id,  
                            p.logintype as logintype,
                            p.loginpageinfo as loginpageinfo, 
                            p.detailgrabtype as detailgrabtype,
                            p.detailgrabinfo as detailgrabinfo,
                            p.programaftergraball as programaftergraball,
                            p.programexternalrun as programexternalrun from Proj_Main p where p.isdeleted = 'N' and p.id = :id";
            Dictionary<string, object> p2vs = new Dictionary<string, object>();
            p2vs.Add("id", id);
            DataTable dt = null;
            try
            {
                dt = SqliteHelper.MainDbHelper.GetDataTable(sql, p2vs);
            }
            catch (Exception ex)
            {
                CommonUtil.Alert("错误提示", "无法获取项目信息. Id = " + id + "\r\n" + ex.Message);
            }
            if (dt != null)
            {
                if (dt.Rows.Count > 0)
                {
                    DataRow row = dt.Rows[0];
                    Proj_Main p = new Proj_Main();
                    p.Id = (string)row["id"];
                    p.Name = (string)row["name"];
                    p.Description = (string)row["description"]; 
                    p.Group_Id = (string)row["group_id"];  
                    p.LoginType = (LoginLevelType)Enum.Parse(typeof(LoginLevelType), (string)row["logintype"]);
                    p.LoginPageInfo = (string)row["loginpageinfo"];
                    p.DetailGrabType = (DetailGrabType)Enum.Parse(typeof(DetailGrabType), (string)row["detailgrabtype"]);
                    p.DetailGrabInfo = (string)row["detailgrabinfo"];
                    p.ProgramAfterGrabAll = (string)row["programaftergraball"];
                    p.ProgramExternalRun = (string)row["programexternalrun"];
                    return p;
                }
            }
            return null;
        }
        #endregion 
    }
}
