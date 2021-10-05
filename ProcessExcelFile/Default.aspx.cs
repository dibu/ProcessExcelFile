using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;
using System.Data;
using System.Xml;

namespace ProcessExcelFile {
    public partial class _Default : Page {
       

        protected void Page_Load(object sender, EventArgs e) {

        }

        protected void btnProcessExcel_Click(object sender, EventArgs e) {
            try {
                string filePath = Server.MapPath("~/ExcelFiles/MenuData.xlsx");
                if (File.Exists(filePath)) {
                    using (var excelWorkBook = new XLWorkbook(filePath)) {
                        var workSheet = excelWorkBook.Worksheet("MenuList");
                        for (int row = 2; row <= 22; row++) {
                            //for (int column = 1; column <= 6; column++) {
                            //    var currentCell = workSheet.Cell(row, column);
                            //}

                            var depthCell = workSheet.Cell(row, 6);

                            var firstCellValue = workSheet.Cell(row, 1).GetValue<string>();
                            var secondCellValue = workSheet.Cell(row, 2).GetValue<string>();
                            var thirdCellValue = workSheet.Cell(row, 3).GetValue<string>();
                            var fourthCellValue = workSheet.Cell(row, 4).GetValue<string>();
                            if (!string.IsNullOrEmpty(fourthCellValue)) {
                                depthCell.SetValue<int>(4).Style.Fill.BackgroundColor= XLColor.LightGreen;
                            } else if (!string.IsNullOrEmpty(thirdCellValue) && string.IsNullOrEmpty(fourthCellValue)) {
                                depthCell.SetValue<int>(3).Style.Fill.BackgroundColor= XLColor.Amber;
                                
                            } else if (!string.IsNullOrEmpty(secondCellValue) && string.IsNullOrEmpty(thirdCellValue) && string.IsNullOrEmpty(fourthCellValue)) {
                                depthCell.SetValue<int>(2).Style.Fill.BackgroundColor = XLColor.Orange;
                            } else {
                                depthCell.SetValue<int>(1).Style.Fill.BackgroundColor = XLColor.Red;
                            }
                        }
                        excelWorkBook.Save();
                    }
                }
            } catch (Exception exp) { 
                throw exp; 
            }
        }

        private List<MenuParent> getMenuParentData(string csvFilePath) {
            List<MenuParent> lstMenuPArent = new List<MenuParent>();
            try {
                if (File.Exists(csvFilePath)) {
                    using (StreamReader reader = new StreamReader(csvFilePath)) {
                        string csvContent = reader.ReadToEnd();

                        string[] lines = csvContent.Split('\n');
                        for (int i = 1; i < lines.Length; i++) {
                            string[] dataContent = lines[i].Split(',');
                            MenuParent menuParent = new MenuParent();
                            menuParent.MenuID = Convert.ToInt16(dataContent[0]);
                            dataContent[1] = dataContent[1].Replace("\r", "");
                            menuParent.ParentID = Convert.ToInt16(dataContent[1]);
                            lstMenuPArent.Add(menuParent);
                        }
                    }
                }
            } catch (Exception exp) {
                throw exp;
            }
            return lstMenuPArent;
        }

        protected void btnFindParent_Click(object sender, EventArgs e) {
            string csvFilePath = Server.MapPath("~/ExcelFiles/MenuParent.csv"); 
            var menuList = getMenuParentData(csvFilePath);

            var parents = getParents(13, ref menuList);

            LinkedList<MenuParent> menuParentList = getParentsLinkedList(23,ref menuList);

            for (int i = menuParentList.Count-1; i > 0; i--) {
                Response.Write("["+ menuParentList.ElementAt(i).MenuID + "|" + menuParentList.ElementAt(i).ParentID +"] -> ");
            }
            
        }
        public LinkedList<MenuParent> getParentsLinkedList(int menuID, ref List<MenuParent> menuParents) {
            LinkedList<MenuParent> menuParentList = new LinkedList<MenuParent>();
            var parentId = menuParents.Where(c => c.MenuID == menuID).FirstOrDefault().ParentID;
            var parentMenu = menuParents.Where(c => c.MenuID == parentId).First();
            while (parentId != 1) {
                 var parent = menuParents.Where(c => c.MenuID == parentId).First();
                menuParentList.AddFirst(parent);
                parentId = parent.ParentID;
            }
            return menuParentList;
        }


        public List<int> getParents(int menuID, ref List<MenuParent> menuParents) {
            List<int> parents = new List<int>();
            var menu = menuParents.Where(c => c.MenuID == menuID).FirstOrDefault();
            var parent = menu.ParentID;
            parents.Add(parent);
            while (parent != 1) {
                parent = menuParents.Where(c => c.MenuID == parent).Select(c => c.ParentID).First();
                parents.Add(parent);
            }
            return parents;
        }


        public class MenuParent {
            public int MenuID { get; set; }
            public int ParentID { get; set; }

            
        }

        protected void btnProxessXML_Click(object sender, EventArgs e) {
            string xmlFilePath = Server.MapPath("~/ExcelFiles/MenuDetails.xml");
            DataTable menuResource = new DataTable();
            menuResource.Columns.Add("menuid", typeof(string));
            menuResource.Columns.Add("menutext", typeof(string));
            menuResource.Columns.Add("resourcename", typeof(string));
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFilePath);

            XmlElement node = doc.SelectNodes("/Root/Menu[@menuid='1']")[0] as XmlElement;

            string menuId = node.GetAttribute("menuid");
            string menutext = node.GetAttribute("menutext");
            var resources = node.GetElementsByTagName("Resources")[0];
            List<string> resourceNames = new List<string>();
            if (resources != null) {
                foreach (XmlElement res in resources.ChildNodes) {
                    string resourceName = res.GetAttribute("resourcename");
                    resourceNames.Add(resourceName);
                }
            }

            /* add self attributes */

            DataRow newRow = menuResource.NewRow();
            newRow["menuid"] = menuId;
            newRow["menutext"] = menutext;
            newRow["resourcename"] = string.Join("|", resourceNames.ToArray());
            menuResource.Rows.Add(newRow);

            GetMenuResource(node, ref menuResource);

            DataTable dt = menuResource;
        }

        private void GetMenuResource(XmlElement node, ref DataTable dataTable) {
            foreach (XmlElement child in node.ChildNodes) {
                string menuId = child.GetAttribute("menuid");
                string menutext = child.GetAttribute("menutext");
                var resources = child.GetElementsByTagName("Resources")[0];
                List<string> resourceNames = new List<string>();
                if (resources != null) {
                    foreach (XmlElement res in resources.ChildNodes) {
                        string resourceName = res.GetAttribute("resourcename");
                        resourceNames.Add(resourceName);
                    }
                }

                if (!string.IsNullOrEmpty(menuId)) {
                    DataRow newRow = dataTable.NewRow();
                    newRow["menuid"] = menuId;
                    newRow["menutext"] = menutext;
                    newRow["resourcename"] = string.Join("|", resourceNames.ToArray());
                    dataTable.Rows.Add(newRow);

                    var menuList = child.GetElementsByTagName("Menu");
                    if (menuList.Count > 0) {
                        GetMenuResource(child, ref dataTable);
                    }
                }

            }
        }
    }
}