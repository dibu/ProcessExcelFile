using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;
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
    }
}