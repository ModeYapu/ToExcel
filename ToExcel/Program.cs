using System;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var list = new List<searchLogInfo>();
            DownloadFileInfo downloadFile = ConvertLogToExcel(list);
            //file = File(downloadFile.FilePath, "application/x-excel", downloadFile.FileName);
        }
        /// <summary>
        /// DataTable转换成Excel文档流(导出数据量超出65535条,分sheet)
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static MemoryStream ExportDataTableToExcel(DataTable sourceTable)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            int dtRowsCount = sourceTable.Rows.Count;
            int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
            int SheetNum = 1;
            int rowIndex = 1;
            int tempIndex = 1; //标示 
            ISheet sheet = workbook.CreateSheet("sheet" + SheetNum);
            for (int i = 0; i < dtRowsCount; i++)
            {
                if (i == 0 || tempIndex == 1)
                {
                    IRow headerRow = sheet.CreateRow(0);
                    foreach (DataColumn column in sourceTable.Columns)
                        headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                }
                HSSFRow dataRow = (HSSFRow)sheet.CreateRow(tempIndex);
                foreach (DataColumn column in sourceTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(sourceTable.Rows[i][column].ToString());
                }
                if (tempIndex == 65535)
                {
                    SheetNum++;
                    sheet = workbook.CreateSheet("sheet" + SheetNum);//
                    tempIndex = 0;
                }
                rowIndex++;
                tempIndex++;
                //AutoSizeColumns(sheet);
            }
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            sheet = null;
            // headerRow = null;
            workbook = null;
            return ms;
        }

        private static DownloadFileInfo ConvertLogToExcel(List<searchLogInfo> list)
        {
            //var lis = list.ConvertAll(l => l as LogEntity).ToList();
            //typeof(T).GetCustomAttributes()
            //删除以前的日志文件
            string[] files = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory);
            string pattern = "\\d{4}\\d{2}\\d{2}.+\\.xls$";
            Regex r = new Regex(pattern);

            foreach (string file in files)
            {
                if (r.IsMatch(file))
                {
                    FileInfo fi = new FileInfo(file);
                    fi.Delete();
                }
            }


            DownloadFileInfo info = new DownloadFileInfo();
            try
            {
                //操作时间 用户名 操作类型 描述 
                IWorkbook workbook = new HSSFWorkbook();


                ISheet sheet = workbook.CreateSheet("前台日志");
                IRow row0 = sheet.CreateRow(0);
                int dtRowsCount = list.Count;
                int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
                int SheetNum = 1;
                int rowIndex = 1;
                int tempIndex = 1; //标示 

                row0.CreateCell(0).SetCellValue("应用模式名称");
                row0.CreateCell(1).SetCellValue("操作业务名称");
                row0.CreateCell(2).SetCellValue("日志信息");
                row0.CreateCell(3).SetCellValue("日志时间");
                row0.CreateCell(4).SetCellValue("日志报错");
                row0.CreateCell(5).SetCellValue("日志状态");

                HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
                HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
                style.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss");


                for (int i = 0; i < list.Count; i++)
                {
                    IRow row = sheet.CreateRow(i + 1);
                    //row.CreateCell(0).SetCellValue(list[i].OperateTime);
                    ICell cell = row.CreateCell(0);
                    cell.CellStyle = style;
                    //cell.SetCellValue(list[i].OperateTime);
                    cell.Sheet.SetColumnWidth(0, 18 * 256);
                    row.CreateCell(0).SetCellValue(list[i].AppName);
                    row.CreateCell(1).SetCellValue(list[i].Name);
                    row.CreateCell(2).SetCellValue(list[i].ObjectInfo);
                    row.CreateCell(3).SetCellValue(list[i].CreatedDateTime);
                    row.CreateCell(4).SetCellValue(list[i].ErrorMessage);
                    row.CreateCell(5).SetCellValue(list[i].Result);
                }
                string fileName = "前台日志" + "(" + DateTime.Now.ToString("yyyyMMdd hhmmss") + ")" + ".xls";
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                info.FilePath = path;
                info.FileName = fileName;
                using (FileStream fs = System.IO.File.OpenWrite(path))
                {
                    info.FileSize = fs.Length;
                    workbook.Write(fs);
                }
            }
            catch (Exception)
            {

                throw;
            }
            return info;

        }

        private DownloadFileInfo LogToExcel(List<searchLogInfo> list)
        {
            //var lis = list.ConvertAll(l => l as LogEntity).ToList();
            //typeof(T).GetCustomAttributes()
            //删除以前的日志文件
            string[] files = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory);
            string pattern = "\\d{4}\\d{2}\\d{2}.+\\.xls$";
            Regex r = new Regex(pattern);

            foreach (string file in files)
            {
                if (r.IsMatch(file))
                {
                    FileInfo fi = new FileInfo(file);
                    fi.Delete();
                }
            }
            DownloadFileInfo info = new DownloadFileInfo();
            try
            {
                //操作时间 用户名 操作类型 描述 
                IWorkbook workbook = new HSSFWorkbook();
                //ISheet sheet = workbook.CreateSheet("前台日志");
                //IRow row0 = sheet.CreateRow(0);
                int dtRowsCount = list.Count;
                int SheetCount = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dtRowsCount) / 65536));
                int SheetNum = 1;
                int rowIndex = 1;
                int tempIndex = 1; //标示 
                ISheet sheet = workbook.CreateSheet("sheet" + SheetNum);
                for (int i = 0; i < dtRowsCount; i++)
                {
                    if (i == 0 || tempIndex == 1)
                    {
                        IRow row0 = sheet.CreateRow(0);
                        row0.CreateCell(0).SetCellValue("应用模式名称");
                        row0.CreateCell(1).SetCellValue("操作业务名称");
                        row0.CreateCell(2).SetCellValue("日志信息");
                        row0.CreateCell(3).SetCellValue("日志时间");
                        row0.CreateCell(4).SetCellValue("日志报错");
                        row0.CreateCell(5).SetCellValue("日志状态");
                    }
                    HSSFRow row = (HSSFRow)sheet.CreateRow(tempIndex);
                    //HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
                    //HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
                    //style.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss");
                    var j = tempIndex + (SheetNum - 1) * tempIndex;
                    row.CreateCell(0).SetCellValue(list[i].AppName);
                    row.CreateCell(1).SetCellValue(list[i].Name);
                    row.CreateCell(2).SetCellValue(list[i].ObjectInfo);
                    row.CreateCell(3).SetCellValue(list[i].CreatedDateTime);
                    row.CreateCell(4).SetCellValue(list[i].ErrorMessage);
                    row.CreateCell(5).SetCellValue(list[i].Result);
                    if (tempIndex == 65535)
                    {
                        SheetNum++;
                        sheet = workbook.CreateSheet("sheet" + SheetNum);
                        tempIndex = 0;
                    }
                    rowIndex++;
                    tempIndex++;
                }

                string fileName = "前台日志" + "(" + DateTime.Now.ToString("yyyyMMdd hhmmss") + ")" + ".xls";
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                info.FilePath = path;
                info.FileName = fileName;
                using (FileStream fs = System.IO.File.OpenWrite(path))
                {
                    info.FileSize = fs.Length;
                    workbook.Write(fs);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return info;
        }
    }
}
