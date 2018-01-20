using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using Tools;

namespace FillExcel
{
    public class FillExcel
    {
        // default value if configure file does not exist
        int groupPicCount = 6;    // 6 pictures in each group
        int groupColumnCount = 2;    // 2 column in each group
        int picBoxRowCount = 14;    // 14 rows in a picture box
        int pageHeight = 50;    //50 rows a page
        int pageHeaderHeight = 5;    // header is 5 rows height
        int pageVerticalOffset = 2;    // move down for 2 rows
        int picBoxColumnCount = 4;    // 4 columns in a picture box
        int titleBoxRowCount = 1;    // title box is 1 row height
        int endBoxColumnCount = 8;    // end box is 8 columns width
        int endBoxRowCount = 3;    // end box is 3 rows height
        int INSPECTION_TIME_ROW_INDEX = 2;
        int INSPECTION_TIME_COLUMN_INDEX = 5;
        int ITEM_NO_ROW_INDEX = 3;
        int ITEM_NO_COLUMN_INDEX = 5;
        int FORTH_LINE_ROW_INDEX = 4;
        int FORTH_LINE_COLUMN_INDEX = 1;
        int DOC_TITLE_ROW_INDEX = 1;
        int DOC_TITLE_COLUMN_INDEX = 1;
        int ALERT_COLOR_CELL_ROW_INDEX = 4;
        int ALERT_COLOR_CELL_COLUMN_INDEX = 1;
        dynamic STANDARD_ROW_HEIGHT = 13.5;
        double PICTURE_MARGIN = 15;

        int currentRow = 6;

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        Application app;
        Workbooks wbks;
        Workbook wbk;
        Sheets sheets;
        _Worksheet sheet1;
        string filename;
        int picCount = 0;
        dynamic alertColor = 0;

        public delegate void ProgressUpdatedHandler(int value, string message);
        public event ProgressUpdatedHandler ProgressUpdated;
        public delegate void OnExitEventHandler(string message);
        public event OnExitEventHandler Exit;

        public FillExcel()
        {
            app = new Application();
            app.DisplayAlerts = false;
            //app.Visible = true;
            wbks = app.Workbooks;
        }

        public void LoadConfig(string path)
        {
            try
            {
                /*
                string[] str = File.ReadAllLines(path);
                string[] str2 = new string[str.Count() * 2];
                for (int i = 0; i < str.Count(); i++)
                {
                    str2[i * 2] = str[i].Split('=')[0];
                    str2[i * 2 + 1] = str[i].Split('=')[1];
                }

                groupPicCount = Convert.ToInt16(str2[1]);
                groupColumnCount = Convert.ToInt16(str2[3]);
                picBoxRowCount = Convert.ToInt16(str2[5]);
                pageHeight = Convert.ToInt16(str2[7]);
                pageHeaderHeight = Convert.ToInt16(str2[9]);
                pageVerticalOffset = Convert.ToInt16(str2[11]);
                picBoxColumnCount = Convert.ToInt16(str2[13]);
                titleBoxRowCount = Convert.ToInt16(str2[15]);
                endBoxColumnCount = Convert.ToInt16(str2[17]);
                endBoxRowCount = Convert.ToInt16(str2[19]);
                INSPECTION_TIME_ROW_INDEX = Convert.ToInt16(str2[21]);
                INSPECTION_TIME_COLUMN_INDEX = Convert.ToInt16(str2[23]);
                ITEM_NO_ROW_INDEX = Convert.ToInt16(str2[25]);
                ITEM_NO_COLUMN_INDEX = Convert.ToInt16(str2[27]);
                FORTH_LINE_ROW_INDEX = Convert.ToInt16(str2[29]);
                FORTH_LINE_COLUMN_INDEX = Convert.ToInt16(str2[31]);
                DOC_TITLE_ROW_INDEX = Convert.ToInt16(str2[33]);
                DOC_TITLE_COLUMN_INDEX = Convert.ToInt16(str2[35]);
                ALERT_COLOR_CELL_ROW_INDEX = Convert.ToInt16(str2[37]);
                ALERT_COLOR_CELL_COLUMN_INDEX = Convert.ToInt16(str2[39]);
                PICTURE_MARGIN = Convert.ToDouble(str2[41]);
                STANDARD_ROW_HEIGHT = Convert.ToDouble(str2[43]);
                */

                INI configFile = new INI(path);
                groupPicCount = Convert.ToInt16(configFile.GetValue("groupPicCount"));
                groupColumnCount = Convert.ToInt16(configFile.GetValue("groupColumnCount"));
                picBoxRowCount = Convert.ToInt16(configFile.GetValue("picBoxRowCount"));
                pageHeight = Convert.ToInt16(configFile.GetValue("pageHeight"));
                pageHeaderHeight = Convert.ToInt16(configFile.GetValue("pageHeaderHeight"));
                pageVerticalOffset = Convert.ToInt16(configFile.GetValue("pageVerticalOffset"));
                picBoxColumnCount = Convert.ToInt16(configFile.GetValue("picBoxColumnCount"));
                titleBoxRowCount = Convert.ToInt16(configFile.GetValue("titleBoxRowCount"));
                endBoxColumnCount = Convert.ToInt16(configFile.GetValue("endBoxColumnCount"));
                endBoxRowCount = Convert.ToInt16(configFile.GetValue("endBoxRowCount"));
                INSPECTION_TIME_ROW_INDEX = Convert.ToInt16(configFile.GetValue("INSPECTION_TIME_ROW_INDEX"));
                INSPECTION_TIME_COLUMN_INDEX = Convert.ToInt16(configFile.GetValue("INSPECTION_TIME_COLUMN_INDEX"));
                ITEM_NO_ROW_INDEX = Convert.ToInt16(configFile.GetValue("ITEM_NO_ROW_INDEX"));
                ITEM_NO_COLUMN_INDEX = Convert.ToInt16(configFile.GetValue("ITEM_NO_COLUMN_INDEX"));
                FORTH_LINE_ROW_INDEX = Convert.ToInt16(configFile.GetValue("FORTH_LINE_ROW_INDEX"));
                FORTH_LINE_COLUMN_INDEX = Convert.ToInt16(configFile.GetValue("FORTH_LINE_COLUMN_INDEX"));
                DOC_TITLE_ROW_INDEX = Convert.ToInt16(configFile.GetValue("DOC_TITLE_ROW_INDEX"));
                DOC_TITLE_COLUMN_INDEX = Convert.ToInt16(configFile.GetValue("DOC_TITLE_COLUMN_INDEX"));
                ALERT_COLOR_CELL_ROW_INDEX = Convert.ToInt16(configFile.GetValue("ALERT_COLOR_CELL_ROW_INDEX"));
                ALERT_COLOR_CELL_COLUMN_INDEX = Convert.ToInt16(configFile.GetValue("ALERT_COLOR_CELL_COLUMN_INDEX"));
                PICTURE_MARGIN = Convert.ToDouble(configFile.GetValue("PICTURE_MARGIN"));
                STANDARD_ROW_HEIGHT = Convert.ToDouble(configFile.GetValue("STANDARD_ROW_HEIGHT"));

            }
            catch (Exception err)
            {
                Exit(err.Message);
                return;
            }
            finally
            {
                ProgressUpdated(2, "配置加载完成");

            }
        }

        public void Fill(string xlsxName, string pdfName = null, string docTitle = null, string forthLine = null,
            string[] pictures = null, string[] titles = null, bool[] markAsRed = null,
            string endText = null, bool markEndTextRed = false, string itemNo = null, string inspectionTime = null)
        {
            string exitCode = "0";
            try
            {
                //wbk = wbks.Add(true);
                wbk = wbks.Add(System.Environment.CurrentDirectory + "\\Template.xlsx");
                filename = xlsxName;
                sheets = wbk.Sheets;
                sheet1 = (Worksheet)sheets.Item[1];
                sheet1.Name = "WorkSheet";

                SetStandardRowHeight(pictures.Count());
                //sheet1.Cells[2, 2] = "test";
                alertColor = GetAlertColor();

                ProgressUpdated(5, "文件创建完成");

                AddPictures(pictures, titles, markAsRed);

                ProgressUpdated(110, "图片填充完成");

                // add end box
                if (pictures.Count() == 6)
                    currentRow = pageHeaderHeight + 6 / groupColumnCount * (picBoxRowCount + titleBoxRowCount) + 1;
                else if (pictures.Count() % 6 == 0)
                    currentRow = (pageHeight * (pictures.Count() / 6 - 1) + (groupPicCount / groupColumnCount) * (picBoxRowCount + titleBoxRowCount)) + pageVerticalOffset + 1;
                AddEndBox(endText, 1, currentRow, markEndTextRed);

                ProgressUpdated(115, "结束语添加完成");

                // add INSPECTION TIME and ITEM NO
                AddInspectionTime(inspectionTime);
                AddItemNo(itemNo);
                AddDocTitle(docTitle);
                AddForthLine(forthLine);

                ProgressUpdated(120, "文件信息添加完成");

                // save and export pdf        
                SaveWorkbook();
                if (!string.IsNullOrEmpty(pdfName))
                {
                    ExportPDF(pdfName);
                }
            }
            finally
            {
                Exit(exitCode);
                QuitExcel();
            }

        }

        void SetStandardRowHeight(int pictureCount)
        {
            // Set row height to standard row height
            int pageCount = pictureCount / groupPicCount + 1;
            int totalRowCount = pageCount * pageHeight;
            for (int i = 1; i <= totalRowCount; i++)
            {
                sheet1.Cells[i, 9].RowHeight = STANDARD_ROW_HEIGHT;
            }
        }

        void AddPictures(string[] pictures, string[] titles, bool[] markAsRed)
        {
            picCount = pictures.Count();

            if(picCount == 5||picCount == 6)
            {
                picBoxRowCount -= 1;
            }

            int x = 0;
            int y = 0;
            for(int i=0;i<picCount;i++)
            {
                x = i % groupColumnCount * picBoxColumnCount + 1;
                AddPicture(pictures[i], titles[i], x, currentRow, markAsRed[i]);
                if(i%groupColumnCount!=0)
                {
                    currentRow += picBoxRowCount + titleBoxRowCount;
                }
                if ((i + 1) % groupPicCount == 0)
                {
                    currentRow = (i + 1) / groupPicCount * pageHeight + pageVerticalOffset + 1;
                    sheet1.HPageBreaks.Add(sheet1.Cells[currentRow - pageVerticalOffset, 9]);    // add page breake
                }
                ProgressUpdated((int)((double)i / (double)picCount * 100.0 + 5.0), string.Format("图片 {0}/{1}", i.ToString(), picCount.ToString()));
            }
            if(pictures.Count()%2==1)
            {
                AddPicture(null, null, x + picBoxColumnCount, currentRow, false);
                currentRow += picBoxRowCount + titleBoxRowCount;
            }            
        }

        void AddPicture(string picture, string title,int x, int y, bool markAsRed)
        {
            Excel.Range titleCell = sheet1.Range[sheet1.Cells[y, x], sheet1.Cells[y, x + picBoxColumnCount - 1]];
            FormatCell(ref titleCell);
            titleCell.WrapText = true;
            Excel.Range picCell = sheet1.Range[sheet1.Cells[y + 1, x], sheet1.Cells[y + picBoxRowCount, x + picBoxColumnCount - 1]];
            FormatCell(ref picCell);

            // write title
            if (!string.IsNullOrEmpty(title))
            {
                sheet1.Cells[y, x] = title;
                if(markAsRed)
                {
                    titleCell.Font.Color = alertColor;
                }
            }

            // insert picture
            if (!string.IsNullOrEmpty(picture))
            {
                Image img = new Bitmap(picture);
                double ratio = (double)img.Width / (double)img.Height;
                double picWidth = 0;
                double picHeight = 0;
                double picTop = picCell.Top + PICTURE_MARGIN;
                double picLeft = picCell.Left + PICTURE_MARGIN;
                //double boxWidth = picCell.Width - 20;
                //double boxHeight = picCell.Height - 20;
                double boxWidth = picCell.Width - PICTURE_MARGIN * 2;
                double boxHeight = picCell.Height - PICTURE_MARGIN * 2;
                if (img.Width >= img.Height)
                {
                    picWidth = boxWidth;
                    picHeight = picWidth / ratio;
                    picTop += (boxHeight - picHeight) / 2;
                }
                else
                {
                    picHeight = boxHeight;
                    picWidth = picHeight * ratio;
                    picLeft += (boxWidth - picWidth) / 2;
                }
                //double picTop = picCell.Top + 5;
                //double picLeft = picCell.Left + 5;
                //double picWidth = picCell.Width - 10;
                //double picHeight = picCell.Height - 10; 
                Log(string.Format(@"{0}: width:{1}, height:{2}", picture, picWidth.ToString(), picHeight.ToString()));
                //picWidth = 195;
                //picHeight = 180;
                sheet1.Shapes.AddPicture(picture, MsoTriState.msoFalse, MsoTriState.msoTrue, (float)picLeft, (float)picTop, (float)picWidth, (float)picHeight);
            }
        }

        void AddEndBox(string value, int x, int y, bool markEndTextRed)
        {
            // define end box
            Excel.Range endBox = sheet1.Range[sheet1.Cells[y, x], sheet1.Cells[y + endBoxRowCount - 1, x + endBoxColumnCount - 1]];
            FormatCell(ref endBox);
            endBox.WrapText = true;
            if(markEndTextRed)
            {
                endBox.Font.Color = alertColor;
            }
            // write title
            sheet1.Cells[y, x] = value;
        }

        void AddInspectionTime(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sheet1.Cells[INSPECTION_TIME_ROW_INDEX, INSPECTION_TIME_COLUMN_INDEX] = value;
            }
        }

        void AddItemNo(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sheet1.Cells[ITEM_NO_ROW_INDEX, ITEM_NO_COLUMN_INDEX] = value;
            }
        }

        void AddForthLine(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sheet1.Cells[FORTH_LINE_ROW_INDEX, FORTH_LINE_COLUMN_INDEX] = value;
            }
        }

        void AddDocTitle(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sheet1.Cells[DOC_TITLE_ROW_INDEX, DOC_TITLE_COLUMN_INDEX] = value;
            }
        }

        void ExportPDF(string file)
        {
            wbk.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, file);
        }

        void FormatCell(ref Excel.Range range)
        {
            // merge cells
            range.Merge();
            // add border
            range.BorderAround(XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());
            // center allign
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            // set data format as text
            range.NumberFormatLocal = "@";
        }

        void SaveWorkbook()
        {
            wbk.SaveAs(filename);            
        }

        dynamic GetAlertColor()
        {
            Range r = sheet1.Range[sheet1.Cells[ALERT_COLOR_CELL_ROW_INDEX, ALERT_COLOR_CELL_COLUMN_INDEX],sheet1.Cells[4,8]];
            return r.Font.Color;
        }

        void QuitExcel()
        {
            app.Quit();
            if (sheet1 != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
            }
            if (sheets != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
            }
            if (wbk != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wbk);
            }
            if (app != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }

            sheet1 = null;
            sheets = null;
            wbk = null;
            app = null;

            GC.Collect();

        }

        void Log(string message)
        {
            string time = DateTime.Now.ToString();
            message = string.Format("[{0}]:\t {1} \r\n", time, message);
            //File.AppendAllText("log.txt", message);
        }

    }
}
