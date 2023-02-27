using ClosedXML.Excel;
using System.Data;
using System.Reflection;
using USA_ExportWorkFlow_Web.Models;
using USA_ExportWorkFlow_Web.ViewModel;

namespace USA_ExportWorkFlow_Web
{
    public static class GeneralFunction
    {
        public static DataTable GetDataFromExcel(string path)
        {
            DataTable dt = new DataTable();
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(path))
                {

                    IXLWorksheet workSheet = workBook.Worksheet(1);
                    bool firstRow = true;
                    int skiprows = 1;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        skiprows = skiprows - 1;
                        if (skiprows <= 0)
                        {
                            //Use the first row to add columns to DataTable.
                            if (firstRow)
                            {
                                int j = 0;
                                foreach (IXLCell cell in row.Cells())
                                {
                                    if (!string.IsNullOrEmpty(cell.Value.ToString()))
                                    {
                                        dt.Columns.Add(cell.Value.ToString());
                                    }
                                    else
                                    {    //string A = "A" + j;
                                         //dt.Columns.Add(A.ToString());
                                    }
                                    j++;
                                }
                                firstRow = false;
                            }
                            else
                            {
                                if (!row.IsEmpty())
                                {
                                    int i = 0;
                                    DataRow toInsert = dt.NewRow();
                                    foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
                                    {
                                        try
                                        {
                                            if (cell.Value.ToString() == "")
                                            {
                                                toInsert[i] = "";
                                            }
                                            else
                                            {
                                                toInsert[i] = cell.Value.ToString();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                        i++;
                                    }
                                    dt.Rows.Add(toInsert);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return dt;
        }

        public static List<FileDetailViewModel> GetListFromExcel(Stream path)
        {
            //DataTable dt = new DataTable();
            List<FileDetailViewModel> fileDetailViewModels = new List<FileDetailViewModel>();
            
            try
            {
                using (XLWorkbook workBook = new XLWorkbook(path))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(1);
                    var lastRow = workSheet.LastRowUsed().RowNumber();
                    int skiprows = 6;

                    for (int row = 6; row<=lastRow; row++)
                    {
                        fileDetailViewModels.Add(
                            new FileDetailViewModel
                            {
                                FileNumber = workSheet.Cell(row, 1).Value.ToString(),
                                BookingNumber = workSheet.Cell(row, 3).Value.ToString(),
                                Location = workSheet.Cell(row, 4).Value.ToString(),
                                HBLNumber = workSheet.Cell(row, 5).Value.ToString(),
                                ContainerType = workSheet.Cell(row, 6).Value.ToString(),
                                Destination = workSheet.Cell(row, 7).Value.ToString(),
                                FileDestination = workSheet.Cell(row, 8).Value.ToString(),
                                CustomerName = workSheet.Cell(row, 10).Value.ToString(),
                                ETD = workSheet.Cell(row, 11).Value.ToString() == "" ? null : Convert.ToDateTime(workSheet.Cell(row, 11).Value),
                                ContainerNumber = workSheet.Cell(row, 12).Value.ToString(),
                                Weight = workSheet.Cell(row, 13).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row, 13).Value.ToString()),
                                Volume = workSheet.Cell(row, 14).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row, 14).Value.ToString()),
                                WHSVolume = workSheet.Cell(row, 15).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row, 15).Value.ToString()),
                                DocReceived = workSheet.Cell(row, 16).Value.ToString() == "0" ? false : true,
                                Posted = workSheet.Cell(row, 17).Value.ToString() == "0" ? false : true,
                                Disposition = workSheet.Cell(row, 18).Value.ToString(),
                                Status = workSheet.Cell(row, 19).Value.ToString(),
                                AES = workSheet.Cell(row, 20).Value.ToString(),
                                ITN = workSheet.Cell(row, 21).Value.ToString()
                            }
                            );
                        //fileDetailViewModels.Add(fileDetailViewModel);

                        //List<IXLCell> cells = row.Cells().ToList();
                        //if (cells.Count() != props.Length -1)
                        //{
                        //	break;
                        //}
                        //else

                        //FileDetailViewModel fileDetailViewModel = new FileDetailViewModel();
                        //fileDetailViewModel.FileNumber = workSheet.Cell(row.RowNumber(), 1).Value.ToString();
                        //fileDetailViewModel.BookingNumber = workSheet.Cell(row.RowNumber(), 3).Value.ToString();
                        //fileDetailViewModel.Location = workSheet.Cell(row.RowNumber(), 4).Value.ToString();
                        //fileDetailViewModel.HBLNumber = workSheet.Cell(row.RowNumber(), 5).Value.ToString();
                        //fileDetailViewModel.ContainerType = workSheet.Cell(row.RowNumber(), 6).Value.ToString();
                        //fileDetailViewModel.Destination = workSheet.Cell(row.RowNumber(), 7).Value.ToString();
                        //fileDetailViewModel.FileDestination = workSheet.Cell(row.RowNumber(), 8).Value.ToString();
                        //fileDetailViewModel.CustomerName = workSheet.Cell(row.RowNumber(), 10).Value.ToString();
                        //fileDetailViewModel.ETD = workSheet.Cell(row.RowNumber(), 11).Value.ToString() == "" ? null : Convert.ToDateTime(workSheet.Cell(row.RowNumber(), 11).Value);
                        //fileDetailViewModel.ContainerNumber = workSheet.Cell(row.RowNumber(), 12).Value.ToString();
                        //fileDetailViewModel.Weight = workSheet.Cell(row.RowNumber(), 13).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row.RowNumber(), 13).Value.ToString());
                        //fileDetailViewModel.Volume = workSheet.Cell(row.RowNumber(), 14).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row.RowNumber(), 14).Value.ToString());
                        //fileDetailViewModel.WHSVolume = workSheet.Cell(row.RowNumber(), 15).Value.ToString() == "" ? 0 : Convert.ToDouble(workSheet.Cell(row.RowNumber(), 15).Value.ToString());
                        //fileDetailViewModel.DocReceived = workSheet.Cell(row.RowNumber(), 16).Value.ToString() == "0" ? false : true;
                        //fileDetailViewModel.Posted = workSheet.Cell(row.RowNumber(), 17).Value.ToString() == "0" ? false : true;
                        //fileDetailViewModel.Disposition = workSheet.Cell(row.RowNumber(), 18).Value.ToString();
                        //fileDetailViewModel.Status = workSheet.Cell(row.RowNumber(), 19).Value.ToString();
                        //fileDetailViewModel.AES = workSheet.Cell(row.RowNumber(), 20).Value.ToString();
                        //fileDetailViewModel.ITN = workSheet.Cell(row.RowNumber(), 21).Value.ToString();
                        //fileDetailViewModels.Add(fileDetailViewModel);

                    }
                }
            }
            catch (Exception ex) { }

            return fileDetailViewModels;
        }
    }
}
