using DesignerEyesService.Entities;
using DesignerEyesService.Interfaces;
using FluentFTP;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace DesignerEyesService.Services
{
    public class ReadOrder : IReadOrder
    {
        private readonly ILogger<ReadOrder> _logger;

        public ReadOrder(ILogger<ReadOrder> logger)
        {
            _logger = logger;
        }
        public void ReadOrdersData()
        {

            var inventorylists = new List<InventoryEntry>();

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            _logger.LogDebug("Creating Order Object");
            _logger.LogDebug("Connecting with SQL SP DesignerEyes");
            var Connection = jsonParamModel.connectionString;
            List<Order> list = new List<Order>();
            using (SqlConnection connection = new SqlConnection(Connection))
            {
                SqlDataAdapter da = new SqlDataAdapter();
                DataSet ds = new DataSet();
                SqlCommand command = new SqlCommand("dbo.Invicta_SP_Qry_DesignerEyes", connection);
                command.CommandType = CommandType.StoredProcedure;
                da = new SqlDataAdapter(command);
                da.Fill(ds);
                var dat = ds.Tables[0];
                list = (from DataRow dr in dat.Rows
                        select new Order()
                        {
                            ContactEmail = getString(dr, "ContactEmail"),
                            EntryID = getString(dr, "EntryID"),
                            FirstName = getString(dr, "FirstName"),
                            LastName = getString(dr, "LastName"),
                            ID = getString(dr, "ID"),
                            isPendingForwarding = getString(dr, "isPendingForwarding"),
                            City = getString(dr, "City"),
                            ItemLookupCode = getString(dr, "ItemLookupCode"),
                            OrderNumber = getString(dr, "OrderNumber"),
                            Country = getString(dr, "Country"),
                            PostCode = getString(dr, "PostCode"),
                            QtyCancelled = getString(dr, "QtyCancelled"),
                            QtyOrdered = getString(dr, "QtyOrdered"),
                            QtyRefunded = getString(dr, "QtyRefunded"),
                            QtyShipped = getString(dr, "QtyShipped"),
                            RealQty = getString(dr, "RealQty"),
                            Region = getString(dr, "Region"),
                            RegionID = getString(dr, "RegionID"),
                            SimpleProdLineNo = getString(dr, "SimpleProdLineNo"),
                            Status = getString(dr, "Status"),
                            Street = getString(dr, "Street"),
                            Street2 = getString(dr, "Street2"),
                            Telephone = getString(dr, "Telephone"),
                            wasForwarded = getString(dr, "wasForwarded")
                        }).ToList();
                _logger.LogDebug("Connection with SQL DB succesful");
            }
            var date = DateTime.Now;

            foreach (Order o in list)
            {
                using (SqlConnection connection = new SqlConnection(Connection))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = $"UPDATE [Merlin].[dbo].[eCommerceOrderEntry] SET wasForwarded= 1 where OrderNumber='{o.OrderNumber}' and ItemLookupCode LIKE '%DE-{o.ItemLookupCode}%'";
                        cmd.Connection = connection;
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        _logger.LogDebug($"eCommerceOrderEntry order:{o.OrderNumber} item:{o.ItemLookupCode} already update ");
                    }
                }

                using (SqlConnection connection = new SqlConnection(Connection))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = $"UPDATE Merlin.dbo.eCommerceOrder SET isPendingForwarding= 0 where OrderNumber= '{o.OrderNumber}' and not exists (select '1' from Merlin.dbo.eCommerceOrderEntry where eCommerceOrderId = Merlin.dbo.eCommerceOrder.ID and wasForwarded = 0)";
                        cmd.Connection = connection;
                        connection.Open();
                        cmd.ExecuteNonQuery();
                        connection.Close();
                        _logger.LogDebug($"eCommerceOrder order:{o.OrderNumber}  already update ");
                    }
                }
                
            }

            _logger.LogDebug("ExcelFile created");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("WorkSheet1");
                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];
                var totalData = list.Count();
                var count = 2;

                excelWorksheet.Cells[1, jsonParamModel.ordIDColumn].LoadFromText("ID").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordOrderNumberColumn].LoadFromText("OrderNumber").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordEntryIDColumn].LoadFromText("EntryID").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordStatusColumn].LoadFromText("Status").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordItemLookupCodeColumn].LoadFromText("ItemLookupCode").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordSimpleProdLineNoColumn].LoadFromText("SimpleProdLineNo").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordisPendingForwardingColumn].LoadFromText("isPendingForwarding").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordwasForwardedColumn].LoadFromText("wasForwarded").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordRealQtyColumn].LoadFromText("RealQty").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordQtyOrderedColumn].LoadFromText("QtyOrdered").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordQtyCancelledColumn].LoadFromText("QtyCancelled").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordQtyRefundedColumn].LoadFromText("QtyRefunded").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordQtyShippedColumn].LoadFromText("QtyShipped").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordCityColumn].LoadFromText("City").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordCountryColumn].LoadFromText("Country").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordRegionColumn].LoadFromText("Region").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordPostCodeColumn].LoadFromText("PostCode").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordStreetColumn].LoadFromText("Street").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordStreet2Column].LoadFromText("Street2").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordTelephoneColumn].LoadFromText("Telephone").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordContactEmailColumn].LoadFromText("ContactEmail").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordFirstNameColumn].LoadFromText("FirstName").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordLastNameColumn].LoadFromText("LastName").Style.Font.Bold = true;
                excelWorksheet.Cells[1, jsonParamModel.ordRegionIDColumn].LoadFromText("RegionID").Style.Font.Bold = true;


                foreach (Order o in list)
                {

                    excelWorksheet.Cells[count, jsonParamModel.ordIDColumn].LoadFromText(o.ID);
                    excelWorksheet.Cells[count, jsonParamModel.ordOrderNumberColumn].LoadFromText(o.OrderNumber);
                    excelWorksheet.Cells[count, jsonParamModel.ordEntryIDColumn].LoadFromText(o.EntryID);
                    excelWorksheet.Cells[count, jsonParamModel.ordStatusColumn].LoadFromText(o.Status);
                    excelWorksheet.Cells[count, jsonParamModel.ordItemLookupCodeColumn].LoadFromText(o.ItemLookupCode);
                    excelWorksheet.Cells[count, jsonParamModel.ordSimpleProdLineNoColumn].LoadFromText(o.SimpleProdLineNo);
                    excelWorksheet.Cells[count, jsonParamModel.ordisPendingForwardingColumn].LoadFromText(o.isPendingForwarding);
                    excelWorksheet.Cells[count, jsonParamModel.ordwasForwardedColumn].LoadFromText(o.wasForwarded);
                    excelWorksheet.Cells[count, jsonParamModel.ordRealQtyColumn].LoadFromText(o.RealQty);
                    excelWorksheet.Cells[count, jsonParamModel.ordQtyOrderedColumn].LoadFromText(o.QtyOrdered);
                    excelWorksheet.Cells[count, jsonParamModel.ordQtyCancelledColumn].LoadFromText(o.QtyCancelled);
                    excelWorksheet.Cells[count, jsonParamModel.ordQtyRefundedColumn].LoadFromText(o.QtyRefunded);
                    excelWorksheet.Cells[count, jsonParamModel.ordQtyShippedColumn].LoadFromText(o.QtyShipped);
                    excelWorksheet.Cells[count, jsonParamModel.ordCityColumn].LoadFromText(o.City);
                    excelWorksheet.Cells[count, jsonParamModel.ordCountryColumn].LoadFromText(o.Country);
                    excelWorksheet.Cells[count, jsonParamModel.ordRegionColumn].LoadFromText(o.Region);
                    excelWorksheet.Cells[count, jsonParamModel.ordPostCodeColumn].LoadFromText(o.PostCode);
                    excelWorksheet.Cells[count, jsonParamModel.ordStreetColumn].LoadFromText(o.Street);
                    excelWorksheet.Cells[count, jsonParamModel.ordStreet2Column].LoadFromText(o.Street2);
                    excelWorksheet.Cells[count, jsonParamModel.ordTelephoneColumn].LoadFromText(o.Telephone);
                    excelWorksheet.Cells[count, jsonParamModel.ordContactEmailColumn].LoadFromText(o.ContactEmail);
                    excelWorksheet.Cells[count, jsonParamModel.ordFirstNameColumn].LoadFromText(o.FirstName);
                    excelWorksheet.Cells[count, jsonParamModel.ordLastNameColumn].LoadFromText(o.LastName);
                    excelWorksheet.Cells[count, jsonParamModel.ordRegionIDColumn].LoadFromText(o.RegionID);
                    count++;
                }
                
                string strCombFile = $"TCO-{date.ToString("yyyy-MM-dd")}.xlsx";
                FileInfo excelFile = new FileInfo($"./Orders/{strCombFile}");
                excel.SaveAs(excelFile);

                using (FtpClient ftp = new FtpClient(jsonParamModel.ftpHost, new System.Net.NetworkCredential { UserName = jsonParamModel.ftpUser, Password = jsonParamModel.ftpPassword }))
                {
                    ftp.Connect();
                    ftp.UploadFile($"./Orders/{strCombFile}", $"./In Orders/{strCombFile}", FtpRemoteExists.Overwrite, true, FtpVerify.Retry);
                    _logger.LogDebug("Orders shipped!");
                }
            }

            Console.WriteLine("Orders shipped!");

        }
        
        #region Gets

        private static long getLong(DataRow data, string param)
        {
            if (data[$"{param}"] == DBNull.Value)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(data[$"{param}"]);
            }
        }

        private static int getInt(DataRow data, string param)
        {
            if (data[$"{param}"] == DBNull.Value)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(data[$"{param}"]);
            }
        }

        private static string getString(DataRow data, string param)
        {
            if (data[$"{param}"] == DBNull.Value)
            {
                return "";
            }
            else
            {
                return Convert.ToString(data[$"{param}"]);
            }
        }
        #endregion
    }
}
