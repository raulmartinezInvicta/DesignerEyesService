using DesignerEyesService.Entities;
using DesignerEyesService.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace DesignerEyesService.Services
{

    public class OrderEntry {
        public string itemNumber {get;set;}

        public int orderedQuantity{get;set;}

        public int lineNumber{get;set;}
    }
    public class ShipConfirm {
        public string orderNumber {get;set;}
        public string customerNumber{get;set;}
        public DateTime orderDate{get;set;}
        public List<Detail> details {get;set;}
    }
    public class Detail {
        public int lineNumber{get;set;}
        public string itemNumber{get;set;}
        public int orderedQuantity{get;set;}
        public int shippedQuantity{get;set;}
        public int canceledQuantity{get;set;}
        public DateTime shippedDate{get;set;}
        public string carrier{get;set;}
        public string trackingNumber{get;set;}
        public bool prePaidReturnLabelUsed{get;set;}
        public Decimal prePaidReturnLabelCost{get;set;}

        public static implicit operator List<object>(Detail v)
        {
            throw new NotImplementedException();
        }
    }
    public class ReadShipConfirm : IReadShipConfirm
    {
        private readonly ILogger<ReadShipConfirm> _logger;

        public ReadShipConfirm(ILogger<ReadShipConfirm> logger)
        {
            _logger = logger;
        }
        private OrderEntry GetOrderQuantity(string sku, string orderNumber,int supplierId) {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            Console.WriteLine("jsonParamModel.connectionString:"+ jsonParamModel.connectionString);
            _logger.LogDebug("jsonParamModel.connectionString:"+ jsonParamModel.connectionString);

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            
            builder.ConnectionString=jsonParamModel.connectionString;

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                connection.Open();
                string SQLstr = String.Format("SELECT a.ItemLookupCode,a.QtyOrdered,a.SimpleProdLineNo FROM Merlin.dbo.eCommerceOrderEntry a,Merlin.dbo.ShippersSupplier b WHERE a.OrderNumber = '{0}' and a.ItemLookUpCode = '{1}' and a.SimpleProdLineNo>0  and b.id={2} ", orderNumber, sku, supplierId);
                Console.WriteLine(SQLstr);
                _logger.LogDebug(SQLstr);
                using (SqlCommand cmd = new SqlCommand(SQLstr, connection)) {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            return new OrderEntry(){
                                itemNumber=reader.GetString(0),
                                orderedQuantity=reader.GetInt32(1),
                                lineNumber=reader.GetInt32(2)
                            };

                        }
                    }
                }

            }
            return new OrderEntry();

        }

        public async Task ReadShipConfirmAsync(int supplierId, string path)
        {

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            _logger.LogDebug(jsonParamModel.ToString());
            var excelFile = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var cultureInfo = new CultureInfo("en-US");
            using (var epPackage = new ExcelPackage(excelFile))
            {
                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                _logger.LogDebug("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow; 
                Console.WriteLine("Param:"+initialRow);
                _logger.LogDebug("Param:"+initialRow);
                for (int ws = 0; ws < wscount; ws++)
                {
                    var worksheet = epPackage.Workbook.Worksheets[ws];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    for (int row = initialRow; row <= rowCount; row = row + 1)
                    {
                        try
                        {
                            ShipConfirm sc = new ShipConfirm();
                            Detail det = new Detail();                            
                            sc.orderNumber = worksheet.Cells[row, jsonParamModel.scOrderNumberColumn].Value.ToString();
                            sc.orderDate = DateTime.Now;

                            det.itemNumber = worksheet.Cells[row,jsonParamModel.scItemNumberColumn].Value.ToString();

                            OrderEntry entry =  GetOrderQuantity(det.itemNumber,sc.orderNumber, supplierId);
              
                            det.shippedQuantity= Convert.ToInt32(worksheet.Cells[row,jsonParamModel.scShippedQuantityColumn].Value.ToString());
                            
                            det.orderedQuantity= det.shippedQuantity;
                            det.lineNumber = entry.lineNumber;

                            if (jsonParamModel.scShippedDateColumn>0) {
                                var shippedDateString = worksheet.Cells[row,jsonParamModel.scShippedDateColumn].Value.ToString();
                                det.shippedDate = DateTime.Parse(shippedDateString);
                            } else {
                                det.shippedDate = DateTime.Now;
                            } 
                            
                            if (jsonParamModel.scCarrierColumn>0) {
                                 det.carrier = worksheet.Cells[row,jsonParamModel.scCarrierColumn].Value.ToString();
                            } else {
                                det.carrier = jsonParamModel.defaultcarrier;
                            }

                            det.trackingNumber=worksheet.Cells[row,jsonParamModel.scTrackingNumberColumn].Value.ToString();

                            List<Detail> details = new List<Detail>();
                            details.Add(det);
                            sc.details=details;       

                            var modelJson = JsonSerializer.Serialize(sc, options);
                            Console.WriteLine("modelJson: "+modelJson);
                            _logger.LogDebug("modelJson: "+modelJson);  
                            Uri u = new Uri(jsonParamModel.tcouri+"/upload/shipping?SupplierID="+supplierId);  
                            HttpClient httpClient = new HttpClient();
                            HttpContent c = new StringContent(modelJson, System.Text.Encoding.UTF8, "application/json");
                            var result = await httpClient.PostAsync(u, c);      
                            if (result.IsSuccessStatusCode) {
                                Console.WriteLine("Shipping Update accepted");
                                _logger.LogDebug("Shipping Update accepted:" + modelJson );
                            } else {
                                Console.WriteLine ("Shipping Updated failed");
                                _logger.LogDebug("Shipping Update failed:" + modelJson );
                            }                                        
                        }   
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            _logger.LogError(ex.ToString());
                            row = rowCount;
                            continue;
                        }
                    }
                }
            }
        }
    }
}
