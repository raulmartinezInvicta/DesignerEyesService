using DesignerEyesService.Entities;
using DesignerEyesService.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;

namespace DesignerEyesService.Services
{
    public class InventoryEntry {
        public string item_name {get;set;}
        public int quantity{get;set;}
    }
    public class ReadInventory : IReadInventory
    {
        private readonly ILogger<ReadInventory> _logger;

        public ReadInventory(ILogger<ReadInventory> logger)
        {
            _logger = logger;
        }
        public  async Task ReadInventoryAsync(int supplierId, string path)
        {

            var inventorylists = new List<InventoryEntry>();

            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);
            Console.WriteLine(jsonParamModel);
            var excelFile = new FileInfo(path);
            _logger.LogDebug("Opening Inventory File");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var epPackage = new ExcelPackage(excelFile))
            {
                var wscount = epPackage.Workbook.Worksheets.Count();
                Console.WriteLine("WS Count " + wscount);
                int initialRow = jsonParamModel.initialRow; 
                int invItemNameColumn = jsonParamModel.invItemNameColumn; 
                int invQuantityColumn = jsonParamModel.invQuantityColumn;
                int invBrandColumn = jsonParamModel.invBrandColumn;
                int invCategoryColumn = jsonParamModel.invCategoryColumn;

                Console.WriteLine("Param:"+initialRow);
                string itemName = "";
                string brand = "";
                string category = "";
                int quantity =0;


                for (int ws = 0; ws < wscount; ws++)
                {
                    var worksheet = epPackage.Workbook.Worksheets[ws];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    for (int row = initialRow; row <= rowCount; row = row + 1)
                    {
                        try
                        {
                            if(worksheet.Cells[row, invItemNameColumn].Value == null)
                            {
                                row = rowCount;
                                continue;
                            }
                            itemName = worksheet.Cells[row, invItemNameColumn].Value.ToString();
                            quantity = int.Parse(worksheet.Cells[row, invQuantityColumn].Value.ToString());
                            Console.WriteLine("Item: " +itemName+" Quantity: "+quantity);
                            _logger.LogDebug("Item: " +itemName+" Quantity: "+quantity);
                            InventoryEntry entry = new InventoryEntry();
                            entry.item_name = itemName;
                            entry.quantity = quantity;
                            if (quantity>0)
                            {
                                inventorylists.Add(entry);
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex);
                            _logger.LogError(ex.ToString());
                            continue;
                        }
                    }
                }

                Console.WriteLine("ListCount:"+inventorylists.Count());
                _logger.LogDebug("ListCount:"+inventorylists.Count());

                var modelJson = JsonSerializer.Serialize(inventorylists, options);
                Console.WriteLine(modelJson);
                _logger.LogDebug(modelJson);
                Console.WriteLine("URI:"+jsonParamModel.tcouri+"/upload/inventory?SupplierID="+supplierId);
                _logger.LogDebug("URI:"+jsonParamModel.tcouri+"/upload/inventory?SupplierID="+supplierId);
                Uri u = new Uri(jsonParamModel.tcouri+"/upload/inventory?SupplierID="+supplierId);
                using (HttpClient httpClient = new HttpClient())
                {
                    HttpContent c = new StringContent(modelJson, System.Text.Encoding.UTF8, "application/json");
                    var result = await httpClient.PostAsync(u, c);
                    Console.WriteLine("Result:" + await result.Content.ReadAsStringAsync());
                    _logger.LogDebug("Result:" + await result.Content.ReadAsStringAsync());
                    if (result.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Inventory Update accepted");
                        _logger.LogDebug("Inventory Update accepted");
                    }
                    else
                    {
                        Console.WriteLine("Inventory Updated failed");
                        _logger.LogDebug("Inventory Updated failed");
                    }
                }
                    
                

            }
        }
    }
}
