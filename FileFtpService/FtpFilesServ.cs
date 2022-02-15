using DesignerEyesService.Entities;
using DesignerEyesService.Interfaces;
using FluentFTP;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace DesignerEyesService.FileFtpService
{
    public class FtpFilesServ : IFileFtpService
    {
        private readonly ILogger<FtpFilesServ> _logger;
        private readonly IReadInventory _inventory;
        private readonly IReadShipConfirm _shipConfirm;

        public FtpFilesServ(ILogger<FtpFilesServ> logger, IReadInventory inventory, IReadShipConfirm shipConfirm)
        {
            _logger = logger;
            _inventory = inventory;
            _shipConfirm = shipConfirm;
        }

        public async Task GetFiles(string action)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true
            };
            var jsonParameters = File.ReadAllText("appsettings.json");
            var jsonParamModel = JsonSerializer.Deserialize<Parameters>(jsonParameters, options);

            _logger.LogInformation("Connecting with ftp service");

            using (FtpClient ftp = new FtpClient(jsonParamModel.ftpHost, new System.Net.NetworkCredential { UserName = jsonParamModel.ftpUser, Password = jsonParamModel.ftpPassword }))
            {
                ftp.Connect();
                _logger.LogInformation("Connection to ftp service successful");
                string path = "";
                var list = ftp.GetListing(action, FtpListOption.Recursive);
                _logger.LogInformation("Obtained files");
                foreach (var ftpItem in list)
                {
                    if (ftpItem.Type != FtpFileSystemObjectType.File)
                        continue;

                    switch (action)
                    {
                        case "Inventory":
                            path = "./Inventory/Inventory.xlsx";
                            ftp.DownloadFile(path, ftpItem.FullName, FtpLocalExists.Overwrite, FtpVerify.Retry);
                            _logger.LogInformation("DownloadFile:Inventory");
                            await _inventory.ReadInventoryAsync(jsonParamModel.supplierId, path);
                            ftp.DeleteFile(ftpItem.FullName);

                            break;

                        case "Out Trackings":
                            path = "./Out Trackings/Tracking.xlsx";
                            ftp.DownloadFile(path, ftpItem.FullName, FtpLocalExists.Overwrite, FtpVerify.Retry);
                            _logger.LogInformation("DownloadFile:Tracking");
                            await _shipConfirm.ReadShipConfirmAsync(jsonParamModel.supplierId, path);
                            ftp.DeleteFile(ftpItem.FullName);
                            break;
                    }
                }
            }
        }
    }
}
