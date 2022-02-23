using DesignerEyesService.FileFtpService;
using DesignerEyesService.Interfaces;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DesignerEyesService
{
    public class MainProcess
    {
        private readonly ILogger<MainProcess> _logger;
        private readonly IFileFtpService _ftp;
        private readonly IReadOrder _order;

        public MainProcess(ILogger<MainProcess> logger, IFileFtpService ftp, IReadOrder order)
        {
            _logger = logger;
            _ftp = ftp;
            _order = order;
        }

        public async Task StartService(string[] args)
        {
            #region CreateDirectories
            _logger.LogInformation("Creating directories");
            if (!System.IO.Directory.Exists(".\\Inventory"))
            {
                System.IO.Directory.CreateDirectory(".\\Inventory");
            }

            if (!System.IO.Directory.Exists(".\\Orders"))
            {
                System.IO.Directory.CreateDirectory(".\\Orders");
            }

            if (!System.IO.Directory.Exists(".\\Tracking"))
            {
                System.IO.Directory.CreateDirectory(".\\Tracking");
            }

            _logger.LogInformation("Directories created successfully");
            #endregion

            #region DeleteFiles
            _logger.LogInformation("Deleting history");
            string[] files = Directory.GetFiles(".\\Inventory\\");
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddDays(-7))
                    fi.Delete();
            }

            string[] files2 = Directory.GetFiles(".\\Orders\\");
            foreach (string file in files2)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddDays(-7))
                    fi.Delete();
            }
            string[] files3 = Directory.GetFiles(".\\Tracking\\");
            foreach (string file in files3)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddDays(-7))
                    fi.Delete();
            }
            _logger.LogInformation("History deleted successfully");
            #endregion

            await LobbyAction(args);
        }

        private async Task LobbyAction(string[] args)
        {
            Console.WriteLine("Select the option you want to implement");
            Console.WriteLine("Action  (-i → Inventory,  -o → Orders, -s → Shipping, -e → Exit )");
            _logger.LogInformation("Entrance to lobby");
            string option;
            option = Console.ReadLine();
            //_logger.LogInformation($"Action executed {args[0]}");
            //if (args.Length > 0 && args[0] == "-i")
            //{
            //    Console.WriteLine("Processing inventory...");
            //    try
            //    {
            //        _logger.LogInformation("Selection: Inventory");
            //        await _ftp.GetFiles("Inventory");
            //        //LobbyAction(args);
            //    }
            //    catch (Exception e)
            //    {
            //        _logger.LogError($"Exception: {e}");
            //    }
            //}
            //else if (args.Length > 0 && args[0] == "-o")
            //{
            //    Console.WriteLine("Processing order...");
            //    try
            //    {
            //        _logger.LogInformation("Selection: Order");
            //        _order.ReadOrdersData();
            //        //LobbyAction(args);
            //    }
            //    catch (Exception e)
            //    {
            //        _logger.LogError($"Exception: {e}");
            //    }
            //}
            //else if (args.Length > 0 && args[0] == "-s")
            //{
            //    Console.WriteLine("Processing ShipConfirm...");
            //    try
            //    {
            //        _logger.LogInformation("Selection: Tracking");
            //        //LobbyAction(args);
            //    }
            //    catch (Exception e)
            //    {
            //        _logger.LogError($"Exception: {e}");
            //    }
            //}


            //Manual execution
            switch (option)
            {
                case "-i":
                    Console.WriteLine("Processing inventory...");
                    try
                    {
                        _logger.LogInformation("Selection: Inventory");
                        await _ftp.GetFiles("Inventory");
                        await LobbyAction(args);
                    }
                    catch (Exception e)
                    {
                        _logger.LogError($"Exception: {e}");
                        throw new Exception(e.Message);
                    }

                    break;

                case "-o":
                    Console.WriteLine("Processing order...");
                    try
                    {
                        _logger.LogInformation("Selection: Order");
                        _order.ReadOrdersData();
                        await LobbyAction(args);
                    }
                    catch (Exception e)
                    {
                        _logger.LogError($"Exception: {e}");
                        throw new Exception(e.Message);
                    }
                    break;

                case "-s":
                    Console.WriteLine("Processing ShipConfirm...");
                    try
                    {
                        _logger.LogInformation("Selection: Tracking");
                        await _ftp.GetFiles("Out Trackings");
                        await LobbyAction(args);
                    }
                    catch (Exception e)
                    {
                        _logger.LogError($"Exception: {e}");
                        throw new Exception(e.Message);
                    }
                    break;

                case "-e":
                    _logger.LogInformation("Selection: Exit");
                    Environment.Exit(0);
                    break;

                default:
                    Console.WriteLine("You must enter allowed values");
                    await LobbyAction(args);
                    break;

            }
        }
    }
}
