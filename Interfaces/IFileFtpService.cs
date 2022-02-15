using System.Threading.Tasks;

namespace DesignerEyesService.Interfaces
{
    public interface IFileFtpService
    {
        Task GetFiles(string action);
    }
}
