using System.Threading.Tasks;

namespace DesignerEyesService.Interfaces
{
    public interface IReadShipConfirm
    {
        Task ReadShipConfirmAsync(int supplierId, string path);
    }
}
