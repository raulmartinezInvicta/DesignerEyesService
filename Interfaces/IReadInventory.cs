using System.Threading.Tasks;

namespace DesignerEyesService.Interfaces
{
    public interface IReadInventory
    {
        Task ReadInventoryAsync(int supplierId, string path);
    }
}
