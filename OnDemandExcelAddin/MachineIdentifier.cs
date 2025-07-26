using System.Linq;
using System.Management;

public static class MachineIdentifier
{
    public static string GetMachineGuid()
    {
        try
        {
            // Windows Yönetim Araçları'nı (WMI) kullanarak bilgisayar sistem ürün bilgilerini sorgularız.
            using (var searcher = new ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct"))
            {
                // Sorgu sonucundan ilk nesnenin "UUID" özelliğini alır ve string olarak döndürürüz.
                using (var collection = searcher.Get())
                {
                    return collection.Cast<ManagementBaseObject>().First()["UUID"].ToString();
                }
            }
        }
        catch
        {
            // Herhangi bir hata durumunda (örn: yetki sorunları), boş bir string döndürerek
            // ana kodun bu durumu yönetmesini sağlarız.
            return string.Empty;
        }
    }
}