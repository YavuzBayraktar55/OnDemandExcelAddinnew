using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

public class RibbonConfig
{
    [JsonProperty("groups")]
    public List<RibbonGroup> Groups { get; set; } = new List<RibbonGroup>();
}

public class RibbonGroup
{
    [JsonProperty("id")] public string Id { get; set; }
    [JsonProperty("label")] public string Label { get; set; }
    [JsonProperty("controls")] public List<RibbonControl> Controls { get; set; } = new List<RibbonControl>();
}

public class RibbonControl
{
    [JsonProperty("id")] public string Id { get; set; }
    [JsonProperty("type")] public string Type { get; set; }
    [JsonProperty("label")] public string Label { get; set; }
    [JsonProperty("items")] public List<RibbonControl> Items { get; set; } = new List<RibbonControl>();
}

[ComVisible(true)]
public class MainRibbon : Office.IRibbonExtensibility
{
    private Office.IRibbonUI ribbon;
    private static RibbonConfig _config;
    private static bool _isLoadAttempted = false;

    public stdole.IPictureDisp GetImage(Office.IRibbonControl control)
    {
        try
        {
            switch (control.Id)
            {
                case "btnOpenForm":
                    return PictureConverter.GetImage("AddDevice.16.png");
                case "btnAdminMain":
                    return PictureConverter.GetImage("Admin32.png");
                case "btnAddDevice":
                    return PictureConverter.GetImage("AddDevice.16.png");
                case "btnSetPermissions":
                    return PictureConverter.GetImage("SetPermissions.png");

                case "btnCreateReport":
                    return PictureConverter.GetImage("Report32.png");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"GetImage metodu içinde bir hata oluştu (ID: {control.Id}):\n\n{ex}", "Resim Yükleme Hatası");
        }
        return null;
    }

    public string GetCustomUI(string ribbonID)
    {
        try
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            string resourceName = $"{assemblyName}.MainRibbon.xml";

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    MessageBox.Show($"Kaynak XML bulunamadı: {resourceName}", "XML Yükleme Hatası");
                    return null;
                }

                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Ribbon XML yüklenirken hata:\n\n" + ex.ToString(), "Ribbon Hatası");
            return null;
        }
    }

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
        this.ribbon = ribbonUI;
        _config = null;
        _isLoadAttempted = false;
    }

    public async void OnLoadRibbon_Click(Office.IRibbonControl _)
    {
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

        if (_isLoadAttempted) return;

        string machineId = MachineIdentifier.GetMachineGuid();
        if (string.IsNullOrEmpty(machineId))
        {
            MessageBox.Show("Makine kimliği alınamadı.", "Kritik Hata");
            return;
        }

        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            string appName = assembly.GetName().Name;
            string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, "settings.json");

            if (!File.Exists(settingsPath))
            {
                MessageBox.Show($"Ayar dosyası eksik: {settingsPath}", "Hata");
                return;
            }

            string settingsJson = File.ReadAllText(settingsPath);
            var settings = JsonConvert.DeserializeObject<dynamic>(settingsJson);
            string supabaseFunctionUrl = settings.SupabaseUrl;
            string supabaseAnonKey = settings.SupabaseAnonKey;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("apikey", supabaseAnonKey);
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {supabaseAnonKey}");

                var payload = new { machine_uuid = machineId };
                var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");

                var response = await client.PostAsync(supabaseFunctionUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string jsonString = await response.Content.ReadAsStringAsync();
                    _config = JsonConvert.DeserializeObject<RibbonConfig>(jsonString);
                }
                else
                {
                    Clipboard.SetText(machineId);
                    MessageBox.Show("Bu bilgisayara yetki verilmemiş.\nCihaz kimliği panoya kopyalandı:\n" + machineId, "Yetkisiz");
                    _config = null;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Ağ hatası oluştu:\n\n" + ex.ToString(), "Bağlantı Hatası");
            _config = null;
        }

        _isLoadAttempted = true;
        ribbon?.Invalidate();
    }

    public bool GetVisible_LoadGroup(Office.IRibbonControl _) => !_isLoadAttempted;

    public bool GetVisible_Dynamic(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || _config == null) return false;
        return FindControlById(control.Id) != null;
    }

    public string GetLabel(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || _config == null) return "";
        var found = FindControlById(control.Id);
        return found?.Label ?? "";
    }

    public void OnAction(Office.IRibbonControl control)
    {
        switch (control.Id)
        {

            // --- YENİ EKLENEN STATİK BUTONLAR ---
            case "btnStatic1":
                MessageBox.Show("Statik Test Menüsü - Buton 1'e tıklandı.", "Statik Test");
                break;
            case "btnStatic2":
                MessageBox.Show("Statik Test Menüsü - Buton 2'ye tıklandı.", "Statik Test");
                break;
            case "btnStatic3":
                MessageBox.Show("Statik Test Menüsü - Buton 3'e tıklandı.", "Statik Test");
                break;
            // ------------------------------------


            case "btnOpenForm":
                using (var frm = new FrmKayitEkrani())
                {
                    frm.ShowDialog();
                }
                break;

            case "btnAdminMain":
            case "btnAddDevice":
                MessageBox.Show("Yeni cihaz ekleme ekranı açılacak.");
                break;

            case "btnSetPermissions":
                MessageBox.Show("Yetki ayar ekranı açılacak.");
                break;

            case "btnCreateReport":
                MessageBox.Show("Rapor oluşturuluyor...");
                break;
        }
    }

    private RibbonControl FindControlById(string id)
    {
        if (_config == null || _config.Groups == null) return null;

        foreach (var group in _config.Groups)
        {
            if (group.Id == id) return new RibbonControl { Id = id, Label = group.Label };

            var found = FindInControlList(group.Controls, id);
            if (found != null) return found;
        }
        return null;
    }

    private RibbonControl FindInControlList(List<RibbonControl> controls, string id)
    {
        foreach (var control in controls)
        {
            if (control.Id == id) return control;

            if (control.Items != null && control.Items.Any())
            {
                var foundInItems = FindInControlList(control.Items, id);
                if (foundInItems != null) return foundInItems;
            }
        }
        return null;
    }

    #region Gallery Callback Methods

    private List<string> galleryItems = new List<string> { "Normal", "Geniş", "Dar" };

    // Galeriye tıklandığında bu metod çalışır.
    public void OnGalleryAction(Office.IRibbonControl control, string id, int index)
    {
        MessageBox.Show($"Galeriden '{galleryItems[index]}' seçeneği seçildi (index: {index})", "Galeri Testi");
    }

    // Galeride kaç öğe olduğunu söyler.
    public int OnGetItemCount(Office.IRibbonControl control)
    {
        return galleryItems.Count;
    }

    // Belirli bir index'teki öğenin etiketini döndürür.
    public string OnGetItemLabel(Office.IRibbonControl control, int index)
    {
        return galleryItems[index];
    }

    // Belirli bir index'teki öğenin resmini döndürür.
    public stdole.IPictureDisp OnGetItemImage(Office.IRibbonControl control, int index)
    {
        switch (index)
        {
            case 0: return PictureConverter.GetImage("MarginNormal.png");
            case 1: return PictureConverter.GetImage("MarginWide.png");
            case 2: return PictureConverter.GetImage("MarginNarrow.png");
            default: return null;
        }
    }

    // Belirli bir index'teki öğenin ipucunu döndürür.
    public string OnGetItemScreentip(Office.IRibbonControl control, int index)
    {
        return $"{galleryItems[index]} Kenar Boşluğu";
    }

    public string OnGetItemSupertip(Office.IRibbonControl control, int index)
    {
        return $"Sayfa kenar boşluklarını '{galleryItems[index]}' olarak ayarlar.";
    }

    #endregion
}

internal class PictureConverter : AxHost
{
    private PictureConverter() : base(null) { }

    public static stdole.IPictureDisp GetImage(string imageFileName)
    {
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var assemblyName = assembly.GetName().Name;

            // Resources klasörüne gömülü tüm alt klasörler dahil
            string fullResourceName = $"{assemblyName}.Resources.{imageFileName.Replace("\\", ".").Replace("/", ".")}";

            using (Stream stream = assembly.GetManifestResourceStream(fullResourceName))
            {
                if (stream != null)
                {
                    using (Bitmap bmp = new Bitmap(stream))
                    {
                        return (stdole.IPictureDisp)GetIPictureDispFromPicture(bmp);
                    }
                }
                else
                {
                    MessageBox.Show($"Gömülü kaynak bulunamadı:\n\n{fullResourceName}", "Resim Kaynak Hatası");
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"PictureConverter içinde hata oluştu:\n\n{ex}", "Resim Dönüştürme Hatası");
        }

        return null;
    }
}