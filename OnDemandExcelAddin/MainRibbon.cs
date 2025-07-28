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

// JSON Model Sınıfları (Yeni mimariye uygun)
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
                case "btnAdminMain":
                    return PictureConverter.GetImage("OnDemandExcelAddin.Resources.Admin32.png");

                case "btnSetPermissions":
                    return PictureConverter.GetImage("OnDemandExcelAddin.Resources.Permissions32.png");

                case "btnCreateReport":
                    return PictureConverter.GetImage("OnDemandExcelAddin.Resources.Report32.png");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"GetImage metodu içinde bir hata oluştu (ID: {control.Id}):\n\n{ex.ToString()}", "Resim Yükleme Hatası");
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
                    MessageBox.Show($"Kaynak akışı (stream) bulunamadı! XML dosya adı ({resourceName}) veya 'Derleme Eylemi' ayarı yanlış olabilir.", "Kritik Eklenti Hatası");
                    return null;
                }
                using (StreamReader resourceReader = new StreamReader(stream))
                {
                    return resourceReader.ReadToEnd();
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("GetCustomUI metodunda kritik bir hata oluştu:\n\n" + ex.ToString(), "Kritik Eklenti Hatası");
            return null;
        }
    }
    //// --- GÖRÜNTÜ YÜKLEME METODU GÜNCELLENDİ ---
    //public stdole.IPictureDisp GetImage(Office.IRibbonControl control)
    //{
    //    // Butonun ID'sine göre hangi resmi yükleyeceğimizi seçiyoruz.
    //    switch (control.Id)
    //    {
    //        case "btnAdminMain":
    //            // Lütfen projenizde "Resources" klasörü altında "Admin32.png" adında bir resim olduğundan
    //            // ve "Derleme Eylemi"nin "Gömülü Kaynak" olduğundan emin olun.
    //            return PictureConverter.GetImage("OnDemandExcelAddin.Resources.Admin32.png");

    //        //case "btnOpenForm":
    //        //    // Lütfen projenizde "Resources" klasörü altında "sayacNe32x32.png" adında bir resim olduğundan
    //        //    // ve "Derleme Eylemi"nin "Gömülü Kaynak" olduğundan emin olun.
    //        //    return PictureConverter.GetImage("OnDemandExcelAddin.Resources.sayacNe32x32.png");
    //    }
    //    return null; // Eşleşen bir ID yoksa resim yükleme.
    //}
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
            MessageBox.Show("Bu bilgisayarın benzersiz kimliği alınamadı!", "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        using (var client = new HttpClient())
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string appName = assembly.GetName().Name;
                string settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), appName, "settings.json");

                if (!File.Exists(settingsPath))
                {
                    MessageBox.Show($"Ayar dosyası bulunamadı!\nBeklenen Konum: {settingsPath}", "Yapılandırma Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string settingsJson = File.ReadAllText(settingsPath);
                var settings = JsonConvert.DeserializeObject<dynamic>(settingsJson);
                string supabaseFunctionUrl = settings.SupabaseUrl;
                string supabaseAnonKey = settings.SupabaseAnonKey;

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
                    _config = null; // Yetki yoksa veya hata varsa config'i boşalt
                    string unauthorizedMessage = "Bu bilgisayar için yetki bulunamadı.\n\nLütfen aşağıdaki cihaz kimliğini sistem yöneticinize iletin. Kimlik panonuza kopyalandı:\n\n" + machineId;
                    Clipboard.SetText(machineId);
                    MessageBox.Show(unauthorizedMessage, "Yetki Reddedildi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                _config = null; // Hata durumunda config'i boşalt
                MessageBox.Show("Ayarlar sunucusuna bağlanırken detaylı bir hata oluştu:\n\n" + ex.ToString(), "Ağ Hatası (Detaylı)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        _isLoadAttempted = true;
        this.ribbon?.Invalidate(); // Arayüzü yeniden çizdir.
    }

    public bool GetVisible_LoadGroup(Office.IRibbonControl _)
    {
        return !_isLoadAttempted;
    }

    public bool GetVisible_Dynamic(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || _config == null) return false;
        return FindControlById(control.Id) != null;
    }

    public string GetLabel(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || _config == null) return "";
        var foundControl = FindControlById(control.Id);
        return foundControl?.Label ?? "";
    }

    public void OnAction(Office.IRibbonControl control)
    {
        switch (control.Id)
        {
            // YENİ EKLENDİ: Bu buton FrmKayitEkrani'nı açacak.
            case "btnOpenForm":
                using (var frm = new FrmKayitEkrani())
                {
                    frm.ShowDialog();
                }
                break;

            case "btnAdminMain":
            case "btnAddDevice":
                MessageBox.Show("Yeni cihaz ekleme formu burada açılacak.");
                break;

            case "btnSetPermissions":
                MessageBox.Show("Yetki düzenleme ekranı burada açılacak.");
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
            if (group.Controls != null)
            {
                var foundControl = FindInControlList(group.Controls, id);
                if (foundControl != null) return foundControl;
            }
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
}
internal class PictureConverter : AxHost
{
    private PictureConverter() : base(null) { }

    public static stdole.IPictureDisp GetImage(string imageName)
    {
        // Bu metodun kendisi zaten try-catch içeriyor, bu iyi.
        try
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName);
            if (stream != null)
            {
                using (Bitmap bmp = new Bitmap(stream))
                {
                    return (stdole.IPictureDisp)GetIPictureDispFromPicture(bmp);
                }
            }
            else
            {
                // Eğer stream null ise, bu kaynak adının yanlış olduğu anlamına gelir.
                MessageBox.Show($"Gömülü kaynak bulunamadı:\n\n{imageName}\n\nLütfen proje ad alanını ve dosya adını kontrol edin.", "Resim Kaynak Hatası");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"PictureConverter içinde hata oluştu (Resim Adı: {imageName}):\n\n{ex.ToString()}", "Resim Dönüştürme Hatası");
        }
        return null;
    }
}
//internal class PictureConverter : AxHost
//{
//    private PictureConverter() : base(null) { }

    //public static stdole.IPictureDisp GetImage(string imageName)
    //{
    //    try
    //    {
    //        // Projenin bütününden (assembly) kaynak akışını (stream) bul.
    //        Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageName);
    //        if (stream != null)
    //        {
    //            // Akıştan bir Bitmap oluştur.
    //            Bitmap bmp = new Bitmap(stream);
    //            // Bitmap'i Office'in anladığı IPictureDisp formatına çevir.
    //            return (stdole.IPictureDisp)GetIPictureDispFromPicture(bmp);
    //        }
    //    }
    //    catch (Exception)
    //    {
    //        // Hata olursa (resim bulunamazsa vb.) boş döndür.
    //        return null;
    //    }
    //    return null;
    //}


// Not: Bu dosyada FrmKayitEkrani sınıfı yok, bu sınıfı ayrı bir dosyada tutmak daha temiz bir yaklaşımdır.
// Eğer hala aynı dosyada tutuyorsanız, buraya ekleyebilirsiniz.

