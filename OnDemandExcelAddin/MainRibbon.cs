using Newtonsoft.Json;dfsfsfsdfsdfsd
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// JSON verisini C# nesnelerine çevirmek için model sınıfları
public class RibbonConfig { [JsonProperty("tabs")] public List<RibbonTab> Tabs { get; set; } = new List<RibbonTab>(); }
public class RibbonTab { [JsonProperty("id")] public string Id { get; set; } [JsonProperty("label")] public string Label { get; set; } [JsonProperty("groups")] public List<RibbonGroup> Groups { get; set; } = new List<RibbonGroup>(); }
public class RibbonGroup { [JsonProperty("id")] public string Id { get; set; } [JsonProperty("label")] public string Label { get; set; } [JsonProperty("controls")] public List<RibbonControl> Controls { get; set; } = new List<RibbonControl>(); }
public class RibbonControl { [JsonProperty("id")] public string Id { get; set; } [JsonProperty("type")] public string Type { get; set; } [JsonProperty("label")] public string Label { get; set; } }

[ComVisible(true)]
public class MainRibbon : Office.IRibbonExtensibility
{
    private Office.IRibbonUI ribbon;
    private static RibbonConfig _config;
    private static bool _isLoadAttempted = false;
    private static bool _isAuthorized = false;

    public string GetCustomUI(string ribbonID)
    {
        try
        {
            string resourceText = null;
            // Projenizin varsayılan ad alanı "OnDemandExcelAddin" ise bu ad doğru çalışacaktır.
            string resourceName = "OnDemandExcelAddin.MainRibbon.xml";

            using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    MessageBox.Show("Kaynak akışı (stream) bulunamadı! XML dosya adı (" + resourceName + ") veya 'Derleme Eylemi' ayarı yanlış olabilir.", "Kritik Eklenti Hatası");
                    return null;
                }
                using (StreamReader resourceReader = new StreamReader(stream))
                {
                    resourceText = resourceReader.ReadToEnd();
                }
            }
            return resourceText;
        }
        catch (Exception ex)
        {
            MessageBox.Show("GetCustomUI metodunda kritik bir hata oluştu:\n\n" + ex.ToString(), "Kritik Eklenti Hatası");
            return null;
        }
    }

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
        try
        {
            this.ribbon = ribbonUI;
            _config = null;
            _isLoadAttempted = false;
            _isAuthorized = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show("Ribbon_Load metodunda kritik bir hata oluştu:\n\n" + ex.ToString(), "Kritik Eklenti Hatası");
        }
    }

    public async void OnLoadRibbon_Click(Office.IRibbonControl control)
    {
        // Supabase'in kullandığı modern güvenlik protokolünü (TLS 1.2) kullanmaya zorla.2
        // Supabase'in kullandığı modern güvenlik protokolünü (TLS 1.2) kullanmaya zorla.
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

        if (_isLoadAttempted) return;

        string machineId = MachineIdentifier.GetMachineGuid();
        if (string.IsNullOrEmpty(machineId))
        {
            MessageBox.Show("Bu bilgisayarın benzersiz kimliği alınamadı! Eklenti devam edemez. Sistem yöneticinizle görüşünüz.", "Kritik Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string supabaseFunctionUrl = "https://gpdjsnwplzqdwrwanzis.supabase.co/functions/v1/get-device-config";
        string supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdwZGpzbndwbHpxZHdyd2FuemlzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTI2ODgwMzcsImV4cCI6MjA2ODI2NDAzN30.i9wL-7T7gPzJ1t_J8wH9Uu-l_d_M8fN8y_Q8w_YxH0M";

        using (var client = new HttpClient())
        {
            try
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
                    _isAuthorized = true;
                }
                else
                {
                    _isAuthorized = false;
                    string unauthorizedMessage = "Bu bilgisayar için yetki bulunamadı.\n\n" +
                                                 "Lütfen aşağıdaki cihaz kimliğini sistem yöneticinize iletin. " +
                                                 "Kimlik panonuza kopyalandı:\n\n" +
                                                 machineId;

                    Clipboard.SetText(machineId);
                    MessageBox.Show(unauthorizedMessage, "Yetki Reddedildi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                _isAuthorized = false;
                MessageBox.Show("Ayarlar sunucusuna bağlanırken detaylı bir hata oluştu:\n\n" + ex.ToString(), "Ağ Hatası (Detaylı)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        _isLoadAttempted = true;
        this.ribbon?.Invalidate();
    }

    public bool GetVisible_LoadButton(Office.IRibbonControl control)
    {
        return !_isLoadAttempted;
    }

    public bool GetVisible_Dynamic(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || !_isAuthorized || _config == null) return false;
        return _config.Tabs.Any(tab => tab.Id == control.Id || tab.Groups.Any(group => group.Id == control.Id || group.Controls.Any(c => c.Id == control.Id)));
    }

    public string GetLabel(Office.IRibbonControl control)
    {
        if (!_isAuthorized || _config == null) return "";
        foreach (var tab in _config.Tabs)
        {
            if (tab.Id == control.Id) return tab.Label;
            foreach (var group in tab.Groups)
            {
                if (group.Id == control.Id) return group.Label;
                foreach (var c in group.Controls)
                {
                    if (c.Id == control.Id) return c.Label;
                }
            }
        }
        return string.Empty;
    }

    public void OnAction(Office.IRibbonControl control)
    {
        switch (control.Id)
        {
            case "btnHello":
                MessageBox.Show("Merhaba Dünya! Bu buton dinamik olarak yüklendi.", "Bilgi");
                break;
            case "btnInfo":
                MessageBox.Show("Raporlama işlemi burada tetiklenecek.", "Bilgi");
                break;
        }
    }
}