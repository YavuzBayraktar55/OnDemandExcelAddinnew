using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// JSON Model Sınıfları
public class RibbonConfig { [JsonProperty("tabs")] public List<RibbonTab> Tabs { get; set; } = new List<RibbonTab>(); }
public class RibbonTab { [JsonProperty("id")] public string Id { get; set; } [JsonProperty("label")] public string Label { get; set; } [JsonProperty("groups")] public List<RibbonGroup> Groups { get; set; } = new List<RibbonGroup>(); }
public class RibbonGroup { [JsonProperty("id")] public string Id { get; set; } [JsonProperty("label")] public string Label { get; set; } [JsonProperty("controls")] public List<RibbonControl> Controls { get; set; } = new List<RibbonControl>(); }
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
    private static bool _isAuthorized = false;

    public string GetCustomUI(string ribbonID)
    {
        try
        {
            string resourceName = "OnDemandExcelAddin.MainRibbon.xml";
            using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    MessageBox.Show("Kaynak akışı (stream) bulunamadı! XML dosya adı (" + resourceName + ") veya 'Derleme Eylemi' ayarı yanlış.", "Kritik Eklenti Hatası");
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

    public void Ribbon_Load(Office.IRibbonUI ribbonUI)
    {
        this.ribbon = ribbonUI;
        _config = null;
        _isLoadAttempted = false;
        _isAuthorized = false;
    }

    public async void OnLoadRibbon_Click(Office.IRibbonControl control)
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
                string supabaseFunctionUrl = "https://gpdjsnwplzqdwrwanzis.supabase.co/functions/v1/get-device-config";
                string supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImdwZGpzbndwbHpxZHdyd2FuemlzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTI2ODgwMzcsImV4cCI6MjA2ODI2NDAzN30.i9wL-7T7gPzJ1t_J8wH9Uu-l_d_M8fN8y_Q8w_YxH0M";

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
                    string unauthorizedMessage = "Bu bilgisayar için yetki bulunamadı.\n\nLütfen aşağıdaki cihaz kimliğini sistem yöneticinize iletin. Kimlik panonuza kopyalandı:\n\n" + machineId;
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
        return FindControlById(control.Id) != null;
    }

    public string GetLabel(Office.IRibbonControl control)
    {
        if (!_isLoadAttempted || !_isAuthorized || _config == null) return "";
        var foundControl = FindControlById(control.Id);
        return foundControl?.Label ?? "";
    }

    public void OnAction(Office.IRibbonControl control)
    {
        switch (control.Id)
        {
            case "btnHello":
                using (var frm = new FrmKayitEkrani()) { frm.ShowDialog(); }
                break;
            case "btnInfo":
                MessageBox.Show("Raporlama işlemi burada tetiklenecek.", "Bilgi");
                break;
            case "splitMain":
                MessageBox.Show("Split butonunun ana kısmına tıklandı!", "Bilgi");
                break;
            case "splitAlt1":
                MessageBox.Show("Menüdeki Alt İşlem 1'e tıklandı!", "Bilgi");
                break;
        }
    }

    private RibbonControl FindControlById(string id)
    {
        if (_config == null || _config.Tabs == null) return null;

        foreach (var tab in _config.Tabs)
        {
            if (tab.Id == id) return new RibbonControl { Id = id, Label = tab.Label };
            if (tab.Groups != null)
            {
                foreach (var group in tab.Groups)
                {
                    if (group.Id == id) return new RibbonControl { Id = id, Label = group.Label };
                    if (group.Controls != null)
                    {
                        var foundControl = FindInControlList(group.Controls, id);
                        if (foundControl != null)
                        {
                            return foundControl;
                        }
                    }
                }
            }
        }
        return null;
    }

    private RibbonControl FindInControlList(List<RibbonControl> controls, string id)
    {
        foreach (var control in controls)
        {
            if (control.Id == id)
            {
                return control;
            }
            if (control.Items != null && control.Items.Any())
            {
                var foundInItems = FindInControlList(control.Items, id);
                if (foundInItems != null)
                {
                    return foundInItems;
                }
            }
        }
        return null;
    }
}

public class FrmKayitEkrani : Form
{
    public TextBox txtAd;
    public TextBox txtSoyad;
    public TextBox txtEmail;
    public Button btnKaydet;
    public Button btnIptal;

    public FrmKayitEkrani()
    {
        this.Text = "Kayıt Ekranı";
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.Size = new Size(400, 320);
        this.BackColor = Color.WhiteSmoke;
        this.Font = new Font("Segoe UI", 10);
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        var lblTitle = new Label { Text = "Kullanıcı Kaydı", Font = new Font("Segoe UI", 16, FontStyle.Bold), ForeColor = Color.FromArgb(44, 62, 80), AutoSize = true, Location = new Point(120, 20) };
        var lblAd = new Label { Text = "Ad:", Location = new Point(50, 80), AutoSize = true };
        var lblSoyad = new Label { Text = "Soyad:", Location = new Point(50, 120), AutoSize = true };
        var lblEmail = new Label { Text = "Email:", Location = new Point(50, 160), AutoSize = true };

        txtAd = new TextBox { Location = new Point(130, 75), Width = 200, BorderStyle = BorderStyle.FixedSingle };
        txtSoyad = new TextBox { Location = new Point(130, 115), Width = 200, BorderStyle = BorderStyle.FixedSingle };
        txtEmail = new TextBox { Location = new Point(130, 155), Width = 200, BorderStyle = BorderStyle.FixedSingle };

        btnKaydet = new Button { Text = "Kaydet", Location = new Point(130, 210), Width = 90, BackColor = Color.FromArgb(52, 152, 219), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        btnIptal = new Button { Text = "İptal", Location = new Point(240, 210), Width = 90, BackColor = Color.FromArgb(231, 76, 60), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };

        btnKaydet.FlatAppearance.BorderSize = 0;
        btnIptal.FlatAppearance.BorderSize = 0;

        btnIptal.Click += (s, e) => this.Close();
        btnKaydet.Click += (s, e) =>
        {
            if (string.IsNullOrWhiteSpace(txtAd.Text) || string.IsNullOrWhiteSpace(txtSoyad.Text) || string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                MessageBox.Show("Lütfen tüm alanları doldurun.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            MessageBox.Show($"Kayıt Başarılı!\nAd: {txtAd.Text}\nSoyad: {txtSoyad.Text}\nEmail: {txtEmail.Text}", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        };

        this.Controls.Add(lblTitle);
        this.Controls.Add(lblAd);
        this.Controls.Add(lblSoyad);
        this.Controls.Add(lblEmail);
        this.Controls.Add(txtAd);
        this.Controls.Add(txtSoyad);
        this.Controls.Add(txtEmail);
        this.Controls.Add(btnKaydet);
        this.Controls.Add(btnIptal);
    }
}