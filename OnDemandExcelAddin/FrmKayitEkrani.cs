using System.Windows.Forms;
using System.Drawing;

// Bu sınıf artık kendi özel dosyasında yaşıyor.
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