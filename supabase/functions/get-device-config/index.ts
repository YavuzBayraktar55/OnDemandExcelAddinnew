import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

Deno.serve(async (req) => {
  try {
    // Güvenli istemciyi, RLS'i aşma gücüne sahip olan ve ortam değişkenlerinden
    // (secrets) okunan 'service_role' anahtarı ile oluştur.
    const supabaseClient = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('MY_SERVICE_KEY') ?? '' // 'secrets set' ile ayarladığımız özel isim
    );

    // C#'tan gelen isteğin gövdesinden (body) machine_uuid'yi al.
    const { machine_uuid } = await req.json();

    if (!machine_uuid) {
      throw new Error("machine_uuid is required in the request body.");
    }

    // Veritabanı sorgusu: İlgili UUID'ye sahip kaydın sadece 'ribbon_config' kolonunu seç.
    const { data, error } = await supabaseClient
      .from('device_configs')
      .select('ribbon_config')
      .eq('machine_uuid', machine_uuid)
      .single(); // Sadece tek bir sonuç bekliyoruz.

    if (error) {
      // Supabase, kayıt bulunamadığında bir hata fırlatır. Bu hatayı yakalayıp
      // istemciye "bulunamadı" anlamına gelen 404 hatası döndürüyoruz.
      console.error(error.message);
      return new Response(JSON.stringify({ error: 'Config not found for this device.' }), {
        headers: { 'Content-Type': 'application/json' },
        status: 404,
      });
    }

    // Başarılı olursa, direkt olarak bulunan 'ribbon_config' JSON nesnesini döndür.
    return new Response(JSON.stringify(data.ribbon_config), {
      headers: { 'Content-Type': 'application/json' },
      status: 200,
    });

  } catch (err) {
    // Diğer tüm beklenmedik hataları yakala (örn: JSON parse hatası).
    return new Response(String(err?.message ?? err), { status: 500 });
  }
});