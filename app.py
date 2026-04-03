import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import io
import traceback
import os
from openai import OpenAI

st.set_page_config(page_title="Excel AI Uzmanı", layout="wide", page_icon="🟢")

st.markdown("""
<style>
    /* Modern Tasarım */
    .stApp {
        background-color: #fcefeƒ;
    }
    
    h1 {
        color: #047857; /* Koyu yeşil - Excel hissiyatı */
        font-weight: 800;
        margin-bottom: -10px;
    }
    
    .stButton button {
        background-color: #10B981 !important;
        color: white !important;
        border-radius: 8px;
        font-weight: bold;
    }
    .stButton button:hover {
        background-color: #059669 !important;
    }
    
    .stDownloadButton button {
        background-color: #3B82F6 !important;
        color: white !important;
        border-radius: 8px;
        font-weight: bold;
    }
    .stDownloadButton button:hover {
        background-color: #2563EB !important;
    }
    
    /* İndir Butonu Konteynırı */
    div[data-testid="stVerticalBlock"] div:has(div.stDownloadButton) {
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

st.title("🟢 Yapay Zeka Excel Uzmanı")
st.markdown("Gerçek Excel dosyaları üretir. Hücre boyama, toplama-çıkarma **formülleri** ve profesyonel analizler.")

# --- API KEY ALANI ---
try:
    api_key = st.secrets["OPENAI_API_KEY"]
except:
    api_key = os.environ.get("OPENAI_API_KEY")

st.sidebar.divider()
st.sidebar.markdown(
    "💡 **İpuçları:**\n"
    "- 'Maaş ile primi toplayıp Toplam sütununa excel formülü olarak yaz.'\n"
    "- 'B sütununu yeşile boya.'\n"
    "- '100'den büyük hücreleri bul ve kalın yap.'"
)

# --- APP STATE & FILES ---
WORKSPACE_FILE = "workspace_excel.xlsx"

if "messages" not in st.session_state:
    st.session_state["messages"] = []

if "file_ready" not in st.session_state:
    st.session_state["file_ready"] = False

# --- DOSYA YÜKLEME / OLUŞTURMA (DOĞRUDAN EKRANIN ORTASI) ---
if not st.session_state["file_ready"]:
    st.info("👋 Başlamak için aşağıdaki seçeneklerden birini kullanın:")
    col_up, col_new = st.columns(2)
    
    with col_up:
        st.subheader("📁 Mevcut Dosyanızı Yükleyin")
        uploaded_file = st.file_uploader("Sadece .xlsx dosyaları:", type=["xlsx"])
        if uploaded_file:
            with open(WORKSPACE_FILE, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state["file_ready"] = True
            st.session_state["messages"] = [{"role": "assistant", "content": "Dosyanızı inceledim. Hangi hücrelere excel formülü yazalım veya nasıl düzenleyelim?"}]
            st.rerun()

    with col_new:
        st.subheader("📝 Sıfırdan Boş Excel Yarat")
        st.write("Kendiniz bir test dosyası üzerinden işlemi deneyimleyin.")
        if st.button("Boş Excel İle Başla (Örnek Tablo)", use_container_width=True):
            wb = openpyxl.Workbook()
            ws = wb.active
            # Başlıklar
            basliklar = ["Ürün Adi", "Aylik Satis", "Birim Fiyat", "KDV Orani", "Net Satis"]
            ws.append(basliklar)
            
            # Biraz örnek dummy veri
            ws.append(["Elma", 120, 15, 0.20, ""])
            ws.append(["Armut", 80, 20, 0.20, ""])
            ws.append(["Kavun", 45, 30, 0.10, ""])
            
            wb.save(WORKSPACE_FILE)
            st.session_state["file_ready"] = True
            st.session_state["messages"] = [{"role": "assistant", "content": "Size 3 satırlık meyve satışları içeren test tablosu hazırladım. Örneğin: **'Net satış sütununu satış ile birim fiyatı çarparak KDV dahil bir excel formülü yazdır.'** diyebilirsiniz!"}]
            st.rerun()
            
else:
    # --- ÇALIŞMA ALANI ---
    col_preview, col_actions = st.columns([5, 2])
    
    with col_preview:
        st.subheader("📊 Tablonun Son Hali")
        try:
            # Sadece okuma/basma amacıyla pandas:
            # Pandas ile formüller (örn: "=B2*C2") ekranda "NaN" veya "text" gözükebilir ama Excel'de düzgün açılacaktır.
            # Ekranda null görünebilecek formül hücreleri için endişelenmeye gerek yok, dilersek string olarak gösterebiliriz (pandas engine calisiyor)
            df_preview = pd.read_excel(WORKSPACE_FILE, engine="openpyxl")
            st.dataframe(df_preview, use_container_width=True, height=200)
        except Exception as e:
            st.error(f"Önizleme gösterilemedi: {e}")
            
    with col_actions:
        st.subheader("💾 İndirme & Ayarlar")
        
        with open(WORKSPACE_FILE, "rb") as f:
            st.download_button(
                label="📥 İndir (Gerçek Formüllerle)",
                data=f,
                file_name="Ai_Duzenlenmis_Umarim_Son.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        st.divider()
        if st.button("❌ Bu Dosyayı Kapat", use_container_width=True):
            st.session_state["file_ready"] = False
            st.session_state["messages"] = []
            st.rerun()
            
    st.divider()
    
    # Sohbet Alanı
    for msg in st.session_state["messages"]:
        if msg["role"] == "user":
            st.chat_message("user").write(msg["content"])
        else:
            st.chat_message("assistant", avatar="🟢").write(msg["content"])
            if "code" in msg:
                with st.expander("🛠 AI'ın Yazdığı OpenPyXL Kodunu İncele"):
                    st.code(msg["code"], language="python")

    user_query = st.chat_input("Hücre rengi değiştirme, Matematiksel Excel formülü hesaplamaları, Sütun silme vb. söyleyin.")
    
    if user_query:
        if not api_key:
            st.error("Sistem Yapılandırma Hatası: OpenAI API Anahtarı sunucuya eklenmemiş.")
            st.stop()
            
        st.session_state["messages"].append({"role": "user", "content": user_query})
        st.chat_message("user").write(user_query)
        
        with st.chat_message("assistant", avatar="🟢"):
            with st.spinner("AI Dosyayı Analiz Ediyor, Formülleri işliyor... 🚀"):
                # Excel bağlamını alma
                try:
                    wb_temp = openpyxl.load_workbook(WORKSPACE_FILE)
                    ws_temp = wb_temp.active
                    max_row = ws_temp.max_row
                    max_col = ws_temp.max_column
                    
                    sample_data = []
                    for row in ws_temp.iter_rows(min_row=1, max_row=min(max_row, 3), values_only=True):
                        sample_data.append(row)
                        
                    context = f"""
                    MEVCUT EXCEL DOSYASI:
                    - Satır Sayısı: {max_row}
                    - Sütun Sayısı: {max_col}
                    - İlk {min(max_row, 3)} Satır Verisi:
                    {sample_data}
                    
                    ÖNEMLİ: İlk satır genellikle başlıkları içerir (1. satır).
                    """
                except Exception as e:
                    context = f"Veriler okunamadı. Hata: {e}"

                system_prompt = f"""
                Sen İLERİ DÜZEY Bir Excel ve Python OpenPyXL Uzmanısın!
                Kesinlikle 'pandas' kütüphanesini KULLANMA. Aksi belirtilmedikçe Tüm işlemleri 'openpyxl' kullanarak mevcut excel dosyası üzerinde gerçekleştireceksin.
                
                GÖREVİN: Kullanıcının isteğini yerine getiren SAF, çalıştırılabilir PYTHON ('openpyxl' kullanarak) kodunu yazmaktır.
                
                ZORUNLU KURALLAR:
                1. Hiçbir zaman kodu markdown block (```python) içine ALMA. Yalnızca salt harfiyen Python komutlarını yaz. 
                2. Kodun EN BAŞINA 'import openpyxl' yaz. (Gerekiyorsa: 'from openpyxl.styles import PatternFill, Font')
                3. Mutlaka Excel'i aç: `wb = openpyxl.load_workbook("{WORKSPACE_FILE}")`
                4. Aktif sayfayı al: `ws = wb.active`
                5. İstekleri hücre bazlı yap. Excel Formülü işleteceksen `=B2*C2` şeklinde metin olarak `ws['D2'] = "=B2*C2"` vs yaz.
                6. Mutlaka iteratif satır işlemlerinde, döngünün 2'den başladığından (for r in range(2, ws.max_row + 1):) ve başlık satırına dokunmadığından emin ol.
                7. Kodun EN SONUNA HER ZAMAN `wb.save("{WORKSPACE_FILE}")` satırını ekle!
                8. Print gibi fonksiyonlar kullanma. Güvenlik için bilgisayardaki diğer hiçbir dosya veya yola erişme!
                
                KRİTİK UYARI - YAPILAMAYACAKLAR:
                - Openpyxl HİÇBİR ZAMAN Excel dosyasına tıklanabilir bir buton, kontrol (ComboBox, Button) veya Makro (VBA) EKLEYEMEZ. Kullanıcı "buton ekle" derse SADECE o hücreye "Çıkart" yazmak gibi görsel işlemler yap, ama metot uydurma (örn ws.add_button vs yoktur!).
                - Bilmediğin openpyxl attribute'larını veya sınıflarını uydurma (hallucinate etme).
                
                DOSYA BİLGİSİ:
                {context}
                """
                
                try:
                    client = OpenAI(api_key=api_key)
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_query}
                        ],
                        temperature=0
                    )
                    
                    code_string = response.choices[0].message.content.strip()
                    
                    # LLM inat edip markdown yazarsa diye güvenlik filtresi:
                    if code_string.startswith("```python"):
                        code_string = code_string[9:]
                    elif code_string.startswith("```"):
                        code_string = code_string[3:]
                        
                    if code_string.endswith("```"):
                        code_string = code_string[:-3]
                        
                    code_string = code_string.strip()
                    
                    # Güvenli ortamda Çalıştırma
                    try:
                        local_env = {}
                        exec(code_string, globals(), local_env)
                        
                        success_msg = "İşlem başarıyla excelinize native(orjinal) olarak uygulandı!"
                        st.session_state["messages"].append({
                            "role": "assistant",
                            "content": success_msg,
                            "code": code_string
                        })
                        st.rerun()
                        
                    except Exception as code_error:
                        st.error("Kod çalıştırılırken bir hata oluştu.")
                        err_msg = f"Sizin için yazdığım excel uzman kodunu çalıştırırken şu hatayı aldım:\n`{str(code_error)}`"
                        st.session_state["messages"].append({"role": "assistant", "content": err_msg, "code": code_string})
                        
                except Exception as api_err:
                    st.error(f"OpenAI API Hatası: {str(api_err)}")
