import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import io

st.set_page_config(page_title="Coretax Kredit Gabungan", page_icon="📑", layout="wide")

st.title("📑 Coretax - Kredit Gabungan")
st.markdown("**Bantu kreditkan faktur masukan sekaligus | Aturan: 2025+ wajib masa sama, 2024 kebawah boleh +3 bulan**")

with st.sidebar:
    st.header("⚙️ Pengaturan")
    masa_tujuan_str = st.selectbox("Kreditkan untuk SPT Masa:", [f"{y}-{m:02d}" for y in range(2024, 2027) for m in range(1,13)][::-1], index=2)
    st.info(f"Masa Tujuan: **{masa_tujuan_str}**")
    st.divider()
    st.caption("Upload CSV dari Coretax: eFaktur > Pajak Masukan > Bulk Process > Download CSV by Period")

uploaded = st.file_uploader("📤 Upload CSV / Excel Coretax di sini", type=["csv","xlsx"])

if uploaded:
    if uploaded.name.endswith('.xlsx'):
        df = pd.read_excel(uploaded)
    else:
        try:
            df = pd.read_csv(uploaded, sep=None, engine='python')
        except:
            df = pd.read_csv(uploaded, sep=';', engine='python')

    st.success(f"File terbaca: {len(df)} faktur")

    def auto_find(keywords):
        for c in df.columns:
            for k in keywords:
                if k.lower() in str(c).lower():
                    return c
        return df.columns[0]

    c1, c2 = st.columns(2)
    col_no = c1.selectbox("Kolom Nomor Faktur", df.columns, index=list(df.columns).index(auto_find(['nomor','faktur'])))
    col_tgl = c2.selectbox("Kolom Tanggal Faktur", df.columns, index=list(df.columns).index(auto_find(['tanggal','tgl'])))

    df[col_tgl] = pd.to_datetime(df[col_tgl], errors='coerce', dayfirst=True)
    df_clean = df.dropna(subset=[col_tgl]).copy()
    
    masa_tujuan = datetime.strptime(masa_tujuan_str, "%Y-%m")

    def cek_legal(tgl):
        if tgl.year >= 2025:
            boleh = (tgl.year == masa_tujuan.year and tgl.month == masa_tujuan.month)
            return boleh, "Baru - Masa Sama", f"{masa_tujuan_str}" if boleh else f"Harus di {tgl.strftime('%Y-%m')}"
        else:
            batas = (tgl + relativedelta(months=3)).replace(day=1)
            tujuan = masa_tujuan.replace(day=1)
            asal = tgl.replace(day=1)
            boleh = asal <= tujuan <= batas
            return boleh, "Lama - +3 Bulan", f"Batas {batas.strftime('%Y-%m')}"

    df_clean[['Boleh','Aturan','Catatan']] = df_clean[col_tgl].apply(lambda x: pd.Series(cek_legal(x)))
    
    df_boleh = df_clean[df_clean['Boleh']==True]
    df_tidak = df_clean[df_clean['Boleh']==False]

    m1, m2, m3 = st.columns(3)
    m1.metric("Total CSV", len(df_clean))
    m2.metric("BOLEH Kredit", len(df_boleh))
    m3.metric("TIDAK BOLEH", len(df_tidak))

    tab1, tab2, tab3 = st.tabs(["✅ BOLEH Kredit", "❌ TIDAK BOLEH", "🗂️ Bagi Per Masa"])

    with tab1:
        st.dataframe(df_boleh[[col_no, col_tgl, 'Aturan','Catatan']], use_container_width=True)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_boleh.to_excel(writer, index=False, sheet_name='BOLEH_KREDIT')
            df_tidak.to_excel(writer, index=False, sheet_name='TIDAK_BOLEH')
        st.download_button("📥 Download Excel Gabungan", output.getvalue(), f"Kredit_Gabungan_{masa_tujuan_str}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        txt = "\n".join(df_boleh[col_no].astype(str).tolist())
        st.download_button("📄 Download TXT Nomor Faktur", txt, f"nomor_boleh_{masa_tujuan_str}.txt")

    with tab2:
        st.dataframe(df_tidak[[col_no, col_tgl, 'Aturan','Catatan']], use_container_width=True)

    with tab3:
        df_boleh['Masa Faktur'] = df_boleh[col_tgl].dt.strftime("%Y-%m")
        for mf in sorted(df_boleh['Masa Faktur'].unique()):
            sub = df_boleh[df_boleh['Masa Faktur']==mf]
            with st.expander(f"📦 Masa Faktur {mf} - {len(sub)} faktur"):
                st.dataframe(sub[[col_no, col_tgl]], use_container_width=True)

    st.warning("⚠️ Cara pakai di Coretax: Filter per Masa Faktur sesuai kelompok di atas, lalu Kreditkan manual. Aplikasi ini hanya bantu gabungkan & filter biar legal.")
else:
    st.info("👆 Upload file dulu untuk mulai")