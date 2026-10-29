import os
print("📦 প্রয়োজনীয় প্যাকেজ ও লাইব্রেরি কনফিগার করা হচ্ছে...")

# প্রয়োজনীয় লাইব্রেরি ইনস্টল করা (python-docx যুক্ত করা হয়েছে)
!pip install -q openpyxl pandas streamlit python-docx

with open("app.py", "w", encoding="utf-8") as f:
    f.write("""import streamlit as st
import pandas as pd
import random
import io
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="আমন ধান সংগ্রহ লটারি অ্যাপ-২০২৬", page_icon="🌾", layout="wide")
st.title("🌾 আমন ধান সংগ্রহ কার্যক্রম-২০২৬")
st.subheader("🎯 ইউনিয়ন ভিত্তিক কৃষক নির্বাচন লটারি অ্যাপ")

uploaded_file = st.file_uploader("কৃষকদের ডাটাবেজ (Excel File) আপলোড করুন", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("✅ ডাটাবেজ সফলভাবে আপলোড হয়েছে!")
    df.columns = df.columns.str.strip()
    
    union_col = None
    possible_names = ['ইউনিয়ন', 'ইউনিয়ন', 'ইউনিয়নের নাম', 'ইউনিয়নেরনাম', 'union', 'union name', 'ইউনি', 'ইউন']
    for col in df.columns:
        col_cleaned = str(col).strip().lower()
        if any(name in col_cleaned for name in possible_names):
            union_col = col
            break
            
    if union_col is None:
        st.error("❌ আপলোডকৃত এক্সেল ফাইলে ইউনিয়ন সংক্রান্ত কোনো কলাম পাওয়া যায়নি!")
    else:
        union_list = sorted(df[union_col].dropna().unique())
        st.sidebar.header("📍 লটারি ফিল্টার ও সেटिंग")
        selected_union = st.sidebar.selectbox("লটারির জন্য ইউনিয়ন নির্বাচন করুন", union_list)
        df_union = df[df[union_col] == selected_union].reset_index(drop=True)
        total_applicants = len(df_union)
        
        st.info(f"📍 **নির্বাচিত ইউনিয়ন:** {selected_union} | 📝 **মোট আবেদনকারী কৃষক:** {total_applicants} জন")
        
        with st.expander(f"👀 {selected_union} ইউনিয়নের আবেদনকারীদের তালিকা"):
            st.dataframe(df_union)
            
        winner_count = st.sidebar.number_input(
            f"কতজন বিজয়ী কৃষক নির্বাচন করবেন? (সর্বোচ্চ {total_applicants} জন)", 
            min_value=1, 
            max_value=max(1, total_applicants), 
            value=min(5, max(1, total_applicants))
        )
        
        if 'lottery_done' not in st.session_state:
            st.session_state['lottery_done'] = False
            st.session_state['df_winners'] = None
            st.session_state['current_union'] = ""
            st.session_state['total_applicants'] = 0
            st.session_state['winner_count'] = 0

        if st.button(f"🎰 {selected_union} ইউনিয়নের লটারি শুরু করুন"):
            if total_applicants == 0:
                st.error("এই ইউনিয়নে কোনো আবেদনকারী নেই!")
            else:
                with st.spinner("🎰 লটারির ড্র চলছে..."):
                    time.sleep(1.2)
                winner_indices = random.sample(range(total_applicants), winner_count)
                df_winners = df_union.iloc[winner_indices].reset_index(drop=True)
                
                st.session_state['df_winners'] = df_winners
                st.session_state['current_union'] = selected_union
                st.session_state['total_applicants'] = total_applicants
                st.session_state['winner_count'] = winner_count
                st.session_state['lottery_done'] = True
                
        if st.session_state['lottery_done'] and st.session_state['current_union'] == selected_union:
            df_winners = st.session_state['df_winners']
            st.balloons()
            st.subheader(f"🎉 লটারিতে নির্বাচিত {selected_union} ইউনিয়নের বিজয়ীদের তালিকা")
            st.dataframe(df_winners)
            
            # ১. এক্সেল ফাইল জেনারেটর (Buffer)
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_winners.to_excel(writer, index=False, sheet_name='বিজয়ী তালিকা')
            excel_data = excel_buffer.getvalue()
            
            # ২. ওয়ার্ড ফাইল (.docx) জেনারেটর ফাংশন
            def generate_word_file(winners_dataframe, u_name, total_app, win_num):
                doc = Document()
                
                # মার্জিন সেট করা (Landscape স্টাইল দিতে চাইলে ইঞ্চি বাড়িয়ে দিতে পারেন)
                sections = doc.sections
                for section in sections:
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                
                # শিরোনাম
                p1 = doc.add_paragraph()
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run1 = p1.add_run("আমন ধান সংগ্রহ কার্যক্রম-২০২৬\\n")
                run1.font.size = Pt(18)
                run1.font.bold = True
                
                run2 = p1.add_run("লটারির মাধ্যমে নির্বাচিত কৃষক বিজয়ীদের অফিসিয়াল তালিকা\\n")
                run2.font.size = Pt(14)
                run2.font.bold = True
                
                run3 = p1.add_run(f"ইউনিয়ন: {u_name} | মোট আবেদনকারী: {total_app} জন | নির্বাচিত বিজয়ী: {win_num} জন")
                run3.font.size = Pt(11)
                
                # টেবিল তৈরি (কলাম সংখ্যা = ১টি ক্রমিঙ্কের জন্য + মূল ডেটার কলাম)
                cols_count = len(winners_dataframe.columns) + 1
                table = doc.add_table(rows=1, cols=cols_count)
                table.style = 'Table Grid'
                
                # টেবিল হেডার
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'ক্রমিক'
                hdr_cells[0].paragraphs[0].runs[0].font.bold = True
                
                for i, col_name in enumerate(winners_dataframe.columns):
                    hdr_cells[i+1].text = str(col_name)
                    hdr_cells[i+1].paragraphs[0].runs[0].font.bold = True
                
                # টেবিল ডাটা যোগ করা
                for idx, row in winners_dataframe.iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(idx + 1) # ক্রমিক সংখ্যা
                    for i, col_name in enumerate(winners_dataframe.columns):
                        val = row[col_name]
                        row_cells[i+1].text = str(val) if pd.notna(val) else ""
                
                # ফুটার সিগনেচার বক্স
                doc.add_paragraph("\\n\\n\\n")
                p_sig = doc.add_paragraph()
                p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # ট্যাব স্পেস দিয়ে অফিসারদের নাম সাজানো
                sig_text = "___________________               ___________________               ___________________\\n"
                sig_text += "উপজেলা খাদ্য নিয়ন্ত্রক             উপজেলা কৃষি অফিসার             উপজেলা নির্বাহী অফিসার"
                run_sig = p_sig.add_run(sig_text)
                run_sig.font.size = Pt(11)
                run_sig.font.bold = True
                
                # ফাইলটি মেমোরি বাফারে সেভ করা
                word_buffer = io.BytesIO()
                doc.save(word_buffer)
                return word_buffer.getvalue()

            # ডাউনলোড বাটন লেআউট
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label=f"🟢 {selected_union} এর এক্সেল ফাইল ডাউনলোড",
                    data=excel_data,
                    file_name=f"বিজয়ী_তালিকা_{selected_union}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            with col2:
                with st.spinner('ওয়ার্ড ফাইল তৈরি হচ্ছে...'):
                    word_data = generate_word_file(
                        df_winners, 
                        st.session_state['current_union'], 
                        st.session_state['total_applicants'], 
                        st.session_state['winner_count']
                    )
                
                st.download_button(
                    label=f"🔵 {selected_union} এর অফিসিয়াল ওয়ার্ড ফাইল ডাউনলোড",
                    data=word_data,
                    file_name=f"বিজয়ী_তালিকা_{selected_union}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            st.success("🎉 আপনার ফাইলগুলো ডাউনলোডের জন্য প্রস্তুত!")

else:
    st.info("ℹ️ অনুগ্রহ করে প্রথমে কৃষকদের মূল এক্সেল ফাইলটি আপলোড করুন। ফাইলে অবশ্যই ইউনিয়ন কলামটি থাকতে হবে।")
""")

print("✅ 'app.py' নতুন এক্সেল ও ওয়ার্ড ফিচার সহ জেনারেট হয়েছে!")
