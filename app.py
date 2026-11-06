import json
import os
import io
from datetime import datetime
from google import genai
import pandas as pd
from PIL import Image, ExifTags
import streamlit as st
import smtplib
from email import message
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage


# ==========================================
# 0. ตั้งค่าหน้าเว็บและการรองรับ HEIC
# ==========================================
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except ImportError:
    pass

st.set_page_config(
    page_title="ระบบลงสินค้า DN ver.001",
    page_icon="📦",
    layout="centered"  # เหมาะสำหรับการเปิดใช้งานผ่านมือถือ
)

# CSS ตกแต่งให้ปุ่มและฟอร์มดูใช้งานง่ายบนมือถือ
st.markdown("""
    <style>
    .stApp { background-color: #F8FAFC !important; }
    h1, h2, h3 { color: #0F172A !important; }
    .stButton > button {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        font-weight: 600;
        width: 100%;
    }
    .stFormSubmitButton > button {
        background-color: #10B981 !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        font-weight: 700;
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. ตั้งค่า Gemini Client
# ==========================================
# ระบบจะดึงค่าจาก Environment Variable อัตโนมัติ หรือใส่คีย์ตรงนี้ได้เช่นกัน
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except Exception:
    # กรณีรันบนคอมพิวเตอร์ส่วนตัว (Local) หากไม่ได้ตั้งค่า secrets ไว้ จะดึงจาก Environment หรือใส่คีย์สำรองตรงนี้ได้
    GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "ใส่_API_Key_ของคุณที่นี่")

# สร้าง Client ของ Google GenAI แบบใหม่
if GEMINI_API_KEY and GEMINI_API_KEY != "GEMINI_API_KEY":
    client = genai.Client(api_key=GEMINI_API_KEY)
else:
    client = genai.Client()

# ==========================================
# 2 วางฟังก์ชันส่งอีเมลไว้ตรงนี้ (ระดับบนสุดของสคริปต์)
# ==========================================
def send_email_with_excel(excel_bytes, filename, supplier_name):
  try:
    sender_email = st.secrets["EMAIL_SENDER"]
    sender_password = st.secrets["EMAIL_PASSWORD"]
    receiver_email = st.secrets["EMAIL_RECEIVER"]

    msg = MIMEMultipart()
    msg["Subject"] = f"[รายงานสต็อก] บิลรับสินค้าจาก {supplier_name}"
    msg["From"] = sender_email
    msg["To"] = receiver_email

    msg.attach(
        MIMEText(
            "สวัสดีครับคุณนิคกี้,\n\nแนบไฟล์ Excel รายการสินค้าจากบิลรับสินค้า"
            f" {supplier_name} มาพร้อมกับอีเมลฉบับนี้ครับ",
            "plain",
        )
    )

    part = MIMEApplication(excel_bytes, Name=filename)
    part["Content-Disposition"] = f'attachment; filename="{filename}"'
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
      server.login(sender_email, sender_password)
      server.sendmail(sender_email, receiver_email, msg.as_string())

    return True
  except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการส่งอีเมล: {e}")
    return False


# ==========================================
# 3. Session State Initialization
# ==========================================
if "step" not in st.session_state:
    st.session_state.step = 1  # 1: อัปโหลดบิลและสกัดตาราง, 2: ถ่ายรูปฉลากทีละชิ้น, 3: ตรวจสอบและบันทึก
if "df_items" not in st.session_state:
    st.session_state.df_items = pd.DataFrame(columns=["ลำดับ", "ชื่อสินค้า", "จำนวน", "ราคา", "Lot", "MFG", "EXP"])
if "current_item_idx" not in st.session_state:
    st.session_state.current_item_idx = 0
if "supplier_name" not in st.session_state:
    st.session_state.supplier_name = ""
if "invoice_no" not in st.session_state:
    st.session_state.invoice_no = ""
if "invoice_date" not in st.session_state:
    st.session_state.invoice_date = datetime.now().date()

# สร้างที่เก็บรูปภาพชั่วคราวใน session_state ถ้ายังไม่มี
if "item_images" not in st.session_state:
    st.session_state.item_images = {}

st.title("📦 ระบบลงสินค้า DN ver.001")
st.markdown("---")



# ==========================================
# 4. ถ่ายรูปบิล & AI สกัดรายการสินค้า (รองรับหลายหน้า)
# ==========================================
if st.session_state.step == 1:
  st.subheader("ขั้นตอนที่ 1: ถ่ายรูปบิลรับสินค้า (รองรับบิลหลายหน้า)")

  st.session_state.supplier_name = st.text_input(
      "ชื่อบริษัท / ผู้จัดจำหน่าย (Supplier)",
      value=st.session_state.supplier_name,
  )
  st.session_state.invoice_no = st.text_input(
      "เลขที่บิล / ใบกำกับภาษี", value=st.session_state.invoice_no
  )
  st.session_state.invoice_date = st.date_input(
      "วันที่ในบิล", value=st.session_state.invoice_date
  )

  # 📌 ปรับให้รับไฟล์ได้หลายไฟล์พร้อมกัน (accept_multiple_files=True)
  bill_images = st.file_uploader(
      "📸 ถ่ายรูปหรืออัปโหลดใบรับสินค้า (เลือกได้หลายหน้าพร้อมกัน)",
      type=["png", "jpg", "jpeg", "heic"],
      accept_multiple_files=True,
  )

  if bill_images:
    st.write(f"📁 อัปโหลด/เลือกบิลมาแล้วทั้งหมด **{len(bill_images)} หน้า**")
    # แสดงรูปภาพตัวอย่างทุกหน้าที่อัปโหลด
    for i, img_file in enumerate(bill_images):
      st.image(
          img_file,
          caption=f"ใบรับสินค้า หน้าที่ {i+1}",
          use_container_width=True,
      )

  # เงื่อนไขปุ่มกดทำงาน จะเช็คจาก list ของ bill_images ว่ามีข้อมูลหรือไม่
  if st.button(
      "🤖 ให้ AI สร้างตารางสินค้าจากบิลทุกหน้า",
      disabled=not bill_images or not st.session_state.supplier_name,
  ):
    with st.spinner(
        "AI กำลังอ่านข้อมูลบิลทุกหน้าและรวบรวมสร้างตารางสินค้า..."
    ):
      try:
        # 1. เปิดไฟล์รูปภาพทุกหน้าเก็บไว้ใน List
        img_list = []
        for img_file in bill_images:
          img_list.append(Image.open(img_file))

        # 2. ปรับ Prompt ให้ AI รองรับหลายภาพและรวมข้อมูลเป็นชุดเดียว
        prompt = """คุณคือ AI ช่วยอ่านบิลสินค้า ซึ่งอาจมีหลายหน้า จงอ่านภาพบิลทั้งหมดที่แนบมานี้ แล้วสกัดรายการสินค้าออกมาเป็น JSON เท่านั้น โดยมีโครงสร้างดังนี้:
                {
                  "items": [
                    {
                      "name": "ชื่อสินค้า",
                      "qty": จำนวน (ตัวเลข),
                      "price": ราคาต่อหน่วย (ตัวเลข)
                    }
                  ]
                }
                *หมายเหตุ: หากมีหลายหน้า ให้รวบรวมรายการสินค้าทั้งหมดจากทุกหน้าเข้าด้วยกันเป็น List เดียว*"""

        # 3. ส่งข้อมูลภาพทั้งหมดพร้อม Prompt เข้าไปยัง Gemini
        response = client.models.generate_content(
            model="gemini-3.5-flash-lite", contents=[*img_list, prompt]
        )
        text_out = response.text.strip()

        # 4. ใช้ Regex ค้นหาและดึงเฉพาะ JSON ออกมาป้องกัน Error ข้อความขยะ
        import re

        json_match = re.search(r"\{.*\}", text_out, re.DOTALL)
        if json_match:
          clean_json_str = json_match.group(0)
          data = json.loads(clean_json_str)
        else:
          data = json.loads(text_out)

        items = data.get("items", [])

        rows = []
        for idx, item in enumerate(items, start=1):
          rows.append({
              "ลำดับ": idx,
              "ชื่อสินค้า": item.get("name", ""),
              "จำนวน": item.get("qty", 1),
              "ราคา": item.get("price", 0.0),
              "Lot": "-",
              "MFG": "-",
              "EXP": "-",
          })

        st.session_state.df_items = pd.DataFrame(rows)
        st.session_state.current_item_idx = 0
        st.session_state.step = 2
        st.rerun()

      except Exception as e:
        st.error(
            f"เกิดข้อผิดพลาดในการอ่านบิลหลายหน้า: {e} | ข้อมูลที่ตอบกลับ:"
            f" {text_out}"
        )





# ==========================================
# 5. ถ่ายรูปฉลากสินค้าทีละรายการ (Item-by-Item)
# ==========================================
elif st.session_state.step == 2:
    df = st.session_state.df_items
    idx = st.session_state.current_item_idx
    total_items = len(df)
    
    st.subheader(f"ขั้นตอนที่ 2: ถ่ายรูปฉลากสินค้า ({idx + 1} จาก {total_items})")
    
    if idx < total_items:
        current_row = df.iloc[idx]
        
        st.info(f"👉 **กำลังสแกนสินค้า:**\n\n📌 **{current_row['ชื่อสินค้า']}**\n- จำนวน: {current_row['จำนวน']} | ราคา: {current_row['ราคา']}")

        st.markdown(
            """
            <style>
            /* ขยายความกว้างของกล่อง camera_input ให้เต็มพื้นที่ */
            [data-testid="stCameraInput"] {
                max-width: 100% !important;
                overflow: hidden !important; /* ซ่อนส่วนเกินจากการซูม */
            }
            /* ซูมภาพวิดีโอในกล้องเข้าไป 30% (scale 1.3) */
            [data-testid="stCameraInput"] video {
                width: 100% !important;
                height: auto !important;
                transform: scale(1.3);
                transform-origin: center center;
            }
            
            /* ปรับแต่งปุ่มถ่ายรูป (Take Photo Button) ให้เป็นสีเหลือง */
            [data-testid="stCameraInput"] button[kind="secondary"], 
            [data-testid="stCameraInput"] button {
                background-color: #FFC107 !important; /* สีเหลือง */
                color: #000000 !important;          /* สีของข้อความ/ไอคอนด้านในเป็นสีดำเพื่อให้ตัดกันและมองเห็นชัดเจน */
                border: none !important;
            }
            
            /* เอฟเฟกต์ตอนเอาเมาส์ชี้หรือกดปุ่ม (Hover/Active) */
            [data-testid="stCameraInput"] button:hover {
                background-color: #FFB300 !important;
                color: #000000 !important;
            }
            
            </style>
            """,
            unsafe_allow_html=True,
        )

        
        # ใช้ camera_input เพื่อให้พนักงานกดถ่ายรูปจากมือถือได้ทันที
        label_image = st.camera_input(f"ถ่ายรูปฉลากข้างกล่อง: {current_row['ชื่อสินค้า']}", key=f"cam_{idx}")
        
        if label_image:
            # 📌 [จุดที่เพิ่มเข้ามา]: บันทึกรูปภาพของสินค้ารายการนี้เก็บไว้ใน session_state โดยอ้างอิงตาม index (idx)
            st.session_state.item_images[idx] = label_image.getvalue()

            with st.spinner("AI กำลังอ่าน Lot, MFG, และ EXP จากฉลาก..."):
                try:
                    img = Image.open(label_image)
                    prompt ="""
                    คุณคือผู้ช่วยจัดการร้านขายยา หน้าที่ของคุณคืออ่านภาพบิลหรือฉลากสินค้า แล้วสกัดข้อมูลออกมาเป็นรูปแบบวันที่ DD/MM/YYYY
                    
                    เงื่อนไขสำคัญสำหรับการแปลงวันที่ในคอลัมน์ MFG และ EXP:
                    1. ต้องแสดงผลในรูปแบบ "DD/MM/YYYY" เสมอ (เช่น 30/07/2026)
                    2. หากฉลากระบุปีเป็น พ.ศ. (เช่น 2569, 2570) ให้แปลงเป็น ค.ศ. (ลบด้วย 543)
                    3. **กรณีที่ฉลากระบุมาเฉพาะ "เดือนและปี" เท่านั้น (เช่น "JUL 26", "07/2026", "ค.ค. 69") ให้แปลงดังนี้:**
                       - **สำหรับวันหมดอายุ (EXP):** ให้ปัดเป็น **วันสุดท้ายของเดือนนั้นๆ** (เช่น ถ้าเป็น เดือนกรกฎาคม 2026 / JUL 26 ให้บันทึกเป็นวันที่สิ้นเดือนคือ **30/07/2026** หรือ 31 ตามจริงของเดือนนั้น)
                       - **สำหรับวันผลิต (MFG):** ให้ปัดเป็น **วันแรกของเดือนนั้นๆ** หรือวันสุดท้ายของเดือนก็ได้ตามความเหมาะสม (เช่น **01/07/2026** หรือวันสิ้นเดือน) เพื่อให้ได้ฟอร์แมต DD/MM/YYYY ที่สมบูรณ์ครบถ้วน

                    รูปแบบ JSON ที่ต้องการตัวอย่างเช่น:
                    {"lot": "ABC1234", "mfg": "01/07/2025", "exp": "30/06/2028"}
                    
                    """ 
                   
                    response = client.models.generate_content(
                        model="gemini-3.5-flash-lite", # หรือโมเดลที่คุณนิคกี้ใช้งานอยู่
                        contents=[img, prompt]
                    )
                    text_out = response.text.strip()
                    if text_out.startswith("```json"):
                        text_out = text_out[7:-3]
                    
                    label_data = json.loads(text_out.strip())
                    
                    # อัปเดตข้อมูลลงใน DataFrame ของชิ้นปัจจุบัน
                    st.session_state.df_items.at[idx, "Lot"] = label_data.get("lot", "-")
                    st.session_state.df_items.at[idx, "MFG"] = label_data.get("mfg", "-")
                    st.session_state.df_items.at[idx, "EXP"] = label_data.get("exp", "-")
                    
                    st.success("✅ อ่านข้อมูลจากฉลากสำเร็จ!")
                except Exception as e:
                    st.error(f"อ่านฉลากไม่สำเร็จ: {e}")
                
        # แสดงค่าที่บันทึกแล้วของสินค้านี้
        st.markdown(f"**ข้อมูลปัจจุบัน:** Lot: `{current_row['Lot']}` | MFG: `{current_row['MFG']}` | EXP: `{current_row['EXP']}`")
        
        col_prev, col_next = st.columns(2)
        with col_prev:
            if idx > 0 and st.button("⬅️ ย้อนกลับ"):
                st.session_state.current_item_idx -= 1
                st.rerun()
        with col_next:
            if st.button("➡️ ถัดไป / บันทึกชิ้นนี้"):
                if idx + 1 < total_items:
                    st.session_state.current_item_idx += 1
                    st.rerun()
                else:
                    st.session_state.step = 3
                    st.rerun()
                    
        if st.button("⏭️ ข้ามรายการนี้ไปก่อน", type="secondary"):
            if idx + 1 < total_items:
                st.session_state.current_item_idx += 1
                st.rerun()
            else:
                st.session_state.step = 3
                st.rerun()





# ==========================================
# 6. ตรวจสอบข้อมูลทั้งหมดและบันทึก Excel
# ==========================================
elif st.session_state.step == 3:
  st.subheader("ขั้นตอนที่ 3: ตรวจสอบข้อมูลและบันทึก")

  st.write(f"📌 **Supplier:** {st.session_state.supplier_name}")
  st.write(
      f"📌 **เลขที่บิล:** {st.session_state.invoice_no} | **วันที่:**"
      f" {st.session_state.invoice_date}"
  )

  # เปิดให้แก้ไขตารางขั้นสุดท้ายก่อนเซฟ
  edited_df = st.data_editor(st.session_state.df_items, use_container_width=True)
  st.session_state.df_items = edited_df

  col_back, col_save = st.columns(2)
  with col_back:
    if st.button("🔄 กลับไปถ่ายรูปใหม่"):
      st.session_state.step = 2
      st.session_state.current_item_idx = 0
      st.rerun()

  with col_save:
    if st.button("✅ ยืนยันและ ส่งข้อมูลเข้าอีเมล"):
      safe_inv = "".join(
          [c if c.isalnum() else "_" for c in str(st.session_state.invoice_no)]
      )
      file_name = f"DN_Receive_{safe_inv}.xlsx"

      try:
        # 1. สร้าง Excel DataFrame เป็น Bytes ปกิติก่อน
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
          edited_df.to_excel(writer, sheet_name="Receive_Items", index=False)
        excel_buffer.seek(0)

        # 📌 2. โหลดไฟล์ด้วย openpyxl เพื่อนำรูปภาพฉลากสินค้ามาแทรกในคอลัมน์ขวาสุด
        wb = openpyxl.load_workbook(excel_buffer)
        ws = wb["Receive_Items"]

        # หาคอลัมน์ขวาสุดสำหรับวางรูปฉลาก
        img_col_letter = openpyxl.utils.get_column_letter(ws.max_column + 1)
        ws[f"{img_col_letter}1"] = "รูปฉลากสินค้า"

        # 📌 [เพิ่มบรรทัดนี้]: ขยายความกว้างของคอลัมน์รูปภาพให้กว้างขึ้น (เช่น 25) เพื่อรับรูปขนาดใหญ่
        ws.column_dimensions[img_col_letter].width = 25
          
        # วนลูปหยิบรูปภาพที่เก็บไว้ใน session_state มาแปะตามแถวสินค้า
        for idx, row in edited_df.iterrows():
          row_num = idx + 2  # แถวที่ 1 เป็น Header หัวตาราง
          if idx in st.session_state.get("item_images", {}):
            img_bytes = st.session_state.item_images[idx]
            img_io = io.BytesIO(img_bytes)

            # สร้างออบเจกต์รูปภาพสำหรับ openpyxl
            img = OpenpyxlImage(img_io)
            img.width = 200  # กำหนดความกว้างรูป (พิกเซล)
            img.height = 150  # กำหนดความสูงรูป (พิกเซล)

            # แปะรูปลงในเซลล์คอลัมน์ขวาสุด เช่น H2, H3, H4...
            ws.add_image(img, f"{img_col_letter}{row_num}")
            ws.row_dimensions[row_num].height = (
                300  # ขยายความสูงแถวให้พอดีกับรูป
            )

        # 3. บันทึกไฟล์ Excel สมบูรณ์พร้อมรูปภาพลง BytesIO เพื่อเตรียมส่งอีเมล
        final_output = io.BytesIO()
        wb.save(final_output)
        excel_bytes = final_output.getvalue()

        # 4. บันทึกไฟล์ลงเครื่อง (เฉพาะกรณีรัน Local ที่มีโฟลเดอร์รองรับ)
        target_dir = r"G:\ร้านยา\web app\excel file from extract"
        try:
          os.makedirs(target_dir, exist_ok=True)
          file_path = os.path.join(target_dir, file_name)
          with open(file_path, "wb") as f:
            f.write(excel_bytes)
          local_saved = True
        except Exception:
          local_saved = (
              False  # ถ้ารันบน Cloud แล้วไม่มี path นี้ ข้ามไปไม่ต้อง Error
          )

        # 5. ส่งอีเมลพร้อมแนบไฟล์ Excel (ที่มีรูปภาพแนบแล้ว) อัตโนมัติ
        email_success = send_email_with_excel(
            excel_bytes, file_name, st.session_state.supplier_name
        )

        # 6. แสดงผลลัพธ์ให้ผู้ใช้ทราบ
        if local_saved:
          st.success(
              f"🎉 บันทึกไฟล์และส่งอีเมลสำเร็จ!\n\n📂 เซฟไว้ที่:"
              f" `{file_path}`"
          )
        else:
          st.success(
              "🎉 ประมวลผลและส่งอีเมลรายงานเข้าสู่ระบบพร้อมรูปภาพฉลากสำเร็จเรียบร้อยแล้ว!"
          )

        if email_success:
          st.info("✉️ ส่งไฟล์รายงานเข้าอีเมลเรียบร้อยครับ")

      except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการบันทึกหรือส่งอีเมล: {e}")


    # ==========================================
    # ส่วนเว้นระยะห่าง และปุ่มเริ่มทำรายการใหม่ (สีเขียว)
    # ==========================================
    
    # เคาะเว้นบรรทัดลงไปประมาณ 3 บรรทัด เพื่อป้องกันการกดโดนโดยไม่ได้ตั้งใจ
    st.markdown("<br><br><br>", unsafe_allow_html=True)

    st.markdown(
    """
    <style>
    /* ดักจับปุ่มโดยอ้างอิงจากคีย์ที่เรากำหนด หรือปุ่มที่เป็นกว้างเต็มจอ */
    div.stButton > button[kind="secondary"] {
        /* ปรับแต่งปุ่มที่เป็นประเภท secondary (ถ้าต้องการให้ปุ่มนี้เป็นสีเขียวเด่นชัด) */
    }
    
    /* หรือใช้วิธีเจาะจงหาปุ่มที่มีคำว่า "เริ่มทำรายการใหม่" โดยตัดอีโมจิออกเพื่อความชัวร์ */
    div.stButton > button:has(p:contains("เริ่มทำรายการใหม่")) {
        background-color: #28A745 !important; /* สีเขียว */
        color: #FFFFFF !important;            /* ตัวหนังสือสีขาว */
        border: none !important;
        width: 100% !important;               /* ขยายให้ปุ่มยาวเต็มตา */
    }
    div.stButton > button:has(p:contains("เริ่มทำรายการใหม่")):hover {
        background-color: #218838 !important; /* สีเขียวเข้มตอนเอาเมาส์ชี้ */
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
    )
    
    # ปุ่มสำหรับเริ่มทำรายการใหม่ (เพิ่ม key="btn_reset_new" เอาไว้เผื่ออ้างอิง)
    if st.button("🔄 เริ่มทำรายการใหม่ (New Bill)", key="btn_reset_new"):
      st.session_state.step = 1
      st.session_state.df_items = pd.DataFrame(
          columns=["ลำดับ", "ชื่อสินค้า", "จำนวน", "ราคา", "Lot", "MFG", "EXP"]
      )
      st.session_state.current_item_idx = 0
      st.session_state.invoice_no = ""
      st.session_state.item_images = {}  # ล้างรูปภาพเก่าของบิลเดิมออก
      st.rerun()
