SGCL - BAN DA SUA

1. Giai nen thu muc nay.
2. Dam bao app.py va TemplateNSW.xlsx nam cung thu muc.
3. Cai dat thu vien Python:
   pip install -r requirements.txt
4. Cai Tesseract OCR tren he dieu hanh.
   - Streamlit Community Cloud: giu file packages.txt trong cung repository.
   - May Windows: cai Tesseract OCR va them thu muc Tesseract vao PATH.
5. Chay:
   streamlit run app.py

Luu y:
- OCR la ket qua nhan dang de nguoi dung kiem tra/chinh sua, khong phai du lieu da duoc xac minh.
- Neu OCR khong chac chan hoac khong tim thay truong, de trong va kiem tra thu cong.
- Buoc 4 co luoi xem truoc Excel va ho tro click mot o de gan cho truong dang chon (can Streamlit >= 1.35).
- File TemplateNSW.xlsx goc duoc dung lam mau; buoc xuat tao workbook rieng.
- Nen sao luu thu muc data/ va sgcl_field_config.json truoc khi thay app cu.
