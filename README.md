# MGMTNA Compare By Date 7.0 (web)

App đối chiếu Excel Lines (BC) vs Sheet1 (TP) theo ngày.

- Không đăng nhập / đăng xuất
- Không lưu file người dùng (xử lý xong là xóa)
- Hai chức năng: **So mã** và **So mã + Unit**

## Chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Cần **Python 3.14** (engine bundled là bytecode 3.14).

## Deploy miễn phí (Streamlit Community Cloud)

1. Đẩy repo này lên GitHub (public).
2. Mở [share.streamlit.io](https://share.streamlit.io) → đăng nhập bằng GitHub.
3. **Create app** → chọn repo / branch `main` / file `app.py`.
4. Advanced settings: Python **3.14** (bắt buộc).
5. Deploy. URL dạng `https://<tên>.streamlit.app`.

App public: người khác mở URL là dùng được, không cần tài khoản Streamlit.
App miễn phí có thể ngủ sau khoảng 12 giờ không ai vào; lần mở sau sẽ khởi động lại.
