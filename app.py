# =========================================================================
    # 4. HIỂN THỊ KẾT QUẢ & XUẤT FILE EXCEL (DẠNG SỔ CHI TIẾT THEO NGÀY BÁN)
    # =========================================================================
    st.divider()
    st.subheader(f"3. KẾT QUẢ PHÂN BỔ (CHI TIẾT THEO NGÀY BÁN)")

    tong_tien_thuc_te = int(np.sum(df_items["THÀNH TIỀN"]))
    chenh_lech = tong_tien_thuc_te - int(tong_tien_muc_tieu)

    m1, m2, m3 = st.columns(3)
    m1.metric("DOANH THU MỤC TIÊU", f"{tong_tien_muc_tieu:,.0f} đ")
    m2.metric("ĐÃ PHÂN BỔ THỰC TẾ", f"{tong_tien_thuc_te:,.0f} đ")
    m3.metric("CHÊNH LỆCH", f"{chenh_lech:,.0f} đ")

    # --- CHUYỂN ĐỔI MA TRẬN SANG DẠNG SỔ CHI TIẾT TỪNG DÒNG (NHƯ ẢNH MẪU) ---
    chi_tiet_records = []
    # Lấy mốc ngày bắt đầu tính từ hôm nay
    base_date = datetime.now().date()

    for d_idx, d_col in enumerate(danh_sach_ngay):
        # Tính chuỗi ngày/tháng/năm định dạng dd/mm/yyyy
        ngay_hd = (base_date + pd.Timedelta(days=d_idx)).strftime("%d/%m/%Y")
        
        for i in range(n):
            sl = matrix_ngay[i, d_idx]
            if sl > 0:
                gia = int(gia_arr[i])
                thanh_tien = int(sl * gia)
                chi_tiet_records.append({
                    "Ngày HĐ": ngay_hd,
                    "Mã hàng (*)": df_items.loc[i, "MÃ VTHH"],
                    "Tên hàng": df_items.loc[i, "TÊN VTHH"],
                    "ĐVT": df_items.loc[i, "ĐVT"],
                    "Số lượng": sl,
                    "Đơn giá": gia,
                    "Thành tiền": thanh_tien
                })

    df_so_chi_tiet = pd.DataFrame(chi_tiet_records)

    st.markdown("#### 📋 BẢNG KÊ CHI TIẾT BÁN HÀNG PHÂN BỔ")
    st.dataframe(
        df_so_chi_tiet,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ngày HĐ": st.column_config.TextColumn("Ngày HĐ"),
            "Mã hàng (*)": st.column_config.TextColumn("Mã hàng (*)"),
            "Tên hàng": st.column_config.TextColumn("Tên hàng"),
            "ĐVT": st.column_config.TextColumn("ĐVT"),
            "Số lượng": st.column_config.NumberColumn("Số lượng", format="%d"),
            "Đơn giá": st.column_config.NumberColumn("Đơn giá", format="%,d đ"),
            "Thành tiền": st.column_config.NumberColumn("Thành tiền", format="%,d đ"),
        }
    )

    # --- TẠO FILE EXCEL ĐỊNH DẠNG MÀU SẮC CHUẨN GIAO DIỆN MẪU ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_so_chi_tiet.to_excel(writer, sheet_name="CHI_TIET_BAN_HANG", index=False)
        ws = writer.sheets["CHI_TIET_BAN_HANG"]

        # 1. Định dạng Tiêu đề Header (Màu xanh rêu cho Ngày HĐ, Màu Cam cho các cột còn lại)
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)
        fill_green = PatternFill(start_color="4F7942", end_color="4F7942", fill_type="solid")  # Cột Ngày HĐ
        fill_orange = PatternFill(start_color="E26B00", end_color="E26B00", fill_type="solid") # Các cột còn lại

        ws.row_dimensions[1].height = 26
        for col_idx, cell in enumerate(ws[1], start=1):
            cell.font = font_header
            cell.alignment = align_header
            cell.fill = fill_green if col_idx == 1 else fill_orange

        # 2. Định dạng Căn lề & Số liệu từng dòng
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        data_font = Font(name="Arial", size=10)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            ws.row_dimensions[row[0].row].height = 20
            # Ngày HĐ
            row[0].alignment = align_center
            row[0].font = data_font
            # Mã hàng (*)
            row[1].alignment = align_center
            row[1].font = data_font
            # Tên hàng
            row[2].alignment = align_left
            row[2].font = data_font
            # ĐVT
            row[3].alignment = align_center
            row[3].font = data_font
            # Số lượng
            row[4].alignment = align_right
            row[4].number_format = '#,##0'
            row[4].font = data_font
            # Đơn giá
            row[5].alignment = align_right
            row[5].number_format = '#,##0'
            row[5].font = data_font
            # Thành tiền
            row[6].alignment = align_right
            row[6].number_format = '#,##0'
            row[6].font = data_font

        # 3. Dòng Tổng cộng cuối bảng
        last_r = ws.max_row + 1
        ws.cell(last_r, 1, "TỔNG CỘNG").font = Font(name="Arial", size=11, bold=True)
        ws.cell(last_r, 1).alignment = align_center
        ws.cell(last_r, 5, int(df_so_chi_tiet["Số lượng"].sum())).font = Font(name="Arial", size=11, bold=True)
        ws.cell(last_r, 5).number_format = '#,##0'
        ws.cell(last_r, 7, int(tong_tien_thuc_te)).font = Font(name="Arial", size=11, bold=True)
        ws.cell(last_r, 7).number_format = '#,##0'
        
        sum_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        for c in range(1, 8):
            ws.cell(last_r, c).fill = sum_fill

        # 4. Tự động căn chỉnh độ rộng cột
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

    st.download_button(
        label="📥 TẢI BẢNG KÊ CHI TIẾT (EXCEL)",
        data=output.getvalue(),
        file_name=f"BANG_KE_BAN_HANG_{so_ngay_int}_NGAY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
