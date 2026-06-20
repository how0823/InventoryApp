import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
from io import BytesIO
from supabase import create_client

DB_NAME = "inventory.db"
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

TABLE_NAME = "inventory"

RAW_SAFETY_STOCK = 300000

def fmt(x):
    return f"{int(x):,}"
RAW_BEGIN = 100000        # 100000
FINISHED_BEGIN = 100000   # 100000


def init_db():
    pass


def insert_data(type_, date_, order_no, qty, return_qty=0, defect_qty=0, memo=""):
    data = {
        "type": type_,
        "date": str(date_),
        "order_no": order_no,
        "qty": int(qty),
        "return_qty": int(return_qty),
        "defect_qty": int(defect_qty),
        "memo": memo
    }

    supabase.table(TABLE_NAME).insert(data).execute()


def load_data():
    response = (
        supabase
        .table(TABLE_NAME)
        .select("*")
        .order("id", desc=True)
        .execute()
    )

    df = pd.DataFrame(response.data)

    if df.empty:
        df = pd.DataFrame(columns=[
            "id", "type", "date", "order_no",
            "qty", "return_qty", "defect_qty", "memo"
        ])

    return df

def create_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            sheet_name="재고이력",
            index=False
        )

    return output.getvalue()

def update_data(id_, type_, date_, order_no, qty,
                return_qty=0, defect_qty=0, memo=""):
    data = {
        "type": type_,
        "date": str(date_),
        "order_no": order_no,
        "qty": int(qty),
        "return_qty": int(return_qty),
        "defect_qty": int(defect_qty),
        "memo": memo
    }

    supabase.table(TABLE_NAME).update(data).eq("id", int(id_)).execute()
    


def delete_data(id_):
    supabase.table(TABLE_NAME).delete().eq("id", int(id_)).execute()

init_db()

st.title("TBP 재고관리 APP")

menu = st.sidebar.radio(
    "메뉴 선택",
    ["발주 입력", "입고 입력", "생산 입력", "출하 입력", "발주현황", "입고현황", "재고현황", "입력이력 관리", "현황 조회"]
)

if menu == "발주 입력":
    st.header("발주 입력")

    order_no = st.text_input("발주NO")
    order_date = st.date_input("발주일자", date.today())

    st.subheader("입고요청 일정 입력")

    request_count = st.number_input(
        "입고요청 라인 수",
        min_value=1,
        max_value=10,
        value=1,
        step=1
    )

    request_lines = []

    for i in range(request_count):
        col0, col1, col2 = st.columns([1, 3, 3])

        with col0:
            st.markdown(f"**요청 {i+1}**")

        with col1:
            request_date = st.date_input(
                "입고요청일자",
                date.today(),
                key=f"request_date_{i}"
            )

        with col2:
            request_qty = st.number_input(
                "요청수량",
                min_value=0,
                step=1,
                key=f"request_qty_{i}"
            )
            st.caption(f"입력수량: {fmt(request_qty)} EA")

        request_lines.append((request_date, request_qty))

    total_request_qty = sum(qty for _, qty in request_lines)

    st.info(f"총 입고요청수량: {total_request_qty:,}")

    if st.button("발주 저장"):
        if order_no.strip() == "":
            st.warning("발주NO를 입력하세요.")
        elif total_request_qty <= 0:
            st.warning("입고요청수량을 입력하세요.")
        else:
            for request_date, request_qty in request_lines:
                if request_qty > 0:
                    insert_data(
                        "발주",
                        order_date,
                        order_no,
                        request_qty,
                        memo=f"입고요청일자: {request_date}"
                    )

            st.success(f"발주정보가 저장되었습니다. 총 {total_request_qty:,}개")

elif menu == "입고 입력":
    st.header("입고 입력")

    df = load_data()

    order_df = df[df["type"] == "발주"]

    if order_df.empty:
        st.info("등록된 발주정보가 없습니다. 먼저 발주를 입력하세요.")
    else:
        order_list = sorted(order_df["order_no"].dropna().unique())

        selected_order_no = st.selectbox(
            "발주NO 선택",
            order_list
        )

        st.subheader("해당 발주의 요청일정 자동조회")

        selected_order_df = order_df[order_df["order_no"] == selected_order_no].copy()

        receive_df = df[
            (df["type"] == "입고") &
            (df["order_no"] == selected_order_no)
        ].copy()

        total_received = receive_df["qty"].sum() - receive_df["return_qty"].sum()

        selected_order_df["입고요청일자"] = selected_order_df["memo"].str.replace(
            "입고요청일자: ",
            "",
            regex=False
        )

        schedule_df = selected_order_df[["입고요청일자", "qty"]].copy()
        schedule_df["입고요청일자"] = pd.to_datetime(schedule_df["입고요청일자"])
        schedule_df = schedule_df.sort_values("입고요청일자")

        schedule_df = schedule_df.rename(columns={
            "qty": "요청수량"
        })

        schedule_df["입고누계"] = 0

        remaining_received = total_received

        for idx in schedule_df.index:
            request_qty = schedule_df.loc[idx, "요청수량"]

            if remaining_received >= request_qty:
                schedule_df.loc[idx, "입고누계"] = request_qty
                remaining_received -= request_qty
            else:
                schedule_df.loc[idx, "입고누계"] = remaining_received
                remaining_received = 0

        st.write(f"**{selected_order_no}**")

        st.dataframe(
            schedule_df,
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        st.subheader("이번 입고수량 입력")

        receive_date = st.date_input("입고일자", date.today())
        receive_qty = st.number_input("입고수량", min_value=0, step=1)
        st.caption(f"입고수량: {fmt(receive_qty)} EA")
        return_qty = st.number_input("입고반품수량", min_value=0, step=1)
        st.caption(f"반품수량: {fmt(return_qty)} EA")

        if st.button("입고 저장"):
            if receive_qty <= 0 and return_qty <= 0:
                st.warning("입고수량 또는 입고반품수량을 입력하세요.")
            else:
                insert_data(
                    "입고",
                    receive_date,
                    selected_order_no,
                    receive_qty,
                    return_qty=return_qty
                )
                st.success("입고정보가 저장되었습니다.")

elif menu == "생산 입력":
    st.header("생산 입력")

    prod_date = st.date_input("생산날짜", date.today())
    qty = st.number_input("양품수량", min_value=0, step=1)
    st.caption(f"양품수량: {fmt(qty)} EA")
    defect_qty = st.number_input("불량수량", min_value=0, step=1)
    st.caption(f"불량수량: {fmt(defect_qty)} EA")

    if st.button("생산 저장"):
        insert_data("생산", prod_date, "", qty, defect_qty=defect_qty)
        st.success("생산정보가 저장되었습니다.")

elif menu == "출하 입력":
    st.header("출하 입력")

    ship_date = st.date_input("출하날짜", date.today())
    qty = st.number_input("출하수량", min_value=0, step=1)
    st.caption(f"출하수량: {fmt(qty)} EA")

    if st.button("출하 저장"):
        insert_data("출하", ship_date, "", qty)
        st.success("출하정보가 저장되었습니다.")

elif menu == "발주현황":
    st.header("발주현황")

    df = load_data()

    if df.empty or df[df["type"] == "발주"].empty:
        st.info("등록된 발주정보가 없습니다.")
    else:
        order_df = df[df["type"] == "발주"].copy()
        receive_df = df[df["type"] == "입고"].copy()

        order_df["입고요청일자"] = order_df["memo"].str.replace(
            "입고요청일자: ",
            "",
            regex=False
        )

        order_df["입고요청일자"] = pd.to_datetime(order_df["입고요청일자"])

        result_rows = []

        for order_no in order_df["order_no"].unique():
            orders = order_df[order_df["order_no"] == order_no].sort_values("입고요청일자")
            receives = receive_df[receive_df["order_no"] == order_no]

            total_received = receives["qty"].sum() - receives["return_qty"].sum()
            remaining_received = total_received

            for _, row in orders.iterrows():
                request_qty = row["qty"]

                if remaining_received >= request_qty:
                    received_qty = request_qty
                    remaining_received -= request_qty
                else:
                    received_qty = remaining_received
                    remaining_received = 0

                not_received = request_qty - received_qty

                delay_days = 0
                if not_received > 0 and pd.Timestamp.today().normalize() > row["입고요청일자"]:
                    delay_days = (pd.Timestamp.today().normalize() - row["입고요청일자"]).days

                result_rows.append({
                    "발주NO": order_no,
                    "입고요청일자": row["입고요청일자"].date(),
                    "요청수량": request_qty,
                    "입고누계": received_qty,
                    "미입고": not_received,
                    "지연일수": delay_days
                })

        result_df = pd.DataFrame(result_rows)

        st.subheader("발주 요청 대비 입고 현황")

        display_df = result_df.copy()
        for col in ["요청수량", "입고누계", "미입고"]:
            display_df[col] = display_df[col].apply(fmt)

        display_df["지연일수"] = display_df["지연일수"].apply(
            lambda x: "-" if x == 0 else f"{x}일"
        )

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

        delayed_count = (result_df["지연일수"] > 0).sum()
        delayed_qty = result_df[result_df["지연일수"] > 0]["미입고"].sum()

        if delayed_count > 0:
            st.error(f"⚠ 입고 지연 {delayed_count}건 / 지연수량 {fmt(delayed_qty)} EA")
        else:
            st.success("입고 지연 건이 없습니다.")

elif menu == "입고현황":
    st.header("입고현황")

    df = load_data()

    if df.empty:
        st.info("입고현황을 표시할 데이터가 없습니다.")
    else:
        order_df = df[df["type"] == "발주"].copy()
        receive_df = df[df["type"] == "입고"].copy()

        if order_df.empty:
            st.info("등록된 발주정보가 없습니다.")
        else:
            order_summary = order_df.groupby("order_no", as_index=False)["qty"].sum()
            order_summary = order_summary.rename(columns={"qty": "총발주량"})

            if receive_df.empty:
                receive_summary = pd.DataFrame(columns=["order_no", "누적입고량"])
            else:
                receive_df["실입고량"] = receive_df["qty"] - receive_df["return_qty"]
                receive_summary = receive_df.groupby("order_no", as_index=False)["실입고량"].sum()
                receive_summary = receive_summary.rename(columns={"실입고량": "누적입고량"})

            result = pd.merge(order_summary, receive_summary, on="order_no", how="left")
            result["누적입고량"] = result["누적입고량"].fillna(0).astype(int)
            result["미입고량"] = result["총발주량"] - result["누적입고량"]
            result["입고율"] = result["누적입고량"] / result["총발주량"]

            def receive_status(row):
                if row["누적입고량"] == 0:
                    return "미입고"
                elif row["누적입고량"] < row["총발주량"]:
                    return "입고중"
                elif row["누적입고량"] == row["총발주량"]:
                    return "완료"
                else:
                    return "초과입고"

            result["상태"] = result.apply(receive_status, axis=1)

            display_df = result.copy()
            display_df["총발주량"] = display_df["총발주량"].apply(fmt)
            display_df["누적입고량"] = display_df["누적입고량"].apply(fmt)
            display_df["미입고량"] = display_df["미입고량"].apply(fmt)
            display_df["입고율"] = result["입고율"].apply(lambda x: f"{x:.1%}")

            display_df = display_df.rename(columns={"order_no": "발주NO"})

            st.dataframe(
                display_df[["발주NO", "총발주량", "누적입고량", "미입고량", "입고율", "상태"]],
                use_container_width=True,
                hide_index=True
            )

            over_count = (result["상태"] == "초과입고").sum()
            if over_count > 0:
                over_qty = (result[result["미입고량"] < 0]["미입고량"].abs()).sum()
                st.warning(f"⚠ 초과입고 {over_count}건 / 초과수량 {fmt(over_qty)} EA")

elif menu == "재고현황":
    st.header("재고현황 조회")

    df = load_data()

    if df.empty:
        st.info("아직 입력된 데이터가 없습니다.")
    else:
        df["date"] = pd.to_datetime(df["date"])

        col1, col2 = st.columns(2)

        with col1:
            start_date = st.date_input("조회 시작일", df["date"].min().date())

        with col2:
            end_date = st.date_input("조회 종료일", df["date"].max().date())

        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)

        before_df = df[df["date"] < start_date]
        period_df = df[(df["date"] >= start_date) & (df["date"] <= end_date)]

        # 시작일 전 누적
        before_receive = before_df[before_df["type"] == "입고"]["qty"].sum()
        before_return = before_df[before_df["type"] == "입고"]["return_qty"].sum()
        before_real_receive = before_receive - before_return

        before_prod = before_df[before_df["type"] == "생산"]["qty"].sum()
        before_defect = before_df[before_df["type"] == "생산"]["defect_qty"].sum()
        before_ship = before_df[before_df["type"] == "출하"]["qty"].sum()

        raw_start = RAW_BEGIN + before_real_receive - before_prod - before_defect
        finished_start = FINISHED_BEGIN + before_prod - before_defect - before_ship

        # 조회기간 내 수량
        period_receive = period_df[period_df["type"] == "입고"]["qty"].sum()
        period_return = period_df[period_df["type"] == "입고"]["return_qty"].sum()
        period_real_receive = period_receive - period_return

        period_prod = period_df[period_df["type"] == "생산"]["qty"].sum()
        period_defect = period_df[period_df["type"] == "생산"]["defect_qty"].sum()
        period_ship = period_df[period_df["type"] == "출하"]["qty"].sum()

        raw_end = raw_start + period_real_receive - period_prod - period_defect
        if raw_end < RAW_SAFETY_STOCK:
            st.error(
                f"⚠ 원재료 재고 부족: 현재고 {fmt(raw_end)} EA / 안전재고 {fmt(RAW_SAFETY_STOCK)} EA"      
            )      
        finished_end = finished_start + period_prod - period_ship

        st.subheader("원재료 재고")

        raw_detail = pd.DataFrame({
            "구분": ["기초재고", "입고", "생산투입", "불량투입", "기말재고"],
            "수량(EA)": [
                fmt(raw_start),
                fmt(period_real_receive),
                fmt(period_prod),
                fmt(period_defect),
                fmt(raw_end)
            ]
        })

        st.table(raw_detail)

        st.subheader("완제품 재고")

        finished_detail = pd.DataFrame({
            "구분": ["기초재고", "양품생산", "출하", "기말재고"],
            "수량(EA)": [
                fmt(finished_start),
                fmt(period_prod),
                fmt(period_ship),
                fmt(finished_end)
            ]
        })

        st.table(finished_detail)

        
        st.divider()

        st.subheader("조회기간 입력 이력")
        st.dataframe(period_df.sort_values("date"), use_container_width=True)

elif menu == "입력이력 관리":
    st.header("입력이력 관리")

    df = load_data()

    excel_data = create_excel(df)

    st.download_button(
        label="📥 엑셀 다운로드",
        data=excel_data,
        file_name="재고관리_전체이력.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if df.empty:
        st.info("수정/삭제할 데이터가 없습니다.")
    else:
        display_df = df.copy()

        display_df["수량"] = display_df["qty"].apply(fmt)
        display_df["반품수량"] = display_df["return_qty"].apply(fmt)
        display_df["불량수량"] = display_df["defect_qty"].apply(fmt)

        display_df = display_df.rename(columns={
            "id": "ID",
            "type": "구분",
            "date": "일자",
            "order_no": "발주NO",
            "memo": "비고"
        })

        st.subheader("전체 입력 이력")
        st.dataframe(
            display_df[["ID", "구분", "일자", "발주NO", "수량", "반품수량", "불량수량", "비고"]],
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        st.subheader("수정 / 삭제할 ID 선택")

        selected_id = st.number_input(
            "ID 입력",
            min_value=1,
            step=1
        )

        target_df = df[df["id"] == selected_id]

        if target_df.empty:
            st.info("선택한 ID의 데이터가 없습니다.")
        else:
            row = target_df.iloc[0]

            st.write(f"선택된 데이터: ID {int(row['id'])} / {row['type']} / {row['date']}")

            type_ = st.selectbox(
                "구분",
                ["발주", "입고", "생산", "출하"],
                index=["발주", "입고", "생산", "출하"].index(row["type"])
            )

            date_ = st.date_input(
                "일자",
                pd.to_datetime(row["date"]).date()
            )

            order_no = st.text_input(
                "발주NO",
                value="" if pd.isna(row["order_no"]) else str(row["order_no"])
            )

            qty = st.number_input(
                "수량",
                min_value=0,
                step=1,
                value=int(row["qty"])
            )
            st.caption(f"수량: {fmt(qty)} EA")

            return_qty = st.number_input(
                "반품수량",
                min_value=0,
                step=1,
                value=int(row["return_qty"])
            )
            st.caption(f"반품수량: {fmt(return_qty)} EA")

            defect_qty = st.number_input(
                "불량수량",
                min_value=0,
                step=1,
                value=int(row["defect_qty"])
            )
            st.caption(f"불량수량: {fmt(defect_qty)} EA")

            memo = st.text_input(
                "비고",
                value="" if pd.isna(row["memo"]) else str(row["memo"])
            )

            col1, col2 = st.columns(2)

            with col1:
                if st.button("수정 저장"):
                    update_data(
                        selected_id,
                        type_,
                        date_,
                        order_no,
                        qty,
                        return_qty,
                        defect_qty,
                        memo
                    )
                    st.success("수정되었습니다.")
                    st.rerun()

            with col2:
                if st.button("삭제"):
                    delete_data(selected_id)
                    st.warning("삭제되었습니다.")
                    st.rerun()


elif menu == "현황 조회":
    st.header("현황 조회")

    df = load_data()

    if df.empty:
        st.info("아직 입력된 데이터가 없습니다.")
    else:
        total_order = df[df["type"] == "발주"]["qty"].sum()
        total_receive = df[df["type"] == "입고"]["qty"].sum()
        total_return = df[df["type"] == "입고"]["return_qty"].sum()
        total_prod = df[df["type"] == "생산"]["qty"].sum()
        total_defect = df[df["type"] == "생산"]["defect_qty"].sum()
        total_ship = df[df["type"] == "출하"]["qty"].sum()

        real_receive = total_receive - total_return
        good_prod = total_prod - total_defect
        inventory = real_receive + good_prod - total_ship
        order_balance = total_order - real_receive

        col1, col2, col3 = st.columns(3)

        col1.metric("발주수량", int(total_order))
        col2.metric("실입고수량", int(real_receive))
        col3.metric("발주잔량", int(order_balance))

        col4, col5, col6 = st.columns(3)

        col4.metric("양품생산수량", int(good_prod))
        col5.metric("출하수량", int(total_ship))
        col6.metric("현재고", int(inventory))

        st.divider()
        st.subheader("입력 이력")
        st.dataframe(df, use_container_width=True)