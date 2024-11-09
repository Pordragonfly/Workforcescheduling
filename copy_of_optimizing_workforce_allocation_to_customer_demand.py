import streamlit as st
import pandas as pd
import pulp
from io import BytesIO

def convert_time_to_minutes(time_obj):
    return time_obj.hour * 60 + time_obj.minute

def adjust_time_interval(start_minutes, end_minutes):
    if end_minutes < start_minutes:
        end_minutes += 24 * 60
    return start_minutes, end_minutes

st.title("โปรแกรมจัดสรรพนักงาน (Workforce Scheduling)")

# อัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("กรุณาอัปโหลดไฟล์ Excel", type=["xlsx"])

if uploaded_file:
    # อ่านข้อมูลจากไฟล์ Excel
    wages_df = pd.read_excel(uploaded_file, sheet_name='cj')
    demands_df = pd.read_excel(uploaded_file, sheet_name='dt')
    shifts_df = pd.read_excel(uploaded_file, sheet_name='shift')
    time_intervals_df = pd.read_excel(uploaded_file, sheet_name='time_interval')

    st.write("ข้อมูลค่าแรงพนักงาน (Wage):")
    st.write(wages_df.head())
    st.write("ข้อมูลความต้องการพนักงาน (Demand):")
    st.write(demands_df.head())
    st.write("ข้อมูลกะทำงาน (Shift):")
    st.write(shifts_df.head())
    st.write("ข้อมูลช่วงเวลา (Time Interval):")
    st.write(time_intervals_df.head())

    # เตรียมข้อมูลสำหรับการคำนวณ
    n = len(wages_df)
    cj = wages_df['Wage'].tolist()
    T = len(time_intervals_df)
    dt = demands_df['Demand_Employee'].tolist()

    time_intervals = [
        adjust_time_interval(convert_time_to_minutes(start), convert_time_to_minutes(end))
        for start, end in zip(time_intervals_df['Start_time'], time_intervals_df['End_time'])
    ]

    shift_times = [
        adjust_time_interval(convert_time_to_minutes(start), convert_time_to_minutes(end))
        for start, end in zip(shifts_df['Shift_start'], shifts_df['Shift_end'])
    ]

    atj = []
    for i in range(n):
        atj_row = []
        for t in range(T):
            if shift_times[i][0] <= time_intervals[t][0] and shift_times[i][1] >= time_intervals[t][1]:
                atj_row.append(1)
            else:
                atj_row.append(0)
        atj.append(atj_row)

    # สร้างปัญหา Minimize LP
    problem = pulp.LpProblem("Workforce_Scheduling", pulp.LpMinimize)
    x = pulp.LpVariable.dicts("x", range(n), lowBound=0, cat='Integer')

    problem += pulp.lpSum(cj[j] * x[j] for j in range(n))
    for t in range(T):
        problem += pulp.lpSum(atj[j][t] * x[j] for j in range(n)) >= dt[t]
    problem.solve()

    # แสดงผลลัพธ์
    results_df = pd.DataFrame({
        "Shift": [f"Shift {j + 1}" for j in range(n)],
        "Number of Employees": [x[j].varValue for j in range(n)],
        "Cost": [cj[j] * x[j].varValue for j in range(n)]
    })

    st.write("สถานะการแก้ปัญหา:", pulp.LpStatus[problem.status])
    st.write("ค่าใช้จ่ายทั้งหมด:", pulp.value(problem.objective))
    st.write("ตารางผลลัพธ์")
    st.dataframe(results_df)
