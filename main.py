import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

def run_feature_1(source_all, ref_all):
    source_sheets = list(source_all.keys())
    ref_sheets = list(ref_all.keys())
    
    # [비교 기준] 참고파일: Sub Item / Entry ID 결합 (첫 번째 시트)
    df_ref = ref_all[ref_sheets[0]]
    df_ref['KEY'] = df_ref["Sub Item"].astype(str) + "/" + df_ref["Entry ID"].astype(str)

    # 원본 2번째 시트 업데이트
    if len(source_sheets) >= 2:
        df = source_all[source_sheets[1]]
        for i, row in df.iterrows():
            if str(row.get('LAYOUT', '')).upper() == "SYSDEFINED":
                df.at[i, 'Mandatory'] = "N"
            else:
                # [비교 기준] 원본: DOMAIN / VARIABLE_ID 결합
                s_key = f"{row.get('DOMAIN', '')}/{row.get('VARIABLE_ID', '')}"
                match = df_ref[df_ref['KEY'] == s_key]
                if not match.empty:
                    f_val = str(match.iloc[0].get('Sub Item', ''))
                    o_val = str(match.iloc[0].get('Default Missing Query', ''))
                    df.at[i, 'Mandatory'] = "Y/N" if "." in f_val else ("Y" if o_val == "Yes" else "N" if o_val == "No" else o_val)
        source_all[source_sheets[1]] = df

    # 원본 3번째 시트 업데이트
    if len(source_sheets) >= 3:
        df_c = source_all[source_sheets[2]]
        for i, row in df_c.iterrows():
            s_key = f"{row.get('DOMAIN', '')}/{row.get('VARIABLE_ID', '')}"
            match = df_ref[df_ref['KEY'] == s_key]
            if not match.empty:
                o_val = str(match.iloc[0].get('Default Missing Query', ''))
                df_c.at[i, 'Mandatory'] = "Y" if o_val == "Yes" else "N" if o_val == "No" else o_val
        source_all[source_sheets[2]] = df_c
    return source_all

def process():
    src_p = label2.cget("text").strip('{}')
    ref_p = label3.cget("text").strip('{}')
    if not os.path.exists(src_p) or not os.path.exists(ref_p):
        messagebox.showwarning("알림", "파일을 모두 넣어주세요!")
        return
    try:
        source_all = pd.read_excel(src_p, sheet_name=None)
        ref_all = pd.read_excel(ref_p, sheet_name=None)
        res = run_feature_1(source_all, ref_all)
        save_p = os.path.join(os.path.dirname(src_p), "업데이트_결과.xlsx")
        with pd.ExcelWriter(save_p) as writer:
            for name, df in res.items():
                if 'KEY' in df.columns: df = df.drop(columns=['KEY'])
                df.to_excel(writer, sheet_name=name, index=False)
        messagebox.showinfo("성공", f"완료! 파일 위치:\n{save_p}")
    except Exception as e:
        messagebox.showerror("오류", str(e))

# UI 설정
root = TkinterDnD.Tk()
root.title("Excel Updater")
root.geometry("500x500")

tk.Label(root, text="기능: DB Specification 업데이트", font=("Arial", 10, "bold")).pack(pady=10)
label2 = tk.Label(root, text="원본파일 드래그(2,3번 시트)", relief="solid", width=50, height=5, bg="#E3F2FD")
label2.pack(pady=5); label2.drop_target_register(DND_FILES); label2.dnd_bind('<<Drop>>', lambda e: label2.config(text=e.data))
label3 = tk.Label(root, text="참고파일 드래그(1번 시트)", relief="solid", width=50, height=5, bg="#FFF9C4")
label3.pack(pady=5); label3.drop_target_register(DND_FILES); label3.dnd_bind('<<Drop>>', lambda e: label3.config(text=e.data))
tk.Button(root, text="실행하기", command=process, bg="#2E7D32", fg="white", width=15, height=2).pack(pady=20)
root.mainloop()
