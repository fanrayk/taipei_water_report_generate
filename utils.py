# utils.py
import os
import re
import glob
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def transform_measurement_method(x):
    """將施測方式或儀器代號轉換成 4 位數字字串，再拆分成四個部分"""
    if pd.isnull(x):
        return None
    try:
        s = str(int(x)).zfill(4)
    except Exception:
        s = "0000"
    return {"part1": s[0], "part2": s[1], "part3": s[2], "part4": s[3]}

def set_cell_width(cell, width):
    """設定 Docx 表格中儲存格的寬度"""
    tc = cell._element
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width))
    tcW.set(qn("w:type"), "dxa")

def chunk_list(data_list, chunk_size):
    """將 data_list 分割成每個區塊不超過 chunk_size 的子列表"""
    return [data_list[i:i+chunk_size] for i in range(0, len(data_list), chunk_size)]

def cleanup_temp_files(output_folder, pattern="temp*"):
    """刪除 output_folder 中符合 pattern 的暫存檔案"""
    temp_file_pattern = os.path.join(output_folder, pattern)
    temp_files = glob.glob(temp_file_pattern)
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
            print(f"已刪除暫存檔案: {temp_file}")
        except Exception as e:
            print(f"刪除暫存檔案 {temp_file} 時發生錯誤: {e}")
