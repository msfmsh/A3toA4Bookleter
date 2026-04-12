# Copyright (c) 2026 msfmsh. All rights reserved.

import os
import glob
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    import sys
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf"])
    from pypdf import PdfReader, PdfWriter

def detect_binding(pdf_path):
    reader = PdfReader(pdf_path)
    if len(reader.pages) == 0:
        return 'left', 'ページなし'
    
    right_keywords = ['国語', '漢字']
    left_keywords = ['算数', '数学', '理科', '社会', '英語', '英', '算']
    
    # 全てのページを順に検索して、最初に見つかったキーワードで判定する
    for i, page in enumerate(reader.pages):
        text = page.extract_text()
        if not text:
            continue
            
        text = text.replace(' ', '').replace('\n', '')
        
        for k in right_keywords:
            if k in text:
                return 'right', f"{k} (P{i+1}で検出)"
                
        for k in left_keywords:
            if k in text:
                return 'left', f"{k} (P{i+1}で検出)"
                
    return 'left', '不明（デフォルト）'

def process_pdf(input_pdf):
    binding, matched_keyword = detect_binding(input_pdf)
    print(f"[{input_pdf}] 検出キーワード: '{matched_keyword}' -> 綴じ方向: {'右綴じ' if binding == 'right' else '左綴じ'}")
    
    # output file name logic
    filename = os.path.basename(input_pdf)
    base, ext = os.path.splitext(filename)
    output_pdf = os.path.join("output", f"{base}_A4_booklet{ext}")
    
    reader_l = PdfReader(input_pdf)
    reader_r = PdfReader(input_pdf)
    writer = PdfWriter()

    num_a3_sides = len(reader_l.pages)
    num_a4_pages = num_a3_sides * 2

    a4_pages = [None] * num_a4_pages

    for i in range(num_a3_sides):
        page_l = reader_l.pages[i]
        page_r = reader_r.pages[i]

        width = float(page_l.mediabox.right) - float(page_l.mediabox.left)
        height = float(page_l.mediabox.top) - float(page_l.mediabox.bottom)

        if width > height:
            # 横長 (Landscape): 横幅(X)を半分に分割
            mid_x = float(page_l.mediabox.left) + width / 2.0
            page_l.cropbox.right = mid_x
            page_r.cropbox.left = mid_x
        else:
            # 縦長 (Portrait): 縦幅(Y)を半分に分割
            # そのままだとA4横長(297x210)になってしまうため、A4縦長(210x297)にするためにページを90度回転させます。
            # （給紙向きにより逆さまになる可能性があるため、その場合はプレビュー等で回転をお願いします）
            mid_y = float(page_l.mediabox.bottom) + height / 2.0
            page_l.cropbox.bottom = mid_y   # 上半分
            page_r.cropbox.top = mid_y      # 下半分

        if binding == "left":
            idx_l, idx_r = (num_a4_pages - i - 1, i) if i % 2 == 0 else (i, num_a4_pages - i - 1)
        else: # binding == "right"
            idx_l, idx_r = (i, num_a4_pages - i - 1) if i % 2 == 0 else (num_a4_pages - i - 1, i)

        a4_pages[idx_l] = page_l
        a4_pages[idx_r] = page_r

    for page in a4_pages:
        if page:
            writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)
        
    print(f"完了: {output_pdf} に保存しました。")

if __name__ == '__main__':
    import shutil
    
    os.makedirs("input", exist_ok=True)
    os.makedirs("input/done", exist_ok=True)
    os.makedirs("output", exist_ok=True)

    pdfs = glob.glob("input/*.pdf")
    processed = 0
    for pdf in pdfs:
        print(f"処理開始: {pdf}")
        try:
            process_pdf(pdf)
            done_path = os.path.join("input/done", os.path.basename(pdf))
            # 同名ファイルが既にある場合は上書きを防ぐため削除してから移動
            if os.path.exists(done_path):
                os.remove(done_path)
            shutil.move(pdf, done_path)
            processed += 1
        except Exception as e:
            print(f"[{pdf}] エラー発生: {e}")
        
    if processed == 0:
        print("処理するPDFが見つかりませんでした。 (./input/ フォルダにPDFを置いてください)")
    else:
        print(f"全ての処理が完了しました。処理件数: {processed}件")
