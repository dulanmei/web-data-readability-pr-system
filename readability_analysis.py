# -*- coding: utf-8 -*-

"""
Readability Analysis for Word Document

Calculates:
1. Flesch Reading Ease
2. Flesch-Kincaid Grade Level
3. Gunning Fog Index
4. SMOG Index
5. Coleman-Liau Index
6. Automated Readability Index
7. Dale-Chall Score
"""

from docx import Document
import textstat
import pandas as pd


# ===========================
# 修改这里为你的Word文件路径
# ===========================
file_path = r"/Users/dulanmei/Desktop/www.DESCRIPTION.docx"
# 例如：
# file_path = r"D:\Books\MyBook.docx"


# ===========================
# 读取Word文档
# ===========================
doc = Document(file_path)

text = ""

for para in doc.paragraphs:
    if para.text.strip():
        text += para.text + "\n"

print("=" * 60)
print("Document loaded successfully!")
print("=" * 60)

print(f"Characters: {len(text)}")
print(f"Words: {textstat.lexicon_count(text)}")
print(f"Sentences: {textstat.sentence_count(text)}")
print()


# ===========================
# 计算可读性指标
# ===========================

results = {
    "Flesch Reading Ease":
        textstat.flesch_reading_ease(text),

    "Flesch-Kincaid Grade":
        textstat.flesch_kincaid_grade(text),

    "Gunning Fog":
        textstat.gunning_fog(text),

    "SMOG Index":
        textstat.smog_index(text),

    "Coleman-Liau Index":
        textstat.coleman_liau_index(text),

    "Automated Readability Index":
        textstat.automated_readability_index(text),

    "Dale-Chall Score":
        textstat.dale_chall_readability_score(text)
}


# ===========================
# 控制台输出
# ===========================

print("=" * 60)
print("READABILITY RESULTS")
print("=" * 60)

for metric, value in results.items():
    print(f"{metric:<35}: {value:.2f}")

print("=" * 60)


# ===========================
# 保存Excel
# ===========================

df = pd.DataFrame([results])

output_file = "Readability_Results.xlsx"

df.to_excel(output_file, index=False)

print()
print(f"Results saved to: {output_file}")
print("Finished!")
