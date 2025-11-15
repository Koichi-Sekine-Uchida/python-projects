import codecs

# 文字化けしたデータ
data_list = [
    "1770435 myasuda",
    "1796950 åç°æ´è¡ãé¢å£ åä¸",
    "110771 åç°æ´è¡Backlogç®¡çè",
    "1084718 å è¶ãæ¶ä¸",
    "140255 å¸å·çå¸",
    "1796951 æ²æ¬ æå­",
    "140258 ç³å³¶æå",
    "¢æ ¹å­ä¸"
]

# UTF-8 にデコード
corrected_list = []
for line in data_list:
    try:
        corrected_text = line.encode("latin1").decode("utf-8")
        corrected_list.append(corrected_text)
    except UnicodeDecodeError:
        corrected_list.append(f"デコード失敗: {line}")

# **UTF-8 でファイル出力**
output_filename = "fixed_list.txt"
with codecs.open(output_filename, "w", encoding="utf-8") as f:
    f.write("\n".join(corrected_list))

print(f"修正されたリストを {output_filename} に保存しました。")
