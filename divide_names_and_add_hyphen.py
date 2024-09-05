# 韵母和声母定义
compound_vowels = ["uang", "iang", "uai", "uan", "uei", "ang", "eng", "ing", "ong", "iao", "iou", "ai", "ei", "ui", "ao", "ou", "iu", "ie", "ve", "an", "en", "in", "un", "er", "ia", "ua", "uo", "ue"]
single_vowels = ["a", "o", "i", "e", "u", "v"]
consonants = ["zh", "ch", "sh", "b", "p", "m", "f", "d", "t", "n", "l", "g", "k", "h", "j", "q", "x", "z", "c", "s", "r", "y", "w"]

def add_hyphen_to_pinyin(pinyin_name):
    # 处理姓和名的不同连接方式
    if "," in pinyin_name:
        parts = pinyin_name.split(",")
        if len(parts) == 2:
            last_name = parts[0].strip().capitalize()
            names = parts[1].strip()
        else:
            return pinyin_name  # 如果分隔符处理失败，则返回原始字符串
    else:
        # 处理空格分隔情况
        parts = pinyin_name.split()
        if len(parts) >= 3:
            last_name = parts[0].strip().capitalize()
            name1 = parts[1].strip()
            name2 = ' '.join(parts[2:]).strip()
            return f"{last_name}, {name1}-{name2}"

        if len(parts) == 2:
            last_name = parts[0].strip().capitalize()
            names = parts[1].strip()
        else:
            return pinyin_name  # 如果分隔符处理失败，则返回原始字符串

    # 处理名中的多个空格（第二个空格及之后的空格替换为短横线）
    if ' ' in names:
        split_names = names.split()
        if len(split_names) > 1:
            names = split_names[0] + '-' + '-'.join(split_names[1:])

    # 特殊处理规则
    special_cases = {
        "dean": "de-an",
        "aoer": "ao-er",
        "anan": "an-an",
        "jiaer": "jia-er",
        "jiaan": "jia-an",
        "aiai": "ai-ai",
        "erer": "er-er"
    }

    # 如果名符合特殊处理规则，直接返回处理后的结果
    names_lower = names.lower()
    if names_lower in special_cases:
        return f"{last_name}, {special_cases[names_lower].capitalize()}"

    # 检查名是否仅由一个声母和一个（复合）韵母组成
    for cons in consonants:
        for vowel in compound_vowels + single_vowels:
            if names.lower() == cons + vowel:
                return f"{last_name}, {names.capitalize()}"

    # 检查名中是否存在重复的声母和韵母组合
    for cons in consonants:
        for vowel in compound_vowels + single_vowels:
            combo = cons + vowel
            if names.lower().count(combo) > 1:
                # 找到第二个出现的组合位置
                second_occurrence = names.lower().find(combo, names.lower().find(combo) + 1)
                # 在第二个组合的前面插入短横线
                names = names[:second_occurrence] + '-' + names[second_occurrence:]
                return f"{last_name}, {names.capitalize()}"

    # 从右到左检查名的部分，寻找可以加短横线的位置
    length = len(names)
    i = length
    while i > 0:
        # 先匹配复合韵母，按照最长的匹配
        for vowel in compound_vowels:
            if i >= len(vowel) and names.lower()[i-len(vowel):i] == vowel:
                # 检查复合韵母左边的字符
                if i - len(vowel) - 1 >= 0:
                    prev_char = names[i-len(vowel)-1]
                    if prev_char not in [',', ' '] and prev_char not in consonants:
                        # 添加短横线的左边不能是空格或逗号
                        if i - len(vowel) - 1 >= 0 and names[i-len(vowel)-1] not in [',', ' ']:
                            names = names[:i-len(vowel)] + '-' + names[i-len(vowel):]
                            return f"{last_name}, {names.capitalize()}"
                # 增加额外的检查规则：an、ao、er 的左边不是空格、逗号或任何声母
                elif vowel in ['an', 'ao', 'er'] and i - len(vowel) > 0:
                    prev_char = names[i-len(vowel)-1]
                    if prev_char not in [',', ' '] and prev_char not in consonants:
                        # 添加短横线的左边不能是空格或逗号
                        if i - len(vowel) - 1 >= 0 and names[i-len(vowel)-1] not in [',', ' ']:
                            names = names[:i-len(vowel)] + '-' + names[i-len(vowel):]
                            return f"{last_name}, {names.capitalize()}"
        # 如果最右边不是复合韵母，则检查单个韵母
        if names[i-1] in single_vowels and i - 2 >= 0 and names[i-2] in consonants:
            if i - 3 < 0 or names[i-3] not in [",", " "]:
                names = names[:i-2] + '-' + names[i-2:]
                return f"{last_name}, {names.capitalize()}"
        i -= 1

    return f"{last_name}, {names.capitalize()}"

# 示例
pinyin_names = [
    "Zhang huan",       # 名中有两个不连续的空格，返回 "Zhang-Huan"
    "Chen, jin",        # 不符合特殊处理规则
    "Li, Changhong",    # 不符合特殊处理规则
    "Wang, Yiming",     # 不符合特殊处理规则
    "Guo, Dean",        # 符合特殊处理规则，返回 "Guo, De-an"
    "He, Ai",           # 不符合特殊处理规则
    "Zhao, shuai",      # 不符合特殊处理规则
    "Wu, lilai",        # 不符合特殊处理规则
    "Li  Zhenan",       # 名中有两个不连续的空格，返回 "Li-Zhenan"
    "Wu  Ao",           # 名中有两个不连续的空格，返回 "Wu-Ao"
    "Li  minjia",         # 名中有两个不连续的空格，返回 "Li-Xuer"
    "Wang Yiming",      # 名中有一个空格，返回 "Wang-Yiming"
    "Li  Jia",          # 名中有两个空格，返回 "Li-Jia"
    "Zheng huangying",   # 名中有两个不连续的空格，返回 "Zheng, Ming-yue"
    "Zheng dean" # 姓 名1 名2，返回 "Zheng, Ming-yue-Zhong"
]

# 应用规则并输出结果
split_pinyin_names = [add_hyphen_to_pinyin(name) for name in pinyin_names]

for original, split in zip(pinyin_names, split_pinyin_names):
    print(f"Original: {original} -> Split: {split}")
