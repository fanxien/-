import os
import json
import xml.etree.ElementTree as ET
from openai import OpenAI

# ===================== 配置区 =====================

# 建议：实际用时用环境变量更安全
# 例如在 PowerShell:
#   setx DEEPSEEK_API_KEY "sk-xxxx"
# 然后这里写 api_key=os.environ.get("DEEPSEEK_API_KEY")
client = OpenAI(
    api_key="sk-a85e15d664c8461c9a98d0048ec4c1fc",   # ← 换成你的 key，别提交到任何仓库
    base_url="https://api.deepseek.com"
)

# 三个文件夹路径（按你的实际路径改）
EN_DIR = r"F:\原文"
V1_DIR = r"F:\第一版翻译"
V2_DIR = r"F:\第二版翻译"

# 输出结果保存目录
OUTPUT_DIR = r"F:\翻译评估结果_XML"

# 使用的模型
MODEL_NAME = "deepseek-reasoner"

# 每个文件最多评估多少条 entry（按 <entry> 计数）
MAX_ENTRIES_PER_FILE = 30

# ===================== 提示词 =====================

SYSTEM_PROMPT = """
你是一个专业的本地化翻译审核员，精通中英文游戏本地化。

输入中会包含游戏文本的三种版本：
- 英文原文（source_with_tags），包含结构标签，例如 <page>、<hpage> 等；
- 第一版中文翻译（zh_v1_with_tags），结构标签应与原文对应；
- 第二版中文翻译（zh_v2_with_tags），通常质量更高，可作为参考。

要求：
1. 判断第一版中文翻译在含义、风格和结构上是否忠实于英文原文。
2. 结构标签（例如 <page>、<hpage>）本身不需要翻译，但需要检查：
   - 是否漏掉、顺序错位或复制错误；
   - 若有此类问题，可归类为 tech.markup。
3. 重点关注文本内容的翻译问题（错译、漏译、风格不符等），其次关注结构标签是否一致。
4. 只分析当前这一条文本（一个 entry），不要引用本条以外的内容。

请严格输出 JSON（不要有任何多余文字，不要在 JSON 外加解释），结构如下：

{
  "overall_comment": "对这一条第一版翻译的简要总体评价（1-3 句）",
  "issues": [
    {
      "type": "错误类型",
      "source_span": "对应的英文片段（可包含标签）",
      "zh_v1_span": "第一版中文中的相关片段（可包含标签）",
      "zh_v2_span": "第二版中文中的相关片段（如有助于说明问题，可填；否则留空字符串）",
      "description": "问题说明，指出哪里不对以及为什么"
    }
  ]
}

type 字段从以下列表中选择（“大类.子类”的形式）：

1. 准确性（Accuracy）
- "accuracy.mistranslation"        // 语义错译：意思翻错、反了、关系错等
- "accuracy.omission"              // 漏译：原文信息缺失
- "accuracy.addition"              // 多译/臆译：原文没有的信息被多加进去
- "accuracy.ambiguity"             // 信息变得模糊或歧义，原文本来是清楚的

2. 术语与一致性（Terminology & Consistency）
- "terminology.wrong_term"         // 术语错误，用词不符合约定 / 官方译名
- "terminology.inconsistent_term"  // 术语不一致，同一概念多种说法
- "terminology.proper_name"        // 人名/地名/专有名词错误或不统一
- "terminology.ui_inconsistent"    // UI 文案不一致（按钮、菜单等用语不统一）

3. 流畅度（Fluency）
- "fluency.grammar"                // 语法错误、结构不完整
- "fluency.word_choice"            // 用词或搭配不当，但大致意思对
- "fluency.word_order"             // 语序不自然或导致轻微歧义
- "fluency.unintelligible"         // 病句或难以理解，需要反复读才能明白
- "fluency.punctuation"            // 标点、全角半角问题影响阅读

4. 风格与语气（Style & Tone）
- "style.character_voice"          // 人物人设不符，说话风格与角色形象不一致
- "style.register"                 // 语体不合适（过于书面/口语，与场景不匹配）
- "style.tone_mismatch"            // 情绪或语气失真（严肃变成搞笑、吐槽变得僵硬等）

5. 本地化与文化（Locale & Culture）
- "locale.cultural_inappropriate"  // 文化不当、可能冒犯本地玩家
- "locale.convention"              // 日期/时间/数字/货币/度量单位等本地化规范错误
- "locale.platform_requirement"    // 未遵守特定地区/平台的用语或审查规范

6. 技术与格式（Technical & Formatting）
- "tech.placeholder"               // 占位符/变量错误，如 %s, {0}, <player_name> 等
- "tech.markup"                    // 标签/富文本/颜色代码等错误，可能导致显示异常
- "tech.truncation"                // 文本过长或设计不当，容易在 UI 中被截断/溢出
- "tech.layout"                    // 换行、空格等排版问题明显影响阅读或 UI 展示

7. 设定与事实（Lore & Mechanics）
- "verity.lore_error"              // 世界观/背景设定相关内容错误
- "verity.mechanics_error"         // 数值、技能效果、冷却时间等机制描述错误
- "verity.context_inconsistency"   // 与前后文矛盾或整体设定不一致

8. 其他（Other）
- "other.other"                    // 以上类别都不适用时使用，并在 description 中说明

如果这一条的第一版翻译几乎没有问题，可以让 issues 为 []。
请务必返回合法的 JSON，且不要在 JSON 外输出任何多余文字。
"""

# ===================== 工具函数 =====================

def build_user_prompt(entry_name: str, en_text: str, v1_text: str, v2_text: str) -> str:
    """把三个版本的文本拼到 user prompt 里"""
    return f"""
当前条目 name = {entry_name}

[英文原文 - source_with_tags]
{en_text}

[第一版中文翻译 - zh_v1_with_tags]
{v1_text}

[第二版中文翻译 - zh_v2_with_tags]
{v2_text}
"""

def parse_entries_with_tags(path: str) -> dict:
    """
    解析类似：
    <entries>
      <entry name="XXX">文本 &lt;page&gt; 文本</entry>
      ...
    </entries>

    返回: { name: text_with_tags }
    注意：文件里是 &lt;page&gt;，解析出来会变成真正的 "<page>" 字符串。
    """
    tree = ET.parse(path)
    root = tree.getroot()
    data = {}
    for entry in root.findall("entry"):
        name = entry.get("name")
        if not name:
            continue
        # entry.text 已经是解码后的文本，包括 <page> <hpage>
        text = entry.text or ""
        # 如果你想更清楚一点，也可以把标签替换成 [PAGE]、[HPAGE]，这里先原样保留：
        # text = text.replace("<page>", "\n[PAGE]\n").replace("<hpage>", "\n[HPAGE]\n")
        data[name] = text.strip()
    return data

def evaluate_entry(entry_name: str, en_text: str, v1_text: str, v2_text: str) -> dict:
    """
    调一次 DeepSeek，评估单个 entry，返回 dict
    """
    user_prompt = build_user_prompt(entry_name, en_text, v1_text, v2_text)

    resp = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ],
        stream=False
    )

    content = resp.choices[0].message.content

    # 解析 JSON，带一个简单兜底
    try:
        result = json.loads(content)
    except json.JSONDecodeError:
        start = content.find("{")
        end = content.rfind("}")
        if start != -1 and end != -1 and end > start:
            result = json.loads(content[start:end+1])
        else:
            print("JSON 解析失败，原始内容如下：")
            print(content)
            raise

    # 把 entry_name 带上，便于后续追踪
    result["entry_name"] = entry_name
    return result

# ===================== 主逻辑 =====================

def process_file_triplet(en_path: str, v1_path: str, v2_path: str):
    """
    处理一组三个文件（EN / V1 / V2），按 entry name 对齐并逐条评估。
    """
    en_name = os.path.basename(en_path)
    print(f"\n===== 处理文件：{en_name} =====")
    print("EN:", en_path)
    print("V1:", v1_path)
    print("V2:", v2_path)

    en_dict = parse_entries_with_tags(en_path)
    v1_dict = parse_entries_with_tags(v1_path)
    v2_dict = parse_entries_with_tags(v2_path)

    en_keys = list(en_dict.keys())

    total_entries = len(en_keys)
    limit = min(total_entries, MAX_ENTRIES_PER_FILE)

    print(f"条目总数: {total_entries}，本次评估前 {limit} 条。")

    results = []

    for i, name in enumerate(en_keys[:limit], start=1):
        en_text = en_dict.get(name, "")
        v1_text = v1_dict.get(name, "")
        v2_text = v2_dict.get(name, "")

        print(f"  [{i}/{limit}] name = {name} (EN 有: {bool(en_text)}, V1 有: {bool(v1_text)}, V2 有: {bool(v2_text)})")

        result = evaluate_entry(name, en_text, v1_text, v2_text)
        results.append(result)

    # 保存结果
    base = os.path.splitext(en_name)[0]
    out_path = os.path.join(OUTPUT_DIR, f"{base}_eval.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    print(f"已保存评估结果 -> {out_path}")

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 按文件顺序对齐：假设三个文件夹中的文件一一对应（第 1 个对第 1 个）
    en_files = sorted([f for f in os.listdir(EN_DIR) if f.lower().endswith((".txt", ".xml"))])
    v1_files = sorted([f for f in os.listdir(V1_DIR) if f.lower().endswith((".txt", ".xml"))])
    v2_files = sorted([f for f in os.listdir(V2_DIR) if f.lower().endswith((".txt", ".xml"))])

    if not (len(en_files) == len(v1_files) == len(v2_files)):
        print("三个文件夹中的文件数量不一致，请检查！")
        print("EN 数量：", len(en_files))
        print("V1 数量：", len(v1_files))
        print("V2 数量：", len(v2_files))
        return

    for en_fname, v1_fname, v2_fname in zip(en_files, v1_files, v2_files):
        en_path = os.path.join(EN_DIR, en_fname)
        v1_path = os.path.join(V1_DIR, v1_fname)
        v2_path = os.path.join(V2_DIR, v2_fname)

        process_file_triplet(en_path, v1_path, v2_path)

    print("\n全部文件评估完成！")

if __name__ == "__main__":
    main()
