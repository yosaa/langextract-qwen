import langextract as lx
from langextract import factory
from langextract.providers.openai import OpenAILanguageModel
import sys
import os
from docx import Document

# 从命令行参数获取文件名
if len(sys.argv) < 2:
    print("使用方法: python test.py 文件名.docx")
    sys.exit(1)

file_path = sys.argv[1]
if not os.path.exists(file_path):
    print(f"错误: 文件 {file_path} 不存在")
    sys.exit(1)

# 读取docx文件内容
print(f"正在读取文件: {file_path}")
doc = Document(file_path)
input_text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
print(f"文件读取成功，共 {len(input_text)} 个字符\n")

# 定义提取提示
prompt_description = "从合同文本中提取关键信息,包括合同编号、甲方名称、乙方名称、身份证号、金额等信息,按照它们在文本中出现的顺序。"

# 定义示例数据
examples = [
    lx.data.ExampleData(
        text="合同编号:HT-2023-088,签约方为赵六(身份证:440101199203031111)和孙七(身份证:500101198806062222),联系人：10086，签订日期2023年12月1日,金额300,000元。",
        extractions=[
            lx.data.Extraction(extraction_class="合同编号", extraction_text="HT-2023-088"),
            lx.data.Extraction(extraction_class="甲方名称", extraction_text="赵六"),
            lx.data.Extraction(extraction_class="乙方名称", extraction_text="赵六"),
            lx.data.Extraction(extraction_class="身份证号", extraction_text="440101199203031111"),
            lx.data.Extraction(extraction_class="乙方联系电话", extraction_text="10086"),
            lx.data.Extraction(extraction_class="金额", extraction_text="300,000元")
        ]
    )
]

# 自定义HTML生成函数
def generate_custom_html(result, original_text, output_file):
    """生成自定义的左右布局HTML可视化页面"""
    import html
    import json
    
    # 准备抽取结果数据
    extractions = []
    for idx, entity in enumerate(result.extractions):
        extraction_data = {
            'id': idx,
            'class': entity.extraction_class,
            'text': entity.extraction_text,
            'start': entity.char_interval.start_pos if entity.char_interval else -1,
            'end': entity.char_interval.end_pos if entity.char_interval else -1,
            'matched': False  # 标记是否在原文中找到匹配
        }
        extractions.append(extraction_data)
    
    # 智能匹配:如果langextract没有提供位置,尝试在原文中查找
    for ext in extractions:
        if ext['start'] < 0:  # 未提供位置信息
            # 尝试在原文中精确查找
            text_to_find = ext['text']
            pos = original_text.find(text_to_find)
            if pos >= 0:
                ext['start'] = pos
                ext['end'] = pos + len(text_to_find)
                ext['matched'] = True
                print(f"智能匹配成功: [{ext['class']}] {text_to_find} -> 位置 {pos}-{ext['end']}")
            else:
                print(f"警告: 未能在原文中找到 [{ext['class']}] {text_to_find}")
        else:
            ext['matched'] = True
    
    # 生成带高亮的原文HTML
    highlighted_text = original_text
    # 按位置倒序排序,避免位置偏移
    sorted_extractions = sorted([e for e in extractions if e['matched']], key=lambda x: x['start'], reverse=True)
    
    for ext in sorted_extractions:
        if ext['start'] >= 0:
            # 使用原始ID
            original_id = ext['id']
            # 为每个抽取的文本添加span标记
            before = highlighted_text[:ext['start']]
            matched = highlighted_text[ext['start']:ext['end']]
            after = highlighted_text[ext['end']:]
            highlighted_text = f"{before}<span class='highlight' data-id='{original_id}' data-class='{html.escape(ext['class'])}'>{html.escape(matched)}</span>{after}"
    
    # 读取HTML模板
    template_path = os.path.join(os.path.dirname(__file__), 'template.html')
    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()
    
    # 准备注入的数据
    extractions_json = json.dumps(extractions, ensure_ascii=False)
    original_text_json = json.dumps(highlighted_text, ensure_ascii=False)
    
    # 替换模板中的占位符
    html_content = html_template.replace('{{EXTRACTIONS_DATA}}', extractions_json)
    html_content = html_content.replace('{{ORIGINAL_TEXT_DATA}}', original_text_json)
    
    # 写入文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

# 使用Qwen模型进行信息抽取
print("正在调用Qwen模型进行信息抽取...")
result = lx.extract(
    text_or_documents=input_text,
    prompt_description=prompt_description,
    examples=examples,
    fence_output=True,
    use_schema_constraints=False,
    model=OpenAILanguageModel(
        model_id='Qwen/Qwen3-Next-80B-A3B-Instruct',
        base_url='https://api-inference.modelscope.cn/v1',
        api_key='填入你的api',
        provider_kwargs={
            'connect_timeout': 60,    # 允许 60 秒完成 SSL 握手
            'timeout': 120            # 保持 120 秒的整体请求超时
        }
    )
)
print("信息抽取完成！\n")

# 显示抽取结果
print(f"输入文本: {input_text}\n")
print("抽取到的实体信息:")
for entity in result.extractions:
        position_info = ""
        if entity.char_interval:
                start, end = entity.char_interval.start_pos, entity.char_interval.end_pos
                position_info = f" (位置: {start}-{end})"
        print(f"• {entity.extraction_class}: {entity.extraction_text}{position_info}")

# 保存结果
lx.io.save_annotated_documents(
    [result], output_name="contract_extraction.jsonl", output_dir=".")

# 生成自定义可视化HTML
print("正在生成可视化文件...")
generate_custom_html(result, input_text, "contract_visualization.html")
print("\n可视化结果已保存到 contract_visualization.html")
