import langextract as lx
from langextract import factory
from langextract.providers.openai import OpenAILanguageModel
import sys
import os
from docx import Document

# 从命令行参数获取文件名
if len(sys.argv) < 2:
    print("使用方法: python error_check.py 文件名.docx")
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
prompt_description = """你是一个专业的文档校对助手。请仔细检查文本中的错误，包括：
1. 错别字：拼写错误或同音字误用
2. 语病：语法不通顺、表达不清晰的地方
3. 标点符号错误：标点符号使用不当或缺失

对于每个错误，请标注：
- 错误类型（错别字/语病/标点符号错误）
- 错误内容（原文中的错误部分）
- 正确内容（应该修改成的内容）

请按照在文本中出现的顺序提取错误。"""

# 定义示例数据
examples = [
    lx.data.ExampleData(
        text="我们需球在下周完成这个项目。会议时间是明天上午9点，请准时参加会议",
        extractions=[
            lx.data.Extraction(extraction_class="错别字", extraction_text="需球 应改为 需求"),
            lx.data.Extraction(extraction_class="标点符号错误", extraction_text="会议， 应改为 会议。")
        ]
    ),
    lx.data.ExampleData(
        text="这个方案虽然很好但是需要进一步的改进和完善。公司将在近期召开会议讨论此事",
        extractions=[
            lx.data.Extraction(extraction_class="标点符号错误", extraction_text="虽然很好但是， 应改为 虽然很好，但是"),
            lx.data.Extraction(extraction_class="标点符号错误", extraction_text="此事， 应改为 此事。")
        ]
    )
]

# 自定义HTML生成函数
def generate_error_check_html(result, original_text, output_file):
    """生成文档纠错的可视化HTML页面"""
    import html
    import json
    import re
    
    # 准备错误检查结果数据
    errors = []
    for idx, entity in enumerate(result.extractions):
        # 解析错误信息，格式如: "需球应该为需求" 或 "xxx应该改成xxxx"
        error_text = entity.extraction_text
        
        # 尝试多种分隔符：应该为、应该改成、应该是、应改为
        original_part = ""
        correct_part = ""
        
        if "应该为" in error_text:
            parts = error_text.split("应该为", 1)
            original_part = parts[0].strip()
            correct_part = parts[1].strip()
        elif "应该改成" in error_text:
            parts = error_text.split("应该改成", 1)
            original_part = parts[0].strip()
            correct_part = parts[1].strip()
        elif "应该是" in error_text:
            parts = error_text.split("应该是", 1)
            original_part = parts[0].strip()
            correct_part = parts[1].strip()
        elif "应改为" in error_text:
            parts = error_text.split("应改为", 1)
            original_part = parts[0].strip()
            correct_part = parts[1].strip()
        else:
            # 如果没有找到分隔符，整个作为原始错误
            original_part = error_text
            correct_part = ""
        
        error_data = {
            'id': idx,
            'type': entity.extraction_class,  # 错别字/语病/标点符号错误
            'original': original_part,  # 错误内容
            'correct': correct_part,    # 正确内容
            'full_text': error_text,    # 完整的描述
            'start': -1,
            'end': -1,
            'matched': False
        }
        
        # 尝试在原文中查找错误位置
        if original_part:
            # 移除可能的标点符号和空格来提高匹配准确性
            search_text = original_part.replace('，', '').replace('。', '').replace('、', '').replace(' ', '')
            
            # 在原文中查找（支持模糊匹配）
            pos = -1
            for i in range(len(original_text)):
                # 检查是否匹配（忽略标点）
                match_len = 0
                j = i
                k = 0
                while j < len(original_text) and k < len(search_text):
                    char = original_text[j]
                    if char not in ['，', '。', '、', ' ', '\n', '\r', '\t']:
                        if char == search_text[k]:
                            match_len += 1
                            k += 1
                        else:
                            break
                    j += 1
                
                if k == len(search_text):
                    # 找到匹配，记录原始位置（包含标点）
                    pos = i
                    # 计算实际结束位置
                    end_pos = j
                    error_data['start'] = pos
                    error_data['end'] = end_pos
                    error_data['matched'] = True
                    print(f"智能匹配成功: [{error_data['type']}] {original_part} -> 位置 {pos}-{end_pos}")
                    break
            
            if pos < 0:
                # 尝试直接精确匹配
                pos = original_text.find(original_part)
                if pos >= 0:
                    error_data['start'] = pos
                    error_data['end'] = pos + len(original_part)
                    error_data['matched'] = True
                    print(f"精确匹配成功: [{error_data['type']}] {original_part} -> 位置 {pos}-{error_data['end']}")
                else:
                    print(f"警告: 未能在原文中找到 [{error_data['type']}] {original_part}")
        
        errors.append(error_data)
    
    # 生成带高亮的原文HTML
    # 按位置倒序排序，避免位置偏移
    sorted_errors = sorted([e for e in errors if e['matched']], key=lambda x: x['start'], reverse=True)
    
    # 定义不同错误类型的样式类
    type_class_map = {
        '错别字': 'error-typo',
        '语病': 'error-grammar',
        '标点符号错误': 'error-punctuation'
    }
    
    # 从原文开始构建高亮文本
    highlighted_text = original_text
    
    # 倒序处理，从后往前插入标签，避免位置偏移
    for err in sorted_errors:
        if err['start'] >= 0:
            # 根据错误类型选择样式类
            style_class = type_class_map.get(err['type'], 'error-other')
            original_id = err['id']
            
            # 为每个错误添加span标记
            before = highlighted_text[:err['start']]
            matched = highlighted_text[err['start']:err['end']]
            after = highlighted_text[err['end']:]
            
            # 只对匹配的文本进行HTML转义，保留之前插入的标签
            highlighted_text = f"{before}<span class='highlight {style_class}' data-id='{original_id}' data-type='{html.escape(err['type'])}'>{html.escape(matched)}</span>{after}"
    
    # 最后对整个文本进行一次完整的HTML转义处理（只转义未被span包裹的部分）
    # 使用更安全的方式：先标记所有span，然后转义剩余部分
    import re
    
    # 将所有span标签临时替换为占位符
    span_pattern = r'(<span[^>]*>.*?</span>)'
    parts = re.split(span_pattern, highlighted_text)
    
    # 对非span部分进行HTML转义
    for i in range(len(parts)):
        if not parts[i].startswith('<span'):
            parts[i] = html.escape(parts[i])
    
    highlighted_text = ''.join(parts)
    
    # 读取HTML模板
    template_path = os.path.join(os.path.dirname(__file__), 'error_check_template.html')
    with open(template_path, 'r', encoding='utf-8') as f:
        html_template = f.read()
    
    # 准备注入的数据
    errors_json = json.dumps(errors, ensure_ascii=False)
    original_text_json = json.dumps(highlighted_text, ensure_ascii=False)
    
    # 替换模板中的占位符
    html_content = html_template.replace('{{ERRORS_DATA}}', errors_json)
    html_content = html_content.replace('{{ORIGINAL_TEXT_DATA}}', original_text_json)
    
    # 写入文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

# 使用Qwen模型进行错误检查
print("正在调用Qwen模型进行文档纠错...")
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
            'connect_timeout': 60,
            'timeout': 120
        }
    )
)
print("文档纠错完成！\n")

# 显示检查结果
print(f"输入文本: {input_text}\n")
print("检测到的错误:")
for entity in result.extractions:
    print(f"• [{entity.extraction_class}] {entity.extraction_text}")

# 保存结果
lx.io.save_annotated_documents(
    [result], output_name="error_check_result.jsonl", output_dir=".")

# 生成可视化HTML
print("\n正在生成可视化文件...")
generate_error_check_html(result, input_text, "error_check_visualization.html")
print("\n可视化结果已保存到 error_check_visualization.html")
