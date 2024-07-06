import os
import requests
import json
from flask import Flask, request, jsonify, send_file, render_template
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from dotenv import load_dotenv

app = Flask(__name__)

load_dotenv()

OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
API_REQUEST_URL = os.getenv('API_REQUEST_URL')


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    topic = data['topic']
    reference = data.get('reference', '')

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {OPENAI_API_KEY}'
    }
    prompt = (
        f'请为主题"{topic}"生成一个详细的PPT内容，包括封面、目录和5个以上的内容页面,以及需要2张以上图片，封面可以不添加图片，目录必须有目录。内容页面必须有相关的内容'
        '使用以下格式和标识符：\n'
        '1. 使用"[SLIDE]"作为每页幻灯片的开始标记。\n'
        '2. 使用"[TITLE]"标记主标题。\n'
        '3. 使用"[SUBTITLE]"标记副标题。\n'
        '4. 使用"[CONTENT]"标记普通内容，每个要点使用"-"开始。\n'
        '5. 使用"[IMAGE]"标记图片位置，后面跟图片描述。\n'
        '例如：\n'
        '[SLIDE]\n'
        '[TITLE]PPT主题\n'
        '[SUBTITLE]副标题\n'
        '[CONTENT]\n'
        '- 内容要点1\n'
        '- 内容要点2\n'
        '[IMAGE]这里是图片描述\n'
        '[SLIDE]\n'
        '... (下一页)'
    )
    if reference:
        prompt += f' 参考内容或者修改意见: {reference}'

    data = {
        'model': 'gpt-3.5-turbo',
        'messages': [
            {'role': 'system', 'content': prompt},
            {'role': 'user', 'content': '请生成完整的PPT内容。'}
        ]
    }

    response = requests.post(API_REQUEST_URL, headers=headers, json=data)
    response_data = response.json()
    ppt_content = response_data['choices'][0]['message']['content']

    ppt = create_ppt(ppt_content)
    os.makedirs('static', exist_ok=True)
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    ppt.save(ppt_path)

    return jsonify({"content": ppt_content, "file_path": ppt_path})


@app.route('/update_ppt', methods=['POST'])
def update_ppt():
    content = request.json['content']
    ppt = create_ppt(content)
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    ppt.save(ppt_path)
    return jsonify({"success": True})


@app.route('/download_ppt')
def download_ppt():
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    if os.path.exists(ppt_path):
        return send_file(ppt_path, as_attachment=True)
    else:
        return "PPT file not found", 404


def add_image_placeholder(slide, description):
    # left = Inches(1)
    # top = Inches(3.5)
    # width = Inches(8)
    # height = Inches(3)
    left = Inches(1)
    top = Inches(4)  # 将图片位置向下移动
    width = Inches(8)
    height = Inches(2.5)  # 减小图片高度

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(240, 240, 240)
    shape.line.color.rgb = RGBColor(200, 200, 200)

    tf = shape.text_frame
    tf.clear()
    p = tf.add_paragraph()
    p.text = f"[图片: {description}]"
    p.alignment = PP_ALIGN.CENTER

    return shape


# def create_ppt(content):
#     ppt = Presentation()
#     slides = content.split('[SLIDE]')
#
#     for slide_content in slides[1:]:  # Skip the first empty split
#         slide = ppt.slides.add_slide(ppt.slide_layouts[1])
#         lines = slide_content.strip().split('\n')
#
#         for line in lines:
#             if line.startswith('[TITLE]'):
#                 title = slide.shapes.title
#                 title.text = line.replace('[TITLE]', '').strip()
#             elif line.startswith('[SUBTITLE]'):
#                 subtitle = slide.placeholders[1]
#                 subtitle.text = line.replace('[SUBTITLE]', '').strip()
#             elif line.startswith('[CONTENT]'):
#                 content = slide.placeholders[1]
#                 content.text = ''
#             elif line.startswith('-'):
#                 p = content.text_frame.add_paragraph()
#                 p.text = line.strip()
#                 p.level = 0
#             elif line.startswith('[IMAGE]'):
#                 description = line.replace('[IMAGE]', '').strip()
#                 add_image_placeholder(slide, description)
#
#         # Set font properties
#         for shape in slide.shapes:
#             if not shape.has_text_frame:
#                 continue
#             for paragraph in shape.text_frame.paragraphs:
#                 for run in paragraph.runs:
#                     run.font.name = '微软雅黑'
#                     run.font.size = Pt(18)
#                     run.font.color.rgb = RGBColor(0, 0, 0)
#
#     return ppt
def create_ppt(content):
    ppt = Presentation()
    slides = content.split('[SLIDE]')

    for i, slide_content in enumerate(slides[1:]):  # Skip the first empty split
        if i == 0:  # Cover slide
            slide = ppt.slides.add_slide(ppt.slide_layouts[0])  # Use the title slide layout
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
        else:
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])  # Use the content slide layout
            title = slide.shapes.title
            content = slide.placeholders[1]


        lines = slide_content.strip().split('\n')

        for line in lines:
            if line.startswith('[TITLE]'):
                title.text = line.replace('[TITLE]', '').strip()
            elif line.startswith('[SUBTITLE]'):
                if i == 0:
                    subtitle.text = line.replace('[SUBTITLE]', '').strip()
                else:
                    content.text += line.replace('[SUBTITLE]', '').strip() + '\n'
            elif line.startswith('[CONTENT]'):
                if i != 0:
                    content.text = ''
            elif line.startswith('-'):
                if i != 0:
                    p = content.text_frame.add_paragraph()
                    p.text = line.strip()
                    p.level = 0
            elif line.startswith('[IMAGE]'):
                description = line.replace('[IMAGE]', '').strip()
                add_image_placeholder(slide, description)

        # Set font properties
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if i == 0 and shape == title:  # Cover slide title
                        run.font.name = '微软雅黑'
                        run.font.size = Pt(44)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 0, 0)
                    elif i == 0 and shape == subtitle:  # Cover slide subtitle
                        run.font.name = '微软雅黑'
                        run.font.size = Pt(32)
                        run.font.color.rgb = RGBColor(89, 89, 89)
                    else:  # Other slides
                        run.font.name = '微软雅黑'
                        run.font.size = Pt(18)
                        run.font.color.rgb = RGBColor(0, 0, 0)

    return ppt

if __name__ == '__main__':
    app.run(debug=True)
