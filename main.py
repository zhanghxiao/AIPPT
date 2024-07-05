import os
import requests
import json
from flask import Flask, request, jsonify, send_file, render_template
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
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
        '你是一个专业的PPT制作助手。请为以下主题生成一个详细的PPT内容，包括封面、目录和至少5页内容。'
        '封面至少应该有主标题和副标题，其余每页应该有详细的内容，包括文字描述和图片建议。'
        '使用markdown格式输出，并在适当位置添加[IMAGE:图片描述]标记来表示图片位置和所需图片的描述。'
        '图片可以插入在文字旁边或其他合适的位置。'
    )
    if reference:
        prompt += f'参考内容: {reference}'

    data = {
        'model': 'gpt-3.5-turbo',
        'messages': [
            {'role': 'system', 'content': prompt},
            {'role': 'user', 'content': f'请为主题"{topic}"生成PPT内容。'}
        ]
    }

    response = requests.post(API_REQUEST_URL, headers=headers, json=data)
    response_data = response.json()
    ppt_content = response_data['choices'][0]['message']['content']

    ppt = create_ppt(ppt_content)
    os.makedirs('static', exist_ok=True)
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    try:
        ppt.save(ppt_path)
    except Exception as e:
        print(f"Error saving PPT: {str(e)}")
        return jsonify({"error": "Unable to save PPT file"}), 500

    return jsonify({"content": ppt_content, "file_path": ppt_path})


@app.route('/update_ppt', methods=['POST'])
def update_ppt():
    content = request.json['content']
    ppt = create_ppt(content)
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    try:
        ppt.save(ppt_path)
        return jsonify({"success": True})
    except Exception as e:
        print(f"Error saving updated PPT: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/download_ppt')
def download_ppt():
    ppt_path = os.path.join('static', 'generated_ppt.pptx')
    if os.path.exists(ppt_path):
        return send_file(ppt_path, as_attachment=True)
    else:
        return "PPT file not found", 404


def set_text_frame_properties(text_frame):
    text_frame.word_wrap = True  # 设置自动换行
    for p in text_frame.paragraphs:
        font = p.font
        font.size = Pt(18)  # 设置字体大小
        font.name = '微软雅黑'  # 设置字体
        font.color.rgb = RGBColor(0, 0, 0)  # 设置字体颜色为黑色


def create_ppt(content):
    ppt = Presentation()
    ppt.slide_width = Inches(13.33)  # 设置宽度为16:9比例
    ppt.slide_height = Inches(7.5)   # 设置高度为16:9比例
    slides = content.split('## ')
    for slide_content in slides:
        if slide_content.strip():
            slide = ppt.slides.add_slide(ppt.slide_layouts[1])
            title = slide.shapes.title
            content_shape = slide.placeholders[1]

            lines = slide_content.split('\n')
            title.text = lines[0].strip()

            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)  # Adjust based on your slide size
            height = Inches(0.5)

            for line in lines[1:]:
                if '[IMAGE:' in line:
                    # Extract image description
                    description = line[line.index(':') + 1:line.index(']')]
                    img_left = Inches(1)
                    img_top = top
                    img_width = Inches(4)
                    img_height = Inches(3)
                    img_placeholder = slide.shapes.add_textbox(
                        img_left, img_top, img_width, img_height)
                    img_placeholder.fill.solid()
                    img_placeholder.fill.fore_color.rgb = RGBColor(
                        240, 240, 240)
                    img_placeholder.line.color.rgb = RGBColor(200, 200, 200)

                    tf = img_placeholder.text_frame
                    p = tf.add_paragraph()
                    p.text = f"[图片: {description}]"
                    p.alignment = PP_ALIGN.CENTER

                    top += img_height + Inches(0.5)
                else:
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    p = tf.add_paragraph()
                    p.text = line
                    set_text_frame_properties(tf)  # 设置文本框属性，包括自动换行
                    top += height + Inches(0.1)

            # Remove default content shape
            sp = content_shape._element
            sp.getparent().remove(sp)

    return ppt


if __name__ == '__main__':
    app.run(debug=True)
