from flask import Flask, render_template, request, jsonify, send_file
import os
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Pt
import re
import json
from io import BytesIO

app = Flask(__name__)

# 設定ファイルを読み込む
def load_config():
    """設定ファイルを読み込む"""
    config_file = 'config.json'
    default_config = {
        'default_keywords': ['Hitachi Astemo', '日立Astemo', '日立アステモ'],
        'default_replacement': 'Astemo',
        'max_file_size_mb': 50,
        'allowed_extensions': ['pptx', 'ppt']
    }
    
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"設定ファイルの読み込みに失敗しました: {str(e)}")
            return default_config
    return default_config

config = load_config()

# 設定
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = set(config['allowed_extensions'])
MAX_FILE_SIZE = config['max_file_size_mb'] * 1024 * 1024

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE


def allowed_file(filename):
    """ファイルが許可されている拡張子かチェック"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def find_keywords_in_presentation(prs, keywords):
    """プレゼンテーション内のキーワードを検出 (OR条件)"""
    results = []
    
    for slide_num, slide in enumerate(prs.slides, 1):
        for shape_num, shape in enumerate(slide.shapes):
            if hasattr(shape, "text"):
                found_keywords = []
                total_count = 0
                
                # すべてのキーワードを検查
                for keyword in keywords:
                    if keyword.lower() in shape.text.lower():
                        count = shape.text.lower().count(keyword.lower())
                        found_keywords.append(keyword)
                        total_count += count
                
                # いずれかのキーワードが見つかった場合
                if found_keywords:
                    results.append({
                        'slide': slide_num,
                        'shape': shape_num,
                        'text': shape.text,
                        'keywords': found_keywords,
                        'count': total_count
                    })
    
    return results


def replace_text_in_shape(shape, keywords, new_text, is_delete=False):
    """シェイプ内のテキストを置換 (複数キーワード対応)"""
    if hasattr(shape, "text_frame"):
        for paragraph in shape.text_frame.paragraphs:
            # パラグラフレベルでテキスト全体を取得
            full_text = ''.join(run.text for run in paragraph.runs)
            
            # いずれかのキーワードがマッチするか確認
            has_match = any(keyword.lower() in full_text.lower() for keyword in keywords)
            
            if not has_match:
                continue
            
            # すべてのキーワードを置換（ループで処理）
            new_full_text = full_text
            for keyword in keywords:
                pattern = re.compile(re.escape(keyword), re.IGNORECASE)
                new_full_text = pattern.sub(new_text, new_full_text)
            
            # すべてのrunをクリアして新しいテキストを設定
            for run in paragraph.runs:
                run.text = ''
            
            # 新しいテキストを最初のrunに設定
            if paragraph.runs:
                paragraph.runs[0].text = new_full_text
            else:
                # runがない場合は新しく作成
                paragraph.text = new_full_text


def process_presentation(prs, keywords, new_keyword=None, is_delete=False):
    """プレゼンテーション全体を処理 (複数キーワード対応)"""
    modified_count = 0
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                original_text = shape.text
                
                # いずれかのキーワードが含まれているか確認
                has_keyword = any(kw.lower() in original_text.lower() for kw in keywords)
                
                if has_keyword:
                    # 置換先のテキストを決定
                    replacement_text = '' if is_delete else (new_keyword or keywords[0])
                    
                    # 置換を実行
                    replace_text_in_shape(shape, keywords, replacement_text, is_delete=is_delete)
                    
                    # 変更があったかを確認
                    if shape.text != original_text:
                        modified_count += 1
    
    return modified_count


@app.route('/')
def index():
    """メインページ"""
    return render_template('index.html', 
                          default_keywords=config['default_keywords'],
                          default_replacement=config['default_replacement'])


@app.route('/api/detect', methods=['POST'])
def detect_keywords():
    """キーワード検出API"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ファイルがアップロードされていません'}), 400
        
        file = request.files['file']
        keywords_json = request.form.get('keywords', '[]')
        
        try:
            keywords = json.loads(keywords_json)
        except json.JSONDecodeError:
            keywords = [keywords_json] if keywords_json else []
        
        if not keywords or len(keywords) == 0:
            return jsonify({'error': 'キーワードを入力してください'}), 400
        
        if file.filename == '':
            return jsonify({'error': 'ファイルを選択してください'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'PPTX形式のファイルをアップロードしてください'}), 400
        
        # ファイルを保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # PowerPointファイルを開く
        prs = Presentation(filepath)
        
        # キーワードを検出
        results = find_keywords_in_presentation(prs, keywords)
        
        return jsonify({
            'success': True,
            'keywords': keywords,
            'total_count': sum(r['count'] for r in results),
            'affected_slides': len(results),
            'results': results
        })
    
    except Exception as e:
        return jsonify({'error': f'エラーが発生しました: {str(e)}'}), 500


@app.route('/api/replace', methods=['POST'])
def replace_keywords():
    """キーワード置換API"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ファイルがアップロードされていません'}), 400
        
        file = request.files['file']
        keywords_json = request.form.get('keywords', '[]')
        new_keyword = request.form.get('new_keyword', '').strip()
        action = request.form.get('action', 'replace')
        
        try:
            keywords = json.loads(keywords_json)
        except json.JSONDecodeError:
            keywords = [keywords_json] if keywords_json else []
        
        if not keywords or len(keywords) == 0:
            return jsonify({'error': 'キーワードを入力してください'}), 400
        
        if file.filename == '':
            return jsonify({'error': 'ファイルを選択してください'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'PPTX形式のファイルをアップロードしてください'}), 400
        
        if action == 'replace' and not new_keyword:
            return jsonify({'error': '置換先のキーワードを入力してください'}), 400
        
        # ファイルを保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # PowerPointファイルを開く
        prs = Presentation(filepath)
        
        # 処理前の検出
        before_results = find_keywords_in_presentation(prs, keywords)
        before_count = sum(r['count'] for r in before_results)
        
        # 置換または削除を実行
        is_delete = (action == 'delete')
        modified_count = process_presentation(
            prs, 
            keywords, 
            new_keyword if not is_delete else None,
            is_delete=is_delete
        )
        
        # 処理後の検出
        after_results = find_keywords_in_presentation(prs, keywords)
        after_count = sum(r['count'] for r in after_results)
        
        # ファイルをメモリに保存
        output = BytesIO()
        prs.save(output)
        output.seek(0)
        
        # ファイルをダウンロード用に返す
        result_filename = f"modified_{filename}"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=result_filename
        )
    
    except Exception as e:
        return jsonify({'error': f'エラーが発生しました: {str(e)}'}), 500


@app.route('/api/preview', methods=['POST'])
def preview_results():
    """置換前後のプレビューAPI"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'ファイルがアップロードされていません'}), 400
        
        file = request.files['file']
        keywords_json = request.form.get('keywords', '[]')
        new_keyword = request.form.get('new_keyword', '').strip()
        action = request.form.get('action', 'replace')
        
        try:
            keywords = json.loads(keywords_json)
        except json.JSONDecodeError:
            keywords = [keywords_json] if keywords_json else []
        
        if not keywords or len(keywords) == 0:
            return jsonify({'error': 'キーワードを入力してください'}), 400
        
        if file.filename == '':
            return jsonify({'error': 'ファイルを選択してください'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'PPTX形式のファイルをアップロードしてください'}), 400
        
        if action == 'replace' and not new_keyword:
            return jsonify({'error': '置換先のキーワードを入力してください'}), 400
        
        # ファイルを保存
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # PowerPointファイルを開く
        prs = Presentation(filepath)
        
        # 処理前の検出
        before_results = find_keywords_in_presentation(prs, keywords)
        before_count = sum(r['count'] for r in before_results)
        
        # 処理を実行（プレビューのみ）
        is_delete = (action == 'delete')
        modified_count = process_presentation(
            prs, 
            keywords, 
            new_keyword if not is_delete else None,
            is_delete=is_delete
        )
        
        # 処理後の検出
        after_results = find_keywords_in_presentation(prs, keywords)
        after_count = sum(r['count'] for r in after_results)
        
        return jsonify({
            'success': True,
            'before': {
                'count': before_count,
                'slides': len(before_results)
            },
            'after': {
                'count': after_count,
                'slides': len(after_results)
            },
            'modified_shapes': modified_count,
            'action': action
        })
    
    except Exception as e:
        return jsonify({'error': f'エラーが発生しました: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
