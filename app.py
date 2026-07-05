
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from transformers import AutoTokenizer, AutoModelForSequenceClassification
import torch
import re

app = Flask(__name__)
CORS(app)

# تحميل AraBERT
print("جاري تحميل نموذج AraBERT...")
model_path = 'arabert_final_model'
tokenizer = AutoTokenizer.from_pretrained(model_path)
model = AutoModelForSequenceClassification.from_pretrained(model_path)
device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
model = model.to(device)
model.eval()
print(f"✅ تم تحميل النموذج على {device}")

def clean_arabic_text(text):
    if not isinstance(text, str):
        return ""
    text = re.sub(r'http\S+|www\S+', '', text)
    text = re.sub(r'@\w+|#\w+', '', text)
    text = re.sub(r'\d+', '', text)
    text = re.sub(r'[\u064B-\u065F]', '', text)
    text = re.sub(r'[أإآا]', 'ا', text)
    text = re.sub(r'(.)\1+', r'\1', text)
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/predict', methods=['POST'])
def predict():
    data = request.get_json()
    text = data.get('text', '')

    if not text:
        return jsonify({'error': 'لا يوجد نص'}), 400

    # معالجة
    cleaned = clean_arabic_text(text)

    # Tokenization
    inputs = tokenizer(cleaned, return_tensors='pt', 
                      truncation=True, padding=True, max_length=128)
    inputs = {k: v.to(device) for k, v in inputs.items()}

    # التنبؤ
    with torch.no_grad():
        outputs = model(**inputs)
        probs = torch.nn.functional.softmax(outputs.logits, dim=1)
        prediction = torch.argmax(probs, dim=1).item()
        confidence = probs[0][prediction].item() * 100

    result = {
        'prediction': 'كاذب' if prediction == 1 else 'صادق',
        'confidence': float(confidence),
        'label': int(prediction)
    }

    return jsonify(result)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
