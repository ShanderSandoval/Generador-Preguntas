import re
from xml.sax.saxutils import escape
from docx import Document

class MoodleXMLGenerator:
    def __init__(self):
        self.xml_output = []
        self._start_document()

    def _start_document(self):
        self.xml_output.append('<?xml version="1.0" encoding="UTF-8"?>\n<quiz>\n')

    def _end_document(self):
        self.xml_output.append('</quiz>')

    def wrap_cdata(self, text):
        return f'<![CDATA[<p>{text}</p>]]>'

    def add_multiple_choice(self, name, text, options):
        xml = f'''
  <question type="multichoice">
    <name><text>{escape(name)}</text></name>
    <questiontext format="html">
      <text>{self.wrap_cdata(text)}</text>
    </questiontext>
    <single>true</single>
    <shuffleanswers>true</shuffleanswers>
    <answernumbering>abc</answernumbering>'''
        
        for opt in options:
            fraction = "100" if opt.get('is_correct') else "0"
            xml += f'''
    <answer fraction="{fraction}" format="html">
      <text>{self.wrap_cdata(opt['text'])}</text>
    </answer>'''
        
        xml += '\n  </question>\n'
        self.xml_output.append(xml)

    def add_true_false(self, name, text, is_true_correct):
        fraction_true = "100" if is_true_correct else "0"
        fraction_false = "0" if is_true_correct else "100"
        
        xml = f'''
  <question type="truefalse">
    <name><text>{escape(name)}</text></name>
    <questiontext format="html">
      <text>{self.wrap_cdata(text)}</text>
    </questiontext>
    <answer fraction="{fraction_true}">
      <text>true</text>
    </answer>
    <answer fraction="{fraction_false}">
      <text>false</text>
    </answer>
  </question>\n'''
        self.xml_output.append(xml)

    def get_final_xml(self):
        self._end_document()
        return "".join(self.xml_output)

class SmartQuizParser:
    def __init__(self):
        self.questions = []
        self.generator = MoodleXMLGenerator()

    def preprocess_text(self, text):
        text = re.sub(r'([a-zñáéíóú])([A-E]\))', r'\1\n\2', text)
        text = re.sub(r'([a-zñáéíóú]\.)\s*([a-e][\)\.])\s', r'\1\n\2 ', text)
        return text

    def parse_document(self, raw_text, log_callback=None):
        raw_text = self.preprocess_text(raw_text)
        lines = [line.strip() for line in raw_text.split('\n') if line.strip()]
        
        current_q = None
        preguntas_en_este_doc = 0
        
        for line in lines:
            # 1. Detectar inicio de pregunta
            q_match = re.match(r'^\d+[\.\-\)]\s*(.*)', line)
            if q_match:
                if current_q:
                    self.questions.append(current_q)
                    preguntas_en_este_doc += 1
                    if log_callback:
                        log_callback(f"  ✓ Detectada: {current_q['text'][:50]}...")
                
                current_q = {
                    'text': q_match.group(1),
                    'options': [],
                    'correct_letter': None
                }
                continue

            if not current_q: continue

            # 2. Detectar si la línea marca la respuesta correcta al final (Incluso si está vacía)
            ans_match = re.match(r'^Respuest[a-z\s]*:\s*([a-zA-Z]?)', line, re.IGNORECASE)
            if ans_match:
                letra_detectada = ans_match.group(1).strip()
                if letra_detectada: 
                    current_q['correct_letter'] = letra_detectada.lower()
                continue

            # 3. Detectar opciones
            opt_match = re.match(r'^\[?\s*([xX\*]?)\s*\]?\s*([a-eA-E])[\.\)]\s*(.*)', line)
            if opt_match:
                is_marked_correct = opt_match.group(1).lower() in ['x', '*']
                letter = opt_match.group(2).lower()
                opt_text = opt_match.group(3)

                current_q['options'].append({
                    'letter': letter,
                    'text': opt_text,
                    'is_correct': is_marked_correct
                })
                
                if is_marked_correct:
                    current_q['correct_letter'] = letter
                continue

            # 4. Destruir encabezados sueltos
            if re.match(r'^[\W_]*(secci[óo]n|bloque|instrucciones|cuestionario|banco|tema)\b', line, re.IGNORECASE):
                continue

            # 5. Continuación de texto
            if not current_q['options']:
                current_q['text'] += "\n" + line
            else:
                current_q['options'][-1]['text'] += "\n" + line

        if current_q:
            self.questions.append(current_q)
            preguntas_en_este_doc += 1
            if log_callback:
                log_callback(f"  ✓ Detectada: {current_q['text'][:50]}...")

        return preguntas_en_este_doc

    def generate_all_xml(self, log_callback=None):
        for q in self.questions:
            correct_letter = q.get('correct_letter')
            for opt in q['options']:
                if correct_letter and opt['letter'] == correct_letter:
                    opt['is_correct'] = True

            if not any(opt.get('is_correct') for opt in q['options']) and q['options']:
                q['options'][0]['is_correct'] = True 
                if log_callback:
                    log_callback(f"  [!] Advertencia: La pregunta '{q['text'][:20]}...' no tenía respuesta. Se asignó la opción A.")

            q_type = 'multichoice'
            is_true_correct = False
            
            if len(q['options']) == 2:
                texts = [o['text'].lower() for o in q['options']]
                if any("verdadero" in t or "true" in t for t in texts):
                    q_type = 'truefalse'
                    for opt in q['options']:
                        if opt.get('is_correct') and ("verdadero" in opt['text'].lower() or "true" in opt['text'].lower()):
                            is_true_correct = True
            
            # Limite de 250 caracteres para aprovechar el nombre en Moodle
            name = q['text'][:250].replace('\n', ' ')

            if q_type == 'multichoice' and q['options']: 
                self.generator.add_multiple_choice(name, q['text'], q['options'])
            elif q_type == 'truefalse':
                self.generator.add_true_false(name, q['text'], is_true_correct)

        return self.generator.get_final_xml()

def read_docx(file_path):
    doc = Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs])