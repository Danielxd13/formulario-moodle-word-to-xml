import pandas as pd
from xml.etree import ElementTree as ET
from xml.dom import minidom
import argparse

def read_excel_file(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print(f"Error: The file {file_path} was not found.")
        return None
    except Exception as e:
        print(f"Error reading the Excel file: {str(e)}")
        return None

def create_xml_question(row):
    """Crea un elemento question con la estructura requerida"""
    question = ET.Element('question')
    question.set('type', 'multichoice')
    
    # Añadir nombre
    name = ET.SubElement(question, 'name')
    name_text = ET.SubElement(name, 'text')
    name_text.text = f'<![CDATA[<p>{row["ID"]}</p>]]>'
    
    # Añadir texto de la pregunta
    questiontext = ET.SubElement(question, 'questiontext')
    questiontext.set('format', 'html')
    question_text = ET.SubElement(questiontext, 'text')
    question_text.text = f'<![CDATA[<p>{row["Preguntas"]}</p>]]>'
    
    # Elementos fijos
    generalfeedback = ET.SubElement(question, 'generalfeedback')
    generalfeedback.set('format', 'html')
    feedback_text = ET.SubElement(generalfeedback, 'text')
    feedback_text.text = ''
    
    defaultgrade = ET.SubElement(question, 'defaultgrade')
    defaultgrade.text = '1.0000000'
    penalty = ET.SubElement(question, 'penalty')
    penalty.text = '0.3333333'
    hidden = ET.SubElement(question, 'hidden')
    hidden.text = '0'
    single = ET.SubElement(question, 'single')
    single.text = 'true'
    shuffleanswers = ET.SubElement(question, 'shuffleanswers')
    shuffleanswers.text = 'false'
    answernumbering = ET.SubElement(question, 'answernumbering')
    answernumbering.text = 'abc'
    correctfeedback = ET.SubElement(question, 'correctfeedback')
    correctfeedback.set('format', 'html')
    correctfeedback_text = ET.SubElement(correctfeedback, 'text')
    correctfeedback_text.text = f'<![CDATA[<p>¡Bien respondido!</p>]]>'
    partiallycorrectfeedback = ET.SubElement(question, 'partiallycorrectfeedback')
    partiallycorrectfeedback.set('format', 'html')
    partiallycorrectfeedback_text = ET.SubElement(partiallycorrectfeedback, 'text')
    partiallycorrectfeedback_text.text = f'<![CDATA[<p>¡Debes prestar más atención!</p>]]>'
    incorrectfeedback = ET.SubElement(question, 'incorrectfeedback')
    incorrectfeedback.set('format', 'html')
    incorrectfeedback_text = ET.SubElement(incorrectfeedback, 'text')
    incorrectfeedback_text.text = f'<![CDATA[<p>¡Puedes hacerlo mejor que esto!</p>]]>'
    ET.SubElement(question, 'shownumcorrect')
    
        # Añadir respuestas
    for i, resp in enumerate(['Respuesta1', 'Respuesta2', 'Respuesta3', 'Respuesta4']):
        if resp in row and row[resp]:
            answer = ET.SubElement(question, 'answer')
            # Si es la respuesta correcta
            if i + 1 == row.get('Respuesta correcta', 0):
                answer.set('fraction', '100.00000')
            else:
                answer.set('fraction', '-25.00000')
            answer.set('format', 'html')
            text = ET.SubElement(answer, 'text')
            text.text = f'<![CDATA[<p>{row[resp]}</p>]]>'
            feedback = ET.SubElement(answer, 'feedback')
            feedback.set('format', 'html')
            ET.SubElement(feedback, 'text')
            
            
    
    return question

def convert_to_xml(data, output_file):
    try:
        # Crear elemento raíz
        root = ET.Element('quiz')
        
        # Añadir cada pregunta
        for row in data:
            question = create_xml_question(row)
            root.append(question)
        
        # Convertir a string y formatear
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent='\t')
        
        # Reemplazar las codificaciones no deseadas
        xml_str = xml_str.replace('&lt;', '<').replace('&gt;', '>')
        
        # Eliminar la declaración XML que añade minidom
        xml_str = xml_str.split('\n', 1)[1]
        
        # Añadir declaración XML al inicio
        xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' + xml_str
        
        # Guardar archivo
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(xml_str)
            
        return xml_str
    except Exception as e:
        print(f"Error converting to XML: {str(e)}")
        return None

def main():
    # Crear el parser de argumentos
    parser = argparse.ArgumentParser(description='Convierte un archivo Excel a formato XML para Moodle.')
    parser.add_argument('--excel_path', '-e', 
                       help='Ruta al archivo Excel que se va a procesar',
                       default='Contenido_docx.xlsx')
    parser.add_argument('--salida', '-s', 
                       help='Nombre del archivo XML de salida',
                       default="output.xml")
    
    # Parsear los argumentos
    args = parser.parse_args()
    
    # Leer el archivo Excel
    data = read_excel_file(args.excel_path)
    if data is not None:
        # Convertir DataFrame a lista de diccionarios
        dict_records = data.to_dict('records')
        
        # Convertir a XML
        xml_data = convert_to_xml(dict_records, args.salida)
        if xml_data:
            print(f"\nArchivo XML creado exitosamente: {args.salida}")

if __name__ == "__main__":
    main()