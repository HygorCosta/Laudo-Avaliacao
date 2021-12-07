from docx import Document

# Source CSV - columns names that must match the * that are {{***}} inside
doc = Document(r"C:\Users\hygorsilva\OneDrive\Compesa\PyValuation\templates\TemplateDesapropriacao.docx")

replacements = {
    '{NUM_LAUDO}': '165/2021',
    '{NUM_SEI}': '0000000.000-11',
    '{SOLICITANTE}': 'HYGOR COSTA',
    '{TIPO}': 'desapropriação',
    '{VTOTAL}': 'R$ 11.000,00',
    }

for paragraph in doc.paragraphs:
    for key in replacements:
        paragraph.text = paragraph.text.replace(key, replacements[key])

doc.save('teste.docx')