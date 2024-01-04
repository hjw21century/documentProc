import docx

doc = docx.Document('demo.docx')
print(f'paragraphs len: {len(doc.paragraphs)}')

print(f'tables len: {len(doc.tables)}')

print(f'sections len: {len(doc.sections)}')

print(f'styles len: {len(doc.styles)}')

for i in range(len(doc.paragraphs)):
    print(f'paragraphs text: {doc.paragraphs[i].text}')
    print(f'paragraphs style: {doc.paragraphs[i].style.name}')
    print(f'paragraphs style font color: {doc.paragraphs[i].style.font.color.rgb}')
    print(f'paragraphs style font italic: {doc.paragraphs[i].style.font.italic}')
    print(f'paragraphs style font bold: {doc.paragraphs[i].style.font.bold}')
    


for style in doc.styles:
    if style.font.italic and style.font.bold:
        print(f'doc.styles name: {style.name}')
        print(f'doc.styles font color: {style.font.color.rgb}')
        print(f'doc.styles font italic: {style.font.italic}')
        print(f'doc.styles font bold: {style.font.bold}')
    