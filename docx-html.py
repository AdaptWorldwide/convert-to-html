import mammoth
import glob
import re

HTMLStrings = ['strong','em']

for file in glob.glob('*.docx'):
    try:
        filename = file.replace('.docx', '.txt')
        result = mammoth.convert_to_html(file)
        html = result.value
        for tag in HTMLStrings:
            startTag = '<{}>'.format(tag)
            endTag = '</{}>'.format(tag)
            html = str(html).replace(startTag,'')
            html = str(html).replace(endTag,'')

        html = re.sub(r'<a.*id.*</a>', '',str(html))
        with open(filename, 'a') as outputfile:
            outputfile.write(html)
    except:
        with open('ErrorLog.txt','a',encoding='utf-8') as ErrorFile:
            ErrorFile.write('{}\n'.format(filename))
