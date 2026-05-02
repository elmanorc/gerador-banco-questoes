import codecs

with codecs.open('geradorBancosDeQuestoesPorTopico.py', 'r', 'utf-8') as f:
    text = f.read()

text = text.replace('\\"\\"\\"', '"""')
text = text.replace('\\n', '\n')

with codecs.open('geradorBancosDeQuestoesPorTopico.py', 'w', 'utf-8') as f:
    f.write(text)
