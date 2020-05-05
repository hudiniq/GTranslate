from os import path
import urllib.request
import pypandoc
import docx
from googletrans import Translator
import docx2txt

def main():
    doc_final = docx.Document()
    translator = Translator()

    print("Vnesite ime .docx datoteke, oz. URL:")
    while True:
        try:
            usersource = init_source(input())
        except FileNotFoundError:
            print("\nDatoteka ne obstaja ali je v napačnem formatu.\nProsim za ponovni vnos:")
            continue
        break

    result = docx2txt.process(usersource)
    translated = translator.translate(result, dest='en')

    doc_final.add_paragraph(translated.text)

    doc_final.save("translated.docx")

def init_source(source):
    try:
        html = urllib.request.urlopen(source).read()
    except ValueError:
        if ".docx" not in source:
            source = source + ".docx"

        if not path.exists(source):
            raise FileNotFoundError

        return source

    while True:
        try:
            pypandoc.convert_text(source=html, format='html', to='docx', outputfile="text_pre_translate.docx", extra_args=['-RTS'])
        except OSError:
            pypandoc.download_pandoc()
            continue
        break

    return "text_pre_translate.docx"


if __name__ == "__main__":
    main()