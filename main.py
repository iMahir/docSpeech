import argparse
import fitz # PyMuPDF
from docx import Document
import asyncio
import edge_tts

def read_docx(file_path):
   doc = Document(file_path)
   return ' '.join([paragraph.text for paragraph in doc.paragraphs])

def read_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""

    for page in doc:
        blocks = page.get_text("blocks")
        blocks.sort(key=lambda block: -block[1])  # sort y-coordinate (top to bottom)

        for block in blocks:
            text += block[4] + '\n'

    return text


def create_arg_parser():
   parser = argparse.ArgumentParser()
   parser.add_argument('file_path', help='Path to the document file')
   parser.add_argument('output_path', help='Path to the output audio file')
   parser.add_argument('--language', default='en-GB-SoniaNeural', help='Language for the text-to-speech conversion')
   return parser

async def main():
   parser = create_arg_parser()
   args = parser.parse_args()

   if args.file_path.endswith('.docx'):
       text = read_docx(args.file_path)
   elif args.file_path.endswith('.pdf'):
       text = read_pdf(args.file_path)
   else:
       print('Unsupported file format')
       return

   communicate = edge_tts.Communicate(text, args.language)
   await communicate.save(args.output_path)

if __name__ == '__main__':
   loop = asyncio.get_event_loop_policy().get_event_loop()
   try:
       loop.run_until_complete(main())
   finally:
       loop.close()
