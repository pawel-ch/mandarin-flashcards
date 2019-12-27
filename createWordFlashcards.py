"""
    To run, have FlashcardTemplate.docx and flashcard-input.txt in the script directory. 
    flashcard-input.txt should have a list of Chinese characters/words with optional example sentences, 
    one entry per line, the characters separated from the example with white space (tab, space, etc) like this:
    高   大人在高高的山。
    本來
    悟空  悟空打了個妖怪。
    
    The output is a .docx file with just the characters/words on one side and pinyin + example on the other side.
    The formatting is meant for printing on pre-perforating business card paper. 
"""
import codecs
import pinyin
import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_ALIGN_VERTICAL

vocab = []
with codecs.open("flashcard-input.txt", "r", "utf-8-sig") as file:
    for line in file.read().splitlines():
        entry = line.split()
        vocab.append(entry)

pinyin_list = []
for zi_with_example in vocab:
    zi = zi_with_example[0]
    example = ""
    if len(zi_with_example) > 1:
        example = zi_with_example[1:]
    pinyin_list.append([pinyin.get(zi),example])

doc = Document('FlashcardTemplate.docx')
doc_table = doc.tables[0]

hanzi_style = doc.styles.add_style('Hanzi Style', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style.font.name = 'DFKai-SB'
hanzi_style_115 = doc.styles.add_style('Hanzi Style 115', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style_115.font.size = Pt(115)
hanzi_style_115.base_style = doc.styles['Hanzi Style']
hanzi_style_100 = doc.styles.add_style('Hanzi Style 100', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style_100.font.size = Pt(100)
hanzi_style_100.base_style = doc.styles['Hanzi Style']
hanzi_style_80 = doc.styles.add_style('Hanzi Style 80', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style_80.font.size = Pt(80)
hanzi_style_80.base_style = doc.styles['Hanzi Style']
hanzi_style_60 = doc.styles.add_style('Hanzi Style 60', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style_60.font.size = Pt(60)
hanzi_style_60.base_style = doc.styles['Hanzi Style']
hanzi_style_45 = doc.styles.add_style('Hanzi Style 45', WD_STYLE_TYPE.PARAGRAPH)
hanzi_style_45.font.size = Pt(45)
hanzi_style_45.base_style = doc.styles['Hanzi Style']
pinyin_style = doc.styles.add_style("Pinyin", WD_STYLE_TYPE.PARAGRAPH)
pinyin_style.base_style = doc.styles['Normal']
pinyin_style.font.size = Pt(25)
example_style = doc.styles.add_style('Example', WD_STYLE_TYPE.PARAGRAPH)
example_style.base_style = doc.styles['Hanzi Style']
example_style.font.size = Pt(20)

char_count_to_style = {1:'Hanzi Style 115',2:'Hanzi Style 100',3:'Hanzi Style 80',4:'Hanzi Style 60',5:'Hanzi Style 45'}

lists_of_10_vocab_items = [vocab[i * 10:(i + 1) * 10] for i in range((len(vocab) + 9) // 10 )]
lists_of_10_pinyin = [pinyin_list[i * 10:(i + 1) * 10] for i in range((len(pinyin_list) + 9) // 10 )]
hanzi_10 = lists_of_10_vocab_items[0]
pinyin_10 = lists_of_10_pinyin[0]

print(doc.tables[0].rows[0].height_rule)

def process_page(hanzi_10, pinyin_10, offset):
    print("Processing word list of length="+ str(len(hanzi_10)))
    for i in range(0,len(hanzi_10)):
        (r,c) = (int((i+offset)/2), i % 2)
        hanzi_p = doc_table.cell(r,c).paragraphs[0]
        hanzi_p.text = hanzi_10[i][0]
        hanzi_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        char_count = len(hanzi_10[i][0])
        new_style = char_count_to_style[char_count]
        hanzi_p.style = new_style
        (r,c) = (int((i+offset)/2) + 5, (i+1) % 2)
        pinyin_cell = doc_table.cell(r,c)
        pinyin_p = pinyin_cell.add_paragraph(pinyin_10[i][0],style=doc.styles['Pinyin'])
        pinyin_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if len(pinyin_10[i]) > 1:
            for p in pinyin_10[i][1:]:
                para = pinyin_cell.add_paragraph(p,style=doc.styles['Example'])
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

process_page(hanzi_10,pinyin_10,0)

offset = 0
if len(lists_of_10_vocab_items) > 1:
    for hanzi_10,pinyin_10 in zip(lists_of_10_vocab_items[1:],lists_of_10_pinyin[1:]):
        for i in range(10):
            new_row = doc_table.add_row()
            new_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            new_row.height = Inches(2)
            for new_cell in new_row.cells:
                new_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        offset += 20
        process_page(hanzi_10,pinyin_10,offset)
    
time_str = str(datetime.datetime.now().strftime("%Y%m%d_%H-%M-%S"))
outputDocName = "flashcards-" + time_str + ".docx"
doc.save(outputDocName)