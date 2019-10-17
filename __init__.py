# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"
    
    pip install <package> -t .

"""
import os
import sys

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'OfficeWord' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import docx2txt
from subprocess import Popen, PIPE

docto = os.path.join(cur_path.replace("libs", "bin"), "docto.exe")


def style_text(text, size, bold, ital, under):

    font = text.font
    font.size = size
    font.bold = bold
    font.italic = ital
    font.underline = under


module = GetParams("module")
global document

if module == "new":
    document = Document()

if module == "open":
    path = GetParams("path")

    document = Document(path)

if module == "read":

    result = GetParams("result")

    read_path = os.path.join(base_path, "tmp")
    try:
        os.mkdir(read_path)
    except:
        pass
    print(read_path)

    document.save(os.path.join(read_path, "tmp.docx")) #create temporal file
    text = docx2txt.process(os.path.join(read_path, "tmp.docx"))
    os.unlink(os.path.join(read_path, "tmp.docx")) #delete file

    if result:
        SetVar(result, text)

if module == "save":

    path = GetParams("path")

    if path:
        document.save(path)

if module == "write":

    text = GetParams("text")
    type_ = GetParams("type")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    ital = GetParams("italic")
    under = GetParams("underline")

    if size:
        size = Pt(int(size))
    if bold:
        bold = eval(bold)
    if ital:
        ital = eval(ital)
    if under:
        under = eval(under)

    if align == "left" or None:
        align = WD_ALIGN_PARAGRAPH.LEFT
    elif align == "center":
        align = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        align = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        align = WD_ALIGN_PARAGRAPH.JUSTIFY

    if type_ == "title":
        print("title")
        t = document.add_heading(level = 0)
        run = t.add_run(text)
        style_text(run, size, bold, ital, under)
        t.alignment = align
    elif type_ == "h1":
        t = document.add_heading(level =1)
        run = t.add_run(text)
        style_text(run, size, bold, ital, under)
        t.alignment = align
    elif type_ == "h2":
        t = document.add_heading(level = 2)
        run = t.add_run(text)
        style_text(run, size, bold, ital, under)
        t.alignment = align
    elif type_ == "p":
        texto = text.split("\\n ")
        for line in texto:
            t = document.add_paragraph()
            run = t.add_run(line)
            style_text(run, size, bold, ital, under)
            t.alignment = align
    elif type_ == "bp":
        texto = text.split("\\n ")
        for line in texto:
            t = document.add_paragraph(style='List Bullet')
            run = t.add_run(line)
            style_text(run, size, bold, ital, under)
            t.alignment = align
    elif type_ == "ln":
        texto = text.split("\\n ")
        for line in texto:
            t = document.add_paragraph(style='List Number')
            run = t.add_run(line)
            style_text(run, size, bold, ital, under)
            t.alignment = align
    else:
        raise Exception("No se ha seleccionado tipo de texto")

if module == "close":
    document = None

if module == "new_page":
    document.add_page_break()

if module == "add_pic":

    img_path = GetParams("img_path")
    document.add_picture(img_path)

if module == "to_pdf":
    try:

        from_ = GetParams("from")
        to_ = GetParams("to")

        from_ = os.path.normpath(from_)
        to_ = os.path.normpath(to_)

        options = ' -f "' + from_ + '" -O "' + to_ + '" -T wdFormatPDF'

        run_ = docto + options
        con = Popen(run_, shell=True, stdout=PIPE, stderr=PIPE)
        a = con.communicate()
    except Exception as e:
        PrintException()
        raise e