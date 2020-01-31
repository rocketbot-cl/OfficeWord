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
from docx.oxml.shared import qn
import docx
from xml.etree import ElementTree
from lxml import etree

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

if module == "readTable":

    result = GetParams("result")
    tablesDoc = []
    for table in document.tables:
        table_ = []
        for row in table.rows:
            array_row = []
            for cell in row.cells:
                if len(array_row) > 0:
                    if array_row[-1] != cell.text:
                        array_row.append(cell.text)
                else:
                    array_row.append(cell.text)
            table_.append(array_row)
        tableDoc.append(table_)
    if result:
        SetVar(result, tableDoc)

if module == "addTextBookmark":

    import copy

    bookmark_searched = GetParams("bookmark")
    text = GetParams("text")
    clean = GetParams("Clean")
    print(clean)

    try:
        ele = document._element[0]
        bookmarks_list = ele.findall('.//' + qn('w:bookmarkStart'))
        for bookmark in bookmarks_list:
            name = bookmark.get(qn('w:name'))
            if name == bookmark_searched:
                # get parent and search value
                next_el = bookmark.getnext()
                if next_el.get(qn('w:name')) == "_GoBack":
                    next_el = next_el.getnext()

                if clean:
                    next_el.find(qn('w:t')).text = text
                else:
                    next_el.find(qn('w:t')).text += str(text)
                break
            else:
                name = False
        if not name:
            raise Exception("Bookmark not found")

    except Exception as e:
        PrintException()
        raise e

if module == "save":

    path = GetParams("path")

    if path:
        if not path.endswith(".docx"):
            path += ".docx"
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

        if not to_.endswith(".pdf"):
            to_ += ".pdf"

        from_ = os.path.normpath(from_)
        to_ = os.path.normpath(to_)

        options = ' -f "' + from_ + '" -O "' + to_ + '" -T wdFormatPDF'

        run_ = docto + options
        con = Popen(run_, shell=True, stdout=PIPE, stderr=PIPE)
        a = con.communicate()
    except Exception as e:
        PrintException()
        raise e

## Modificado por Mijahil Franchi: Ubica el parrafo en el que se encuentra un texto
if module == "search_text":
    parrafos = document.paragraphs
    text_buscar = GetParams("text_search")
    variable = GetParams("variable")
    posicion = 0
    posiciones = list()
    try:
        for parrafo in parrafos:
            if text_buscar in parrafo.text:
                posiciones.append(posicion)
            posicion = posicion+1
    except Exception as e:
        PrintException()
        raise e

    if posiciones:
        SetVar(variable, posiciones)

## Modificado por Mijahil Franchi: Cuenta los parrafos de un documento
if module == "count_paragraphs":
    variable = GetParams("variable")
    
    try:
        parrafos = document.paragraphs
        cantidad = len(parrafos)
    except Exception as e:
        PrintException()
        raise e

    if cantidad:
        SetVar(variable, cantidad)

## Modificado por Mijahil Franchi: Busca y remplaza el contenido de un texto
if module == "search_replace_text":
    
    variable = GetParams("variable")
    parrafos = GetParams("parrafos")
    buscar = GetParams("text_search")
    remplazar = GetParams("text_replace")
    resultado = False
    posicion = 0
    parrafos_ini = document.paragraphs

    try:
        if parrafos:
            parrafos = parrafos.split(',')
            for parrafo in parrafos:
                parrafo = int(parrafo)
                text_parrafo = parrafos_ini[parrafo].text
                if buscar in text_parrafo:
                    texto = text_parrafo
                    texto = texto.replace(buscar, remplazar)
                    parrafos_ini[parrafo].text = texto
                    resultado = True
        else:
            print("esta vacio el string")
            for parrafo in parrafos_ini:
                if buscar in parrafo.text:
                    texto = parrafo.text
                    texto = texto.replace(buscar, remplazar)
                    parrafos_ini[posicion].text = texto
                    resultado = True
                posicion = posicion+1

        SetVar(variable, resultado)
            
    except Exception as e:
        SetVar(variable, False)
        PrintException()
        raise e
    
    