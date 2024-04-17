import docx
import lxml.etree as ET


class Node:
    def __init__(self, id, tag, name, parent=None, value=None):
        self.id = id
        self.tag = tag
        self.name = name
        self.parent = parent
        self.value = value
        self.children = []


def parse_content_description_to_tree(doc):

    tocs = styles['toc 1'] + styles['toc 2'] + styles['toc 3']


    root = Node(0, 'main', 'ROOT')
    first_level_id = 1

    in_content_description = False
    for i in range(len(doc.paragraphs)):
        para = doc.paragraphs[i]
        print(para.style.name)
        if in_content_description:
            s_name = para.style.name

            if s_name not in tocs:
                return root
            if s_name in styles['toc 1']:
                node = Node(first_level_id, 'sec', ' '.join(para.text.split()[:-1]), root)
                root.children.append(node)
                first_level_id += 1
            else:
                id = para.text.split()[0].split('.')
                parent = root
                for i in range(len(id) - 1):
                    for ch in parent.children:
                        if ch.id == id[i]:
                            parent = ch
                            break
                node = Node(id[-1], 'subsec', ' '.join(para.text.split()[:-1]), parent)
                parent.children.append(node)

        if para.style.name in styles['content_description']:
            in_content_description = True


def parse_content_description_to_etree(doc):

    tocs = styles['toc 1'] + styles['toc 2'] + styles['toc 3']


    root = ET.Element('content_description')
    last_first_level = None
    last_second_level = None

    in_content_description = False
    for i in range(len(doc.paragraphs)):
        para = doc.paragraphs[i]
        print(para.style.name)
        if in_content_description:
            s_name = para.style.name

            if s_name not in tocs:
                return root

            if s_name in styles['toc 1']:
                node = ET.SubElement(root, 'sec')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

                last_first_level = node

            if s_name in styles['toc 2']:
                node = ET.SubElement(last_first_level, 'subsec1')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

                last_second_level = node

            if s_name in styles['toc 3']:
                node = ET.SubElement(last_second_level, 'subsec2')
                name = ' '.join(para.text.split()[:-1])
                if name[0].isdigit():
                    name = ' '.join(name.split()[1:])
                node.set('name', name)

        if para.style.name in styles['content_description']:
            in_content_description = True
    return root


def parse_em_all(doc):
    root = parse_content_description_to_etree(doc)
    prev = root

    tocs = styles['toc 1'] + styles['toc 2'] + styles['toc 3']

    levels = {
        'sec': 1,
        'subsec1': 2,
        'subsec2': 3,
    }

    doc_index = 0

    for elem in root.iter():
        if elem == root:
            continue
        if prev == root or levels[prev.tag] < levels[elem.tag]:
            prev = elem
            continue

        print(elem.get('name'))
        doc_index = 0

        while doc_index < len(doc.paragraphs):
            if prev.get('name') in doc.paragraphs[doc_index].text and doc.paragraphs[doc_index].style.name not in tocs:
                in_iter = False
                s_name = doc.paragraphs[doc_index].style.name
                iter_start = prev
                iter1_start = prev
                numpstart = prev
                doc_index += 1
                for i in range(doc_index, len(doc.paragraphs)):
                    para = doc.paragraphs[i]
                    print(para.style.name)
                    if para.style.name == s_name:
                        doc_index = i
                        break

                    if para.style.name in styles['nump 2']:
                        node = ET.SubElement(prev, 'levelledPara')
                        node.text = para.text
                        numpstart = node
                        continue
                    elif para.style.name in styles['nump 3']:
                        node = ET.SubElement(prev, 'levelledPara')
                        node.text = para.text
                        numpstart = node
                        continue
                    else:
                        numpstart = prev

                    if para.style.name in styles['iter 1']:
                        if not in_iter:
                            in_iter = True
                            iter_start = ET.SubElement(prev, 'sequentialList')
                        node = ET.SubElement(iter_start, 'listItem')
                        node.text = para.text
                        iter1_start = node
                        continue

                    if para.style.name in styles['iter 2']:
                        node = ET.SubElement(iter1_start, 'listItem')
                        node.text = para.text
                        continue

                    if para.style.name in styles['iter alt']:
                        if not in_iter:
                            in_iter = True
                            iter_start = ET.SubElement(prev, 'sequentialList')
                        if iter_start.tag in 'iter alt':
                            node = ET.SubElement(iter_start, 'listItem')
                        else:
                            node = ET.SubElement(iter1_start, 'listItem')
                        node.text = para.text
                        continue

                    in_iter = False

                    if para.style.name in styles['normal']:
                        node = ET.SubElement(numpstart, 'para')
                        node.text = para.text
                        continue

                    if styles['extra'][0] in para.style.name:
                        node = ET.SubElement(prev, 'extra')
                        node.text = para.text
                        numpstart = node

                    node = ET.SubElement(prev, para.style.name[:3])
                    node.text = para.text
                break
            doc_index += 1
        prev = elem

    return root

styles = {
    'content_description': ['Название-caps'],  # Содержание
    'toc 1': ['toc 1'],
    'toc 2': ['toc 2'],
    'toc 3': ['toc 3'],  # Перечисления в содержании (индексация длины 1, 2 и 3 соответственно)
    'subsec v': ['Нумерованный заголовок 3'],
    'iter 1': ["Перечисление-1"],
    'iter 2': ["Перечисление 2 уровень"],
    'iter alt': ["Перечень"],
    'nump 2': ["Нумерованный абзац 2"],
    'nump 3': ["Нумерованный абзац 3"],
    'normal': ["Normal"],
    'extra': ["Приложение"],
}

module_codees = {
    "DMC-VBMA-A-46-20-01-00A-018A-A_000_01_ru_RU.xml": "Введение",
    "DMC-VBMA-A-46-20-01-00A-020A-A_000_01_ru_RU.xml": "Назначение",
    "DMC-VBMA-A-46-20-01-00A-030A-A_000_01_ru_RU.xml": "Технические характеристики",
    "DMC-VBMA-A-46-20-01-00A-034A-A_000_01_ru_RU.xml": "Состав изделия",
    "DMC-VBMA-A-46-20-01-00A-041A-A_000_01_ru_RU.xml": "Устройство и работа",
    "DMC-VBMA-A-46-20-01-00A-044A-A_000_01_ru_RU.xml": "Описание и работа составных частей изделия",
    "DMC-VBMA-A-46-20-01-00A-122A-A_000_01_ru_RU.xml": "Указания по включению и опробованию работы изделия",
    "DMC-VBMA-A-46-20-01-00A-123A-A_000_01_ru_RU.xml": "Установка и настройка программного обеспечения",
    "DMC-VBMA-A-46-20-01-00A-410A-A_000_01_ru_RU.xml": "Перечень возможных неисправностей в процессе использования изделия и рекомендации по действиям при их возникновении",
}


def run_parser(doc):
    root = parse_em_all(doc)

    et = ET.ElementTree(root)
    et.write('output.xml', encoding="utf-8", pretty_print=True)

    root = parse_content_description_to_etree(doc)
    et = ET.ElementTree(root)
    et.write('output2.xml', encoding="utf-8", pretty_print=True)


doc = docx.Document("input.docx")
run_parser(doc)
