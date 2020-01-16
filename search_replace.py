import glob
import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile as ZF


def replace_namespaces_method(hardcoded_namespaces):
    old_method = None

    def _namespace_replacement(elem, default_namespace=None):
        qnames, namespaces = old_method(elem, default_namespace)
        flipped = {}
        for key, val in hardcoded_namespaces.items():
            flipped[val] = key

        return qnames, flipped

    old_method = ET._namespaces
    ET._namespaces = _namespace_replacement


def update_zip(zipname, zip_info, data):
    # generate a temp file
    tmp_fd, tmp_name = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmp_fd)

    # create a temp copy of the archive without filename
    with ZF.ZipFile(zipname, 'r') as z_in:
        with ZF.ZipFile(tmp_name, 'w') as z_out:
            z_out.comment = z_in.comment  # preserve the comment
            for item in z_in.infolist():
                if item.filename != zip_info.filename:
                    z_out.writestr(item, z_in.read(item.filename))

    # replace with the temp archive
    os.remove(zipname)
    os.rename(tmp_name, zipname)

    # now add filename with its new data
    with ZF.ZipFile(zipname, mode='a') as zf:
        zf.writestr(zip_info, data)


class DocxRedactor:
    def __init__(self):
        self.is_open = False
        self.namespaces = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'o': 'urn:schemas-microsoft-com:office:office',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w10': 'urn:schemas-microsoft-com:office:word',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
            'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
            'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
            'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
        }
        self.path = None
        self.doc_info = None
        self.root = None
        self.parent_map = None

        replace_namespaces_method(self.namespaces)

        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)

    def open(self, file_path):
        if not os.path.isfile(file_path):
            raise Exception('File \'%s\' does not exist'.format(file_path))

        self.path = file_path

        with ZF.ZipFile(self.path, 'r') as zip_file:
            self.doc_info = zip_file.getinfo('word/document.xml')
            with zip_file.open(self.doc_info) as document:
                self.root = ET.fromstring(document.read())

        self.parent_map = dict((c, p) for p in self.root.getiterator() for c in p)
        self.is_open = True

    def check_open(self):
        if not self.is_open:
            raise Exception("You have to open a file first")

    def save(self):
        self.check_open()
        data = ET.tostring(self.root)

        update_zip(self.path, self.doc_info, data)

    def get_all_colors(self):
        self.check_open()
        
        highlights = self.get_highlights()
        colors = sorted(frozenset(map(lambda h: h.attrib[self.expand('w:val')], highlights)))
        return colors

    def redact(self, color, replacement):
        self.check_open()

        for highlight in self.get_highlights(color=color):
            run = self.get_run_or_paragraph_for_highlight(highlight)
            if run is None:
                print('Failed to successfully redact a highlight')
                continue
            self.replace_text_in_run_or_paragraph(run, replacement)

    def get_highlights(self, color=None):
        attribute = self.expand('w:val')
        highlights = list(self.root.iter(self.expand('w:highlight')))

        if color is None:
            return highlights
        else:
            return list(filter(lambda h: h.attrib[attribute] == color, highlights))

    def get_run_or_paragraph_for_highlight(self, highlight):
        elem = highlight
        while True:
            try:
                elem = self.parent_map[elem]
                if elem.tag == self.expand('w:r') or elem.tag == self.expand('w:p'):
                    return elem
            except KeyError:
                return None

    def replace_text_in_run_or_paragraph(self, elem, text):
        if elem.tag == self.expand('w:r'):
            rPrs = list(elem.iter(self.expand('w:rPr')))
            for rPr in rPrs:
                elem.remove(rPr)

        if elem.tag == self.expand('w:p'):
            pPrs = list(elem.iter(self.expand('w:pPr')))
            for pPr in pPrs:
                elem.remove(pPr)

        for text_node in elem.iter(self.expand('w:t')):
            text_node.text = text

    def expand(self, tag):
        prefix, uri = tag.split(":", 1)
        try:
            prefix, uri = tag.split(":", 1)
            return "{%s}%s" % (self.namespaces[prefix], uri)
        except KeyError:
            raise SyntaxError("prefix %r not found in prefix map" % prefix)


def clear():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


def choose_action():
    print('(l) Choose file to redact by list')
    print('(p) Choose file to redact by path')
    print('(q) To quit (press CTRL+C) to quit anytime')

    cmd = input('> ')
    if cmd == 'l':
        return choose_by_list
    elif cmd == 'p':
        return choose_by_path
    elif cmd == 'q':
        return None
    else:
        print()
        choose_action()


def choose_by_list():
    files = glob.glob("*.docx")
    for i, file in enumerate(files):
        print('{0}. {1}'.format(i, file))

    choice = None
    while choice is None:
        cmd = input('> ')
        try:
            choice = int(cmd)
            if choice < 0 or choice > len(files):
                choice = None
        except ValueError:
            pass

    return files[choice]


def choose_by_path():
    file_path = None
    while file_path is None:
        file_path = input('Enter file name: ')
        if os.path.isfile(file_path):
            break
        continue
    return file_path


def redact_menu(file_path):
    redactor = DocxRedactor()
    redactor.open(file_path)

    print('(l) List all used highlighting colors')
    print('(r) Redact highlights')
    print('(s) Save your changes')
    print('(c) Close the current file')

    cmd = input('> ')
    if cmd == 'l':
        colors = redactor.get_all_colors()
        for color in colors:
            print(color)
        redact_menu(file_path)
    elif cmd == 'r':
        replacement = input('> Replacement? ')
        color = input('> Color? ')
        redactor.redact(color, replacement)
        redact_menu(file_path)
    elif cmd == 's':
        print('Warning: the original file will be overwritten! Do you want to proceed? [N/y]', end=' ')
        proceed = input()

        if proceed.lower() == 'y':
            redactor.save()
            print('Saved')
        else:
            print('Canceled')
        redact_menu(file_path)
    elif cmd == 'c':
        return
    else:
        print()
        redact_menu(file_path)


def main():
    clear()
    print('*** Docx Highlight Redactor ***\n')

    while True:
        action = choose_action()
        if action is None:
            break
        file_path = action()
        redact_menu(file_path)


if __name__ == '__main__':
    main()
