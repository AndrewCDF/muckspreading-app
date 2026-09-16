"""Small, dependency-free XLSX settings store. Values are literal text, not formulas."""
import os
import posixpath
import tempfile
import threading
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
LOCK = threading.RLock()


def column(number):
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def worksheet(rows):
    width = max((len(row) for row in rows), default=1)
    result = ['<?xml version="1.0" encoding="UTF-8"?>', '<worksheet xmlns="%s">' % NS,
              '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
              '<cols><col min="1" max="%s" width="28" customWidth="1"/></cols><sheetData>' % width]
    for index, row in enumerate(rows, 1):
        result.append('<row r="%s">' % index)
        for col, value in enumerate(row, 1):
            if isinstance(value, bool):
                value = "1" if value else "0"
            text = str(value if value is not None else "")
            if any(ord(char) < 32 and char not in '\t\n\r' for char in text):
                raise ValueError("Settings contain an invalid control character")
            result.append('<c r="%s%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (column(col), index, escape(text).replace('\r', '&#13;')))
        result.append('</row>')
    result.append('</sheetData><autoFilter ref="A1:%s%s"/></worksheet>' % (column(width), max(1, len(rows))))
    return ''.join(result).encode('utf-8')


class Workbook:
    def __init__(self, path):
        self.path = path

    def _archive(self):
        with zipfile.ZipFile(self.path) as archive:
            return {name: archive.read(name) for name in archive.namelist()}

    @staticmethod
    def _targets(files):
        root = ET.fromstring(files['xl/workbook.xml'])
        rels = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
        targets = {rel.get('Id'): rel.get('Target') for rel in rels}
        result = {}
        for sheet in root.iter():
            if sheet.tag.rsplit('}', 1)[-1] != 'sheet':
                continue
            relation = next((value for key, value in sheet.attrib.items() if key.rsplit('}', 1)[-1] == 'id'), '')
            target = targets[relation]
            result[sheet.get('name')] = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl', target))
        return result

    def rows(self, name):
        files = self._archive()
        target = self._targets(files).get(name)
        if not target:
            raise ValueError('settings.xlsx is missing the %s tab' % name)
        shared = []
        if 'xl/sharedStrings.xml' in files:
            for item in ET.fromstring(files['xl/sharedStrings.xml']):
                shared.append(''.join(node.text or '' for node in item.iter() if node.tag.rsplit('}', 1)[-1] == 't'))
        rows = []
        for row in ET.fromstring(files[target]).iter():
            if row.tag.rsplit('}', 1)[-1] != 'row':
                continue
            row_number = int(row.get('r', len(rows) + 1))
            while len(rows) < row_number:
                rows.append([])
            values = rows[row_number - 1]
            for cell in row:
                letters = ''.join(c for c in cell.get('r', '') if c.isalpha())
                index = 0
                for letter in letters.upper():
                    index = index * 26 + ord(letter) - 64
                if not index:
                    index = len(values) + 1
                while len(values) < index:
                    values.append('')
                if cell.get('t') == 'inlineStr':
                    value = ''.join(node.text or '' for node in cell.iter() if node.tag.rsplit('}', 1)[-1] == 't')
                else:
                    value = next((node.text or '' for node in cell if node.tag.rsplit('}', 1)[-1] == 'v'), '')
                    if cell.get('t') == 's' and value:
                        value = shared[int(value)]
                values[index - 1] = value
        return rows

    def records(self, name):
        rows = self.rows(name)
        if not rows:
            return []
        headers = [str(value).strip().lower() for value in rows[0]]
        return [{header: row[index] if index < len(row) else '' for index, header in enumerate(headers) if header}
                for row in rows[1:] if any(str(value).strip() for value in row)]

    def _write(self, files):
        parent = os.path.dirname(self.path) or '.'
        os.makedirs(parent, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='.settings-', suffix='.xlsx', dir=parent)
        try:
            with os.fdopen(fd, 'wb') as handle:
                with zipfile.ZipFile(handle, 'w', zipfile.ZIP_DEFLATED) as archive:
                    for name, data in files.items():
                        archive.writestr(name, data)
            os.replace(temporary, self.path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)

    def create(self, sheets):
        with LOCK:
            if os.path.exists(self.path):
                return False
            files = {'_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>' % REL}
            workbook = ['<workbook xmlns="%s" xmlns:r="%s"><sheets>' % (NS, REL)]
            relationships = ['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
            types = ['<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
            for index, (name, rows) in enumerate(sheets.items(), 1):
                workbook.append('<sheet name="%s" sheetId="%s" r:id="rId%s"/>' % (escape(name, {'"': '&quot;'}), index, index))
                relationships.append('<Relationship Id="rId%s" Type="%s/worksheet" Target="worksheets/sheet%s.xml"/>' % (index, REL, index))
                types.append('<Override PartName="/xl/worksheets/sheet%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % index)
                files['xl/worksheets/sheet%s.xml' % index] = worksheet(rows)
            files['xl/workbook.xml'] = ''.join(workbook) + '</sheets></workbook>'
            files['xl/_rels/workbook.xml.rels'] = ''.join(relationships) + '</Relationships>'
            files['[Content_Types].xml'] = ''.join(types) + '</Types>'
            self._write(files)
            return True

    def update(self, name, transform):
        with LOCK:
            rows = self.rows(name)
            updated = transform(rows)
            if updated == rows:
                return
            files = self._archive()
            target = self._targets(files)[name]
            files[target] = worksheet(updated)
            self._write(files)

    def set_rows(self, name, rows):
        self.update(name, lambda previous: rows)

    def add_sheet(self, name, rows):
        with LOCK:
            files = self._archive()
            if name in self._targets(files):
                return False
            workbook_root = ET.fromstring(files['xl/workbook.xml'])
            rels_root = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
            types_root = ET.fromstring(files['[Content_Types].xml'])
            sheets = next(node for node in workbook_root if node.tag.rsplit('}', 1)[-1] == 'sheets')
            sheet_ids = [int(node.get('sheetId', '0')) for node in sheets]
            relation_ids = []
            for node in rels_root:
                value = node.get('Id', '')
                if value.startswith('rId') and value[3:].isdigit():
                    relation_ids.append(int(value[3:]))
            sheet_number = 1
            while 'xl/worksheets/sheet%s.xml' % sheet_number in files:
                sheet_number += 1
            relation_id = 'rId%s' % (max(relation_ids or [0]) + 1)
            ET.SubElement(sheets, '{%s}sheet' % NS, {
                'name': name, 'sheetId': str(max(sheet_ids or [0]) + 1), '{%s}id' % REL: relation_id,
            })
            package_rel = rels_root.tag[1:].split('}', 1)[0]
            ET.SubElement(rels_root, '{%s}Relationship' % package_rel, {
                'Id': relation_id, 'Type': REL + '/worksheet', 'Target': 'worksheets/sheet%s.xml' % sheet_number,
            })
            content_ns = types_root.tag[1:].split('}', 1)[0]
            ET.SubElement(types_root, '{%s}Override' % content_ns, {
                'PartName': '/xl/worksheets/sheet%s.xml' % sheet_number,
                'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            })
            ET.register_namespace('', NS)
            ET.register_namespace('r', REL)
            files['xl/workbook.xml'] = ET.tostring(workbook_root, encoding='utf-8', xml_declaration=True)
            files['xl/_rels/workbook.xml.rels'] = ET.tostring(rels_root, encoding='utf-8', xml_declaration=True)
            files['[Content_Types].xml'] = ET.tostring(types_root, encoding='utf-8', xml_declaration=True)
            files['xl/worksheets/sheet%s.xml' % sheet_number] = worksheet(rows)
            self._write(files)
            return True


def active(row):
    return str(row.get('active', '1')).strip().lower() not in ('0', 'false', 'no', 'n', 'off')
