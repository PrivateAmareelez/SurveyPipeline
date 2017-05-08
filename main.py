from lxml import etree
import openpyxl


def xlsx_to_list(doc):
    wb = openpyxl.load_workbook(doc)
    sheet = wb['Sheet1']
    return [[cell.value for cell in row[1:3]] for row in sheet.rows]


class LsgBuilder:
    def __init__(self, data) -> None:
        self.data = data
        self.index = 0

        example = open("limesurvey_group_1.lsg", "r")
        self.tree = etree.parse(example, parser=etree.XMLParser(strip_cdata=False))
        self.root = self.tree.getroot()

    @staticmethod
    def cdata(text):
        # return "<![CDATA[{0}]]>".format(text)
        from lxml.etree import CDATA
        return CDATA(str(text))

    def make_survey(self):
        for i in range(len(data)):
            self.make_one_lsg()

    def make_one_lsg(self):
        self.index = self.index + 1
        for child in self.root:
            if child.tag == 'groups':
                self.make_group(child)
            elif child.tag == 'questions':
                self.make_questions(child)

        output = open("groups/group_{0}.lsg".format(self.index), "wb")
        self.tree.write(output, pretty_print=True, xml_declaration=True, encoding='UTF-8')

    def make_group(self, root):
        row = root[1][0]
        for child in row:
            if child.tag == 'gid':
                child.text = self.cdata(self.index)
            elif child.tag == 'group_name':
                child.text = self.cdata(self.get_title(data[self.index - 1][0]))
            elif child.tag == 'group_order':
                child.text = self.cdata(self.index - 1)

    def make_questions(self, root):
        row = root[1][0]
        for child in row:
            if child.tag == 'qid':
                child.text = self.cdata(self.index * 2 - 1)
            elif child.tag == 'gid':
                child.text = self.cdata(self.index)
            elif child.tag == 'title':
                child.text = self.cdata(self.get_title(data[self.index - 1][0]))
            elif child.tag == 'question':
                child.text = self.cdata(data[self.index - 1][1])

        row = root[1][1]
        for child in row:
            if child.tag == 'qid':
                child.text = self.cdata(self.index * 2)
            elif child.tag == 'gid':
                child.text = self.cdata(self.index)
            elif child.tag == 'title':
                child.text = self.cdata('feasible' + str(self.index))

    @staticmethod
    def get_title(text):
        title = ''.join(filter(lambda x: x.isalnum(), text))
        if title[0].isnumeric():
            return 'title' + title
        elif title[0].islower():
            return title[0].upper() + title[1:]
        return title


data = xlsx_to_list('Problem Proposals 2018.xlsx')
builder = LsgBuilder(data)
builder.make_survey()
