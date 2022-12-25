import pandas as pd
import xml.etree.ElementTree as ET


# classe con la tabella
# è un df con qualche attributo in più, come il nome del foglio
class Table(pd.DataFrame):
    """Class that inherits from pandas.DataFrame then customizes it with additonal methods."""

    def __init__(self, name, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.name = name

    @property
    def _constructor(self):
        """
        Creates a self object that is basically a pandas.Dataframe.
        self is a dataframe-like object inherited from pandas.DataFrame
        self behaves like a dataframe + new
        """
        return Table


    def to_excel(self, excel_writer, na_rep='', float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None, merge_cells=True, inf_rep='inf',
                 freeze_panes=None, storage_options=None):
        super(Table, self).to_excel(excel_writer, self.name, na_rep, float_format, columns, header, index, index_label, startrow, startcol, engine, merge_cells, inf_rep, freeze_panes, storage_options)
    def append(self, other, **kwargs):
        return Table(self.name, pd.concat([pd.DataFrame(self), pd.DataFrame(other)]))

# classe che è una collezione di Table. Deve gestire gli id e le reference.
# è l'unica che può accedere alla classe Table, che quindi sarà una classe privata

class Data(object):
    #apri il file in xml
    def __init__(self, fileName, path = ''):
        self._fieName = fileName
        self._listTable = []
        self._contatori = {}
        self._pathName = path
        p = path + fileName
        if '.xml' in fileName:
            self._data = DataConverter.fromXML(p)
    @property
    def listTable(self):
        return self._listTable

    def _countIstance(self, element):
        tagIstance = {}
        for e in element:
            if e.tag in list(tagIstance.keys()):
                tagIstance[e.tag] += 1
            else:
                tagIstance[e.tag] = 1
        return tagIstance

    def analyze(self):
        self._treeAnalysys(self._data)


    def _treeAnalysys(self, root, ref=None):
        tIstance = self._countIstance(root)
        maxTpl = max(list(tIstance.items()), key=lambda tup: tup[1])
        t = {i: [] for i in tIstance.keys()}
        if ref != None:
            t['ID_INTERNAL'] = [ref] * maxTpl[1]
        for e in root:
            val = ''
            ##gestione della ricorsività
            if len(e) == 0:
                val = e.text
            else:
                code = 1

                if e.tag in self._contatori.keys():

                    code = self._contatori[e.tag]
                    self._contatori[e.tag] += 1

                else:
                    self._contatori[e.tag] = code + 1
                val = e.tag + '__' + str(code)
                self._treeAnalysys(e, val)
            # gestione dei campi multipli o meno
            if e.tag == maxTpl[0]:
                t[e.tag].append(val, )
            else:
                t[e.tag] = [val] * maxTpl[1]

        df = Table(root.tag, t)
        self._listTable.append(df)

    def _createSheet(self):
        listSheet = []
        listSheetName = []
        i2 = 0
        for tab in self.listTable:

            if tab.name in listSheetName:

                i = listSheetName.index(tab.name)
                tmp = listSheet[i]
                i2 +=1
                listSheet[i] = tmp.append(tab)

            else:

                listSheet.append(tab)
                listSheetName.append(tab.name)
        return listSheet

    def convertTo(self, format = 'xlsx'):
        sheets = self._createSheet()
        if format == 'xlsx':
            savePath = self._pathName + self._fieName.split('.')[0] + '.xlsx'
            DataConverter.toExcel(sheets, savePath)



class DataConverter(object):
    def __init__(self):
        pass
    @staticmethod
    def toExcel(fogli, fileName):
        writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
        for i in fogli:
            try:
                i.to_excel(writer)
            except ValueError:
                print(ValueError)
        writer.save()
    @staticmethod
    def fromXML(filePath):
        tree = ET.parse(filePath)
        root = tree.getroot()
        return root