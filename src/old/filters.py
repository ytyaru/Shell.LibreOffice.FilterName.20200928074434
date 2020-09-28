#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import uno
import unohelper
from pythonscript import ScriptContext
import os.path
from com.sun.star.awt.FontWeight import BOLD
from functools import reduce
from com.sun.star.beans import PropertyValue

def connect_script_context(host='localhost', port='2002', namedpipe=None):
    UNO_RESOLVER = "com.sun.star.bridge.UnoUrlResolver"
    UNO_DESKTOP = "com.sun.star.frame.Desktop"
    localCtx = uno.getComponentContext()
    localSmgr = localCtx.ServiceManager
    resolver = localSmgr.createInstanceWithContext(UNO_RESOLVER, localCtx)
    if namedpipe is None:
        uno_string = 'uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext' % (host, port)
    else:
        uno_string = 'uno:pipe,name=%s;urp;StarOffice.ComponentContext' % namedpipe
    ctx = resolver.resolve(uno_string)
    smgr = ctx.ServiceManager
    XSCRIPTCONTEXT = ScriptContext(ctx,
                                   smgr.createInstanceWithContext(UNO_DESKTOP, ctx),
                                   None)
    return XSCRIPTCONTEXT

def makeFilterOptionsSheet(*args, **kwds):
    # https://qiita.com/shota243/items/b9003a17cf136381dfed
    ALL_FIELDS = False
    FIELDS = ['DocumentService', 'UIName', 'Flags']
    ALL_FLAGS = False
    FLAGS = ['IMPORT', 'EXPORT', 'DEFAULT', 'PREFERRED']
    KEY_FIELD = 'DocumentService'

    SHOW_TYPE_FIELDS = True
    ALL_TYPE_FIELDS = False
    TYPE_FIELDS = ['Extensions', 'MediaType']

# Flags are defined in <source directory>/include/comphelper/documentconstants.hxx
# with exception of3RDPARTYFILTER is STARONEFILTER and READONLY is OPENREADONLY there
    FLAG_LABELS = [(0x00000001, 'IMPORT'),
                   (0x00000002, 'EXPORT'),
                   (0x00000004, 'TEMPLATE'),
                   (0x00000008, 'INTERNAL'),
                   (0x00000010, 'TEMPLATEPATH'),
                   (0x00000020, 'OWN'),
                   (0x00000040, 'ALIEN'),
                   (0x00000100, 'DEFAULT'),
                   (0x00000400, 'SUPPORTSSELECTION'),
                   (0x00001000, 'NOTINFILEDIALOG'),
                   (0x00010000, 'READONLY'),
                   (0x00020000, 'MUSTINSTALL'),
                   (0x00040000, 'CONSULTSERVICE'),
                   (0x00080000, '3RDPARTYFILTER'),
                   (0x00100000, 'PACKED'),
                   (0x00200000, 'EXOTIC'),
                   (0x00800000, 'COMBINED'),
                   (0x01000000, 'ENCRYPTION'),
                   (0x02000000, 'PASSWORDTOMODIFY'),
                   (0x04000000, 'GPGENCRYPTION'),
                   (0x10000000, 'PREFERRED'),
                   (0x20000000, 'STARTPRESENTATION'),
                   (0x40000000, 'SUPPORTSSIGNING'),
    ]


    XSCRIPTCONTEXT = connect_script_context(namedpipe='librepipe')
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager
    # get filter information from filter factory
    SERVICE_FILTERFACTORY = 'com.sun.star.document.FilterFactory'
    SERVICE_TYPEDETECTION = 'com.sun.star.document.TypeDetection'
    filter_factory = smgr.createInstanceWithContext(SERVICE_FILTERFACTORY, ctx)
    filter_dict = {}
    for filter_name in filter_factory.getElementNames():
        try:
            property_dict = {}
            for property in filter_factory.getByName(filter_name):
                property_dict[property.Name]  = property.Value
            filter_dict[filter_name] = property_dict
        except:
            continue
    fields = sorted(list(reduce(lambda x,y:x|y,
                                (set(v.keys()) for v in filter_dict.values()),
                                set())))
    if not ALL_FIELDS:
        fields = [f for f in FIELDS if f in fields]
        if SHOW_TYPE_FIELDS and 'Type' not in fields:
            fields.append('Type')
    #
    if SHOW_TYPE_FIELDS:
        type_detection = smgr.createInstanceWithContext(SERVICE_TYPEDETECTION, ctx)
        type_dict = {}
        for type_name in type_detection.getElementNames():
            try:
                property_dict = {}
                for property in type_detection.getByName(type_name):
                    property_dict[property.Name]  = property.Value
                type_dict[type_name] = property_dict
            except:
                continue
        type_fields = sorted(list(reduce(lambda x,y:x|y,
                                         (set(v.keys()) for v in type_dict.values()),
                                         set())))
        if not ALL_TYPE_FIELDS:
            type_fields = [f for f in TYPE_FIELDS if f in type_fields]
    else:
        type_dict = {}
        type_fields = []
    #
    head = ['Name'] + fields + type_fields
    filter_table = []
    for filter_name, property_dict in filter_dict.items():
        filter_values = [filter_name] + [property_dict.get(f) for f in fields]
        if 'Flags' in fields:
            idx = head.index('Flags')
            if filter_values[idx] is not None:
                flag_value = filter_values[idx]
                flag_names = []
                for n, name in FLAG_LABELS:
                    if flag_value & n == n:
                        if ALL_FLAGS or name in FLAGS:
                            flag_names.append(name)
                        flag_value &= ~n
                if flag_value != 0: # found an unknown flag
                    flag_names.append('%8.8x' % flag_value)
                filter_values[idx] = ' '.join(flag_names)
        type_values = [type_dict.get(property_dict.get('Type'), {}).get(n) for n in type_fields]
        if 'Extensions' in type_fields:
            idx = type_fields.index('Extensions')
            if type_values[idx] is not None:
                type_values[idx] = ' '.join(type_values[idx])
        filter_table.append(filter_values + type_values)
    #
    if KEY_FIELD in fields:
        idx = head.index(KEY_FIELD)
        key = lambda r:(r[idx], r)
    else:
        key = None
    filter_table.sort(key=key)
    filter_table.insert(0, head)
    # table to spreadsheet
    NEWSHEET_URL = 'private:factory/scalc'
    desktop = XSCRIPTCONTEXT.getDesktop()
    spreadsheet = desktop.loadComponentFromURL(NEWSHEET_URL, '_blank', 0, ())
    sheet = spreadsheet.getSheets().getByIndex(0)
    for i, row in enumerate(filter_table):
        for j, col in enumerate(row):
            cell = sheet.getCellByPosition(j, i)
            cell.String = str(col)
    # set column width optimal
    columns = sheet.getColumns()
    for i in range(len(head)):
        columns.getByIndex(i).OptimalWidth = True
    # make head line characters bold
    for i in range(len(head)):
        cell = sheet.getCellByPosition(i, 0)
        cell.CharWeight = BOLD
    # freeze head row (column 0, row 1)
    spreadsheet.CurrentController.freezeAtPosition(0, 1)
    # set auto filter
    cell_range = sheet.getCellRangeByPosition(0, 0, len(head) - 1, len(filter_table) - 1)
    range_name = 'filter range'
    spreadsheet.DatabaseRanges.addNewByName(range_name, cell_range.getRangeAddress())
    filter_range = spreadsheet.DatabaseRanges.getByName(range_name)
    filter_range.AutoFilter = True
    return spreadsheet

def saveSpreadSheet(doc, url='', filters=()):
#    doc.storeToURL(unohelper.systemPathToFileUrl(os.path.join(os.path.dirname(os.path.abspath(__file__)),'filters.ods')),())
    if doc.isModified:
        if doc.hasLocation and (not doc.isReadonly):
            doc.store()
        else:
            doc.storeAsURL(unohelper.systemPathToFileUrl(url), filters)

def _getFilterName(ext='ods'):
    ext = ext.lower()
    value = 'calc8'
    if   'ods' == ext: value = 'calc8'
    elif 'fods' == ext: value = 'OpenDocument Spreadsheet Flat XML'
    elif ext in ('csv','tsv','tab','txt'): value = 'Text - txt - csv (StarCalc)'
    elif ext in ('slk','sylk'): value = 'SYLK'
    else: value = 'calc8'
    return value
def getFilterName(ext='ods'):
    return PropertyValue(Name='FilterName', Value=_getFilterName(ext))
def getFilterOptions(value='44,34,76,,,,,,true'):
    return PropertyValue(Name='FilterOptions', Value=value)
def getOverwrite(value=True):
    return PropertyValue(Name='Overwrite', Value=str(value).lower())
def getCharacterSet(value='UTF-8'):
    return PropertyValue(Name='CharacterSet', Value=value)
 
# https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
def getDsvFilterOptionsValue(separator=',', quotation='"'):
    return '%s,%s,%s,,,,,,true' % (ord(separator), ord(quotation), getCharsetCodeUtf8())
#    return '' + ord(separator) + ',' + ord(quotation) + ',' + getCharsetCodeUtf8() + ',,,,,,true'
def getCharsetCodeUtf8(): return 76
def getCsvProperties():
    return (getFilterName(ext='csv'), 
            getFilterOptions(value=getDsvFilterOptionsValue(separator=',',quotation='"')),
            getOverwrite(), 
            getCharacterSet())
#    return (PropertyValue(Name='FilterName', Value='Text - txt - csv (StarCalc)'),
#        PropertyValue(Name='FilterOptions', Value='44,34,76,,,,,,true'))
##        PropertyValue(Name='FilterOptions', Value=getDsvFilterOptionsValue(separator='')))
def getTsvProperties():
    return (getFilterName(ext='tsv'), 
            getFilterOptions(value=getDsvFilterOptionsValue(separator='\t',quotation='"')),
            getOverwrite(), 
            getCharacterSet())
#    return (PropertyValue(Name='FilterName', Value='Text - txt - csv (StarCalc)'),
#        PropertyValue(Name='FilterOptions', Value='9,34,76,,,,,,true'))
##        PropertyValue(Name='FilterOptions', Value='9,34,76,,,,,,true'))
def getFodsProperties():
    return (getFilterName(ext='fods'), getOverwrite(), getCharacterSet()) 
#    return (PropertyValue(Name='FilterName', Value='OpenDocument Spreadsheet Flat XML'))
def getSylkProperties():
    return (getFilterName(ext='sylk'), getOverwrite(), getCharacterSet())

if __name__ == '__main__':
    doc = makeFilterOptionsSheet()
    here = os.path.dirname(os.path.abspath(__file__))
    saveSpreadSheet(doc, url=os.path.join(here,'filters.ods'))
    saveSpreadSheet(doc, url=os.path.join(here,'filters.csv'), filters=getCsvProperties())
    saveSpreadSheet(doc, url=os.path.join(here,'filters.tsv'), filters=getTsvProperties())
    saveSpreadSheet(doc, url=os.path.join(here,'filters.fods'), filters=getFodsProperties())
    saveSpreadSheet(doc, url=os.path.join(here,'filters.sylk'), filters=getSylkProperties())
#    saveSpreadSheet(makeFilterOptionsSheet(), url=os.path.join(os.path.dirname(os.path.abspath(__file__)),'filters.ods'))
#    saveSpreadSheet(makeFilterOptionsSheet(), url=unohelper.systemPathToFileUrl(os.path.join(os.path.dirname(os.path.abspath(__file__)),'filters.ods')))
#    main()

