import xlrd
from datetime import datetime
from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import (Statement, StatementLine,
                                    generate_transaction_id)

class WidibaPlugin(Plugin):
    """Italian Bank Widiba, it parses xlsx file
    """

    def get_parser(self, filename):
        parser = WidibaParser(filename)
        parser.statement.bank_id = self.settings.get('bank', 'Widiba')
        parser.statement.currency = self.settings.get('currency', 'EUR')
        parser.statement.account_type = self.settings.get(
            'account_type', 'CHECKING')

        if self.settings.get('info2name', 'False') == 'True':
            parser.info2name = True
        if self.settings.get('info2memo', 'False') == 'True':
            parser.info2memo = True
        return parser


class WidibaParser(StatementParser):
    # fill ofx <NAME> field with "Causale" column of xlsx file
    info2name = False
    # concat ofx <MEMO> field with "Causale" column of xlsx file
    info2memo = False

    tpl = {
        'th' : [
            u"DATA CONT.",
            u"DATA VAL.",
            u"CAUSALE",
            u"DESCRIZIONE",
            u"IMPORTO (€)(€)",
        ],
    }
    common_footer_marker = 'Totale (€)'
    th_separator_idx = 0

    def __init__(self, filename):
        self.filename = filename
        self.statement = Statement()

    def parse(self):
        """Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        workbook = xlrd.open_workbook(self.filename)
        sheet = workbook.sheet_by_index(0)
        heading, rows = [], []

        # split heading from current statement
        for rowidx in range(sheet.nrows):
            row = sheet.row_values(rowidx)
            
            # remove columns 0 and 5 that are always empty in widiba xlsx
            del row[5]
            del row[0]

            # add row to array
            if self.th_separator_idx > 0:
                if row[0] != '' and row[4] != self.common_footer_marker:
                    rows.append(row)
            else:
                heading.append(row)
            
            # check transaction header row and set th_separator_idx
            if row == self.tpl['th']:           
                self.th_separator_idx = rowidx

        # check if transaction header is recognized
        if self.th_separator_idx == 0:
            raise ValueError('Widiba xlsx file not recognized!')

        self.rows = rows

        self.statement.account_id = self._get_account_id()                
        return super(WidibaParser, self).parse()

    def split_records(self):
        """Return iterable object consisting of a line per transaction
        """
        for row in self.rows:
            yield row

    def xls_date(self,excel_serial_date):
        # parse xlsx date using mode 0 (start from jan 1 1900)
        date = datetime(*xlrd.xldate_as_tuple(excel_serial_date, 0))
        return date

    def parse_record(self, line):
        """Parse given transaction line and return StatementLine object
        """
        stmt_line = StatementLine()        

        # date field
        stmt_line.date = self.xls_date(int(line[0]))

        # amount field
        stmt_line.amount = line[4]
        
        # transaction type field
        if(stmt_line.amount < 0):
            stmt_line.trntype = "DEBIT"
        else:
            stmt_line.trntype = "CREDIT"

        # name field
        # set <NAME> field with content of column 'CAUSALE'
        # only if proper option is active
        if self.info2name:
            stmt_line.payee = line[2]

        # memo field
        stmt_line.memo = line[3]
        # concat "CAUSALE" column at the end of <MEMO> field
        # if proper option is present
        if self.info2memo:
            if stmt_line.memo != '' and line[2] != '':
                stmt_line.memo+= ' - ' 
            stmt_line.memo+= line[2]

        # id field
        stmt_line.id = generate_transaction_id(stmt_line)


        #print(str(stmt_line))
        return stmt_line

    def _get_account_id(self):
        workbook = xlrd.open_workbook(self.filename)
        sheet = workbook.sheet_by_index(0)
        return str(sheet.cell_value(9, 3))
