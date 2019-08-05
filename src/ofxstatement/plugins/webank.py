import pandas as pd
from datetime import datetime
from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import (Statement, StatementLine,
                                    generate_transaction_id)

class WebankPlugin(Plugin):
    """Italian Bank Webank, it parses xlsx file
    """

    def get_parser(self, filename):
        parser = WebankParser(filename)
        parser.statement.bank_id = self.settings.get('bank', 'Webank')
        parser.statement.currency = self.settings.get('currency', 'EUR')
        parser.statement.account_id = self.settings.get('account_id', '00000 - 0000000000')
        parser.statement.account_type = self.settings.get(
            'account_type', 'CHECKING')

        return parser


class WebankParser(StatementParser):
    # WeBank date format
    date_format = '%d/%m/%Y'

    # pandas data frame containing data.
    df = pd.DataFrame()
    df_row_idx = 0

    def __init__(self, filename):
        self.filename = filename
        self.statement = Statement()

    def parse(self):
        """Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        self.df = pd.read_html(self.filename, decimal=',', thousands='.')[0]

        return super(WebankParser, self).parse()

    def split_records(self):
        """Return iterable object, true for all rows, row index is updated 
            internally
        """
        for row_num in range(len(self.df)):
            self.df_row_idx = row_num
            yield True

    def xls_date(self,html_string_date):
        # parse xlsx date using mode 0 (start from jan 1 1900)
        date = datetime.strptime(html_string_date, self.date_format)
        return date

    def parse_record(self, df_row):
        """Parse given transaction line and return StatementLine object
        """
        stmt_line = StatementLine()        

        # date field
        stmt_line.date = self.xls_date(
            self.df['Data Contabile'][self.df_row_idx])

        # amount field
        stmt_line.amount =self.df['Importo'][self.df_row_idx]
        
        # transaction type field
        if(stmt_line.amount < 0):
            stmt_line.trntype = "DEBIT"
        else:
            stmt_line.trntype = "CREDIT"

        # memo field
        stmt_line.memo = self.df['Causale / Descrizione'][self.df_row_idx]

        # id field
        stmt_line.id = generate_transaction_id(stmt_line)
        #print(str(stmt_line))
        return stmt_line
