import xlrd
import getpass
import passpy
import mechanize
import time
import asyncio
import urllib
import os

from datetime                   import datetime
from ofxstatement.plugin        import Plugin
from ofxstatement.downloader    import Downloader
from ofxstatement.exceptions    import DownloadError
from ofxstatement.parser        import StatementParser
from ofxstatement.statement     import (Statement, StatementLine,
                                        generate_transaction_id)                                    

class FinecoPlugin(Plugin):
    """Italian Bank Fineco, it parses xls file
    """

    def get_parser(self, filename):
        parser = FinecoParser(filename)
        parser.statement.bank_id = self.settings.get('bank', 'Fineco')
        parser.statement.currency = self.settings.get('currency', 'EUR')
        parser.statement.account_type = self.settings.get(
            'account_type', 'CHECKING')

        if self.settings.get('info2name', 'False') == 'True':
            parser.info2name = True
        if self.settings.get('info2memo', 'False') == 'True':
            parser.info2memo = True
        return parser

    def get_downloader(self, filename, start_date, end_date):
        downloader = FinecoDownloader(filename, start_date, end_date)
        downloader.ZX2C4passname = self.settings.get('zx2c4', 'None')
        if downloader.ZX2C4passname is not "None":
            downloader.useZX2C4pass = True             
        return downloader

class FinecoParser(StatementParser):
    # fill ofx <NAME> field with "Descrizione" column of xls file
    info2name = False
    # concat ofx <MEMO> field with "Descrizione" column of xls file
    info2memo = False

    tpl = {
        'th' : [
            u"Data Operazione",
            u"Data Valuta",
            u"Entrate",
            u"Uscite",
            u"Descrizione",
            u"Descrizione Completa",
        ],
    }
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

            # add row to array
            if self.th_separator_idx > 0:
                if row[0] != '':
                    rows.append(row)
            else:
                heading.append(row)
            
            # check transaction header row and set th_separator_idx
            if row == self.tpl['th']:           
                self.th_separator_idx = rowidx

        # check if transaction header is recognized
        if self.th_separator_idx == 0:
            raise ValueError('Fineco xls file not recognized!')

        self.rows = rows

        self.statement.account_id = self._get_account_id()                
        return super(FinecoParser, self).parse()

    def split_records(self):
        """Return iterable object consisting of a line per transaction
        """
        for row in self.rows:
            yield row

    def xls_date(self,excel_serial_date):
        # parse xls date using mode 0 (start from jan 1 1900)
        date = datetime(*xlrd.xldate_as_tuple(excel_serial_date, 0))
        return date

    def parse_record(self, line):
        """Parse given transaction line and return StatementLine object
        """
        stmt_line = StatementLine()        

        # date field
        stmt_line.date = self.xls_date(int(line[0]))

        # amount field
        if line[2]:
            income = line[2]
            outcome = 0
        elif line[3]:
            outcome = line[3]
            income = 0
        stmt_line.amount = income - outcome
        
        # transaction type field
        if(stmt_line.amount < 0):
            stmt_line.trntype = "DEBIT"
        else:
            stmt_line.trntype = "CREDIT"

        # name field
        # set <NAME> field with content of column 'Descrizione'
        # only if proper option is active
        if self.info2name:
            stmt_line.payee = line[4]

        # memo field
        stmt_line.memo = line[5]
        # concat "Descrizione" column at the end of <MEMO> field
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
        return str(sheet.cell_value(0, 0).replace("Conto Corrente: ", ''))

class FinecoDownloader(Downloader):
    def __init__(self, filename, start_date, end_date):
        self.filename = filename
        self.start_date = start_date.strftime("%d/%m/%Y")
        self.end_date = end_date.strftime("%d/%m/%Y")
        self.ZX2C4passname = "None"
        self.useZX2C4pass = False
        

    def download(self):
        """Main entry point for downloader
        """

        # url's for access
        FINECO_ROOT = 'https://finecobank.com'
        FINECO_LOGIN = FINECO_ROOT + '/portalelogin'
        FINECO_0 = FINECO_ROOT + '/conto-e-carte/movimenti/movimenti-conto'
        FINECO_AUTH_1 = FINECO_ROOT + '/myfineco-auth/sca/consents?r=rmvc'
        FINECO_AUTH_2 = FINECO_ROOT + '/myfineco-auth/sca/consents/inizia-transazione'
        FINECO_AUTH_3 = FINECO_ROOT + '/myfineco-auth/sca/consents/verifica-stato-transazione'
        FINECO_AUTH_4 = FINECO_ROOT + '/myfineco-auth/sca/consents/confirm'
        FINECO_AUTH_OFFL_1 = FINECO_ROOT + '/myfineco-auth/overlays/ovl-transazione?isOtpInOverlay=true'
        FINECO_AUTH_OFFL_2 = FINECO_ROOT + '/myfineco-auth/sca/consents/conferma-transazione-offline'
        FINECO_XLS = FINECO_ROOT + '/conto-e-carte/movimenti/movimenti-conto/excel'
        FINECO_LOGOUT = FINECO_ROOT + '/public/logout'

        # get credentials
        if self.useZX2C4pass:
            if os.name == 'nt':
                store = passpy.Store(gpg_bin='gpg.exe')
            else:
                store = passpy.Store()
            key = store.get_key(self.ZX2C4passname).split('\n')
            pwd = key[0]
            for k in key:
                if "login: " in k:
                    user = k.replace("login: ", "")
                    break
        else:
            user = input("User: ")
            pwd = getpass.getpass()

        # Init browser
        mech = mechanize.Browser()
        mech.addheaders = [('User-agent', 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:68.0) Gecko/20100101 Firefox/68.0')]
        mech.set_handle_robots(False)
        
        # Open login page
        mech.open(FINECO_LOGIN)

        # Select & submit form that has id = "loginForm" (no "name" attrib.)
        mech.select_form(nr=0)
        mech["LOGIN"] = user
        mech["PASSWD"] = pwd
        mech.submit()

        # compile browser header to generate an XHR request
        header_XHR = {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:69.0) Gecko/20100101 Firefox/69.0', 
                    'Referer': FINECO_0, 
                    'X-Requested-With': 'XMLHttpRequest'}

        # Auth Step 1
        url = FINECO_AUTH_1 + '&_=' + str(int(time.time() * 1000))
        req = mechanize.Request(url, headers=header_XHR, method='GET')
        response = mech.open(req)

        # Auth Step 2
        url = FINECO_AUTH_2
        req = mechanize.Request(url, headers=header_XHR, method='POST')
        response = mech.open(req)

        # Auth Step 3
        # poll on autorization url to check if auth is ok.
        auth_ok = False
        attempts = 15
        url = FINECO_AUTH_3
        req = mechanize.Request(url, headers=header_XHR, method='POST')
        
        while auth_ok == False and attempts > 0:
            time.sleep(1)
            response = mech.open(req)
            response_str = response.read().decode('utf-8')
            if '"stato":"confirmed"' in response_str:
                auth_ok = True
            attempts = attempts - 1
      
        # if timeout is expired use manual authorization
        if not auth_ok:
            print("App authorization failed, use PIN and mobile generated code..")
            # submit PIN code
            url = FINECO_AUTH_OFFL_1
            response = mech.open(url)            
            mech.select_form("pinOfflineForm")
            mech.set_all_readonly(False)
            mech["PIN"] = input("Insert PIN: ")
            response = mech.submit()
            response_str = response.read().decode('utf-8')

            if '"PIN":"OK"' in response_str:
                # submit generated code
                url = FINECO_AUTH_OFFL_2
                mobile_generated_pin = input("Insert generated code: ")
                params = {'PIN': mobile_generated_pin}
                data = urllib.parse.urlencode(params)
                req = mechanize.Request(url, headers=header_XHR, method='POST')
                response = mech.open(req, data=data)
                response_str = response.read().decode('utf-8')
                print(response_str)
                if '"stato":"confirmed"' in response_str:
                    auth_ok = True

        if auth_ok:            
            # Auth Step 4
            url = FINECO_AUTH_4
            req = mechanize.Request(url, headers=header_XHR, method='POST')
            response = mech.open(req)
            
            # Open movement page
            response = mech.open(FINECO_0)
            # select form twice to enable custom search
            mech.select_form("frmKeywordSearch")
            response = mech.submit()
            mech.select_form("frmKeywordSearch")
            mech["dataDal"] = self.start_date
            mech["dataAl"] = self.end_date
            response = mech.submit()
            
            # Get xls file
            response = mech.open(FINECO_XLS)

            # Write xls file
            results = response.read()
            f = open(self.filename, 'wb')
            f.write(results)
            f.close()

            # Logout
            response = mech.open(FINECO_LOGOUT)        
        else:
            # Logout
            response = mech.open(FINECO_LOGOUT)        

            raise DownloadError("Authorization denied")
        




