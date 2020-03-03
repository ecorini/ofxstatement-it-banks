import xlrd
import getpass
import passpy
import mechanize
import time
import pandas as pd
import os

from datetime                   import datetime
from ofxstatement.plugin        import Plugin
from ofxstatement.downloader    import Downloader
from ofxstatement.exceptions    import DownloadError
from ofxstatement.parser        import StatementParser
from ofxstatement.statement     import (Statement, StatementLine,
                                        generate_transaction_id)     

class WebankPlugin(Plugin):
    """Italian Bank Webank, it parses xls file
    """

    def get_parser(self, filename):
        parser = WebankParser(filename)
        parser.statement.bank_id = self.settings.get('bank', 'Webank')
        parser.statement.currency = self.settings.get('currency', 'EUR')
        parser.statement.account_id = self.settings.get('account_id', '00000 - 0000000000')
        parser.statement.account_type = self.settings.get(
            'account_type', 'CHECKING')

        return parser

    def get_downloader(self, filename, start_date, end_date):
        downloader = WebankDownloader(filename, start_date, end_date)
        downloader.ZX2C4passname = self.settings.get('zx2c4', 'None')
        if downloader.ZX2C4passname is not "None":
            downloader.useZX2C4pass = True             
        return downloader        


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
        stmt_line.amount = self.df['Importo'][self.df_row_idx]
        
        # transaction type field
        if(stmt_line.amount < 0):
            stmt_line.trntype = "DEBIT"
        else:
            stmt_line.trntype = "CREDIT"

        # memo field        
        stmt_line.memo = self.df['Causale / Descrizione'][self.df_row_idx]
        if(pd.isnull(stmt_line.memo)):
            stmt_line.memo = ''

        # id field
        stmt_line.id = generate_transaction_id(stmt_line)
        #print(str(stmt_line))
        return stmt_line

class WebankDownloader(Downloader):
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
        WEBANK_ROOT             = 'https://www.webank.it'
        WEBANK_LOGIN            = WEBANK_ROOT + '/webankpub/wbresp/home.do'
        WEBANK_LOGIN_KEY        = WEBANK_ROOT + '/WEBWB/jsp/ht/loginKey.jsp'
        WEBANK_SET_OTP_MODE     = WEBANK_ROOT + '/WEBWB/cambioStatoOperazioneDaAutorizzare.do'
        WEBANK_CHECK_OTP_MODE   = WEBANK_ROOT + '/WEBWB/statoOperazioneDaAutorizzare.do'   
        WEBANK_OTP_LOGIN_ESITO  = WEBANK_ROOT + '/WEBWB/otpLoginEsito.do'   
        WEBANK_MAIN             = WEBANK_ROOT + '/WEBWB/homepage.do'
        WEBANK_HOME             = WEBANK_ROOT + '/WEBEXT/wbOnetoone/fpMyHome.action'
        WEBANK_STATEMENTS       = WEBANK_ROOT + '/WEBWB/cc/movimentiConto.xls'
        WEBANK_LOGOUT           = WEBANK_ROOT + '/WEBWB/logout.do'

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

        # Open Login page
        response = mech.open(WEBANK_LOGIN)
        login_key = mech.open(WEBANK_LOGIN_KEY).read().decode()
        response = mech.open(WEBANK_LOGIN)

        # Submit credentials and login_key
        mech.select_form(name='toplogin')
        mech.set_all_readonly(False)
        mech["j_password"] = pwd
        mech["j_username"] = user
        mech["loginKey"] = login_key
        response = mech.submit()

        # Get hashOtp
        formcount=0
        for frm in mech.forms():  
            if str(frm.attrs["id"])=="otpLoginEsito":
                break
        formcount=formcount+1
        mech.select_form(nr=formcount)
        mech.set_all_readonly(False)
        hashOtp=mech.find_control(name="hashOtp", nr=0).value
        
        # poll on autorization url to check if auth is ok.
        auth_ok = False
        attempts = 15
       
        while auth_ok == False and attempts > 0:
            time.sleep(1)
            response = mech.open(WEBANK_CHECK_OTP_MODE+'?hashOtp='+hashOtp)
            response_str = response.read().decode('utf-8')
            if '"esito" : "VP"' in response_str:
                auth_ok = True
                # open main page
                mech.open(WEBANK_MAIN)
                mech.select_form(nr=formcount)
                mech.set_all_readonly(False)
                response = mech.submit()
            attempts = attempts - 1

        # if timeout is expired use manual authorization
        if not auth_ok:
            print("App authorization failed, use generated OTP code..")
            OTPCode = input("Insert OTP Code: ")
            
            response = mech.open(WEBANK_SET_OTP_MODE+'?hashOtp='+hashOtp)
            response = mech.open(WEBANK_CHECK_OTP_MODE+'?hashOtp='+hashOtp)
            # open main page
            response = mech.open(WEBANK_MAIN)
            mech.select_form(nr=formcount)
            mech.set_all_readonly(False)
            mech["codiceOTP"] = OTPCode
            response = mech.submit()
            response_str = response.read().decode('cp1252')
            if "Codice errore:" not in response_str:
                auth_ok = True

        if auth_ok:            
            # open movement page
            xls_url = WEBANK_STATEMENTS +\
                '?&tipoIntervallo=periodo&dataInizio=' +\
                "%2F".join(self.start_date.split("/")) +\
                '&dataFine=' + "%2F".join(self.end_date.split("/")) +\
                '&ultimiMovimenti=400'

            response = mech.open(xls_url)            
            
            # Write xls file
            results = response.read()
            f = open(self.filename, 'wb')
            f.write(results)
            f.close()

            # Logout
            response = mech.open(WEBANK_LOGOUT)        
        else:
            # Logout
            response = mech.open(WEBANK_LOGOUT)        

            raise DownloadError("Authorization denied")