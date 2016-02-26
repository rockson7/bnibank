#!/usr/bin/env /usr/bin/python

from selenium import webdriver
from selenium.common.exceptions import TimeoutException,NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from decimal import Decimal
from datetime import datetime as dt
from datetime import date 
import datetime
import json
import csv
from piecash import Account, Transaction, Split, open_book
import re
import cv2
import traceback
import  subprocess

from IPython import embed

progname = 'ibank_pull'
date_format = "%d-%m-%Y"
dict_file = "desc_dict.json"
default_config = {}
config = {}


class Ibank:
    def __init__(self):
        self.trxnum = 0
        self.balance = 0
        self.driver = webdriver.Firefox()
        self.date_format="%d-%m-%Y"
        self.transactions = []

    def dbg(self,txt,*params):
        print txt.format(*params)

    def getby(self, by=By.ID, ident=None,is_list=False):
        if is_list:
            getter_lambda = lambda x: x.find_elements(by,ident)
        else:
            getter_lambda = lambda x: x.find_element(by,ident)

        try:
            el = WebDriverWait(self.driver, 60).until(getter_lambda)
        except NoSuchElementException as e:
            print "element {} not found".format(ident)
            return None

        self.dbg("DBG: found {}: {}", ident, el.__str__())
        return el


    def login(self, username, password):
        #self.driver.implicitly_wait(30) # seconds

        self.driver.get('https://ibank.bni.co.id/')
        el_captcha = None
        el_user_id = self.getby(By.ID,'AuthenticationFG.USER_PRINCIPAL')
        el_pwd = self.getby(By.ID,'AuthenticationFG.ACCESS_CODE')
        el_captcha = self.getby(By.ID,'AuthenticationFG.VERIFICATION_CODE')
        el_login_id = 'VALIDATE_CREDENTIALS'


        el_user_id.click()
        el_user_id.send_keys(username)
        el_pwd.click()
        el_pwd.send_keys(password)
        if el_captcha is not None:
            screenshot = '/tmp/screenshot.png'
            self.driver.save_screenshot(screenshot)
            img = self.getby(By.ID,'IMAGECAPTCHA')
            captcha_text = decaptcha(screenshot,img.location)
            el_captcha.send_keys(captcha_text)
            el_captcha.click()
            #raw_input("Enter captcha end press Enter")

        self.getby(By.ID,el_login_id).click()


    def pull_trx(self,acc,start_date,end_date):
        self.getby(By.ID,'REKENING').click()
        print "REKENING"
        #raw_input("Enter captcha end press Enter")
        self.getby(By.ID,'Informasi-Saldo--Mutasi_Mutasi-Tabungan--Giro').click()
        print "Informasi-Saldo--Mutasi_Mutasi-Tabungan--Giro"
        #raw_input("Enter captcha end press Enter")
        search_fld = self.getby(By.ID,'AccountSummaryFG.ACCOUNT_NUMBER')
        colapsible= self.getby(By.CLASS_NAME,'collapsiblelink')
        if not search_fld.is_displayed():
            colapsible.click()
        search_fld.clear()
        #raw_input("Enter captcha end press Enter")
        print "Fetching  account: ", acc
        search_fld.send_keys(acc)
        self.getby(By.ID,'LOAD_ACCOUNTS').click()
        self.getby(By.ID,'VIEW_TRANSACTION_HISTORY').click()

        # TRX history page
        colapsible= self.getby(By.CLASS_NAME,'collapsiblelink')
        els = self.getby(By.XPATH, "//input[@id='TransactionHistoryFG.SELECTED_RADIO_INDEX']", is_list=True)
        if not els[0].is_displayed():
            colapsible.click()
        els[0].click()

        el = self.getby(By.ID,'TransactionHistoryFG.FROM_TXN_DATE')
        el.send_keys(dt.strftime(start_date,self.date_format))
        el = self.getby(By.ID,'TransactionHistoryFG.TO_TXN_DATE')
        el.send_keys(dt.strftime(end_date,self.date_format))
        el = self.getby(By.ID,'SEARCH')
        #embed()
        el.click()
        txthist_table = self.getby(By.ID,'txnHistoryList')

        #get no of pages
        try:
            pagination_el = self.driver.find_element(By.ID,'paginationtxt1')
            paginationtxt1 = pagination_el.strip().lower()
            no_of_pages = int(paginationtxt1.split(' ')[3])
        except NoSuchElementException:
            no_of_pages = 1
        trxnum = 0
        balance = 0

        for pageno in range(1,no_of_pages+1):
            print "rows from page {} of {} ".format(pageno,no_of_pages)
            rows = txthist_table.find_elements(By.XPATH, ".//tbody/tr[@id]")
            self.dbg("rows {}",rows.__str__())
            self.append_transactions(rows)
            if  no_of_pages > 1:
                next_pg = self.getby(By.ID,'Action.OpTransactionListing.GOTO_NEXT__')
                if next_pg is not None and next_pg.get_attribute("disabled") is not None:
                    break
                next_pg.click()

        return self.transactions

    def append_transactions(self, rows):
        for row in rows:
            self.trxnum +=  1
            fields = row.find_elements_by_xpath(".//td/span")
            tmp_val = Decimal(fields[4].text.strip().replace(',',''))
            value  = tmp_val if 'Cr' in fields[3].text.strip() else -tmp_val
            stm_bal = Decimal(fields[5].text.strip().replace(',',''))
            if self.trxnum == 1:
                self.balance = stm_bal + value
            self.balance  -=  value
            #print trxnum,fields[0],fields[1].text.strip(),tmp_val,value,stm_bal,balance
            trx = {
                'num'  : self.trxnum,
                #'date' : dt.strptime(fields[0].text.strip(), date_format),
                'date' : fields[0].text,
                'desc' : fields[1].text.strip(),
                'value': value,
                'stm_bal' : stm_bal,
                'balance' : self.balance
            }
            self.transactions.append(trx)
            #trx['desc']=trx['desc'][:45]
            template = "{num:>2} {date:10} {desc:50} {value:>12,} {stm_bal:>14,} {balance:>14,}"
            print template.format(**trx)

##### end class Ibank ###############################################


def check_balance_ok(trx):
    template = "{i:>2}|{num:>2}|{date:10}|{desc:50}|{value: 16,}|{stm_bal: 16,}|{balance: 16,}|{ST:2}"
    balance = 0
    t_out = {}
    result = True
    for i,t in enumerate(trx,1):
        if i == 1:
            balance = t['stm_bal'] - t['value']
        balance += t['value']
        if balance !=  t['stm_bal']:
            result = False
            break

        t['balance'] = balance
        t_out = t
        t_out['ST'] = '>' if balance > t['stm_bal'] else '<' if balance < t['stm_bal'] else '='

        t_out['i']=i
        print template.format(**t_out)

    return result

refno_patterns = (
    re.compile("\[ref:(\d{4,5})\]",flags=re.IGNORECASE),
    re.compile("TRANSFER KE.*(\d{4,5})$",flags=re.IGNORECASE),
    re.compile("ECHANNEL.*(\d{4,5})$",flags=re.IGNORECASE)
)

def get_refno_from_desc(desc):
    ref_no = None
    for pat in refno_patterns:
        m = pat.search(desc)
        if m is not None:
            ref_no = int(m.group(1)) if m.group(1).isdigit() else None
            break
    return ref_no

def assign_target_accounts(acc_conf,account,transactions):
    patterns = []
    with open('ibank_pull_matcher.csv','r') as f:
        reader = csv.reader(f)
        for row in reader:
            patterns.append((row[0].strip(),row[1].strip(),row[2].strip()))

    for trx in transactions:
        trx['target_acc_name'] = "Imbalance-IDR"

        matched = False
        for pat in patterns:
            if  're' in pat[0]:
                m = re.search(pat[1],trx['desc'].upper(),flags=re.IGNORECASE)
                if m is not None:
                    matched = True
            elif 'in' in pat[0]:
                if pat[1] in trx['desc'].upper():
                    matched = True
            if matched:
                trx['target_acc_name'] = pat[2]
                break
        # cash to cardholder's account
        if matched and trx['target_acc_name'] in "Assets:Cash":
            trx['target_acc_name'] =  acc_conf['cash_target_account']
        trx['ref_no'] = get_refno_from_desc(trx['desc'])
        print "{0:10}|{1:50}|{2:5}|{3:30}".format(trx['date'],trx['desc'],trx['ref_no'],trx['target_acc_name'])


#### Utils ######
def decaptcha(screenshot,loc):

    mImgFile = "/tmp/out.jpg"
    ss = cv2.imread(screenshot,0)
    img = ss[  loc['y']:loc['y']+22, loc['x']:loc['x']+120 ]
    cv2.threshold(img,160,255,cv2.THRESH_BINARY)
    cv2.imwrite(mImgFile,img)
    tesseract = subprocess.Popen(['/usr/bin/tesseract', mImgFile, '-', '-psm','8','digit'], stdout=subprocess.PIPE)
    res = tesseract.communicate()[0]
    return res.strip()

def load_config(config_file=progname+'.config'):
    config = {}
    with open(config_file,'r') as f:
        config = json.load(f)
    return config

def save_config(config,config_file=progname+'.config'):
    with open(config_file,'w') as f:
        json.dump(config,f,indent=4)

def import_into_gc(book,current_config,acc,transactions,args):
    dryrun = args.no_insert
    #last_import_data = current_config[acc.code]['last_import']
    last_import_data = current_config[args.book][acc.code]['last_import']
    if len(acc.splits) > 0:
        last_num = max([ (int(s.action) if s.action.isdigit() else 0 ) for s in acc.splits ])
    else:
        last_num = 0
    last_post_date = dt.strptime(last_import_data['post_date'],date_format)

    # make a dict of splits by t-num (reference) key
    splits_by_tnum = dict([
        (int(s.transaction.num) if s.transaction.num.isdigit() else 0,s.guid) for s in acc.splits
    ])
    if 0 in splits_by_tnum:
        del splits_by_tnum[0]

    #start_num = int(last_import_data['last_tnum']) +1
    num = start_num = last_num +1
    count = existing_count = 0
    try:
        for trx in transactions:
            post_date = dt.strptime( trx['date'], date_format)
            if last_post_date >= post_date:
                continue

            target_account = book.accounts.get(fullname= trx['target_acc_name'])
            if len(target_account.splits) > 0:
                target_num = max(
                    #[ (int(s.action) if s.action.isdigit() else 0 ) for s in target_account.splits  ]
                    [ (s.action if type(s.action) is int else (int(s.action) if s.action.isdigit() else 0 )) for s in target_account.splits  ]
                ) +1
            else:
               target_num = 1
            # TODO try to find if already posted
            ref_split = transaction = None
            found_reference = False
            if trx['ref_no'] in  splits_by_tnum:
                ref_split = acc.splits.get(guid=splits_by_tnum[trx['ref_no']])
                if ref_split.value == trx['value']:
                    found_reference = True
                    transaction = ref_split.transaction
                    ref_split.action = num
                    existing_count += 1
            else:
                trx['ref_no'] = ""

            if transaction is None:
                transaction = Transaction(
                        currency = acc.commodity,
                        num = trx['ref_no'],

                        description = trx['desc'],
                        post_date = post_date,
                        splits = [
                            Split(
                                value  = trx['value'],
                                account = acc,
                                action = num+count
                            ),
                            Split(
                            value = -trx['value'],
                            account = target_account,
                            action = target_num
                            )
                        ]
                )
            print "{0:15}|{1:12}|{2:8}|{3:35}|{4:3} |{5:15,}|{6:5}-{7:1}|{8:20}".format(
                    transaction.splits[0].action, transaction.splits[1].action,
                    dt.strftime(transaction.post_date,"%d-%m-%y"),
                    transaction.description[:35],
                    ('CRD' if transaction.splits[0].value > 0 else 'DEB'),
                    abs(transaction.splits[0].value),
                    transaction.num, '*' if found_reference else '',
                    transaction.splits[1].account.fullname)

            if not dryrun:
                book.session.commit()
            count += 1
            if not dryrun:
                last_import_data['post_date'] = transaction.post_date.strftime(date_format)
                last_import_data['last_tnum'] = num
                last_import_data['imported_at'] = dt.now().isoformat()
                last_import_data['count'] = count
                last_import_data['balance'] = float(trx['balance'])
    except Exception as e:
        print "ROLLBACK",sys.exc_info()[0]
        book.session.rollback()
        book.session.close()
        traceback.print_exc()
        raise e
    finally:
        if not dryrun:
            save_config(current_config)

    print "IMPORTED {} of {} TRANSACTIONS start num: {} TO ACCOUNT <{}>".format(
        count,len(transactions),start_num,acc.fullname)


#####################################################################
### MAIN                                                   ##########
#####################################################################
if __name__ == "__main__":
    import sys
    import pickle
    import argparse

    reload(sys);
    sys.setdefaultencoding("utf8")

    #last day of prevoius month
    date_format="%d-%m-%Y"
    default_start_date = date.today().replace(day=1) - datetime.timedelta(days=1)
    default_end_date = date.today()
    #usage
    # python  ibank_pull.py --book=../test_data/test.sqlite.gnucash --gnc-account=BNI308994356 --account=308994356 --user-id=dreamboat15 --password=modellismo15 --start-date=13-03-2015  --dump trs-bni56.pickle



    parser = argparse.ArgumentParser(description='Pull Transaction history from ibank.co.id.')
    parser.add_argument('--book',required=True,help='Gnucash sqlite database file')
    parser.add_argument('--gnc-account-code',help='Gnucash BNI account code')
    parser.add_argument('--account-no',help='BNI account number')
    parser.add_argument('--user-id',help='BNI ibank user-id')
    parser.add_argument('--password',help='BNI ibank password')
    parser.add_argument('--start-date',default=default_start_date.strftime(date_format),
                        help='Pull transactions from this date. '+
                             'Default [{}]'.format(default_start_date.strftime(date_format))
                    )
    parser.add_argument('--end-date',default=date.today().strftime(date_format),
                        help='Pull transactions until this date. '+
                             'Default [{}]'.format(date.today().strftime(date_format))
                    )
    parser.add_argument('--dump',action='store_true',help='Dump loaded transactionsto the file')
    parser.add_argument('--from-dump',help='Load transactions from dump file')
    parser.add_argument('--save-ibank-info',action='store_true',help='save password and login')
    parser.add_argument('--init',action='store_true',help='TODO Initial self configuration')
    parser.add_argument('--no-insert',action='store_true',help='Do not insert into the book')

    args = parser.parse_args()
    #print args
    saved_config = load_config()
    current_config = saved_config.copy()

    book_config = current_config[args.book]
    book_sqlite_file = current_config[args.book]['dbfile']
    conn_uri = current_config[args.book]['conn_uri']
    gnc_account_code = args.gnc_account_code
    if gnc_account_code not in current_config[args.book]:
        current_config[args.book][gnc_account_code] = {"account_no":args.account, "last_import":{}}
    acc_conf = current_config[args.book][gnc_account_code]
    if 'account_no' not in acc_conf:
        acc_conf['account_no'] = args.account
    if 'user_id' not in acc_conf:
        acc_conf['user_id'] = args.user_id

    if 'password' not in acc_conf:
        acc_conf['password'] = args.password

    account_no = args.account_no if args.account_no is not None else acc_conf['account_no']
    username = args.user_id if args.user_id is not None else acc_conf['user_id']
    password = args.password if args.password is not None else acc_conf['password']
    start_date = dt.strptime(args.start_date,date_format)
    end_date = dt.strptime(args.end_date,date_format)
    transactions = []
    book = open_book(uri_conn=conn_uri, readonly = args.no_insert, do_backup=False,open_if_lock = True)
    account = book.accounts.get(code=gnc_account_code)

    if  args.from_dump is None:
        print "==================================================================================="
        print "Pulling transactions for period {} - {} from account {}".format(start_date,end_date,gnc_account_code)
        ibank = Ibank()
        try:
            ibank.login(username, password)
            transactions = ibank.pull_trx(account_no, start_date, end_date)
            transactions.reverse()
        except Exception as e:
            traceback.print_exc()
            raise e
        finally:
            ibank.driver.close()

        if args.dump:
            dumpfile = "{}-{}_{}.dump".format(progname,
                            gnc_account_code,
                            dt.today().strftime("%Y-%m-%d"))
            with open(dumpfile,'w') as df:
                pickle.dump(transactions,df)

            print "Dumped {} transactions into {}".format(len(transactions),dumpfile)
        print "==================================================================================="
    else:
        with open(args.from_dump,'r') as df:
            transactions = pickle.load(df)


    print "Comparing account balances with the bank statement..."
    if check_balance_ok(transactions):
        print "OK"
    else:
        print "FAILED"
        args.no_insert = True
    print "==================================================================================="
    print "Assigning target acounts..."
    assign_target_accounts(acc_conf,account,transactions)
    print "==================================================================================="
    print "Importing into gnucash book {} {}".format(args.book,
                        "No Insert" if args.no_insert else "Real insert")
    import_into_gc(book,current_config,account,transactions,args)
    print "==================================================================================="
    print "DONE"
    save_config(current_config)
