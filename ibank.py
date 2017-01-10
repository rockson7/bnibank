#!/usr/bin/env /usr/bin/python

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from decimal import Decimal
from datetime import datetime as dt
from datetime import date
import logging
import datetime
import json
import csv
# from piecash import Account, Transaction, Split, open_book
import re
import cv2
import traceback
import subprocess
import requests
from bs4 import BeautifulSoup

from gncutils import Gncutils, AccMap, Account, Split, Transaction
# import time

from IPython import embed

progname = 'ibank'
date_format = "%d-%m-%Y"
dict_file = "desc_dict.json"
cfg = {}
tzinfo = datetime.tzinfo

class Ibank:
    def __init__(self,book=None,dump=True):
        self.trxnum = 0
        self.balance = 0
        self.driver = webdriver.Firefox()
        self.date_format = "%d-%m-%Y"
        self.transactions = []
        self.logged_in = False
        self.dump = dump
        self.gnc = book if book else Gncutils('rok', readonly=True)

    def dbg(self, txt, *params):
        pass
        #print txt.format(*params)

    def getby(self, by=By.ID, ident=None, is_list=False, timeout=60):
        if is_list:
            getter_lambda = lambda x: x.find_elements(by, ident)
        else:
            getter_lambda = lambda x: x.find_element(by, ident)

        try:
            el = WebDriverWait(self.driver, timeout=timeout).until(getter_lambda)
        except NoSuchElementException as e:
            print "element {} not found".format(ident)
            return None

        self.dbg("DBG: found {}: {}", ident, el.__str__())
        return el

    def get_account_selector_value(self,acc):
        initatoracct = {
                    '387014784': "00000000387014784|ROK-SAVING|SBA",
                    '373352345': "00000000373352345|BNI345-DBI|SBA",
                    '289979573': "00000000289979573|BNI573-ROK|SBA"
        }
        return initatoracct[acc]

    def login(self, username, password):
        # self.driver.implicitly_wait(30) # seconds
        self.driver.get('https://ibank.bni.co.id/')
        el_captcha = None
        el_user_id = self.getby(By.ID, 'AuthenticationFG.USER_PRINCIPAL')
        el_pwd = self.getby(By.ID, 'AuthenticationFG.ACCESS_CODE')
        el_captcha = self.getby(By.ID, 'AuthenticationFG.VERIFICATION_CODE')
        el_login_id = 'VALIDATE_CREDENTIALS'

        el_user_id.click()
        el_user_id.send_keys(username)
        el_pwd.click()
        el_pwd.send_keys(password)
        if el_captcha is not None:
            screenshot = '/tmp/screenshot.png'
            self.driver.save_screenshot(screenshot)
            img = self.getby(By.ID, 'IMAGECAPTCHA')
            captcha_text = decaptcha(screenshot, img.location)
            el_captcha.send_keys(captcha_text)
            el_captcha.click()
            # raw_input("Enter captcha end press Enter")

        el = self.getby(By.ID, el_login_id)
        el.click()
        if self.getby(By.ID, 'REKENING'):
            self.logged_in = True
            return True

        return False

    def click_rekening(self):
        el = self.getby(By.ID, 'REKENING')
        if el.is_displayed():
            el.click()

    def click_mutasi(self,recurse=0):
        mutasi_btn = self.getby(By.ID, 'Informasi-Saldo--Mutasi_Mutasi-Tabungan--Giro')
        saldo_btn = self.getby(By.ID, 'Informasi-Saldo--Mutasi')
        rek_btn = self.getby(By.ID, 'REKENING')
        if mutasi_btn.is_displayed():
            mutasi_btn.click()
        elif saldo_btn.is_displayed():
                saldo_btn.click()
                self.click_mutasi(recurse+1)
        elif rek_btn.is_displayed():
                rek_btn.click()
                saldo_btn.click()
                self.click_mutasi(recurse+1)
        else:
            raw_input("click end press Enter")

    def goto_account_page(self):
        self.click_rekening()
        self.click_mutasi()
        # raw_input("Enter captcha end press Enter")

    def pull_trx(self, acc, start_date, end_date):
        # raw_input("Enter captcha end press Enter")
        search_fld = self.getby(By.ID, 'AccountSummaryFG.ACCOUNT_NUMBER')
        colapsible = self.getby(By.CLASS_NAME, 'collapsiblelink')
        if not search_fld.is_displayed():
            colapsible.click()

        search_fld.clear()
        print "Fetching  account: ", acc
        search_fld.send_keys(acc)
        el = self.getby(By.ID, 'LOAD_ACCOUNTS')
        #raw_input("Enter captcha end press Enter")

        el.click()
        #raw_input("Enter captcha end press Enter")
        el = self.getby(By.ID, 'VIEW_TRANSACTION_HISTORY')
        el.click()
        #raw_input("Enter captcha end press Enter")

        # TRX history page
        colapsible = self.getby(By.CLASS_NAME, 'collapsiblelink')
        els = self.getby(By.XPATH, "//input[@id='TransactionHistoryFG.SELECTED_RADIO_INDEX']", is_list=True)
        if not els[0].is_displayed():
            colapsible.click()
        els[0].click()

        """
        initacc = self.getby(By.ID, 'TransactionHistoryFG.INITIATOR_ACCOUNT',is_list=False,timeout=6)
        initacc.click()
        opt_toclick  = initacc.find_elements_by_tag_name("option")[2]
        opt_toclick.click()

        #for o in initacc.find_elements_by_tag_name("option"):
        #    self.dbg("{} displayed:{}",o.text, o.is_displayed())
        #    if acc in o.text and o.is_displayed() :
        #        opt_toclick = o
        """

        el = self.getby(By.ID, 'TransactionHistoryFG.FROM_TXN_DATE')
        el.send_keys(dt.strftime(start_date, self.date_format))
        el = self.getby(By.ID, 'TransactionHistoryFG.TO_TXN_DATE')
        el.send_keys(dt.strftime(end_date, self.date_format))
        el = self.getby(By.ID, 'SEARCH')
        raw_input("Enter captcha end press Enter")
        el.click()

        txthist_table = self.getby(By.ID, 'txnHistoryList')
        self.trxnum = 0
        self.balance = 0
        pageno = 0
        self.transactions = []
        while True:
            #raw_input(" press Enter")

            pageno += 1
            # print "rows from page {} ".format(pageno)
            txthist_table = self.getby(By.ID, 'txnHistoryList')
            rows = txthist_table.find_elements(By.XPATH, ".//tbody/tr[@id]")
            self.dbg("rows {}", rows.__str__())
            self.append_transactions(rows)
            try:
                next_pg = self.driver.find_element(By.ID, 'Action.OpTransactionListing_custom.GOTO_NEXT__')

                if next_pg is not None and next_pg.get_attribute("disabled") is not None:
                    break
            except NoSuchElementException:
                break
            next_pg.click()

        self.transactions.reverse()
        self.click_mutasi()
        return self.transactions

    def logout(self):
        if self.logged_in:
            el = self.getby(By.ID, 'HREF_Logout')
            el.click()
            el = self.getby(By.ID, 'LOG_OUT')
            el.click()


    def append_transactions(self, rows):
        for row in rows:
            self.trxnum += 1
            fields = row.find_elements_by_xpath(".//td/span")
            tmp_val = Decimal(fields[4].text.strip().replace(',', ''))
            value = tmp_val if 'Cr' in fields[3].text.strip() else -tmp_val
            stm_bal = Decimal(fields[5].text.strip().replace(',', ''))
            if self.trxnum == 1:
                self.balance = stm_bal + value
            self.balance -= value
            # print trxnum,fields[0],fields[1].text.strip(),tmp_val,value,stm_bal,balance
            trx = {
                'num': self.trxnum,
                # 'date' : dt.strptime(fields[0].text.strip(), date_format),
                'date': fields[0].text,
                'desc': fields[1].text.strip(),
                'value': value,
                'stm_bal': stm_bal,
                'balance': self.balance
            }
            self.transactions.append(trx)
            # trx['desc']=trx['desc'][:45]
            template = "{num:>2} {date:10} {desc:50} {value:>12,} {stm_bal:>14,} {balance:>14,}"
            print template.format(**trx)




    def process_account(self,acc,start_date=None):
        account_no = acc.code[3:]
        (last_split, last_tnum) = self.gnc.last_txnum(acc)
        last_trx = last_split.transaction
        if start_date is None:
            start_date = last_trx.post_date.date()
        end_date = start_date + datetime.timedelta(days=30)
        if end_date >= date.today():
            end_date = date.today()-datetime.timedelta(days=0)

        print "Pulling for period {} - {} from  {} last tnum={}" .format(
                start_date, end_date, account_no,last_tnum)

        transactions = self.pull_trx(account_no, start_date, end_date)
        return transactions

    def process_accounts(self,accts_to_process,start_date=None):
        transactions = {}
        self.login(username, password)
        self.goto_account_page()
        for acc in accts_to_process:
            print "process_account({},{})".format(acc,start_date)
            transactions[acc.code] = self.process_account(acc,start_date)
        return transactions


##### end class Ibank ###############################################

class Importer:
    refno_patterns = (
        re.compile("\[ref:(\d{4,5})\]", flags=re.IGNORECASE),
        re.compile("TRANSFER KE.*(\d{4,5})$", flags=re.IGNORECASE),
        re.compile("ECHANNEL.*(\d{4,5})$", flags=re.IGNORECASE)
    )

    def __init__(self, dryrun=False,book=None,accounts=[], **kwargs):
        self.gnc = book if book else Gncutils('rok', readonly=dryrun,do_backup=False)
        self.session = self.gnc.session
        self.patterns = []
        self.dryrun = dryrun

        if len(accounts) > 0:
            self.accts_to_process = accounts
        else:
            self.accts_to_process = gnc.query(Account).filter(
                Account.parent == gnc.accounts_by_fullname['Assets:Bank']).all()

        self.trxs = []


    def check_balance(self,transactions):
        passed = True
        for code,trxs in transactions.items():
            if not self.check_balance_acct(trxs):
                passed = False
        return passed

    def check_balance_acct(self,transactions):
        template = "{i:>2}|{num:>2}|{date:10}|{desc:50}|{value: 16,}|{stm_bal: 16,}|{balance: 16,}|{ST:2}"
        balance = 0
        t_out = {}
        result = True
        for i, t in enumerate(transactions, 1):
            if i == 1:
                balance = t['stm_bal'] - t['value']
            balance += t['value']
            if balance != t['stm_bal']:
                result = False
                break

            t['balance'] = balance
            t_out = t
            t_out['ST'] = '>' if balance > t['stm_bal'] else '<' if balance < t['stm_bal'] else '='

            t_out['i'] = i
            print template.format(**t_out)

        return result

    def get_refno_from_desc(self, desc):
        ref_no = None
        for pat in Importer.refno_patterns:
            m = pat.search(desc)
            if m is not None:
                ref_no = int(m.group(1)) if m.group(1).isdigit() else None
                break
        return ref_no

    @staticmethod
    def get_patterns_from_csv():
        patterns = []
        with open('ibank_pull_matcher.csv', 'r') as f:
            reader = csv.reader(f)
            for row in reader:
                patterns.append((row[0].strip(), row[1].strip(), row[2].strip()))

        return patterns

    def assign_target_accounts(self):
        patterns = Importer.get_patterns_from_csv()
        for trx in self.trxs:
            trx['target_acc_name'] = "Imbalance-IDR"

            matched = False
            for pat in patterns:
                if 're' in pat[0]:
                    m = re.search(pat[1], trx['desc'].upper(), flags=re.IGNORECASE)
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
                trx['target_acc_name'] = cfg['cash_target_account']
            # trx['ref_no'] = get_refno_from_desc(trx['desc'])
            print "{0:10}|{1:50}|{2:5}|{3:30}".format(trx['date'], trx['desc'], trx['ref_no'], trx['target_acc_name'])

    class BalanceError(ValueError):
        pass

    def import_into_gc(self,acc,transactions,**kwargs):

        (last_split,last_num) = self.gnc.last_txnum(acc)
        last_post_date = last_split.transaction.post_date.date()
        # make a dict of splits by t-num (reference) key
        splits_by_tnum = dict([
                            (int(s.transaction.num) if s.transaction.num.isdigit() else 0, s.guid) for s in acc.splits
                            ])
        if 0 in splits_by_tnum:
            del splits_by_tnum[0]

        # start_num = int(last_import_data['last_tnum']) +1
        num = start_num = last_num + 1
        count = existing_count = 0
        balance = acc.get_balance()
        print "last_post_date: {:%Y/%m/%d %H:%M} \nlast_num:{} \nbalance:{}  ".format(last_post_date,last_num,balance)
        try:
            if not self.dryrun:
                self.session.autoflush = False
            for trx in transactions:
                post_date = dt.strptime(trx['date'], date_format).date()
                new_bal = trx['stm_bal']-trx['value']
                # new_bal = trx['stm_bal']
                # print("- {:%Y/%m/%d %H:%M} >{} stm_bal:{: 16,} val:{: 16,}  bal:{: 16,}"
                #      .format(post_date, last_post_date > post_date, trx['stm_bal'], trx['value'], new_bal))

                if last_post_date > post_date:
                    continue
                elif last_post_date == post_date and new_bal <> balance:
                    # print "SAME DATE  <{}>\nlastsplit<{}> <{}>\n   newbal:<{}> bal:<{}>"\
                    #        .format(trx,last_split,last_split.transaction,new_bal,balance)
                    continue
                elif  new_bal <> balance:
                    # print "\n trx <{}>\nlastsplit<{}> <{}>\n   newbal:<{}> bal:<{}>"\
                    #        .format(trx,last_split,last_split.transaction,new_bal,balance)
                    # print "\n trx <{}>\nlastsplit<{}> <{}>\n   newbal:<{}> bal:<{}>"\
                    #       .format(trx,last_split,last_split.transaction,new_bal,balance)
                    raise Importer.BalanceError

                target_account = self.gnc.imbalance_acct
                acc_code = AccMap.search_description_pattern(trx['desc'])
                if acc_code is not None and acc_code in self.gnc.accounts_by_code.keys():
                    target_account = self.gnc.accounts_by_code[acc_code]


                tnum = trx['ref_no'] = ""
                found_reference = False

                # TODO try to find if already posted
                """
                ref_split = transaction = None
                found_reference = False
                if trx['ref_no'] in splits_by_tnum:
                    ref_split = acc.splits.get(guid=splits_by_tnum[trx['ref_no']])
                    if ref_split.value == trx['value']:
                        found_reference = True
                        transaction = ref_split.transaction
                        ref_split.action = num
                        existing_count += 1
                else:
                    trx['ref_no'] = ""
                """


                transaction = Transaction(
                    currency=acc.commodity,
                    num=tnum,
                    description=trx['desc'],
                    post_date=dt.combine(post_date,datetime.time(12,0)),
                    splits = [
                        Split(
                            value=trx['value'],
                            account=acc,
                            action=num + count
                        ),
                        Split(
                            value=-trx['value'],
                            account=target_account

                        )
                    ]

                )

                self.session.add(transaction)
                balance += trx['value']
                print "{0:15}|{1:12}|{2:8}|{3:35}|{4:3} |{5:15,}|{6:5}-{7:1}|{8:20}".format(
                    transaction.splits[0].action, transaction.splits[1].action,
                    dt.strftime(transaction.post_date, "%d-%m-%y"),
                    transaction.description[:35],
                    ('CRD' if transaction.splits[0].value > 0 else 'DEB'),
                    abs(transaction.splits[0].value),
                    transaction.num, '*' if found_reference else '',
                    transaction.splits[1].account.fullname)

                count += 1
            # END for trx

            if not self.dryrun:
                self.session.commit()

        except Exception as e:
            print "ROLLBACK", sys.exc_info()[0]
            self.session.rollback()
            self.session.close()
            traceback.print_exc()
            raise e
        finally:
            pass

        print "IMPORTED {} of {} TRANSACTIONS start num: {} TO ACCOUNT <{}>".format(
            count, len(self.trxs), start_num, acc.fullname)

    def do_import(self,transactions):
        for code,trxs in transactions.items():
            acc = self.gnc.accounts_by_code[code]
            print "Account %s\nComparing account balances with the bank statement..." % acc.fullname
            if self.check_balance_acct(trxs):
                print "OK"
            else:
                print "FAILED"
                self.dryrun = True

            print "Importing into gnucash book {} {}".format(gnc.book,
                                                    "DRYRUN" if self.dryrun else "Real insert")

            self.import_into_gc(acc, trxs)


#### Utils ######
def decaptcha(screenshot, loc):
    mImgFile = "/tmp/out.jpg"
    ss = cv2.imread(screenshot, 0)
    img = ss[loc['y']:loc['y'] + 22, loc['x']:loc['x'] + 120]
    cv2.threshold(img, 160, 255, cv2.THRESH_BINARY)
    cv2.imwrite(mImgFile, img)
    tesseract = subprocess.Popen(['/usr/bin/tesseract', mImgFile, '-', '-psm', '8', 'digit'], stdout=subprocess.PIPE)
    res = tesseract.communicate()[0]
    return res.strip()


def load_config(config_file=progname + '.config'):
    config = {}
    with open(config_file, 'r') as f:
        config = json.load(f)
    return config


def save_config(config, config_file=progname + '.config'):
    with open(config_file, 'w') as f:
        json.dump(config, f, indent=4)

def get_from_saved_html_history(url):
    trxnum = 0
    balance = 0
    trxs = []
    r = requests.get(url)
    soup = BeautifulSoup(r.text,'lxml')
    table  = soup.find(id="txnHistoryList")
    for row in table.find_all('tr')[2:]:
        fields = row.find_all('span')
        trxnum += 1
        tmp_val = Decimal(fields[4].text.strip().replace(',', ''))
        value = tmp_val if 'Cr' in fields[3].text.strip() else -tmp_val
        stm_bal = Decimal(fields[5].text.strip().replace(',', ''))
        if trxnum == 1:
            balance = stm_bal + value
        balance -= value
        # print trxnum,fields[0],fields[1].text.strip(),tmp_val,value,stm_bal,balance
        trx = {
            'num': Decimal(fields[0].text),
            'seq': trxnum,
            # 'date' : dt.strptime(fields[0].text.strip(), date_format),
            'date': fields[1].text,
            'desc': fields[2].text.strip(),
            'value': value,
            'stm_bal': stm_bal,
            'balance': stm_bal
        }
        trxs.append(trx)

    trxs.reverse()
    # trx['desc']=trx['desc'][:45]
    template = "{seq:>2}: {num:>2} {date:10} {desc:80} {value:>14,} {stm_bal:>14,} {balance:>14,}"
    for seq,tx in enumerate(trxs,1):
        tx['num'] = seq
        print template.format(**tx)
    return trxs

def pull_from_web():
    ibank = Ibank(dump=True)
    try:
        return ibank.process_accounts(accts_to_process,start_date=None)

    except Exception as e:
        traceback.print_exc()
        raise e
    finally:
        ibank.logout()
        ibank.driver.close()
    return None

def transactions_dump(trxs):
    dumpfile = "trx-{}.pickle".format(dt.today().strftime("%Y-%m-%d"))
    with open(dumpfile, 'w') as df:
        pickle.dump(trxs, df)
        print "Dumped {} transactions into {}".format(len(trxs), dumpfile)

def get_from_dumpfile(dumpfile):
    trxs = {}
    with open(dumpfile, 'r') as df:
        trxs =  pickle.load(df)
    return trxs

#####################################################################
### MAIN                                                   ##########
#####################################################################
if __name__ == "__main__":
    import sys
    import pickle
    import argparse
    from gncutils import *

    reload(sys)
    sys.setdefaultencoding("utf8")

    # last day of prevoius month
    date_format = "%d-%m-%Y"
    cfg = load_config()
    default_accounts = ['BNI289979573']
    default_start_date = date.today().replace(day=1) - datetime.timedelta(days=1)
    default_end_date = date.today()

    parser = argparse.ArgumentParser(description='Pull Transaction history from ibank.co.id.')
    parser.add_argument('--book', help='Gnucash DB connection URI [{}]' \
                        .format(cfg['conn_uri']))
    parser.add_argument('--accounts', default=default_accounts, nargs='+', help='Gnucash BNI account code')
    parser.add_argument('--start-date',help='Pull transactions from this date. ')
    parser.add_argument('--end-date',help='Pull transactions until this date. ')
    parser.add_argument('--dump', action='store_true', help='Dump loaded transactions to json file')
    parser.add_argument('--from-dump', help='Load transactions from dump file')
    parser.add_argument('--from-html',  help='Load transactions from saved html')
    parser.add_argument('--no-insert', action='store_true', help='Do not insert into the book')

    args = parser.parse_args()
    print args.accounts

    conn_uri = args.book if args.book else cfg['conn_uri']
    username = cfg['user_id']
    password = cfg['password']
    accounts = cfg['accounts']

    transactions = {}
    gnc = Gncutils('rok', readonly=args.no_insert, do_backup=False)
    gnc.session.no_autoflush
    if args.accounts:
        accts_to_process = gnc.query(Account).filter(
            Account.parent == gnc.accounts_by_fullname['Assets:Bank'],
            Account.code.in_(args.accounts)
        ).all()
    else:
        #accts_to_process = gnc.query(Account).filter(Account.parent == gnc.accounts_by_fullname['Assets:Bank']).all()
        accts_to_process = [gnc.bank_private]

    if args.from_dump:
        transactions = get_from_dumpfile(args.from_dump)
    elif args.from_html:
        transactions = {default_accounts[0]: get_from_saved_html_history(args.from_html)}
        #print transactions
    else:
        transactions = pull_from_web()
        transactions_dump(transactions)
    importer = Importer(dryrun=args.no_insert,book=gnc)
    importer.do_import(transactions)