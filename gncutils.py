from __future__ import print_function
import sys
#import piecash
from piecash import Account,Transaction, Split, open_book
import traceback
from sqlalchemy import inspect, func, text
from sqlalchemy.orm import object_mapper, ColumnProperty
import sqlalchemy
from datetime import datetime as dt
from datetime import date
import re
import logging



class Gncutils:
    BOOKS = {
        # 'rok': 'sqlite:////home/rok/Documents/accounts/rok/rok-test.sqlite.gnucash',
        # 'dbi': 'sqlite:////home/rok/Documents/accounts/dbi/gnucash-book.sqlite.gnucash',
        'test': 'mysql://root:bluemoon@localhost/gnc_dbitest',
        'rok': "mysql://gnc:karamba@localhost/gnc_rok",
        'dbi': "mysql://gnc:karamba@localhost/gnc_dbi"
    }

    def __init__(self, bookname='rok', readonly=True, open_if_lock=True, do_backup=True):
        self.bookname = bookname
        self.book = open_book(uri_conn=Gncutils.BOOKS[bookname], readonly=readonly, open_if_lock=open_if_lock,
                              do_backup=do_backup)
        self.session = self.book.session
        self.connection = self.session.connection()
        self.currency = self.book.currencies.get(mnemonic='IDR')

        self.accounts = self.book.accounts
        self.transactions = self.book.transactions
        self.splits = self.book.splits
        self.query = self.book.session.query
        self.trxs = []
        self.accounts_by_fullname = dict([(a.fullname, a) for a in self.book.accounts])
        self.accounts_by_code = dict([(a.code.strip(), a) for a in self.book.accounts if a.code.strip() not in ''])

        self.imbalance_acct = self.accounts_by_code['IMBALANCE-IDR']
        self.bank_work = self.accounts_by_code['BNI373352345']
        if self.bookname in 'rok':
            self.bank_private = self.accounts_by_code['BNI289979573']
            self.bank_saving = self.accounts_by_code['BNI387014784']

        self.logger = logging.getLogger('sqlalchemy.engine')
        # self.logger.setLevel(logging.WA)

    def last_txnum(self,account):
        txnum = None
        try:
            split = self.query(Split).filter(
                Split.account_guid == account.guid).order_by(text('splits.action+0 DESC')).first()
            txnum = int(split.action if split else 0)
            return split,txnum
        except ValueError:
            pass
        return None, int(0)


    def transactions_query(self, *filter):
        q = self.query(Split, Transaction, Account)
        q = q.filter(Split.transaction_guid == Transaction.guid, Split.account_guid == Account.guid)
        if filter:
            q = q.filter(*filter)
        return q

    def deposits_query(self):
        myaccts = [self.accounts_by_code[code] for code in 'BNI373352345', 'BNI387014784', 'BNI289979573']

    def latest_no_action(self):
        q = self.transactions_query.filter(
            Split.action == '', Transaction.post_date > dt(2015, 12, 1)
        )
        return q

    def fix_imbalances(self,acc=None):
        trxs = []
        splits = []
        if acc is None:
            acc = self.bank_private

        imbalance_guid = self.imbalance_acct.guid
        print("Fixing imbalanced splits on account", acc.fullname)
        for row in self.imbalanced_splits_query().all():
            tx = row.Transaction
            for s in tx.splits:
                print(s)
                if s.account_guid == acc.guid:
                    print("tx {}      -   {}".format(tx, s))
                    acc_code = AccMap.search_description_pattern(tx.description)
                    if acc_code is not None and acc_code in self.accounts_by_code.keys() :
                        row.Split.account_guid = self.accounts_by_code[acc_code].guid
                        splits.append(row.Split)
        return splits

    def imbalanced_splits_query(self):
        return self.transactions_query().filter(Split.account == self.imbalance_acct)


    def find_target_account(self, oldspl):
        target_acc = None
        action = oldspl.action.upper().strip()
        desc = oldspl.transaction.description.upper().strip()
        res_desc = AccMap.search_description_pattern(desc)

        acctype = 'a' if 'ASSET' in  oldspl.account.type else 'exp'
        res_customer = AccMap.search_customer_pattern(desc,acctype )
        res_accnamemap = AccMap.accname.get(oldspl.account.fullname.upper())
        trying = 'ACTION'
        code = None
        if (oldspl.account.type in 'EXPENSE') and (action in AccMap.action.keys()):
            code = AccMap.action[action]
        elif res_desc is not None:
            trying += '|DESC'
            code = res_desc
        elif res_customer is not None:
            trying += '|CUST'
            code = res_customer
        elif res_accnamemap is not None:
            trying += '|ACCMAP'
            code = res_accnamemap
        elif oldspl.account.code in self.accounts_by_code.keys():
            trying += '|CODE'
            code = oldspl.account.code
        else:
            code = self.imbalance_acct.code

        print(oldspl.account.fullname,'  -> ',code,trying)

        if code in self.accounts_by_code.keys():
            target_acc = self.accounts_by_code[code]
        else:
            target_acc = self.imbalance_acct
        return target_acc

    def get_trx_to_fix(self):
        saccts = [
            #self.accounts_by_fullname['Liabilities:DBI'],
            self.accounts_by_fullname['Expenses:DBI']
        ]
        #saccts.extend(self.accounts_by_fullname['Liabilities:DBI'].children)
        saccts.extend(self.accounts_by_fullname['Expenses:DBI'].children)
        saccts_guid = [a.guid for a in saccts]
        cash_accounts = [self.accounts_by_code[code].guid for code in ['cash-2015-out','cash-rok'] ]
        qry1 = self.query(Transaction,Split).join(Transaction.splits).join(Account) \
            .filter(
                Account.type.in_(['EXPENSE']),
                Split.account_guid.notin_(saccts_guid),
                Transaction.splits.any(Split.account_guid.in_(cash_accounts))
                #Account.parent_guid.in_([a.guid for a in saccts])
            )
        return qry1.all()


    def move_dbi_expenses(self):
        self.session.autoflush = False

        transactions = self.get_trx_to_fix()

        banks = []
        target_accounts = {}
        try:
            for trx, spl in transactions:
                skip = False
                exists = False

                tx = Transaction(
                    currency=self.currency,
                    description=trx.description,
                    enter_date=trx.enter_date,
                    post_date=trx.post_date,
                    num=trx.num,
                    notes=trx.notes
                )

                #print(tx.post_date,tx.description,tx.guid,trx.guid)
                #splits = []
                for s in trx.splits:
                    action = ''
                    if s is spl:
                        if s.account is self.accounts_by_code['cash-2015-out']:
                            s.account = self.accounts_by_code['cash-rok']
                        continue
                    else:
                        target_acc = self.find_target_account(s)
                        if target_acc is None:
                            target_acc = self.imbalance_acct

                    #print('     ',s.value,target_acc.fullname,target_acc.commodity)
                    if s.account.type in 'BANK':
                        if s.action.isalnum():
                            action = s.action
                        else:
                            action = '???'
                        banks.append((trx, s, action))
                    else:
                        action = ''

                    if (skip | exists) is True:
                        break;

                    tx.splits.append(Split(
                        account=target_acc ,
                        memo=s.memo,
                        value=s.value,
                        quantity=s.quantity,
                        action=action
                    ))

                    if target_accounts.has_key(target_acc.fullname):
                        target_accounts[target_acc.fullname]  += 1
                    else:
                        target_accounts[target_acc.fullname]  = 1
                # END for

                if skip or exists or (len(tx.splits)<2):
                    status = "EXISTS" if exists else "SKIP" if skip else "NO WAY!"
                    del tx.splits
                    del tx
                else:
                    #self.trxs.append(tx)
                    print("------  -----------------------")
                    print(tx.post_date,tx.description,tx.guid,trx.guid)
                    print(*trx.splits, sep='\n')
                    self.book.session.add(tx)

            #self.book.session.commit()
        except Exception as e:
            print("ROLLBACK", sys.exc_info()[0])
            self.book.session.rollback()
            traceback.print_exc()
            raise e
        finally:
            print("target_accounts:\n--------",sep='\n')
            for v,k in sorted([(v,k) for k,v in target_accounts.iteritems()]):
                print("%s %d" % (k,v))

            print("\n-------missing from Accmap:", AccMap.get_missing())
            #print("\n-------banks:", *banks, sep='\n')

            self.book.session.close()


def print_trx(transactions):
    for tr in transactions:
        print("- {:%Y/%m/%d} : {}".format(tr.post_date, tr.description))
        for spl in tr.splits:
            print("\t{amount}  {direction}  {account} : {memo}".format(
                amount=abs(spl.value),
                direction="-->" if spl.value > 0 else "<--",
                account=spl.account.fullname,
                memo=spl.memo)
            )


class AccMap(dict):
    action = {
        # action : target acc code
        'TOM': 'exp-dbi-tom',
        'SIM': 'exp-dbi-sim',
        'MAN': 'exp-dbi-man',
        'DBI': 'exp-dbi-int',
        'ADAM': 'exp-adam',
        'OLLE': 'exp-olle',
        'RIC': 'exp-rich'
    }

    descriptions = {
        # desc patterns : target acc code
        'BY ADM|BY TRX|JASA GIRO|BY ADMIN|BIAYA ATM|BIAYA CIRRUS|BY GANTI|BY INQ': 'exp-bankcharges',
        'HOSTINGER' : 'exp-hosting',
        'PLN|TUTI' : 'exp-home',
        'PPH': 'exp-taxes',
        'MAYA': 'exp-gear',
        '87740347134|81236111512|VCR TLKOMSEL': 'exp-phone',
        'BUNGA': 'inc-interest',
        'SKT/' : 'exp-shakti',
        'PENARIKAN TUNAI': 'cash-atm'
    }

    asset_targets = {
        # desc patterns : target acc code
        re.compile("TRANSFER KE.*(\d{4,5})$", flags=re.IGNORECASE): 'a-dbi',
        re.compile("TRANSFER DARI.*(\d{4,5})$", flags=re.IGNORECASE): 'dummy',
        'RAUL BOSCARINO|NURDIANSYAH': 'a-dbi',
        'OLLE|ADAM|RICH|JJ': 'a-rec'
    }

    customer = {
        # action : target acc code suffix
        'TOM': 'dbi-tom',
        'SIM': 'dbi-sim',
        'MAN': 'dbi-man',
        'DBI': 'dbi-int',
        'ADAM': 'exp-adam',
        'JJ': 'exp-jj',
        'OLLE': 'exp-olle',
        'RIC': 'exp-rich'
    }
    accname = {
        # old acc name : target acc code
        'ASSETS:LOANS': 'a-loans',
        'ASSETS:BANK:BNI56': 'a-dbi-bni56',
        'INCOME:INTEREST INCOME': 'inc-interest',
        'ASSETS:TRANSIT': 'a-dbi-transit',
        'ASSETS:CASH:BASO': 'a-dbi-baso',
        'LIABILITIES:CREDIT': 'l-credit',
        'ASSETS:RECEIVABLE:OLLE': 'a-rec-olle',
        'EXPENSES:COMMUNICATIONS': 'exp-dbi-int',
        'EXPENSES:TRAVEL': 'exp-dbi-int',
        'EXPENSES:PROJECTS:TOM': 'exp-dbi-tom',
        'EXPENSES:PROJECTS:MANTRA': 'exp-dbi-man',
        'EXPENSES:PROJECTS:SIMON': 'exp-dbi-sim',
        'EXPENSES:PROJECTS:ADAM': 'exp-adam',
        'EXPENSES:PROJECTS:OLLE': 'exp-olle',
        'EXPENSES:PROJECTS:RICH': 'exp-rich'
    }

    refno_patterns = (
        re.compile("\[ref:(\d{4,5})\]", flags=re.IGNORECASE),
        re.compile("TRANSFER KE.*(\d{4,5})$", flags=re.IGNORECASE),
        re.compile("TRANSFER DARI.*(\d{4,5})$", flags=re.IGNORECASE),
        re.compile("ECHANNEL.*(\d{4,5})$", flags=re.IGNORECASE)
    )

    missing = {}

    @staticmethod
    def get_missing():
        return AccMap.missing

    @staticmethod
    def search_description_pattern(description):
        result = None
        desc = description.upper().strip()
        for k in AccMap.descriptions.keys():
            for p in k.split('|'):
                if p in desc.upper():
                    result = AccMap.descriptions[k]
                    break

        return result

    @staticmethod
    def get_by_accmap(accname):
        result = None
        name = accname.upper().strip()
        for k in AccMap.accname.keys():
            for p in k.split('|'):
                if p in name.upper():
                    result = AccMap.accname[k]
                    break
        if result is None:
            AccMap.missing[name] = AccMap.missing[name] + 1 if AccMap.missing.has_key(name) else 1
        return result

    @staticmethod
    def search_customer_pattern(description, acctype='exp'):

        result = None
        desc = description.upper().strip()
        for k in AccMap.customer.keys():
            for p in k.split('|'):
                if p in desc.upper():
                    result = acctype.lower().strip() + '-' + AccMap.customer[k]
                    break

        return result


def fix_old345_target_accounts():
    book = Gncutils('rok', readonly=False, do_backup=False)
    oldbook = Gncutils('dbi', readonly=True, do_backup=False)

    old345 = oldbook.accounts_by_code['BNI373352345']
    new345 = book.accounts_by_code['BNI373352345']
    imb_acc = book.accounts_by_code['imb-old']
    fixed_count = 0
    count = 0
    try:
        for tx, imb_split in [(s.transaction, s) for s in imb_acc.splits]:
            desc = tx.description.upper()
            count += 1
            print("TRANSACTION {} {}".format(tx.description, tx.post_date))
            result = None

            # try by action first
            target_code = search_pattern(imb_split.action, action_targets)
            if target_code is not None:
                targetacc = book.accounts_by_code[target_code]
                result = 'OK'
                print("\tACTION:{}  {}".format(imb_split.action, targetacc))
                imb_split.account = targetacc
                continue

            for s in tx.splits:
                print("\tSPLIT:<{:40}  {: 16,}  ACT:{}".format(s.account.fullname, s.value, s.action))

                if s.account == new345:
                    # find target in old book
                    try:
                        old_split = old345.splits.get(action=s.action)
                        old_trx = old_split.transaction
                        if old_trx.post_date <> s.transaction.post_date:
                            result = 'DIFFDATE'
                            print("DATE   !!!!! {}  {}\n".format(old_split, s))
                            # raise ValueError

                            break
                        if old_split.value <> s.value:
                            result = 'DIFFVAL'
                            print("VALUE   !!!!! {}  {}\n".format(old_split, s))
                            raise ValueError
                            break

                        if fix_imb_split(book, imb_split, old_trx):
                            result = 'OK'
                        else:
                            result = 'NOFIX'
                    except KeyError:
                        result = 'NOTFOUND'
                        break
            if result not in 'OK':

                print("{:10}  {}\n".format(result, tx))
                print('------------------------')
            else:
                fixed_count += 1
        print("fixed: {} of {}".format(fixed_count, count))
        book.session.commit()
    except Exception as e:
        print("ROLLBACK", sys.exc_info()[0])
        book.session.rollback()
        traceback.print_exc()
        raise e
    finally:
        # pass
        book.session.close()


def remove_dbi2015_transactions():
    book = Gncutils('rok', readonly=False, do_backup=False)
    myaccts = book.get_myaccounts()
    transactions = [tx for tx in book.book.transactions if any(s.account in myaccts for s in tx.splits)]
    try:
        for tx in transactions:
            print("deleting {}".format(tx))
            book.session.delete(tx)

        book.session.commit()
    except Exception as e:
        print("ROLLBACK", sys.exc_info()[0])
        book.session.rollback()
        traceback.print_exc()
        raise e
    finally:
        # pass
        book.session.close()


if __name__ == "__main__":
    from gncutils import *

    logging.basicConfig()
    logging.getLogger('sqlalchemy.engine').setLevel(logging.WARN)
    #engine = sqlalchemy.create_engine()
    gnc = Gncutils('rok', readonly=True, do_backup=False)
    acc = gnc.bank_work
    tnum = gnc.last_txnum(acc)
    print('LAST_TXNUM',acc, tnum)
