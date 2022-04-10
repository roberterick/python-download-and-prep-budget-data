import _mssql,socket,uuid,pymssql
import datetime,csv,os,calendar
import copy
import r.myNumber
import r.myExcel2
import r.myDerive2
import r.myReportList
import r.myTotals
import r.myDate

FY='2023'
CY='2022'
PY='2021'

HEADER='%s-%s Expense Budget'%(CY,FY)
FYC=['2021-07-01','2022-03-31']
FYP=['2020-07-01','2021-06-30']
MAX_CHRS=22
FUNDS={'0800':'Payroll',
       '0810':'Executive',
       '0820':'Advancement Services',
       '0830':'Investments',
       '0840':'Technology Services',
       '0850':'Projects',
       '0860':'Operations',
       '0870':'Accounting',
       '0880':'Human Resources',
       }


class app(object):
    def __init__(self):
        pass
    def begFY(self,dt):
        if dt.month>=7:
            return datetime.date(dt.year,7,1)
        else:
            return datetime.date(dt.year-1,7,1)
    def dtToDate(self,dt):
        return datetime.date(dt.year,dt.month,dt.day)
    def bDate(self,dt):
        return datetime.date(dt.year,dt.month,1)
    def setupDb(self):
        try:
            self.conn=pymssql.connect("ABIDB")
            self.cursor=self.conn.cursor()
        except:
            raw_input('a connection to the database cannot be established...aborting...')
            sys.exit()
    def __enter__(self):return self
    def __exit__(self, exc_type, exc_value, traceback):
        print 'cleaning up...'
        try:
            self.conn.close()
        except:pass
    def getCsvFileList(self,fileName):
        print 'getting %s'%fileName
        flist=[]
        header=''
        with open(fileName, 'rU') as f:
            reader=csv.reader(f)
            i=0
            for row in reader:
                if i==0:
                    header=row
                else:
                    flist+=[row]
                i+=1
        return header,flist
    def promptBatch(self,batch):
        while True:
            res=raw_input('Input maximum batch number:')
            break
    def promptDate(self,prompt,dt):
        while True:
            res=raw_input(prompt%dt)
            if res=='':
                dt=datetime.datetime.strptime(dt,'%m/%d/%Y').date()
                break
            else:
                try:
                    dt=datetime.datetime.strptime(res,'%m/%d/%Y').date()
                except:
                    print 'Please input a date in d/m/y format.'
                break
        if type(dt)==type(''):
            dt=datetime.datetime.strptime(dt,'%m/%d/%Y').date()
        return dt
    def convertStringToDate(self,dt,fmt='%Y-%m-%d'):
        return datetime.datetime.strptime(dt,fmt).date()
    def convertDateSql(self,dt):
        return "'%s'"%dt.strftime('%Y-%m-%d')
    def convertDateToStr(self,dt):
        if dt==None:
            return ''
        else:
            return "'%s'"%dt.strftime('%d/%m/%Y')
    def fixSets(self):
        keys=self.dd.keys()
        for k in keys:
            tmp=list(self.dd[k]['ref'])
            tmp.sort()
            self.dd[k]['ref']=' '.join(tmp)
            
class app2(app):
    def __init__(self):
        app.__init__(self)
        self.zero=r.myNumber.myNumber()
        
        self.setupDb()
        self.setup()
        self.getExpenseData(FYC[0],FYC[1],'%s partial'%CY)
        self.getExpenseData(FYP[0],FYP[1],'%s'%PY)
        self.getAccounts()
        self.getBudgets()
        self.getBudgetData()

##        self.postColumns()
        self.toExcel()
        print 'done'

    def setup(self):
        self.dd={}
        fl=['Fund','Account',
            'Fund Name','Account Name',
            'Description',
            '%s partial'%CY,
            '%s'%PY,
            ]
        self.md,self.fl=r.myDerive2.myDerive2(fl)
        for itm in fl[5:]:self.md[itm]=r.myNumber.myNumber()

    def getAccounts(self):
        self.accounts={}
        sql=self.sql_accounts()
        self.cursor.execute(sql)
        for i,row in enumerate(self.cursor):
            acct,acctname=row
            if not acct in self.accounts:
                self.accounts[acct]=acctname

    def getBudgets(self):
        self.budgets={}
        sql=self.sql_budgets()
        self.cursor.execute(sql)
        for i,row in enumerate(self.cursor):
            fund,acct,desc,amt=row
            amt=r.myNumber.myNumber(amt)
            key=fund,acct
            if not key in self.budgets:
                self.budgets[key]={'fund':fund,'acct':acct,'desc':desc,'amt':r.myNumber.myNumber()}
            self.budgets[key]['amt']+=amt
            
    def getBudgetData(self):
        print 'getting budget data'
        self.ddp={}
        flp=['Fund','Fund Name','Account','Account Name','Description',
             'CY Budget','YTD Actual','Next Year\'s Proposed Budget','Notes']
        self.mdp,self.flp=r.myDerive2.myDerive2(flp)
        for itm in ['CY Budget','YTD Actual']:
            self.mdp[itm]=r.myNumber.myNumber()
            
        self.pages={}
        #post the budget column
        for k in self.budgets.keys():
            fund=self.budgets[k]['fund']
            acct=self.budgets[k]['acct']
            desc=self.budgets[k]['desc']
            amt=self.budgets[k]['amt']
            self.postMainDetail(fund,acct,desc,amt,'CY Budget')
        #post the actuals column
        for k in self.dd.keys():
            fund,acct,desc=k
            amt=self.dd[k]['%s partial'%CY]
            if amt==self.zero:continue
            self.postMainSummary(fund,acct,desc,amt,'YTD Actual')
            
        #squash small rows
        keys=self.ddp.keys()
        for k in keys:
            budget=self.ddp[k]['CY Budget']
            if float(budget)!=0.0:continue
            actual=self.ddp[k]['YTD Actual']
            if float(actual)<-500.00 or float(actual)>500.00:continue
##            print k,self.ddp[k]
##            print k
##            fund,1,acct,2,summarydesc,1=k
            if len(k)!=6:continue
            fund,v1,acct,v2,summarydesc,v3=k
            nk=fund,1,acct,2,'other < +-$500.00',1
            if not nk in self.ddp:
                self.ddp[nk]=copy.deepcopy(self.mdp)
                self.ddp[nk]['Fund']=fund
                self.ddp[nk]['Fund Name']=FUNDS[fund]
                self.ddp[nk]['Account']=acct
                self.ddp[nk]['Description']='other < +-$500.00'
                self.ddp[nk]['Next Year\'s Proposed Budget']=''
                self.ddp[nk]['Notes']=''
            self.ddp[nk]['YTD Actual']+=actual
            self.ddp.pop(k)

    def postMainSummary(self,fund,acct,desc,amt,col='CY Budget'):
        summarydesc=desc.split()[:3]
        summarydesc=' '.join(summarydesc)
        headerkey=fund,1,acct,1
        key=fund,1,acct,2,summarydesc,1 #note the 1
        totalkey=fund,1,acct,3,(None,)
        tabkey=fund,2,(None,),(None,),(None,)
        if not self.ddp.has_key(key):
            self.ddp[key]=copy.deepcopy(self.mdp)
            self.ddp[key]['Fund']=fund
            self.ddp[key]['Fund Name']=FUNDS[fund]
            self.ddp[key]['Account']=acct
            self.ddp[key]['Description']=summarydesc
            self.ddp[key]['Next Year\'s Proposed Budget']=''
            self.ddp[key]['Notes']=''
        self.ddp[key][col]+=amt

        try:
            acctlookup=self.accounts[acct]
        except:
            acctlookup='problem'
            print 'fund:',fund,' account:',acct,' lookup has a problem!'
            
        if not self.ddp.has_key(headerkey):
            self.ddp[headerkey]=copy.deepcopy(self.mdp)
            self.ddp[headerkey]['Account Name']='%s  %s'%(acct,acctlookup)
##            self.ddp[headerkey]['Budget']=''

        if not self.ddp.has_key(totalkey):
            self.ddp[totalkey]=copy.deepcopy(self.mdp)
            self.ddp[totalkey]['Account Name']='%s  %s Total'%(acct,acctlookup) 
        self.ddp[totalkey][col]+=amt

        if not self.ddp.has_key(tabkey):
            self.ddp[tabkey]=copy.deepcopy(self.mdp)
            self.ddp[tabkey]['Account Name']='Total Department Expenses'
        self.ddp[tabkey][col]+=amt

    def postMainDetail(self,fund,acct,desc,amt,col='CY Budget'):
        headerkey=fund,1,acct,1
        key=fund,1,acct,2,desc
        totalkey=fund,1,acct,3,(None,)
        tabkey=fund,2,(None,),(None,),(None,)
        if not self.ddp.has_key(key):
            self.ddp[key]=copy.deepcopy(self.mdp)
            self.ddp[key]['Fund']=fund
            self.ddp[key]['Fund Name']=FUNDS[fund]
            self.ddp[key]['Account']=acct
            self.ddp[key]['Description']=desc
            self.ddp[key]['Proposed']=''
            self.ddp[key]['Notes']=''
        self.ddp[key][col]+=amt

        try:
            acctlookup=self.accounts[acct]
        except:
            acctlookup='problem'
            print 'fund:',fund,' account:',acct,' lookup has a problem!'
            
        if not self.ddp.has_key(headerkey):
            self.ddp[headerkey]=copy.deepcopy(self.mdp)
            self.ddp[headerkey]['Account Name']='%s  %s'%(acct,acctlookup)
##            self.ddp[headerkey]['Budget']=''

        if not self.ddp.has_key(totalkey):
            self.ddp[totalkey]=copy.deepcopy(self.mdp)
            self.ddp[totalkey]['Account Name']='%s  %s Total'%(acct,acctlookup) 
        self.ddp[totalkey][col]+=amt

        if not self.ddp.has_key(tabkey):
            self.ddp[tabkey]=copy.deepcopy(self.mdp)
            self.ddp[tabkey]['Account Name']='Total Department Expenses'
        self.ddp[tabkey][col]+=amt 

    def getExpenseData(self,bdate,edate,col):
        print 'getting data'
##        self.fundaccountamount={}
        self.totals={}
        sql=self.sql_data(bdate,edate)
        self.cursor.execute(sql)
        for i,row in enumerate(self.cursor):
            fund,fundname,acct,acctname,ref,amount=row
            ref=ref[:MAX_CHRS]
            amount=r.myNumber.myNumber(amount)
            key=fund,acct,ref
            if not self.dd.has_key(key):
                self.dd[key]=copy.deepcopy(self.md)
                self.dd[key]['Fund']=fund
                self.dd[key]['Account']=acct
                self.dd[key]['Fund Name']=fundname               
                self.dd[key]['Account Name']=acctname
                self.dd[key]['Description']=ref
            self.dd[key][col]+=amount

    def toExcel(self):
        self.me=r.myExcel2.myExcel2('budget_prep.xlsx')
        rl=r.myReportList.myReportList(self.fl,self.dd)
##        self.createTabs(self.me)

        funds=list(set([x[0] for x in self.dd.keys()]))
        funds.sort()
        for f in funds:
            self.createTab(self.me,f)
            rltmp=filter(lambda x:x[0] in ('Fund',f),rl)
            self.me.printList(rltmp,'%s_prior_year_expense_details'%f)
        self.me.close()
        self.me.openResult()

    def createTab(self,me,fund):
        fl=['Account Name','Description','CY Budget','YTD Actual','Next Year\'s Proposed Budget','Notes']
        items=self.ddp.items()
        items.sort()
        data=filter(lambda itm:itm[0][0]==fund,items)
        nd={}
        for k,v in data:
            nd[k]=v
        rl=r.myReportList.myReportList(fl,nd)
        addons=[HEADER,FUNDS[fund],'']
        addons.reverse()
        for a in addons:rl.insert(0,[a])
        me.printList(rl,fund)        

    def sql_data(self,bdate,edate):
        return '''
select
	trx.fund_number
	,fund.fund_description
	,trx.account_number
	,acct.account_description
	,trx.reference
	,sum(trx.amount) amount
from acctetl.dbo.uo_gl_transaction trx
join acctetl.dbo.uo_fund fund on fund.fund_id=trx.fund_id
join acctetl.dbo.uo_account acct on acct.account_id=trx.account_id
join acctetl.dbo.uo_batch batch on batch.batch_id=trx.batch_id
where
	post_date between '{bdate}' and '{edate}'
	and left(acct.account_number,1) in ('5')
	and left(fund.fund_number,2) in ('08')
	and trx.post_status='Posted'
group by 
	trx.fund_number
	,fund.fund_description
	,trx.account_number
	,acct.account_description
	,trx.reference
having sum(trx.amount)!=0
        '''.format(bdate=bdate,edate=edate)

    def sql_budgets(self):
        return '''
select
	bt.fund_number
	,bt.account_number
	,acct.account_description
	,bt.amount
from acctetl.dbo.uo_fund_budget bt
join acctetl.dbo.UO_ACCOUNT acct on acct.account_number=bt.account_number
where
	scenario_short='Main'
	and fiscal_year='{year}'
        '''.format(year=CY)
    
    def sql_accounts(self):
        return '''
select
	acct.account_number
	,acct.account_description
from acctetl.dbo.uo_account acct
        '''


        
with app2() as s:pass
