# Double-Entry-Bookkeeper
A simple double-entry bookkeeping program based on Excel, using VBA



There are two actions: Enter A
Transaction, and Generate A Report

Entering A Transaction is the more complicated of the two. This starts at Sub Button3_Click() and
also includes Userform1.

In double-entry book-keeping, a transaction involves two or more accounts. The accounts  included in this version are:

Cash

Receivables

Other asset

Payroll exp

Admin exp

Other Opex

Taxes, etc

Freelancers


Revs

Loans

Payables

Deferred Revenues

Accrued Payroll

Other accrued/uncashed checks 


A worksheet is generated for each account. If the worksheets don't already exist, they are created in the
InitializeSheets() subroutine. 

The accounts are classified as Asset/Expense accounts or Liability/Revenue accounts.  The worksheet names of Asset/Expense accounts must begin with "A-", because it is based on this worksheet name that the program deduces the account type.


 


The following rules apply in double-entry book-keeping:



In Asset/Expense accounts, an increase is called a Debit (DR) and a decrease is called a Credit (CR). Traditionally,
in listing the DR and CR for this type of account, the DR is on the left and the CR on the right. The columns in the "A-" sheets are titled in this traditional way. In Liability/Revenue accounts, an increase is a CR (listed on the left) and a decrease is a DR (right).



An increase in an asset account (a DR) must be balanced by either an increase in liabilities/revenues (a CR) or a decrease
in another asset account (a CR), or some combination of these. Ultimately, in any transaction the sum of the CRs must equal the sum of the DRs.
 


The system is fairly intuitive. For
instance, if you receive cash from a sale, you enter a positive amount in the
Cash account (this would be a DR) and a positive amount in the Revenue account
(this would be a CR). For the second entry, the default amount is equal to the
amount of the first entry; but make sure the sign – positive or negative – is correct.



 


The program will make entries
correspondingly in the A-cash worksheet and the L-revs worksheet. Note that the
DR amount is equal to the CR amount. You can also, optionally, put in a Memo
explaining what the transaction is. For instance, "One hour training Ms. Lawrence".
And the transaction includes a date.


 


There is also an entry made in the
J_Entry worksheet. This is what is known as a Journal Entry.It easier to see
how the entries balance because they are in one worksheet rather than two.
Unlike the account worksheets, the J_Entry worksheet does not automatically
arrange the transactions in chronological order.


 


In some cases, there will be more than
two entries. For instance, suppose you make a sale (Revenues) for $10, but only
$5 is in Cash while the rest will be paid later. The remaining $5 goes in the
Receivables account. The system prompts for additional entries if the DRs and
the CRs are not equal.


 


The Generate A Report procedure begins with Sub Button4_Click(), which
uses Userform2. This generates a Balance Sheet for any particular date, or an
Income Statement, given a beginning and ending date, or both.


 


More accounts can be added with a few
modifications.  To an account:


 


(a)  In the InitializeSheets subroutine, increment
the number in SheetNameArray() and add a worksheet name.


 


Dim SheetNameArray(1 To 14) As String


 


SheetNameArray(1) = "A-cash"


SheetNameArray(2) =
"A-rcbls"


SheetNameArray(3) =
"A-other"


SheetNameArray(4) =
"A-payrollexp"


SheetNameArray(5) =
"A-adminexp"


SheetNameArray(6) =
"A-taxetc"


SheetNameArray(7) =
"A-otheropex"


SheetNameArray(8) =
"A-freelancexp"


SheetNameArray(9) = "L-revs"


SheetNameArray(10) =
"L-defdrevs"


SheetNameArray(11) =
"L-payables"


SheetNameArray(12) =
"L-accpayroll"


SheetNameArray(13) =
"L-loans"


SheetNameArray(14) =
"L-otheraccrued"


 


The
only requirement is that any new Asset/Expense accounts be named with a
"A-" prefix, and the size of the array should be
incremented for each new account. Also, increment the number in the For NameArrayCell
= 1 To 14… Next loop, increment this number also.


 


(b) The worksheet called "options"
has a range within it also named "options". Userform1, called by Sub
mySForm, refers to this range when presenting account choices. So it is
necessary to add the name of the new account to this range.


 


(c) The Function DirectSheet returns
the name of the worksheet belonging to the account. So add two lines of code
for each account added:


 


Case Is = "Name_Of_Account"


   
Destination_Sheet = "Name_Of_Worksheet"


 


Remember that Asset/Expense worksheet
names must begin with "A-".


 


(d) Update the Generate A Report functions.
This is quite intuitive. Update the titles in column A and update the formulae accordingly.



 


For a balance sheet, Equity is Total
assets minus Total liabilities. For an Income Statement, Income is Revenues
minus Expenses. 


