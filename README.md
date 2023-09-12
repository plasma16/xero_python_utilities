# xero_python_utilities
This repo contains a list of utilities written in Python that does some common post-processing of Xero generated reports and/or Chart Of Accounts. 
They include some common accounting processes that is not natively supported by Xero. 

Account_Payable_Remove_Matching
-------------------------------
Removes the matching pair of debit and credit entries to retain only the outstanding payable amount from "Account Payable". 
Generates 2 files:
1. Outstanding.xlsx : contains only the outstanding amount
2. removed.xlsx     : contains matched debit & credit pairs for cross-checking, sum of debit column should match sum of credit column
