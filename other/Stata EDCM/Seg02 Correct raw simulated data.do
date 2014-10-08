* Copyright licence and disclaimer
* 
* Copyright 2012-2014 Reckon LLP, Pedro Fernandes and others. All rights reserved.
* 
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
* 
* 1. Redistributions of source code must retain the above copyright notice,
* this list of conditions and the following disclaimer.
* 
* 2. Redistributions in binary form must reproduce the above copyright notice,
* this list of conditions and the following disclaimer in the documentation
* and/or other materials provided with the distribution.
* 
* THIS SOFTWARE IS PROVIDED BY AUTHORS AND CONTRIBUTORS "AS IS" AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL AUTHORS OR CONTRIBUTORS BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
* THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

***************
*Aim: 
*1. Import CSV files with raw input data 
*2. Runs program to check validity of data (Prog_DataCheck)
*2. Make changes to raw data to ensure numerical variables are treated as such
*3. Save corrected data files as .dta files
***************

***************
*Datasets used:
*11.csv
*911.csv
*913.csv
*935.csv
***************

***************
*Programs invoked:
*Prog_DataCheckNew
*Prog_CompaniesInvalidData.do
*Prog_HashValueToMissing
*Prog_BlankToZero
***************

clear

set mem 500m
set type double
*------------------------------

*0.1 Importing Table 11

clear
insheet using 11.csv, c n
replace company= subinstr(company," ","-",.)

*Ensure numerical variables really are numerical (involves replacing "#VALUE!", "" and "#N/A" with missing values)

HashValueToMissing t1113c1  t1113c2 t1113c3  t1113c4 t1113c5 t1113c6 t1113c7 t1113c8 t1113c9 t1113c10 t1113c11 t1113c12
HashValueToMissing t1122c6 
HashValueToMissing t1131c1 t1131c2 t1131c3 t1131c4 t1131c5 t1131c6  t1131c7 t1131c8 t1131c9 t1131c10 t1131c11
HashValueToMissing t1132c1  
HashValueToMissing t1133c1
HashValueToMissing t1134c1

*Transform missing values to 0
BlankToZero t1113c1  t1113c2 t1113c3  t1113c4 t1113c5 t1113c6 t1113c7 t1113c8 t1113c9 t1113c10 t1113c11 t1113c12 t1131c2 t1131c3 t1131c4 t1131c5 t1131c6 t1131c7 t1131c8 t1131c9 t1131c10 t1131c11 t1122c6  t1131c1  t1134c1 t1133c1

sort company

save 11.dta, replace
*------------------------------

*0.2 Importing Table 911

clear
insheet using 911.csv, c n
replace company= subinstr(company," ","-",.)

*Drop rows where location is given as "Not used" or is given as ""

drop if t911c1=="Not used"|t911c1==""

*Check variables are numeric and then transform missing values into 0

HashValueToMissing t911c4  t911c6 t911c7 t911c8 t911c9
BlankToZero t911c4 t911c6 t911c7 t911c8 t911c9 

save 911.dta, replace
*------------------------------

*0.3 Importing Table 913

clear
insheet using 913.csv, c n
replace company= subinstr(company," ","-",.)

*Check variables are numeric and then transform missing values into 0

HashValueToMissing t913c4 t913c5 t913c8 t913c9
BlankToZero t913c4 t913c5 t913c8 t913c9

save 913.dta, replace
*------------------------------

*0.4 Importing Table 935

clear
insheet using 935.csv, c n
replace company= subinstr(company," ","-",.)

*(a) Transform missing values into 0

HashValueToMissing t935c7 t935c10 t935c11 t935c12 t935c13 t935c14 t935c15 t935c16 t935c17 t935c18 t935c19 t935c20 t935c21 t935c22 t935c23 t935c24 t935c25
BlankToZero  t935c7 t935c9 t935c10 t935c11 t935c12 t935c13 t935c14 t935c15 t935c16 t935c17 t935c18 t935c19 t935c20 t935c21 t935c22 t935c23 t935c24 t935c25 

* (b) Deal with "VOID" in t935c2 - t935c6

*Pick-up "VOID" in import capacity (t935c2) and in non-exempt export capacity (t935c5, t935c6, t935c7)

gen ImportCapNA=0
replace ImportCapNA= 1 if t935c2=="VOID"

gen ExportCapNA=0
replace ExportCapNA= 1 if t935c4=="VOID"&t935c5=="VOID"&t935c6=="VOID"

HashValueToMissing t935c2 t935c3 t935c4 t935c5 t935c6
BlankToZero t935c2 t935c3 t935c4 t935c5  t935c6

*Force t935c9 to be in number format
destring t935c9, force replace

*ALERT: Changes below on 11Sep2014 - to keep a dataset with just rows with no import or export capacity data for purpose of reconciling results with table 4601.
*And a dataset that only has obs with either import and/or export capacity

save All935.dta, replace

*Drop rows where there are no export or import capacities; as these will not be considered to be tariifs covered by EDCM
drop if ExportCapNA==1&ImportCapNA==1
save 935.dta, replace

use All935.dta, clear
keep if ExportCapNA==1&ImportCapNA==1
save NoTariffs935.dta, replace

erase All935.dta

***************
*Data files kept:
*11 = Table 11 corrected for identified problems in input data
*911 = Table 911 corrected for identified problems in input data
*913 = Table 913 corrected for identified problems in input data
*935 = Table 935 corrected for identified problems in input data, and without tariffs where import and export capacity is 0
***************
