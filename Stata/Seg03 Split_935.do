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
*1. To split table 935 between FCP and LRIC and save these as 935FCP.dta and 935LRIC.dta respectively
***************

***************
*Datasets used:
*935
***************

***************
*Programs invoked:
*BlankToZero
*NameChange
***************

*Import 935

clear
use 935

*Split data into two datasets: one for FCP and the other for LRIC

gen app="FCP" if regexm(company, "FCP") == 1
replace app="LRIC" if regexm(company, "LRIC") == 1

save 935_v1, replace

keep if app=="FCP"
save 935FCP, replace
clear

use 935_v1.dta

keep if app=="LRIC"
save 935LRIC, replace

*Erase temporary data files
erase 935_v1.dta

***************
*Data files kept:
*935FCP.dta = Table 953 for FCP
*935LRIC.dta =Table 953 for LRIC
***************
