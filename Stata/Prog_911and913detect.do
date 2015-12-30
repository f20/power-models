* Copyright licence and disclaimer
*
* Copyright 2015 Pedro Fernandes and others. All rights reserved.
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

quietly {

*Identify whether set of files include 911 and 913 tables and define macros

foreach tab in 911 913{
    capture confirm file "`tab'.csv"
    if _rc==0 {
        global f`tab' = 1
        }
    else {
        global f`tab' = 0
        }
    }

*Create empty datasets for cases where 911.csv or 913.csv are not included

*Case where there is no 911.csv (ie dataset does not include FCP companies)

if $f911==0{

    clear
    set obs 0
    gen company=""
    gen FCPRevenue=.
    gen  AggFCPSuperRedGenCredit=.

    sort company
    save FCPRevenue.dta, replace

    clear
    set obs 0
    gen company=""
    gen line=.
    gen SuperRedRate_FCP_Demand=.
    gen FCPCapCharge1=.
    gen FCPCapCharge1Exceeded=.
    gen SuperRedRateFCPExceeded=.
    gen CreditGenFCPTariff=.

    sort company line
    save Super_Red_Rate_FCP_Demand_Final.dta, replace

    }

*Case where there is no 913.csv (ie dataset does not include LRIC companies)

if $f913==0{

    clear
    set obs 0
    gen company=""
    gen LRICRevenue=.
    gen  AggLRICSuperRedGenCredit=.

    sort company
    save LRICRevenue.dta, replace

    clear
    set obs 0
    gen company=""
    gen line=.
    gen LRICLocalCharge1=.
    gen    LRICSuperRedRate=.
    gen LRICLocalCharge1Exceeded=.
    gen LRICSuperRedRateExceeded=.
    gen CreditGenLRICTariff=.
    gen appLRIC=""

    sort company line
    save LRICCharge1Final.dta, replace

    }

}
