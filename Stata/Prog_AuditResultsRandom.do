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

capture program drop auditresultsrandom
program auditresultsrandom

quietly {

clear
use ResultsToAudit.dta
local current_data = c(filename)

*Defining set of variables in each table

local StataGrp4501 "SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r  GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r"
local ExcelGrp4501 "t4501c2 t4501c3 t4501c4 t4501c5 t4501c6 t4501c7 t4501c8 t4501c9"

local StataGrp4601 "SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r AnnualChargeCapDem AnnualChargeSuperRedDem AnnualChargeFixedDem AnnualCreditCapGen AnnualChargeFixedGen AnnualCreditSuperRedGen AnnualChargeDemR AnnualChargeDemPrevious ChangeAnnualChargeDem PerChangeAnnualChargeDem  AnnualCreditGenR AnnualCreditCapGenPrevious ChangeAnnualCreditGen PerChangeAnnualCreditGen  AnnualSoleUseAssetCharge AnnualTransExitCharge AnnualDirectCostAllocation AnnualIndirectCostAllocation AnnualNetworkRatesAllocation AnnualFCPLRICCharge AnnualScalingFixedAdder AnnualScalingAssetbased Check"
local ExcelGrp4601 "t4601c2 t4601c3 t4601c4 t4601c5 t4601c6 t4601c7 t4601c8 t4601c9 t4601c10 t4601c11 t4601c12 t4601c13 t4601c14 t4601c15 t4601c16 t4601c17 t4601c18 t4601c19 t4601c20 t4601c21 t4601c22 t4601c23 t4601c24 t4601c25 t4601c26 t4601c27 t4601c28 t4601c29 t4601c30 t4601c31 t4601c32"

if "$Optionchecksums"~="0" {
                local StataGrp4501 = "`StataGrp4501'"  + " Checksum5 Checksum7"
                local ExcelGrp4501 = "`ExcelGrp4501'" + " t4501c10 t4501c11"
                }

*A loop across the four tables in the Results worksheet

local TableGrp "4501 4601"

local m: word count `TableGrp'

capture log close

forvalues j =1/`m' {

    local CSVTable: word `j' of `TableGrp'

    *Merging the Stata with the CSV file

    clear
    insheet using "`CSVTable'.csv"

*Dealing with blanks within company names (to make it consistent with program in model)

    sort company line
    save TempAll.dta, replace

    clear
    use "`current_data'"

    sort company line
    merge company line using TempAll
    drop _merge

    sort company line
    save MergedData, replace

*Loop across each of the variables within relevant CSV Table

    local PickTable  StataGrp`CSVTable'

    local n: word count `StataGrp`CSVTable''

*Have a check that number in two lists is the same

    log using Res`CSVTable', replace
    noisily: list company if company[_n-1]!=company

    forvalues i = 1/`n' {

        local StataVar: word `i' of `StataGrp`CSVTable''
        local ExcelVar: word `i' of `ExcelGrp`CSVTable''

*Destringing ExcelVar for cases where Excel produces #NUM

        destring `ExcelVar', force replace

        gen diff`ExcelVar'= `ExcelVar'/`StataVar'
        replace diff`ExcelVar'=1 if `ExcelVar'==0&`StataVar'==0
        replace diff`ExcelVar'=1 if `ExcelVar'==.&`StataVar'==.
        replace diff`ExcelVar'=1 if `ExcelVar'==0&`StataVar'==.
        replace diff`ExcelVar'=1 if `ExcelVar'==.&`StataVar'==0

        gen Match`ExcelVar'="OK" if (diff`ExcelVar'>0.999999&diff`ExcelVar'<1.000001)
        sort company line

        display as error "``StataVar' `ExcelVar'"
        noisily: list comp line `StataVar' `ExcelVar' diff`ExcelVar' if Match`ExcelVar'~="OK"
        save ResultsAudit.dta, replace

        }

    log close

    }

}

end
