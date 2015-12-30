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
*Aim: Bring together different elements to define the four tariff components for import tariffs
***************

***************
*Datasets used:
*DemandScaling.dta (from "8 Demand scaling.do")
*TransmissionExitChargingRate.dta (from "6 Transmission exit charges.do"
*Super_Red_Rate_FCP_Demand_Final.dta (from "2 FCP Charge1 Demand.do"
*LRICCharge1Final.dta (from "4 LRIC Charge1 Demand.do)

***************
*Input variables used:

*t1113c1 = Number of days in year
*t1113c3 = Annual hours in super red

*t953c3 = Demand or Generation
*t935c22 = Days for which not a customer
*t953c7 = Maximum import capacity or maximum export capacity (kVA)
*t953c8 = Capacity subject to DSM/GSM constraints (kVA)
*t935c15 = Peak-time kW divided by kVA capacity
***************

***************
*Programs invoked:
*BlankToZero
***************

set type double
clear
use DemandScaling

merge company line using TransmissionExitChargingRate
drop _merge

sort company line
merge company line using Super_Red_Rate_FCP_Demand_Final

drop _merge
sort company line
merge company line using LRICCharge1Final

replace app=appLRIC if appLRIC=="LRIC"

*keep  company line app ImportCapFixedAdder AssetBasedChargingRateResidRev  EDCMImportCapINDOCCharge  FixedChargeSoleUseAsset NRDOCImportCapCharge LRICLocalCharge1 LRICSuperRedRate SuperRedRate_FCP_Demand FCPCapCharge1 TransmissionExitCapCharge

*EDCM demand tariffs
********************

BlankToZero LRICLocalCharge1 FCPCapCharge1 LRICSuperRedRate SuperRedRate_FCP_Demand

*BlankToZero FixedChargeSoleUseAsset
ren DemFixedChargeSolePenceDay  DemFixedCharge

BlankToZero ImportCapFixedAdder AssetBasedChargingRateResidRev EDCMImportCapINDOCCharge NRDOCImportCapCharge LRICLocalCharge1 FCPCapCharge1 TransmissionExitCapCharge

gen ImportCapacityCharge= (ImportCapFixedAdder + AssetBasedChargingRateResidRev + EDCMImportCapINDOCCharge +NRDOCImportCapCharge) + LRICLocalCharge1 + FCPCapCharge1 + TransmissionExitCapCharge

gen SuperRedDem = SuperRedRate_FCP_Demand if app=="FCP"
replace SuperRedDem=LRICSuperRedRate if app=="LRIC"

*Define SuperRedPreAdjustment for purpose of calculating AnnualFCPLRICCharge in file "10 Result output.do"
gen SuperRedPreAdjustmentDem=SuperRedDem

replace SuperRedDem=SuperRedDem+ImportCapacityCharge*(t1113c1-t935c22)/t935c15/(t1113c3-t935c23) if ImportCapacityCharge<0&t935c15~=0&t935c23~=t1113c3
gen OriginalImportCapChage=ImportCapacityCharge

replace SuperRedDem = 0 if SuperRedDem<0
replace ImportCapacityCharge= 0 if ImportCapacityCharge<0

*Calculating exceeded capacity

BlankToZero FCPCapCharge1Exceeded LRICLocalCharge1Exceeded SuperRedRateFCPExceeded LRICSuperRedRateExceeded

gen ExceededDemCapCharge = ImportCapacityCharge + (( (FCPCapCharge1Exceeded + LRICLocalCharge1Exceeded) + ((SuperRedRateFCPExceeded+ LRICSuperRedRateExceeded)* t935c15*(1-t935c23/t1113c3)* t1113c3)/ (t1113c1-t935c22))) * (1 - ((t935c2-t935c18)/t935c2)) if t935c18~=0

replace ExceededDemCapCharge=ImportCapacityCharge if t935c18==0

*EDCM generation tariffs
********************

ren GenFixedChargeSolePenceDay  GenFixedCharge
ren ExportCapCharge ExportCapacityCharge

gen SuperRedCreditGen = CreditGenFCPTariff if app=="FCP"
replace SuperRedCreditGen =CreditGenLRICTariff if app=="LRIC"

*See revisions to "6a Export capacity charges.do", done on 11Sept2014
*gen ExceededExpCapCharge=ExportCapacityCharge

*House keeping
drop _merge

*Ensuring that demand charges are presented only for tariffs where import capacity is not #NA

replace SuperRedDem=. if  ImportCapNA==1
replace DemFixedCharge=. if ImportCapNA==1
replace ImportCapacityCharge=. if ImportCapNA==1
replace ExceededDemCapCharge=. if ImportCapNA==1

*Need to re-define ChargeableExportCap as this variable comes from FCP or LRIC;

replace ChargeableExportCap = t935c4+t935c5+t935c6

replace SuperRedCreditGen=. if ExportCapNA==1
replace GenFixedCharge=. if ExportCapNA==1
replace ExportCapacityCharge =. if ExportCapNA==1
replace ExceededExpCapCharge=. if ExportCapNA==1

*Presenting all tariffs
***********************

gen SuperRedDem_r=round(SuperRedDem,0.001)
gen DemFixedCharge_r=round(DemFixedCharge,0.01)
gen ImportCapacityCharge_r=round(ImportCapacityCharge, 0.01)
gen ExceededDemCapCharge_r=round(ExceededDemCapCharge,0.01)

gen SuperRedCreditGen_r=round(SuperRedCreditGen,0.001)
gen GenFixedCharge_r=round(GenFixedCharge,0.01)
gen ExportCapacityCharge_r=round(ExportCapacityCharge,0.01)
gen ExceededExpCapCharge_r=round(ExceededExpCapCharge,0.01)

*Append data on rows in 935 that are neither import nor export: for completeness purpose of auditing table 4601. See "0 Correct raw simulated data"

append using NoTariffs935.dta

sort comp line
order comp line SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r  GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r

save Tariffs, replace
