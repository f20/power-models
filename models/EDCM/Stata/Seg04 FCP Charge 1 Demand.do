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
*Aim: Calculate FCP charge 1 - covers (a) super red rate FCP, (b) local capacity charge 1 FCP (c) FCP exceeded capacity charge
***************

***************
*Datasets used:
*935FCP
*911
*11
***********

***************
*Programs invoked:
*BlankToZero
*CheckZero
***************

clear
use 911

*1. Calculate ActiveFlow and ReactiveFlow 

gen ActiveFlow=(t911c6+t911c8)
gen ReactiveFlow=(t911c7+t911c9)

*2. Set to zero any negative Charge 1 

replace t911c4=0 if t911c4<0

rename t911c1 Location
sort company Location

save 911_v2a, replace

*3. Create dataset that contains: Company, Location and Parent_location

use 935FCP
rename t935c8 Location
sort company Location

merge company Location using 911_v2a
ren t911c3 Parent_Location

keep  company Location Parent_Location
sort company Location

*4. Create dataset that contains: Company, Location and Grandparent location

ren Location True_Location
rename Parent_Location Location
sort company Location

merge company Location using 911_v2a

ren t911c3 Grandparent_Location
ren Location Parent_Location
ren True_Location Location

keep company Location Parent_Location Grandparent_Location

sort company Location
save Ancester_Location, replace

*5. Combine Table 935FCP with information on identity of parents and grandparents as worked out above

clear
use 935FCP

rename t935c8 Location
sort company Location

merge company Location using Ancester_Location
keep if _merge==3
drop _merge
save Combined_1, replace

*5. Combine 911 with parent and grandparent in order to pick up relevant variables about parents and grandparent

*5a  An initial detour: Create two dataset where information is renamed as if it relates to parents, and subsequently, grandparents.

*5a.i Creating a dataset for parents

clear
use 911_v2a

ren Location Parent_Location
ren t911c4 Parent_Charge1_kVAyear
ren ActiveFlow ParentActiveFlow
ren ReactiveFlow ParentReactiveFlow

keep company Parent_Location Parent_Charge1_kVAyear ParentActiveFlow ParentReactiveFlow

sort company Parent_Location
save Parent_details, replace

*5a.ii - Creating a dataset for grandparents

clear
use 911_v2a

ren Location Grandparent_Location
ren t911c4 Grandparent_Charge1_kVAyear
ren ActiveFlow GrandparentActiveFlow
ren ReactiveFlow GrandparentReactiveFlow

keep company Grandparent_Location Grandparent_Charge1_kVAyear GrandparentActiveFlow GrandparentReactiveFlow

sort company Grandparent_Location
save Grandparent_details, replace

*6. Merge Combined_1 with information about parents and grandparents

clear
use Combined_1

sort company Parent_Location
merge company Parent_Location using Parent_details
keep if _merge==1|_merge==3
drop _merge

sort company Grandparent_Location
merge company Grandparent_Location using Grandparent_details
keep if _merge==1|_merge==3
drop _merge

sort company
save Combined_2, replace

*7. Match with data on number of hours in super-red time band in year

clear
use 11.dta
keep company t1113c1 t1113c3
sort company

merge company using Combined_2
keep if _merge==3

save Combined_3, replace

*8. Compute super-red rate. Need to calculate for parents and grand-parents; and then make adjustments with respect to DSM and new customers

*Create variable to pick up restriction on kVAr/kVA

gen CappedkVAr_kVA= t935c16
replace CappedkVAr_kVA = sign(t935c16)*(1-t935c15^2)^0.5 if  (t935c15^2+ t935c16^2> 1)& t935c15^2<= 1
replace CappedkVAr_kVA = 0 if  t935c15^2> 1

*10a - Parent: 

CheckZero t1113c3
gen TempVarP=ParentActiveFlow^2+ParentReactiveFlow^2

*Two cases where t935c15~=0; distinguish between cases where TempVarP~=0 and when TempVarP==0

gen SuperRedRate_Parent = (Parent_Charge1_kVAyear/t1113c3)* 100*(abs(ParentActiveFlow)-ParentReactiveFlow*(CappedkVAr_kVA/t935c15))/((ParentActiveFlow^2+ParentReactiveFlow^2)^0.5) if t935c15~=0&TempVarP~=0
replace SuperRedRate_Parent= (Parent_Charge1_kVAyear/t1113c3)* 100 if t935c15~=0&TempVarP==0

*Two cases where t935c15==0; distinguish between cases where TempVarP~=0 and when TempVarP==0

replace SuperRedRate_Parent = (Parent_Charge1_kVAyear/t1113c3)* 100*(abs(ParentActiveFlow))/((ParentActiveFlow^2+ParentReactiveFlow^2)^0.5) if t935c15==0&TempVarP~=0
replace SuperRedRate_Parent = (Parent_Charge1_kVAyear/t1113c3)* 100 if t935c15==0&TempVarP==0

*10.b - Grandparent

gen TempVarG=GrandparentActiveFlow^2+GrandparentReactiveFlow^2

*Two cases where t953c9~=0; distinguish between cases where TempVarG~=0 and when TempVarG==0

gen SuperRedRate_Grandparent = (Grandparent_Charge1_kVAyear/t1113c3)* 100*(abs(GrandparentActiveFlow)-GrandparentReactiveFlow*(CappedkVAr_kVA/t935c15))/((GrandparentActiveFlow^2+GrandparentReactiveFlow^2)^0.5) if t935c15~=0&TempVarG~=0
replace SuperRedRate_Grandparent= (Grandparent_Charge1_kVAyear/t1113c3)* 100 if t935c15~=0&TempVarG==0

*Two cases where t935c15==0; distinguish between cases where TempVarG~=0 and when TempVarG==0

replace SuperRedRate_Grandparent= (Grandparent_Charge1_kVAyear/t1113c3)* 100*(abs(GrandparentActiveFlow))/((GrandparentActiveFlow^2+GrandparentReactiveFlow^2)^0.5) if t935c15==0&TempVarG~=0
replace SuperRedRate_Grandparent = (Grandparent_Charge1_kVAyear/t1113c3)* 100 if t935c15==0&TempVarG==0

*10.c Replace missing values with 0 (missing values are generated when there are no parents or grandparents)

BlankToZero SuperRedRate_Parent SuperRedRate_Grandparent 

*10.d - Compute super-red charge

*Zero-out any negative values in the parent and in trhe grand-parent components of the super-red 

replace SuperRedRate_Parent = 0 if SuperRedRate_Parent < 0
replace SuperRedRate_Grandparent = 0 if SuperRedRate_Grandparent<0

gen SuperRedRate_FCP_Demand =SuperRedRate_Parent + SuperRedRate_Grandparent 

*10.f  - Adjust for cases with DSM agreements

*To check exceeded capacity calculation
gen SuperRedRateFCPExceeded=SuperRedRate_FCP_Demand
replace SuperRedRate_FCP_Demand  = SuperRedRate_FCP_Demand *((t935c2-t935c18)/t935c2) if t935c18 ~=0&t935c2~=0

*11 - Calculate  FCP capacity charge 

*11.a - General case

drop _merge
sort company Location
merge company Location using 911_v2a

keep if _merge==3

CheckZero t1113c1
gen FCPCapCharge1=100*t911c4/t1113c1

*11b Calculating additional charges to be added onto capacity charge with zero average kW/kVA

gen FCPCapCharge1AddParent=(100/t1113c1)*Parent_Charge1_kVAyear*(-ParentReactiveFlow*CappedkVAr_kVA)/((ParentActiveFlow^2+ParentReactiveFlow^2)^0.5) if TempVarP~=0&t935c15==0
gen FCPCapCharge1AddGrandparent=(100/t1113c1)*Grandparent_Charge1_kVAyear*(-GrandparentReactiveFlow*CappedkVAr_kVA)/((GrandparentActiveFlow^2+GrandparentReactiveFlow^2)^0.5) if TempVarG~=0&t935c15==0

replace FCPCapCharge1AddParent=0 if TempVarP==0
replace FCPCapCharge1AddGrandparent=0 if TempVarG==0

BlankToZero FCPCapCharge1AddParent FCPCapCharge1AddGrandparent

replace FCPCapCharge1AddParent = 0 if FCPCapCharge1AddParent<0
replace FCPCapCharge1AddGrandparent=0 if FCPCapCharge1AddGrandparent<0

replace FCPCapCharge1 = FCPCapCharge1 +  FCPCapCharge1AddParent +FCPCapCharge1AddGrandparent if t935c15==0

*11.c DSM adjustment for FCP capacity charge

gen FCPCapCharge1Exceeded=FCPCapCharge1
replace FCPCapCharge1 = FCPCapCharge1 *((t935c2-t935c18)/t935c2) if t935c18 ~=0 & t935c2~=0

*12 Calculating Charge 1 applied to generation

*Note: need to zero out cases where Parent and Grandparent charges are missing (because they don't have parents); had not been necessry until now

BlankToZero Parent_Charge1_kVAyear Grandparent_Charge1_kVAyear

*Calculation of ChargeableExportCap and of Maximum export capacity

gen ChargeableExportCap=t935c4+t935c5+t935c6
gen MaximumExportCap=t935c3+ChargeableExportCap

gen ShareChargeableExportCap=ChargeableExportCap/MaximumExportCap if MaximumExportCap~=0

*Dealing with option "lowerIntermittentCredit"

if "$OptionlowerIntermittentCredit"=="0" {
gen CreditGenFCPTariff=-100*(t935c21*t911c4+ Parent_Charge1_kVAyear+Grandparent_Charge1_kVAyear)*(ShareChargeableExportCap)/t1113c3 if MaximumExportCap~=0}
}

if "$OptionlowerIntermittentCredit"=="1" {
gen CreditGenFCPTariff=-100*t935c21*(t911c4+ Parent_Charge1_kVAyear+Grandparent_Charge1_kVAyear)*(ShareChargeableExportCap)/t1113c3 if MaximumExportCap~=0
}

replace CreditGenFCPTariff = 0 if MaximumExportCap==0

*Calculating generation credit in �, after rounding the credit to three-decimal places first

gen CreditGenFCP=(round(CreditGenFCPTariff,0.001)/100)*t935c19

*13 - Calculating revenues from FCP charges in �

gen TariffRevFCPSuperRed=(SuperRedRate_FCP_Demand/100)*t935c15*t935c2*t1113c3*(1-t935c23/t1113c3)
gen TariffRevFCPCapCharge1=(FCPCapCharge1/100)*t1113c1*t935c2*(1-t935c22/t1113c1)

BlankToZero TariffRevFCPSuperRed TariffRevFCPCapCharge1 CreditGenFCP
gen TariffRevFCPTot=TariffRevFCPSuperRed + TariffRevFCPCapCharge1
by company, sort: egen FCPRevenue=sum(TariffRevFCPTot)

*14. Calculate sum of generation super-red credits

by company, sort: egen AggFCPSuperRedGenCredit=sum(CreditGenFCP)

drop _merge
sort comp line
save Super_Red_Rate_FCP_Demand_Final, replace

*13. Keep dataset just with total FCP revenue and too super-red generation credit

sort comp
keep if comp~=comp[_n-1]
keep comp FCPRevenue AggFCPSuperRedGenCredit
sort comp
save FCPRevenue, replace

*Erase temporary data files
erase Combined_1.dta
erase Combined_2.dta
erase Combined_3.dta
erase Parent_details.dta
erase Grandparent_details.dta
erase Ancester_Location.dta
erase 911_v2a.dta 
*erase 935_v2a_FCP.dta

***************
*Data files kept:
*Super_Red_Rate_FCP_Demand_Final = Super red rate for FCP demand
*FCPRevenue = Company level revenue from FCP charges (super-red and FCP capacity charge 1)
***************
