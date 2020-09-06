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
*Aim: Calculate transmission exit charges (demand)
***************

***************
*Datasets used:
*SharedAssetsMEAV.dta
*935
*11
***********

*Preparing dataset to contain calculation on transmission loss factors
clear
use SharedAssetsMEAV.dta
keep company line LossAdjustFactorConnection
sort company line
save Temp5, replace

*Input data from 935
clear
use 935
sort company

*Merge with data from Table 11
merge company using 11.dta
drop _merge

*Merge with data containing LossAdjustFactorConnection
sort company line
merge company line using Temp5.dta

*1. Calculate total EDCM peak time consumption across all EDCM demand

CheckZero t1113c3

BlankToZero t935c23
gen EDCMPeakTimeConsumption = t935c2*(t935c15*(1-t935c23/t1113c3))*LossAdjustFactorConnection

by company, sort: egen AggEDCMPeakTimeConsumption=sum(EDCMPeakTimeConsumption)

gen TempVar6=(t1122c1+AggEDCMPeakTimeConsumption)
CheckZero TempVar6
drop TempVar6

*2. Calculate transmission exit charging rate

gen TransmissionExitChargingRate = 100/t1113c1*t1113c5/(t1122c1+AggEDCMPeakTimeConsumption)

*3. Converting charging rate from p/kW/day into a p/kVA/day import capacity based charge.  Need to make part-time adjustment

gen TempVar6=(1-t935c22/t1113c1)
CheckZero t1113c1 t1113c3 TempVar6
drop TempVar6

gen TransmissionExitCapCharge=TransmissionExitChargingRate*t935c15*LossAdjustFactorConnection *(1-t935c23/t1113c3)/(1-t935c22/t1113c1)

*4. Calculate transmission exit credit in p/kVA/day

gen ChargeableExportCap=t935c4+t935c5+t935c6

gen TransmissionExitCreditRate= -TransmissionExitChargingRate *t935c20/ChargeableExportCap if ChargeableExportCap~=0
replace TransmissionExitCreditRate = 0 if  ChargeableExportCap==0

*5. Calculate expected revenue from capacity based exit charges, and transmission exit credit - in pounds per year

gen TransmissionExitChargeRev=(TransmissionExitCapCharge/100)*t1113c1*t935c2*(1-t935c22/t1113c1)
gen TransmissionExitCreditValue=(round(TransmissionExitCreditRate,0.01)/100)*t1113c1*ChargeableExportCap*(1-t935c22/t1113c1)

by company, sort: egen AggTransExitChargeRevenue=sum(TransmissionExitChargeRev)
by company, sort: egen AggTransmissionExitCreditValue=sum(TransmissionExitCreditValue)

keep  company line TransmissionExitChargingRate TransmissionExitCapCharge TransmissionExitChargeRev AggTransExitChargeRevenue TransmissionExitCreditRate TransmissionExitCreditValue AggTransmissionExitCreditValue

sort company line
save TransmissionExitChargingRate, replace

*Deleting temporary data files
erase Temp5.dta

***************
*Data files kept:
*TransmissionExitChargingRate = transmission exit charging rate
***************
