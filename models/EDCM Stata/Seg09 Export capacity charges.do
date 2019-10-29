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
*Aim: Calculate Export capacity charges for generation
***************

clear
use 935
sort company

*Merge with data from Table 11

merge company using 11.dta
drop _merge

*Merge with TransmissionExitChargingRate
sort company line
merge company line using TransmissionExitChargingRate
drop _merge

*t1113c9 Average adjusted GP (�/year)
*t1113c10 GL term from the DG incentive revenue calculation (�/year)
*t1113c11 Total CDCM generation capacity 2005-2010 (kVA)
*t1113c12 Total CDCM generation capacity Post-2010 (kVA)

*1.  Calculate EDCM DG revenue target

by comp, sort: egen TotPre2005GenCap=sum( t935c4*(1-t935c22/t1113c1))
by comp, sort: egen Tot2005_2010GenCap=sum( t935c5*(1-t935c22/t1113c1))
by comp, sort: egen TotPost2010GenCap=sum( t935c6*(1-t935c22/t1113c1))

gen DGRevTarget = t1113c10*Tot2005_2010GenCap/(Tot2005_2010GenCap+t1113c11)+t1113c9*TotPost2010GenCap/(TotPost2010GenCap+t1113c12 )+ 0.2*(TotPre2005GenCap+TotPost2010GenCap)

*2. Calculate export capacity charge: remember necessary to net out the transmission connection (exit) credits

gen ChargeableExportCap=t935c4+t935c5+t935c6
BlankToZero ChargeableExportCap

by comp, sort: egen TotAdjChargeableExportCap=sum( ChargeableExportCap*(1-t935c22/t1113c1))

gen ExportCapCharge=((100/t1113c1)*DGRevTarget /TotAdjChargeableExportCap)+TransmissionExitCreditRate

gen ExceededExpCapCharge=((100/t1113c1)*DGRevTarget /TotAdjChargeableExportCap)

BlankToZero ExportCapCharge ExceededExpCapCharge

*3. Calculate transmission exit recovery in �/year, after rounding charges to 2-decimal places

by comp, sort: egen ExportCapChargeRecovery=sum((t1113c1/100)*(round(ExportCapCharge,0.01)*ChargeableExportCap)*(1-t935c22/t1113c1))

keep comp line ExportCapCharge ExceededExpCapCharge ExportCapChargeRecovery

sort comp line
save ExportCapacityCharges, replace
