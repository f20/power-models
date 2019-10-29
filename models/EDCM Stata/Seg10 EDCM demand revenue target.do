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
*Aim: Calculate EDCM demand revenue target (Annex 3)
***************

***************
*Datasets used:
*SharedAssetsMEAV (from "6. Shared assets MEAV")
***********

***************
*Input variables used:

*t1112c1 = Allowed revenues excluding transmission exit (3/year)
*t1113c6 = DNO expenditure on direct operating costs
*t1113c7 = DNO expenditure on indirect operating costs
*t1113c8 = DNO expenditure on network rates
*t1131cX = Assets in CDCM model (�) (from CDCM table 2705) - at different network levels

*t935c7 = Sole use asset MEAV (�)
*t953c14 = Proportion of sole use asset MEAV not chargeable to this tariff
***************

***************
*Programs invoked:
*BlankToZero
*CheckZero
***************

set type double
clear
use SharedAssetsMEAV
drop _merge

*1. Calculate network rates contribution rate

*Calculating sole use assets: adjust for part-time customers and exclude sole use assets associated with exempt export capacity
BlankToZero t935c2 t935c3 t935c4 t935c5 t935c6

gen TotExportImportMaxCap= t935c2+t935c3+t935c4+t935c5+t935c6

CheckZero TotExportImportMaxCap

gen SoleUseAssetsImport=t935c7*(t935c2/TotExportImportMaxCap)
gen SoleUseAssetsExport=t935c7 - SoleUseAssetsImport

gen TempVar7 = (t935c3+t935c4+t935c5+t935c6)

*Deal with case where sum of export capacity is zero.
gen SoleUseAssetsExportNotExempt=(SoleUseAssetsExport*(t935c4+t935c5+t935c6)/TempVar7) if TempVar7~=0
replace SoleUseAssetsExportNotExempt=0 if TempVar7==0
drop TempVar7

CheckZero t1113c1
gen SoleUseImportPartTime=SoleUseAssetsImport *(1-t935c22/t1113c1)

gen SoleUseExportNotExemptPartTime=SoleUseAssetsExportNotExempt*(1-t935c22/t1113c1)

gen AdjSoleUseAssetParttime=SoleUseImportPartTime+SoleUseExportNotExemptPartTime

by company, sort: egen TotalSoleUseAsset= sum(AdjSoleUseAssetParttime)

gen EHVAssets=t1131c2+t1131c3+t1131c4+t1131c5+t1131c6
gen HVandLVAssets=t1131c7+t1131c8+t1131c9
gen HVandLVService=t1131c10+t1131c11

gen TempVar7=(RawAggDemSS+TotalSoleUseAsset+EHVAssets+HVandLVAssets+HVandLVService)
CheckZero TempVar7
drop TempVar7

gen NetworkRatesContributionRate=t1113c8/(RawAggDemSS+TotalSoleUseAsset+EHVAssets+HVandLVAssets+HVandLVService)

*2. Calculate direct operating costs contribution rate

gen DirectOpCostsContributionRate = t1113c6/(RawAggDemSS + TotalSoleUseAsset+EHVAssets+(HVandLVAssets+HVandLVService)/0.68)

*3. Calculate indirect costs contribution rate

gen IndirectOpCostsContributionRate = t1113c7/(RawAggDemSS + TotalSoleUseAsset+EHVAssets+(HVandLVAssets+HVandLVService)/0.68)

*4 - Calculating fixed charge on sole use assets in p/day  and associated revenue

gen DemFixedChargeSolePenceDay=(100/t1113c1)*SoleUseAssetsImport*( NetworkRatesContributionRate+ DirectOpCostsContributionRate)

if "$Optiondcp189"=="proportionsplitShortfall" {
    rename DemFixedChargeSolePenceDay DemFixedChargeSolePDayNodcp189
    gen DemFixedChargeSolePenceDay=(100/t1113c1)*SoleUseAssetsImport*( NetworkRatesContributionRate+ (1-t935dcp189)*DirectOpCostsContributionRate)
    }

gen GenFixedChargeSolePenceDay=(100/t1113c1)*SoleUseAssetsExportNotExempt*( NetworkRatesContributionRate+ DirectOpCostsContributionRate)

*Calculate annual revenue from fixed charges in �,
*For generation calculate after rounding the charge to 2-decimal places first

gen DemFixedChargePoundAnnual=(t1113c1-t935c22)*DemFixedChargeSolePenceDay/100
by company, sort: egen DemFixedChargeRecovery=sum(DemFixedChargePoundAnnual)

if "$Optiondcp189"=="proportionsplitShortfall" {
    gen DemFixedChargePdAnnualNodcp189=(t1113c1-t935c22)*DemFixedChargeSolePDayNodcp189/100
    by company, sort: egen DemFixedChargeRecoveryNodcp189=sum(DemFixedChargePdAnnualNodcp189)
    }


gen GenFixedChargePoundAnnual=(t1113c1-t935c22)*round(GenFixedChargeSolePenceDay, 0.01)/100
by company, sort: egen GenFixedChargeRecovery=sum(GenFixedChargePoundAnnual)

*5. Calculate GCN, the total forecast revenue in �/year from application of EDCM export tariffs, including the EDCM generation fixed charge

*Bring in data on revenue from export capacity charges and data on super-red generation credits

sort comp line
merge company line using ExportCapacityCharges
drop _merge

sort company
merge company using FCPRevenue
drop _merge

sort company
merge company using LRICRevenue
drop _merge

BlankToZero AggFCPSuperRedGenCredit AggLRICSuperRedGenCredit

gen GCN = ExportCapChargeRecovery+GenFixedChargeRecovery+(AggFCPSuperRedGenCredit+AggLRICSuperRedGenCredit)

*6. Residual revenue contribution rate

gen TempVar7=(RawAggDemSS + EHVAssets + HVandLVAssets )
CheckZero TempVar7
drop TempVar7

gen ResidualRevContributionRate=(t1113c4 -  t1113c6 - t1113c7 - t1113c8 - GCN)/(RawAggDemSS + EHVAssets + HVandLVAssets )

*7. Calculating import capacity based contribution

gen ImportCapNetRatesContr=RawTotSSAssets*t935c2 * (1-t935c22/t1113c1)* NetworkRatesContributionRate

gen ImportCapDirectContr=RawTotSSAssets*t935c2 *  (1-t935c22/t1113c1)*DirectOpCostsContributionRate

gen ImportCapIndirectContr=RawTotSSAssets*t935c2 * (1-t935c22/t1113c1)* IndirectOpCostsContributionRate

gen ImportCapResidualRevContr=RawTotSSAssets*t935c2 * (1-t935c22/t1113c1)* ResidualRevContributionRate

*8. Calculating demand sole use asset MEAV based contribution

gen DemSoleUseAssetNetworkRatesContr = SoleUseImportPartTime*NetworkRatesContributionRate

gen DemSoleUseAssetDirectContr = SoleUseImportPartTime* DirectOpCostsContributionRate

gen DemSoleUseAssetIndirectContr = SoleUseImportPartTime*IndirectOpCostsContributionRate


*9. Calculating aggregate EDCM demand revenue target

BlankToZero ImportCapNetRatesContr ImportCapDirectContr ImportCapIndirectContr ImportCapResidualRevContr DemSoleUseAssetNetworkRatesContr DemSoleUseAssetDirectContr DemSoleUseAssetIndirectContr

gen TempVar=(ImportCapNetRatesContr+ImportCapDirectContr+ImportCapIndirectContr+ImportCapResidualRevContr)+(DemSoleUseAssetNetworkRatesContr+DemSoleUseAssetDirectContr+DemSoleUseAssetIndirectContr)

by company, sort: egen AggDemandRevenueTarget=  sum(TempVar)

*10. Calculating fixed charge reduction assoicated with DCP189

if "$Optiondcp189"=="proportionsplitShortfall" {

    gen OMR = DemFixedChargeRecoveryNodcp189 - DemFixedChargeRecovery
    gen FCR = OMR*(EHVAssets + HVandLVAssets)/(RawAggDemSS + EHVAssets + HVandLVAssets)
    replace AggDemandRevenueTarget=AggDemandRevenueTarget - FCR

    }

drop TempVar

sort comp line
save AggDemandRevenueTarget, replace

***************
*Data files kept:
*AggDemandRevenueTarget
***************
