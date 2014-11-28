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
*Aim: Going through annex 6 of Schedules 17 and 18
***************

***************
*Datasets used:
*AggDemandRevenueTarget (from "7. EDCM demand revenue target")
***********

***************
*Programs invoked:
*BlankToZero
*CheckZero
***************

clear 
use AggDemandRevenueTarget

*Merge with data containing FCP and LRIC revenue

sort company 
merge company using FCPRevenue
drop _merge

sort company
merge company using LRICRevenue
drop _merge

BlankToZero FCPRevenue LRICRevenue

*Steps 1 through to 2.25 are used for later stage in file "9 Tariff summary"

*1. Calculate charging rates for network rates 
by company, sort: egen AggNRContribution=sum(ImportCapNetRatesContr)
gen EDCMNetworkRatesChargingRate=AggNRContribution/(CCAggDemSS)

*2. Calculate charging rates for direct operating costs 

by company, sort: egen AggDOCContribution=sum(ImportCapDirectContr) 
gen EDCMDOCChargingRate=AggDOCContribution/(CCAggDemSS) 

*2.25. Converting charging rates for network rates and direct operating costs into p/kVA/day import capacity based charges (para 338)

gen NRDOCImportCapCharge= (100/t1113c1)*(CCTotSSAssets)*(EDCMDOCChargingRate+EDCMNetworkRatesChargingRate)

*3. Calculate charging rates for indirect operating costs

gen CoincidenceFactor8=t935c15*(1-(t935c23/t1113c3))*(t1113c1/(t1113c1-t935c22))

gen LDNOfactor=1
replace LDNOfactor=0.5 if t935c17==0.5

gen VolumeScaling=(0.5+CoincidenceFactor8)*t935c2*(1-t935c22/t1113c1)
by company, sort: egen AggVolumeScaling=sum(VolumeScaling)

gen VolumeScalingwLDNO=(0.5+CoincidenceFactor8)*t935c2*(1-t935c22/t1113c1)*LDNOfactor
by company, sort: egen AggVolumeScalingwLDNO=sum(VolumeScalingwLDNO)

*Only include SoleUseAssetIndirectContr from demand, ie DemSoleUseAssetIndirectContr

gen INDOCContribution = ImportCapIndirectContr+DemSoleUseAssetIndirectContr
by company, sort: egen AggINDOCContribution=sum(INDOCContribution)

gen EDCMINDOCChargingRate=(100/t1113c1)*(AggINDOCContribution/AggVolumeScalingwLDNO)

*4. Calculate import capacity INDOC based charge for each demand user

gen EDCMImportCapINDOCCharge= EDCMINDOCChargingRate*(0.5+CoincidenceFactor8)*LDNOfactor

if "$Optiondcp185"=="2" {
    replace EDCMINDOCChargingRate=(100/t1113c1)*(AggINDOCContribution/AggVolumeScaling) if (AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)<0
    replace EDCMImportCapINDOCCharge= EDCMINDOCChargingRate*(0.5+CoincidenceFactor8) if (AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)<0
    }

*5. Residual revenue charging rate

gen ResidualRevenueChargingRate=0.8*(AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)/CCAggDemSS

*6. Calculate asset based charging rate for residual revenue

gen AssetBasedChargingRateResidRev = (100/t1113c1)*CCTotSSAssets*ResidualRevenueChargingRate

*7. Fixed adder in p/kVA/day
*8. Converting fixed adder into an importcapacity based charge for each demand user

gen FixedAdder=(100/t1113c1)*0.2*(AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)/AggVolumeScaling
gen ImportCapFixedAdder=FixedAdder*(0.5+CoincidenceFactor8)

if "$Optiondcp185"=="1"|"$Optiondcp185"=="2" {

                replace FixedAdder=(100/t1113c1)*0.2*(AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)/AggVolumeScalingwLDNO if (AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)>0
                replace ImportCapFixedAdder=FixedAdder*(0.5+CoincidenceFactor8)*LDNOfactor if (AggDemandRevenueTarget - AggNRContribution-AggDOCContribution- AggINDOCContribution - DemFixedChargeRecovery- FCPRevenue-LRICRevenue)>0 

                }

sort company line
save DemandScaling, replace
