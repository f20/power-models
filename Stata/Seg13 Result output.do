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
*Tariffs.dta

clear
set mem 500m
use Tariffs

*--------------------------
*1. Calculate annual charges by customer (�/year)

gen AnnualChargeCapDem=ImportCapacityCharge_r*(t1113c1-t935c22)*t935c2/100

gen AnnualChargeSuperRedDem=SuperRedDem_r*(t1113c3-t935c23)*t935c15*t935c2/100

gen AnnualChargeFixedDem=DemFixedCharge_r*(t1113c1-t935c22)/100

gen AnnualChargeFixedGen=GenFixedCharge_r*(t1113c1-t935c22)/100

BlankToZero t935c19
gen AnnualCreditSuperRedGen=SuperRedCreditGen_r*t935c19/100

gen AnnualCreditCapGen= (t1113c1/100)*ExportCapacityCharge_r*ChargeableExportCap*(1-t935c22/t1113c1)

*--------------------------

*2. Totals

BlankToZero AnnualChargeCapDem AnnualChargeSuperRedDem AnnualChargeFixedDem
gen AnnualChargeDemR=AnnualChargeCapDem+AnnualChargeSuperRedDem+AnnualChargeFixedDem

BlankToZero AnnualChargeFixedGen AnnualCreditSuperRedGen AnnualCreditCapGen
gen AnnualCreditGenR=AnnualChargeFixedGen+AnnualCreditSuperRedGen+AnnualCreditCapGen

BlankToZero AnnualChargeDemR
gen AnnualChargeDemPrevious = t935c24
gen ChangeAnnualChargeDem= AnnualChargeDemR - AnnualChargeDemPrevious
gen PerChangeAnnualChargeDem= (AnnualChargeDemR / AnnualChargeDemPrevious)-1

BlankToZero AnnualCreditGenR
gen AnnualCreditCapGenPrevious=t935c25
gen ChangeAnnualCreditGen= AnnualCreditGenR-AnnualCreditCapGenPrevious
gen PerChangeAnnualCreditGen= (AnnualCreditGenR / AnnualCreditCapGenPrevious) - 1

*--------------------------

*3.First. Detour to break done NRDOCImportCapCharge into NR and DOCCapacityCharge (in calculations the two were calculated together) (Look up file "8 Demand scaling.do")

gen DOCImportCapCharge=  (100/t1113c1)*(CCTotSSAssets)*(EDCMDOCChargingRate)

gen NRImportCapCharge=  (100/t1113c1)*(CCTotSSAssets)*(EDCMNetworkRatesChargingRate)

*3.Second.

gen AnnualChargeRemoteFCPLRIC=SuperRedPreAdjustment*(t1113c3-t935c23)*t935c15*t935c2/100

*3.Third

*a. Have already defined "AnnualChargeRemoteFCPLRIC" as recovery from SuperRed in �, before any adjustment for cases where import capacity<0

*b. Calculate annual recovery from import capacity charges based on contributions other than asset based adder, NR and DOC
gen ImportCapacityChargePre= ImportCapFixedAdder + AssetBasedChargingRateResidRev + EDCMImportCapINDOCCharge +NRDOCImportCapCharge + LRICLocalCharge1 + FCPCapCharge1 + TransmissionExitCapCharge

gen Block1=ImportCapFixedAdder + EDCMImportCapINDOCCharge + LRICLocalCharge1 + FCPCapCharge1 + TransmissionExitCapCharge
gen RevBlock1=Block1/100 * (t1113c1-t935c22)*t935c2

*c. Calculate recovery from asset based adder, NR and DOC
gen RevBlock2=(AssetBasedChargingRateResidRev + DOCImportCapCharge+NRImportCapCharge)/100 * (t1113c1-t935c22)*t935c2

*3. Fourth

*The condition "& RevBlock2!=0" was added on 18Dec2015.
*This was coupled with commenting out the line "*replace CCTotSSAssets = 1e-100 if CCTotSSAssets==0" in "Seg07 Shared assets MEAV.do"


gen Ratio=1
replace Ratio = -(RevBlock1+AnnualChargeRemoteFCPLRIC)/RevBlock2 if (RevBlock1+AnnualChargeRemoteFCPLRIC+RevBlock2)<0 & RevBlock2!=0

*3 Fifth
* Impose cap on Block3Recovery so that Block3Recovery+ImportStartRecovery+AnnualChargeRemoteFCPLRIC cannot be negative;

gen AssetBasedRateResidRevAdj =AssetBasedChargingRateResidRev * Ratio
gen DOCImportCapAdj=DOCImportCapCharge * Ratio
gen NRImportCapAdj=NRImportCapCharge * Ratio

*4. List of components

gen AnnualSoleUseAssetCharge= DemFixedCharge*(t1113c1-t935c22)/100

gen AnnualTransExitCharge=TransmissionExitCapCharge /100 * (t1113c1-t935c22)*t935c2

gen AnnualScalingFixedAdder=ImportCapFixedAdder /100 * (t1113c1-t935c22)*t935c2

gen AnnualDirectCostAllocation=DOCImportCapAdj /100  * (t1113c1-t935c22)*t935c2

gen AnnualNetworkRatesAllocation=NRImportCapAdj/100 * (t1113c1-t935c22)*t935c2

gen AnnualScalingAssetbased=AssetBasedRateResidRevAdj /100 * (t1113c1-t935c22)*t935c2

gen AnnualIndirectCostAllocation=EDCMImportCapINDOCCharge/100 * (t1113c1-t935c22)*t935c2

gen AnnualFCPLRICCharge=(LRICLocalCharge1 + FCPCapCharge1)/100 * (t1113c1-t935c22)*t935c2+AnnualChargeRemoteFCPLRIC

*5. Checking sum of components against AnnualCharge

BlankToZero AnnualFCPLRICCharge

gen Check=AnnualChargeDemR- (AnnualSoleUseAssetCharge+AnnualTransExitCharge+ AnnualDirectCostAllocation+ AnnualNetworkRatesAllocation +AnnualIndirectCostAllocation +AnnualFCPLRICCharge+ AnnualScalingFixedAdder +AnnualScalingAssetbased)

*6. Ordering variables as required in email sent by Franck 28Oct 10.32

order comp line   SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r AnnualChargeCapDem AnnualChargeSuperRedDem AnnualChargeFixedDem AnnualCreditCapGen AnnualChargeFixedGen AnnualCreditSuperRedGen AnnualChargeDemR AnnualChargeDemPrevious ChangeAnnualChargeDem PerChangeAnnualChargeDem  AnnualCreditGenR AnnualCreditCapGenPrevious ChangeAnnualCreditGen PerChangeAnnualCreditGen  AnnualSoleUseAssetCharge AnnualTransExitCharge AnnualDirectCostAllocation AnnualIndirectCostAllocation AnnualNetworkRatesAllocation AnnualFCPLRICCharge AnnualScalingFixedAdder AnnualScalingAssetbased Check

save Results, replace
