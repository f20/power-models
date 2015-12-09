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
*Aim: Check last two columns of Table 4501 where Optionchecksum is not 0
***************

***************
*Datasets used:
*Results.dta

clear
set mem 500m
use Results
*--------------------------

BlankToZero SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r

if "$Optionchecksums"~="0" {

*1. Calculate checksum5

    gen Checksum5=mod(49222*(round(ExceededExpCapCharge_r*100)+mod(49222*(round(ExportCapacityCharge_r*100)+mod(49222*(round(GenFixedCharge_r *100)+mod(49222*(round(SuperRedCreditGen_r *1000)+mod(49222*(round(ExceededDemCapCharge_r *100)+mod(49222*(round(ImportCapacityCharge_r *100)+mod(49222*(round(DemFixedCharge_r *100)+mod(49222*(round(SuperRedDem_r*1000)),99991)),99991)),99991)),99991)),99991)),99991)),99991)),99991)

*2. Calculate checksum7

    gsort company -line
    gen Checksum7=mod(4922236*(round(ExceededExpCapCharge_r*100)+mod(4922236*(round(ExportCapacityCharge_r*100)+mod(4922236*(round(GenFixedCharge_r*100)+mod(4922236*(round(SuperRedCreditGen_r*1000)+mod(4922236*(round(ExceededDemCapCharge_r *100)+mod(4922236*(round(ImportCapacityCharge_r *100)+mod(4922236*(round(DemFixedCharge_r *100)+mod(4922236*(round(SuperRedDem_r*1000)),9999991)),9999991)),9999991)),9999991)),9999991)),9999991)),9999991)),9999991) if company~=company[_n-1]
    replace Checksum7=mod(4922236*(round(ExceededExpCapCharge_r*100)+mod(4922236*(round(ExportCapacityCharge_r*100)+mod(4922236*(round(GenFixedCharge_r*100)+mod(4922236*(round(SuperRedCreditGen_r*1000)+mod(4922236*(round(ExceededDemCapCharge_r *100)+mod(4922236*(round(ImportCapacityCharge_r *100)+mod(4922236*(round(DemFixedCharge_r *100)+mod(4922236*(round(SuperRedDem_r*1000)+Checksum7[_n-1]),9999991)),9999991)),9999991)),9999991)),9999991)),9999991)),9999991)),9999991) if company==company[_n-1]

    sort company line
    order comp line  Checksum5 Checksum7  SuperRedDem_r DemFixedCharge_r ImportCapacityCharge_r ExceededDemCapCharge_r SuperRedCreditGen_r GenFixedCharge_r ExportCapacityCharge_r ExceededExpCapCharge_r AnnualChargeCapDem AnnualChargeSuperRedDem AnnualChargeFixedDem AnnualCreditCapGen AnnualChargeFixedGen AnnualCreditSuperRedGen AnnualChargeDemR AnnualChargeDemPrevious ChangeAnnualChargeDem PerChangeAnnualChargeDem  AnnualCreditGenR AnnualCreditCapGenPrevious ChangeAnnualCreditGen PerChangeAnnualCreditGen  AnnualSoleUseAssetCharge AnnualTransExitCharge AnnualDirectCostAllocation AnnualIndirectCostAllocation AnnualNetworkRatesAllocation AnnualFCPLRICCharge AnnualScalingFixedAdder AnnualScalingAssetbased Check

    }
save ResultsToAudit, replace
