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
*Aim: Calculate site-specific assets (raw and cap-and-collared)
***************

***************
*Datasets used:
*935
*11
***************

***************
*Invoking programs that are run
*do Prog_Check
****************

clear
use 935

sort company
merge company using 11.dta

*1. Generating set of dummy variables to pick up category of tariffs

*Level 0

gen CapacityFlagLevel0 = 0
replace CapacityFlagLevel0 = 1 if t935c9==0000

gen DemandFlagLevel0=0
replace DemandFlagLevel0=1 if t935c9==0000

*Level 1

gen CapacityFlagLevel1 = 0
replace CapacityFlagLevel1 = 1 if t935c9==1000

gen DemandFlagLevel1=0
replace DemandFlagLevel1=1 if (t935c9==1100|t935c9==1110|t935c9==1001|t935c9==1101|t935c9==1111)

*Level 2

gen CapacityFlagLevel2 = 0
replace CapacityFlagLevel2 = 1 if (t935c9==1100|t935c9==0100)

gen DemandFlagLevel2=0
replace DemandFlagLevel2=1 if (t935c9==1110|t935c9==0110|t935c9==0111|t935c9==0101|t935c9==1101|t935c9==1111)

*Level 3

gen CapacityFlagLevel3 = 0
replace CapacityFlagLevel3 = 1 if (t935c9==1110|t935c9==110|t935c9==0010)

gen DemandFlagLevel3=0
replace DemandFlagLevel3=1 if (t935c9==0011|t935c9==0111|t935c9==1111)

*Level 4

gen CapacityFlagLevel4 = 0
replace CapacityFlagLevel4 = 1 if (t935c9==0002|t935c9==0011|t935c9==0111|t935c9==0101|t935c9==1101|t935c9==1111)

gen DemandFlagLevel4=0

*Level 5

gen CapacityFlagLevel5 = 0
replace CapacityFlagLevel5 = 1 if (t935c9==0001|t935c9==1001)

gen DemandFlagLevel5=0

*2. Calculating average network asset value per kVA in respect of each network level

*Deriving NetworkAssetValuepervKVA

*Runs program NetworkAssetValue
NetworkAssetValue

*3. Calculating total site-specific shared assets of each demand user,  TNA in statement notation

gen tempvar5=(t1113c1-t935c22)
CheckZero t1113c3 tempvar5

gen RawTotSSAssets=RawSSValCapAllL + (RawSSValDemAllL* (1-(t935c23/t1113c3))*(t1113c1/(t1113c1-t935c22)))
gen CCTotSSAssets=CCSSValCapAllL + (CCSSValDemAllL* (1-(t935c23/t1113c3))*(t1113c1/(t1113c1-t935c22)))

*Line below commented out on 18Dec2015. Deemed unnecessary given revision in "Seg13 Results output.do", also dated 18Dec2015.
*replace CCTotSSAssets = 1e-100 if CCTotSSAssets==0

*4. Producing aggregate across all EDCM demand users

CheckZero t1113c1
gen RawVarTemp=(RawTotSSAssets)*t935c2* (1-t935c22/t1113c1)
gen CCVarTemp=(CCTotSSAssets)*t935c2* (1-t935c22/t1113c1)

by company, sort: egen RawAggDemSS= sum (RawVarTemp)
by company, sort: egen CCAggDemSS= sum (CCVarTemp)

sort company line
save SharedAssetsMEAV, replace

***************
*Data files kept:
*SharedAssetsMEAV = includes all variables calculated within Annex 2 of the methodology statement
*Abbreviation in variable names:
*SS = Site specific;
*Raw = Using network use factors that have not been capped and collared (with exception of treatment of mixed import-export that are generation dominated)
*CC = Using  network use factors that have been capped and collared
***************
