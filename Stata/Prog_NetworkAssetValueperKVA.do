* Copyright licence and disclaimer
*
* Copyright 2012-2017 Reckon LLP, Pedro Fernandes and others. All rights reserved.
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


*NOTE
*14Dec2015 Revisions
*Script changed to deal with change in tables t1133 and t1134, namely fact that first cell that existed in older versions of that table was deleted, so that these tables now only have five columns.


capture program drop NetworkAssetValue
program NetworkAssetValue

quietly{

*Preliminary loop to extract correct loss adjustment factor to use in formula for AvNetAssetValueDemandLevel`i'  further below

    gen LossAdjustFactorConnection= .

    local i = 0
    while `i'<=5 {

    local j = `i'+ 1
    replace LossAdjustFactorConnection= t1135c`j' if  CapacityFlagLevel`i'==1

    local i = `i' + 1
    }

* Loop to calculate notional assets

    local i = 1
    while `i' <=5 {

*Calculating average network asset value for capacity and demand at different levels

        local j = `i'+1
*       Replacing missing values in input data with zeros
        replace t1122c`j' = 0 if t1122c`j'==.
        replace t1131c`j' = 0 if t1131c`j'==.
        replace t1135c`j' = 0 if t1135c`j'==.

        gen NetworkAssetRateperKVALevel`i' = (t1131c`j'/t1122c`j')/t1135c`j' if CapacityFlagLevel`i'==1|DemandFlagLevel`i'==1

* NOTE: It is important that missing values in t1132c1 are not replaced as 0 in earlier stages of program


        if `i'==5 {

            if "$Optionlegacy201"=="1" {
                replace NetworkAssetRateperKVALevel5 = t1132c1 if (CapacityFlagLevel5==1|DemandFlagLevel5==1) & t1132c1~=.
                }

            if "$Optionlegacy201"=="0" {
                 replace NetworkAssetRateperKVALevel5 = t1132c1 if (CapacityFlagLevel5==1|DemandFlagLevel5==1) & t1132c1~=.&t1132c1~=0
                }
                 }

        gen TempVar`i' = (1+t1105c`j') if CapacityFlagLevel`i'==1
        replace TempVar`i' = 99 if CapacityFlagLevel`i'~=1
        CheckZero TempVar`i'
        drop TempVar`i'

*Hard-coded value of power factor in the 500MW model. What was an input data item is now set to 0.95.

        gen ActivePowerEquivalentLevel`i' = 0.95*t1135c`j' if CapacityFlagLevel`i'==1

        gen AvNetAssetValueCapacityLevel`i' = NetworkAssetRateperKVALevel`i' * ActivePowerEquivalentLevel`i' /(1+t1105c`j') if CapacityFlagLevel`i'==1

        gen AvNetAssetValueDemandLevel`i' = NetworkAssetRateperKVALevel`i'*t935c15*LossAdjustFactorConnection  if DemandFlagLevel`i'==1

*Calculating site-specific asset value for capacity and demand at different levels

*Defining AdjNetworkFactor`i' = network use factor
*CHECK: no adjustment to NUF for mixed import-export is needed

        local k = `i'+9

        gen AdjNetworkUseFactor`i'= t935c`k' if (CapacityFlagLevel`i'==1|DemandFlagLevel`i'==1)

		*        replace AdjNetworkUseFactor`i'= t1134c`j' if (CapacityFlagLevel`i'==1|DemandFlagLevel`i'==1) & t953c25=="TRUE"

*Defining site-specific assets for demand users, (a)using raw network use-factors and (b) using capped-and-collared use factors

*(a) Using raw network-use factors after making mixed import-export adjustment

        gen RawSSValCapL`i' = AdjNetworkUseFactor`i' * AvNetAssetValueCapacityLevel`i' if CapacityFlagLevel`i'==1
        gen RawSSValDemL`i' = AdjNetworkUseFactor`i' * AvNetAssetValueDemandLevel`i' if DemandFlagLevel`i'==1

*(b) Using cap and collared network-use factors

*Three lines below revise 21Jul2017

        gen CCNetworkUseFactor`i'= AdjNetworkUseFactor`i'

        replace CCNetworkUseFactor`i'= t1134c`j' if (CapacityFlagLevel`i'==1|DemandFlagLevel`i'==1) & AdjNetworkUseFactor`i'<t1134c`j'

        replace CCNetworkUseFactor`i'= t1133c`j' if (CapacityFlagLevel`i'==1|DemandFlagLevel`i'==1) & AdjNetworkUseFactor`i'>t1133c`j'

        gen CCSSValCapL`i' = CCNetworkUseFactor`i' * AvNetAssetValueCapacityLevel`i' if CapacityFlagLevel`i'==1
        gen CCSSValDemL`i' = CCNetworkUseFactor`i' * AvNetAssetValueDemandLevel`i' if DemandFlagLevel`i'==1

    local i = `i' + 1
    }

*Replacing missing values with 0 so that they can be aggregated

    local i = 1

    while `i'<=5 {
        replace RawSSValCapL`i' =0 if RawSSValCapL`i'==.
        replace RawSSValDemL`i' =0 if RawSSValDemL`i' ==.
        replace CCSSValCapL`i' = 0 if CCSSValCapL`i'==.
        replace CCSSValDemL`i'  = 0 if CCSSValDemL`i'==.

    local i = `i'+1
    }

*Aggregating values across all levels to produce total site-specific shared assets for each tariff

    local i =1

        gen RawSSValCapAllL = 0
        gen RawSSValDemAllL = 0

        gen CCSSValCapAllL = 0
        gen CCSSValDemAllL = 0

    while `i'<=5 {

        replace RawSSValCapAllL =  RawSSValCapAllL + RawSSValCapL`i'
        replace RawSSValDemAllL =  RawSSValDemAllL + RawSSValDemL`i'

        replace CCSSValCapAllL = CCSSValCapAllL + CCSSValCapL`i'
        replace CCSSValDemAllL = CCSSValDemAllL + CCSSValDemL`i'

    local i = `i' + 1
    }

}
end
