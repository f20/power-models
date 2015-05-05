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

global Home `c(pwd)'

capture program drop EDCMCombined
program EDCMCombined

quietly {

*Runs set of relevant programs that are called on within the model

run "$Home/Prog_BlankToZero.do"
run "$Home/Prog_CheckZero.do"
run "$Home/Prog_HashValueToMissing.do"
run "$Home/Prog_Mean_Charge1_Clusters.do"
run "$Home/Prog_NetworkAssetValueperKVA.do"
run "$Home/Prog_AuditResultsRandom"     

*Run the various files that actually make up the model
*Change directory so that output is saved in appropriate folder

do "$Home/Seg01 Options.do"

if $OptionProblem==1 {
            exit
            }
do "$Home/Seg02 Correct raw simulated data.do"
do "$Home/Seg03 Split_935.do"
do "$Home/Seg04 FCP Charge 1 Demand"
do "$Home/Seg05 Linkages.do"
do "$Home/Seg06 LRIC Charge 1 Demand.do"
do "$Home/Seg07 Shared assets MEAV.do"
do "$Home/Seg08 Transmission exit charges.do"
do "$Home/Seg09 Export capacity charges.do"
do "$Home/Seg10 EDCM demand revenue target.do"
do "$Home/Seg11 Demand scaling.do"
do "$Home/Seg12 Tariff summary.do"
do "$Home/Seg13 Result output.do"
do "$Home/Seg14 Checksums.do"

}
auditresultsrandom

end
