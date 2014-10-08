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

global SimulationHome `c(pwd)'

capture program drop simulatedata
program simulatedata

*Generates random data to use in FCP and LRIC models

quietly {

local Vers =  subinstr(subinstr(c(current_date)+ "-" + c(current_time)," ","",.),":","",.)

mkdir Rand`Vers'

run "$SimulationHome/Prog_Jumble.do"

local LRICName "Littleneck Lobster Lamprey Ling Limpet "
local FCPName "Faial Flores Fiji Fogo Faaite"
local CompName = "`LRICName'"+ "`FCPName'"

local l: word count `LRICName'
local f: word count `FCPName'

*Defining variables and their parameters

local DatasetGroup "11 911 913 935"

local Variables11 " t1105c1 t1105c2 t1105c3 t1105c4 t1105c5 t1105c6 t1113c1 t1113c2 t1113c3 t1113c4 t1113c5 t1113c6 t1113c7 t1113c8 t1113c9 t1113c10 t1113c11 t1113c12 t1122c1 t1122c2 t1122c3 t1122c4 t1122c5 t1122c6 t1131c1 t1131c2 t1131c3 t1131c4 t1131c5 t1131c6 t1131c7 t1131c8 t1131c9 t1131c10 t1131c11 t1132c1 t1133c1 t1133c2 t1133c3 t1133c4 t1133c5 t1133c6 t1134c1 t1134c2 t1134c3 t1134c4 t1134c5 t1134c6 t1135c1 t1135c2 t1135c3 t1135c4 t1135c5 t1135c6"
matrix input t11min = (0.01, 0.01, 0.01, 0.05, 0.05, 0.028195, 365, 0.2, 258, 338000000, 15500000, 23000000, 94600000, 14700000, -10000, 283816.1, 0, 0, 2444632, 2432161, 2413469, 2388167, 2368141, 0, 0, 724000000, 60700000, 627000000, 110000000, 0, 924000000, 314000000, 158000000, 610000000, 9671516, 0, 0, 2.246, 1.558, 3.29, 2.38, 2.768, 0, 0.273, 0.677, 0.332, 0.631, 0.697, 1, 1.002, 1.009, 1.013, 1.027777, 1.027777)
matrix input t11max = (0.025, 0.15, 0.15, 0.75, 0.75, 0.1275, 365, 0.2, 258, 485000000, 34100000, 53400000, 142000000, 37700000, 15060.65, 750000, 169931.5, 119256.8, 6371715, 6358997, 5414669, 5393288, 5116179, 832866.3, 0, 1600000000, 246000000, 1340000000, 222000000, 57600000, 1710000000, 866000000, 1470000000, 2150000000, 34400000, 0, 0, 2.246, 1.558, 3.29, 2.38, 2.768, 0, 0.273, 0.677, 0.332, 0.631, 0.697, 1, 1.005128, 1.012912, 1.019158, 1.062, 1.062)
matrix input t11modval = (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
matrix input t11modshare = (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

local Variables911 " t911c4 t911c6 t911c7 t911c8 t911c9"
matrix input t911min = (0, -2500000, -1500000, 0, -1500000)
matrix input t911max = (90, 0, 1500000, 2500000, 1500000)
matrix input t911modval = (0, 0, 0, 0, 0)
matrix input t911modshare = (0.5, 0.1, 0.1, 0.1, 0.25)

local Variables913 " t913c4 t913c5 t913c8 t913c9"
matrix input t913min = (-10, -10, -1000000, -1000000)
matrix input t913max = (100, 100, 1000000, 1000000)
matrix input t913modval = (0, 0, 0, 0)
matrix input t913modshare = (0.05, 0.05, 0.05, 0.05)

local Variables935 " t935c2 t935c3 t935c4 t935c5 t935c6 t935c7 t935c10 t935c11 t935c12 t935c13 t935c14 t935c15 t935c16 t935c17 t935c19 t935c21 t935c22 t935c24 t935c25"
matrix input t935min = (100, 100, 100, 100, 100, 0, 0, 0, 0, 0, 0, 0, -1, 1, 0, 0, 0, -10, -10)
matrix input t935max = (200000, 200000, 200000, 200000, 200000, 100000000, 20, 20, 20, 20, 20, 1.15, 1, 1, 10000000, 1, 365, 10, 10)
matrix input t935modval = (., ., ., ., ., 0, 0, 0, 0, 0, 0, 0, 0, 0.5, 0, 0, 0, 0, 0)
matrix input t935modshare = (0.4, 0.6, 0.6, 0.6, 0.6, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.3, 0, 0.2, 0.5, 0.8, 0.8, 0, 0)

local DatasetGroup "11 911 913 935"

local table = 1
while `table' <=4  {

local TableName: word `table' of `DatasetGroup'
local Count: word count `Variables`TableName''

clear
if `table'<=3 {
            insheet using "$SimulationHome/Shell`TableName'.csv", c n
          }

*Table 11

if `table'==1 {

    gen t1100c1=company
    gen t1100c2="Empty"
    gen t1100c3=""

    local n=1
    while `n'<=`l'+`f' {
        replace t1100c3="Rand`Vers'" in `n'
        local n=`n'+1
        }
    }

*Table 911

if `table'==2 {

    jumble t911c1

    gen t911c3=""
    gen line=_n
    gen root=1 in 1

    local counter=0
    while `counter'<=_N {

        local cluster=int(3*runiform())
        replace t911c3=t911c1[_n+1] if line>`counter'&line<=`counter'+`cluster'
        local counter=`counter'+`cluster'+1
        replace root=1 if line==`counter'+1
        }

    preserve
    keep if root==1
    keep t911c1
    save Rand`Vers'/t911root.dta, replace
    restore

    cross using "$SimulationHome/companies.dta"
    keep if regexm(company,"FCP")==1

    }

*Table 913

if `table'==3 {

    jumble t913c1

    gen t913c3=""
    gen line=_n

    gen root= 1 in 1

    local counter=0
    while `counter'<=_N {

        local cluster=int(1+7*runiform())
        replace t913c3=t913c1[_n+1] if line>`counter'&line<=`counter'+`cluster'
        local counter=`counter'+`cluster'+1
        replace root=1 if line==`counter'+1
        }

    preserve
    keep if root==1
    keep t913c1
    save Rand`Vers'/t913root.dta, replace
    restore

    drop root
    cross using "$SimulationHome/companies.dta"
    keep if regexm(company,"LRIC")==1
    }

*Table 935

if `table'==4 {

    foreach x of numlist 911 913 {

        clear
        insheet using "$SimulationHome/Shell`TableName'.csv", c n
        gen line=_n

        cross using Rand`Vers'/t`x'root.dta
        ren t`x'c1 t935c8
        sample 1, count by(line)

        cross using "$SimulationHome/companies.dta"

        if `x'==911 {
                keep if regexm(company,"FCP")==1
                save Rand`Vers'/Shell935FCP, replace
                }

        else {
                keep if regexm(company,"LRIC")==1
                save Rand`Vers'/Shell935LRIC, replace
                }
        }        
    append using Rand`Vers'/Shell935FCP

    matrix input t935c9cat =(0, 110, 1000, 1001, 1100, 1101, 1110, 1111)

    gen t935c9=t935c9cat[1,int(1+7*runiform())]

    }

*Simulate data values

    local j=1
    while `j'<=`Count' {
        local Var: word `j' of `Variables`TableName''
        replace `Var' = cond(runiform()<=t`TableName'modshare[1,`j'], t`TableName'modval[1,`j'], t`TableName'min[1,`j']+(t`TableName'max[1,`j']-t`TableName'min[1,`j'])*(1-rbeta(4,1)))

        local j=`j'+1
        }

*Dealing with some correlations and other special cases

if `table'==1 {

*Bug log: lines 6 and 7. 
*Setting t1132c1 to missing value. Ensures 'modern' model does same as 'legacy'. Alternatively, could have set t1132c1 to take values greater than 0 t

        replace t1132c1=.
        }

if `table'==2 {
        replace t911c7=cond(runiform()<.8,0,t911c7) if t911c6==0
        }

if `table'==3 {
        replace t913c9=cond(runiform()<.8,0,t913c9) if t913c8==0
        }

if `table'==4 {
        replace t935c15=cond(runiform()<.025, 10, t935c15)
        replace t935c16=cond(runiform()<.025, 10*sign(t935c16), t935c16)
        replace t935c18=cond(runiform()<0.75,0,t935c2*runiform())
*Special: making missing values of t935c18 as 0, to make it compatible with Excel requirements
        replace t935c18=0 if t935c18==.

        gen tempt935c2=cond(t935c2==.,0,t935c2)
        gen tempt935c3=cond(t935c3==.,0,t935c3)
        gen tempt935c4=cond(t935c4==.,0,t935c4)
        gen tempt935c5=cond(t935c5==.,0,t935c5)
        gen tempt935c6=cond(t935c6==.,0,t935c6)
        replace t935c20=cond(runiform()<0.75,0,(tempt935c3+tempt935c4+tempt935c5+tempt935c6)*runiform())

*Bug log: lines 4 and 5. For legacy models
*Set sole use asset values to 0 if import plus export capacities add up to 0

        replace t935c7=0 if tempt935c2+tempt935c3+tempt935c4+tempt935c5+tempt935c6==0

        drop tempt935c2 tempt935c3 tempt935c4 tempt935c5 tempt935c6

}

sort company

if `table'~=1 {
        sort company line
        order company line
        }

save Rand`Vers'/`TableName'.dta, replace
local table=`table'+1
}

*Fixing t935c23: Hours in super-red for which not a customer

clear
use Rand`Vers'/11.dta
keep company t1113c3
merge company using Rand`Vers'/935.dta
drop _merge

replace t935c23=cond(uniform()<0.75,0,t1113c3*uniform())

drop t1113c3
save Rand`Vers'/935.dta, replace

erase Rand`Vers'/t911root.dta
erase Rand`Vers'/t913root.dta
erase Rand`Vers'/Shell935FCP.dta
erase Rand`Vers'/Shell935LRIC.dta

}    

end
