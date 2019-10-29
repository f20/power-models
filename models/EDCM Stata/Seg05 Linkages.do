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
*Aim: Defining set of clusters of linked locations (for LRIC)
***************

***************
*Datasets used:
*913
***********

***************
*Input variables used:
*t913c1 = Location name/ID
*t913c3 = Linked location (if any)
***************

clear

use 913

*1. Building a dataset with name of those locations with a linked location and with name of those observations that are linked locations to others
drop if t913c1=="Not used"|t913c1==""
keep company line t913c1 t913c3

sort company t913c1
save 913_vLinkA, replace

save 913_vLinkC, replace

*2. Defining program to establish set of clusters

capture program drop LinkagesNew
program LinkagesNew

quietly{

*a.Generate list of linked locations for each t913c1. Impose maximum of 7 linked locations

local i = 1

    while `i' <= 7 {

    clear
    use 913_vLinkA
    ren t913c1 Link`i'
    sort comp Link`i'
    save 913_vLinkB, replace

    clear
    use 913_vLinkC
    ren t913c3 Link`i'
    sort comp Link`i'

    merge comp Link`i' using 913_vLinkB
    drop if _merge==2
    drop _merge

    save 913_vLinkC, replace

    local i = `i' + 1
    }

*drop t913c3

*b. Create a variable picking up concatenated names and one picking up number of observations for each company

gen ConcatLocation="||"+t913c1+"||"

local i = 1

    while `i'<=_N {
        local j = 1

        while `j'<=7{

        local RelevantLink = Link`j' in `i'

        if "`RelevantLink'"~=""{
        replace ConcatLocation=ConcatLocation + Link`j'+"||" in `i'
        }
        local j = `j'+1
        }

    local i = `i'+1
    }

gen cluster=_n

*Encode company names so that can run program
encode company, gen(IdComp)

*Create variable to pick up number of observations for each company

gen Temp=1
by company, sort: egen ObsComp = sum(Temp)
drop Temp

gen FirstObsComp=1

sort comp line
local i = 1

while `i'<=_N{

    local CompInterest= IdComp in `i'
    if `CompInterest'~=1 {

        local j = `i' - 1

        local CurrComp=comp in `i'
        local PrevComp=comp in `j'

        local FirstPrev=FirstObsComp in `j'

        if "`CurrComp'"~="`PrevComp'"{
            local ObservPrev=ObsComp in `j'
            replace FirstObsComp=`FirstPrev' + `ObservPrev' in `i'
            }

        else {
            replace FirstObsComp=`FirstPrev' in `i'
            }
    }
local i = `i'+1
}
gen LastObsComp = ObsComp + FirstObsComp-1

*c. Loop to go through ConcatLocation to see to which group it belongs
*Note: Need to ensure that observations within company are sorted in order of number of characters in CocatLocation, so that right cluster is picked

gen ConcatLength=length(ConcatLocation)

gsort comp -ConcatLength

local i = 1
    while `i'<=_N {

    local m = FirstObsComp in `i'
    local p=  LastObsComp  in `i'

    local j = `m'

    while `j'<=`p' {

    local LookIn = ConcatLocation in `j'
    local LookFor = ConcatLocation in `i'

        if strpos("`LookIn'", "`LookFor'") == 0 {
        local j = `j'+1
        }

        if strpos("`LookIn'", "`LookFor'") > 0 {
        replace cluster = `j' in `i'
        local j = `j'+ `p'
        }
    }
    local i = `i'+1
    }

}

end

LinkagesNew

gen st_cluster=string(cluster)

*Check whether there are clusters with more than 8 linked locations

capture program drop OversizedClusters
program OversizedClusters

quietly{

local i =1
    while `i'<=_N {

    local LinkedLoc=t913c3 in `i'

    if "`LinkedLoc'"~="" {
    local BadComp= company in `i'
    display as error "Invalid input data
    display as error "`BadComp' : There are clusters with more than 8 linked locations"
    }

    local i = `i'+1
    }
}
end

OversizedClusters

*8. Keeping useful variables

*keep comp t913c1 st_cluster
drop cluster
sort  company t913c1
save 913_vCluster, replace

*Erase temporary data files
erase 913_vLinkA.dta
erase 913_vLinkB.dta
erase 913_vLinkC.dta

***************
*Data files kept:
*913_vClusterNew
***************
