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

*Program is invoked within file Linkages.do

*It takes on a dataset with three variables:
*t913c1 - a list of all locations in Table 913 that (a) have a linked location and/or (b) are a linked location of some other locations
*t913c3 - location reported to be a linked location of t913c1 
*link - a variable that is 1 or 0. At start of program:
    * link = 1 if t913c1==t913c3l. 
    * link = 1 if t913c3 is the location linked to t913c1 according to original Table 913
    * link = 0 otherwise

*The program updates the value of link, so that it picks up all linkages within clusters

*The dataset contains all pairwise combinations of t913c1 and t913c3 for each company
*NB: The list of locations under t913c1 is not the same as the set of locations under t913c1 in the original Table 913. Here, the list is
*augmented by those locations that are reported in the original Table 913 as being linked locations of other locations.

capture program drop linkages
program linkages

*quietly {    

    local max_iterations = 10
    local number_companies = comp_id[_N]
    local iteration = 1

    while `iteration' <=`max_iterations'  { 

    display `iteration'
quietly {    

    egen old_link_sum=sum(link)

    local comp_selected = 1
    local initial_pos = 1

    while `comp_selected'<= `number_companies' { 

        local c_obs = obs_comp in `initial_pos' 

        local i = 1  

        while `i'<= `c_obs' {

            local j = 1       
            while `j'<= `c_obs' { 

                local pos = (`i' - 1)*`c_obs' +(`j'-1) + `initial_pos'
                local check = link in `pos'

                if `check' ==1  { 

                    local k = 1
                    while `k'<= `c_obs' {

                        local pos_i = (`i'-1)*`c_obs' + `k' - 1 + `initial_pos'
                        local pos_j = (`j'-1)*`c_obs' + `k' - 1 + `initial_pos'

                        local m = link in `pos_i'
                        local n = link in `pos_j'

                        replace link = min(1, `m'+`n') in `pos_i'
                        replace link = min(1, `m'+`n') in `pos_j'

                     local k =`k'+1
                    }
                    }
            local j = `j' + 1
            }
        local i = `i' + 1
        }

    local initial_pos = `initial_pos' + `c_obs'^2
    local comp_selected=`comp_selected'+1
    }

    egen new_link_sum=sum(link)

    local old_link_sum = old_link_sum in 1
    local new_link_sum = new_link_sum in 1

    if `old_link_sum' ~= `new_link_sum' {
        local iteration = `iteration' + 1

        drop old_link_sum
        drop new_link_sum
        }
    else {
        local iteration = `iteration' + `max_iterations'
        }
}
}
display `iteration'
end
