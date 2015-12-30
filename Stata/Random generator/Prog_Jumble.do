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

capture program drop jumble
program jumble

args v

quietly{

*Define proportion of times that variable is jumbled up (at each stage of jumbing)
local prop = 0.05
****

*Defining matrix ascii which contains ascii values of characters being jumbled in to name of location

*Start off with ascii values 9 and 10, relate to backspace and horizontal tab.

matrix input ascii = (9, 10)

*Picking up ascii characters 32 to 125
*Not including quotation mark (ascii 34) because it messes up Stata code
*Not including dollar sign (ascii 36) because it messes up Stata code (global macro)
*Not including asterisk (ascii 42) because of a feature of Excel's MATCH
*Not including equal sign (ascii 61) because it messes up Perl code
*Not including question mark (ascii 63) because of a feature of Excel's MATCH
*Not including tilde (ascii 126) because of a feature of Excel's MATCH

local i = 32

while `i'<126 {
    if `i'~=34&`i'~=36&`i'~=42&`i'~=61&`i'~=63 {
        matrix ascii=(ascii,`i')
    }
    local i = `i'+1
    }

replace `v'=cond(runiform()<`prop',`v'+char(ascii[1,int(98*runiform())]),`v')
replace `v'=cond(runiform()<`prop',char(ascii[1,int(98*runiform())]),`v')
replace `v'=cond(runiform()<`prop',char(ascii[1,int(98*runiform())])+`v',`v')

*Just on off chance of having two locations with same name, following the above replacement
sort `v'
drop if `v'==`v'[_n-1]

}
end
