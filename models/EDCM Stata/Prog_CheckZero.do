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

capture program drop CheckZero
program CheckZero

quietly{

    local i = 1
    while "``i''" ~="" {

        local k = 1
        while `k' <=_N {

        local n  = "``i''" in `k'

        if `n'== . {
                global GuiltyCompany = comp in `k'
                display as error "Invalid input data for $GuiltyCompany"
                display as error "``i''" ": Observation:"  "`k'" ":  Dividing by a missing value"
                global r = $r +1
                exit
        }

        if `n'==0 {
                global GuiltyCompany = comp in `k'
                display as error "Invalid input data for $GuiltyCompany"
                display as error "``i''"  ": Observation:"  "`k'" ":  Dividing by zero"
                global r = $r +1
                exit
        }

        local k = `k' + 1
        }

    local i = `i' + 1
}
}
end
