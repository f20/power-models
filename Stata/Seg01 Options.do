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
*1. Import CSV file with information on options
*2. Check that values for some of the options are in line with "base"  Stata model
*3. Define global variables to reflect options: these can be called later by $Optionxxx, where xxx is the name of the option itself 
***************

*1 Importing Table Options

clear
insheet using 0.csv, c n case

*Deleting blank spaces within values of variables

foreach var of varlist * {

        capture confirm string variable `var'
            if !_rc {
            replace `var'=subinstr(`var'," ","",.)
            }                   
        }

*2 Check options used in Stata model are defined within the Options.csv, and that values given for options match values recognised by the model
*Deal with option "checksums" separately, as allowed 

global OptionProblem = 0

*************************************************
*****Defining options and values of options******

local Options "lowerIntermittentCredit checksums dcp185 legacy201 dcp189"

local Val_lowerIntermittentCredit "0 1"
local Val_checksums "0 Tariffchecksum5;Modelchecksum7 Linechecksum5;Tablechecksum7"
local Val_dcp185 "0 1 2"
local Val_legacy201 "0 1"
local Val_dcp189 "0 proportionsplitShortfall"

*************************************************
*************************************************

local m: word count `Options'

local j= 1
while `j'<=`m'{

    local OptionVariable: word `j' of `Options'
    capture confirm variable `OptionVariable'
    if _rc !=0 {
            noisily display as error "`OptionVariable' was not defined in Options.csv file. See file 0.csv."
            global OptionProblem = 1 
            exit
            }
    else {
        local OptionValue = `OptionVariable' in 1
        if subinword("`Val_`OptionVariable''", "`OptionValue'","XXXXXX",.)=="`Val_`OptionVariable''" {
            noisily display as error "The value `OptionValue' which was given for option `OptionVariable' is not valid. See file 0.csv."
            global OptionProblem = 1  
            exit
            }
*3. Define global variables

        global Option`OptionVariable' `OptionValue'
        }
    local j = `j'+1
    }

*4. Deal with fact that option "Tariffchecksum5;Modelchecksum7" is called "Linechecksum5;Tablechecksum7" in some models (and script uses former designation)

	if "$Optionchecksums" == "Linechecksum5;Tablechecksum7" {
		global Optioncheksums Tariffchecksum5;Modelchecksum7 
		}

