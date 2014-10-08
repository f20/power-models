===========================
github.com/f20/power-models
===========================

This repository contains:

* Perl code to build Microsoft Excel spreadsheet models that implement parts of the methods used
by the regional electricity distribution companies in England, Scotland and Wales to set their
use of system charges.

* Data published by these companies to populate these models, in a structured plain text format
(YAML) designed for use with the Perl code above.

To see this code in action, go to http://dcmf.co.uk/models/.

To create models on your own computer using this code:

Step 1. Ensure that you have Perl 5 installed and working (at least version 5.8.8).

Step 2. Download https://github.com/f20/power-models/archive/master.zip and extract all the files from it
(or, if you have Git, clone this repository using git clone https://github.com/f20/power-models.git).

Step 3. Change to the root of the repository and try this sample command:

    perl pmod.pl CDCM/Current/%-after163.yml CDCM/Data-2014-02/SPEN-SPM.yml

Step 4. If this fails, examine the error messages.  Usually the problem is a missing module
which can be installed from CPAN.  Once you have solved the problem, re-run the test command in
Step 3; and repeat until it works.

Step 5. Once everything seems to be working, you can try any of the following sample commands to
explore some of the functionality of this code:

    perl pmod.pl -xls CDCM/Current/%-clean132.yml CDCM/Current/Blank1001.yml
    perl pmod.pl ModelM/Current/%-postDCP118.yml ModelM/Data-2014-02/SSEPD-SEPD.yml
    perl pmod.pl EDCM/Issue-70/%-i70-FCP.yml EDCM/Issue-70/%-i70-LRIC.yml EDCM/Data-2013-02/ENWL.yml EDCM/Data-2013-02/WPD-WestM.yml other/Blank.yml
    perl pmod.pl -rtf -text -html -perl -yaml -graphviz CDCM/Current/%-model132.yml CDCM/Current/Blank1001.yml
    perl pmod.pl -template=All-DNOs-2014-02-postDCP163 CDCM/Current/%-micro163.yml CDCM/Data-2014-02/*.yml
    perl pmod.pl -template=All-DNOs-2014-02-postDCP118 ModelM/Current/%-postDCP118.yml ModelM/Data-2014-02/*.yml
    perl pmod.pl -template=DCP-123-2014-02 CDCM/Current/%-micro163.yml CDCM/Scaling/%-micro123.yml CDCM/Data-2014-02/*.yml
    perl pmod.pl -template=%%-DCPARPtest -pickbest CDCM/DCP-137/%-micro*.yml CDCM/Current/%-micro*.yml CDCM/Previous/%-micro*.yml CDCM/Data-2014-02/ENW*.yml CDCM/Future/Data-201*-02/ENW*.yml
    perl pmod.pl -template=%%-ARP -pickbest CDCM/Current/%-micro*.yml CDCM/Previous/%-micro*.yml CDCM/Data-20??-02/*Wales*.yml CDCM/Future/Data-20??-02/*Wales*.yml

For information on how to use the Stata tools for EDCM verification, see Stata EDCM help.txt in the "other" folder.

This software is licensed under open source licences. Check the source code for details.

THIS SOFTWARE AND DATA ARE PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ANY AUTHOR OR CONTRIBUTOR BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Franck Latrémolière, 8 October 2014.
