github.com/f20/power-models
===========================

This repository contains:

* Perl code to build various Microsoft Excel spreadsheet models, including
models that implement the methods used by the regional electricity distribution
companies in England, Scotland and Wales to set their use of system charges.

* Data published by these companies to populate these models, in a structured
plain text format (YAML) designed for use with the Perl code above.

To download spreadsheets built using this code, go to http://dcmf.co.uk/models/.

To build spreadsheet models on your own computer using this code, follow the
instructions below.

Step 1.  Check or set-up Perl 5.
--------------------------------

You need Perl 5, v5.8.8 or later.

A suitable version of Perl is pre-installed on Apple macOS systems.
For Microsoft Windows systems, Strawberry Perl (strawberryperl.com) is
usually a good choice.  For FreeBSD, Linux and similar systems, Perl is
either pre-installed or readily available from the ports/packages system.

To test whether you have a suitable version of Perl, try this at the Terminal
or command line:

    perl --version

Step 2.  Download the code.
---------------------------

Either download https://github.com/f20/power-models/archive/master.zip and
extract all the files from it, or use a git client to clone this repository.

Step 3.  Install any missing modules.
-------------------------------------

Change to the root of the repository and try this sample command:

    perl pmod.pl CDCM/2017-02-Baseline/%-extras227.yml CDCM/2017-02/SPEN-SPM.yml

If this fails, examine the error messages.  Usually the problem is a missing
module, which can be installed from CPAN (www.cpan.org).  Once you have solved
the problem, re-run the test command and repeat the process until it works.

Step 4.  Start using the code.
------------------------------

Once everything seems to be working, you can try any of the following sample
commands to explore some of the functionality of this code:

    perl pmod.pl ModelM/2014-02-Baseline/%-cleancombo118.yml ModelM/2015-02/SSEPD-SEPD.yml
    perl pmod.pl -rtf -text -html CDCM/2017-02-Baseline/%-clean227.yml Blank.yml

Other code in the repository
----------------------------

There is some VBA code (Excel macros) in the VBA folder, and some Stata code in
the Stata folder. The VBA code is not currently documented. See "How to use.txt"
in the "Stata" folder for information on Stata tools to test EDCM spreadsheets.

Licensing
---------

All the components of this software are licensed under open source licences.
Check the source code for details.

THIS SOFTWARE AND DATA ARE PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED
WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO
EVENT SHALL ANY AUTHOR OR CONTRIBUTOR BE LIABLE FOR ANY DIRECT, INDIRECT,
INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Franck Latrémolière, 29 November 2016.
