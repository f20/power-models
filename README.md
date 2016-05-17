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

A suitable version of Perl is usually pre-installed on Mac OS X and Linux
systems. For Windows systems, Strawberry Perl (strawberryperl.com) usually
works.

There is no support for running Perl on iOS (iPhone/iPad), Android, Chrome OS or
Windows Mobile, but spreadsheet models created by a server such as
dcmf.co.uk/models can be used on these devices.

To check whether you have a suitable version of Perl, try this at the Terminal
or command line:

    perl --version

Step 2.  Download the code.
---------------------------

Either download https://github.com/f20/power-models/archive/master.zip and
extract all the files from it, or clone this repository using git with the
command:

    git clone https://github.com/f20/power-models.git

Step 3.  Install any missing modules.
-------------------------------------

Change to the root of the repository and try this sample command:

    perl pmod.pl CDCM/Current/%-after179.yml CDCM/Data-2015-02/SPEN-SPM.yml

If this fails, examine the error messages.  Usually the problem is a missing
module, which can be installed from CPAN (www.cpan.org).  Once you have solved
the problem, re-run the test command.  Repeat until it works.

Step 4.  Start using the code.
------------------------------

Once everything seems to be working, you can try any of the following sample
commands to explore some of the functionality of this code:

    perl pmod.pl -xls CDCM/Current/%-after179.yml CDCM/Data-2015-02/WPD-SWest.yml
    perl pmod.pl ModelM/Current/%-postDCP118.yml ModelM/Data-2014-02/SSEPD-SEPD.yml
    perl pmod.pl EDCM/Current/%-clean*.yml EDCM/Data-2014-02/UKPN-EPN.yml
    perl pmod.pl -rtf -text -html CDCM/Current/%-after227.yml Blank.yml

Other code in the repository
----------------------------

There is some VBA code (Excel macros) in the VBA folder, and some Stata code in
the Stata folder. See "How to use.txt" in the "Stata" folder for information on
Stata tools to test EDCM spreadsheets.

Licensing
---------

This software is licensed under open source licences.
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

Franck Latrémolière, 17 May 2016.
