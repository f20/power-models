github.com/f20/power-models
===========================

This repository contains an open source Perl 5 system to construct
Microsoft Excel spreadsheet models that address business problems.

This project's first task was to implement the methods used by the regional
electricity distribution companies in England, Scotland and Wales to set
their use of system charges. It has subsequently expanded to explore other
areas to which the Perl-managed spreadsheet model methodology pioneered for
distribution charging models can make a useful contribution.

The repository also contains data used by regional electricity distribution
companies in England, Scotland and Wales in models to set their use of
system charges, in a form suitable for use with the Perl code above, and
some (currently undocumented) tools to manage these data.

To download some of the use of system charging workbooks that can be built
using this code, go to http://dcmf.co.uk/models/.

To get started with building spreadsheet models on your own computer using
this code, follow the instructions below.

Step 1. Set-up a Perl 5 development environment.
------------------------------------------------

You need a terminal or console interface, and Perl 5 (v5.8.8 or later).

This is normally easy to set-up on desktop and server computing platforms.

On Apple macOS, the built-in Terminal.app and Perl 5 installations are good.

On Microsoft Windows 7-10, the built-in Command Prompt and the Strawberry
Perl package available from strawberryperl.com are good.

On FreeBSD and many Linux distributions, console applications and Perl 5 are
either pre-installed or available from the ports/packages system.

On mobile operating systems, setting up a suitable environment is much more
troublesome; the dcmf.co.uk/models website might better meet your needs.

To test whether you have a suitable version of Perl, try this at the
Terminal or command line:

    perl --version

Step 2. Download the code.
---------------------------

Either download https://github.com/f20/power-models/archive/master.zip and
extract all the files from it, or use a git client to clone this repository.

Step 3. Install any missing modules.
-------------------------------------

Change to the root of the repository and try this sample command:

    perl pmod.pl Sampler/%-short.yml Blank.yml

If this fails, examine the error messages. Often the problem is a missing
module, which can be installed from CPAN (www.cpan.org). Once you have
solved the problem, re-run the test command and repeat until it works.

Step 4. Start using the code.
------------------------------

Once everything seems to be working, you can try any of the following
sample commands to explore some of the functionality of this code:

    perl pmod.pl CDCM/2017-02-Baseline/%-extras227.yml CDCM/2017-02/SPEN-SPM.yml
    perl pmod.pl ModelM/2014-02-Baseline/%-cleancombo118.yml ModelM/2015-02/SSEPD-SEPD.yml
    perl pmod.pl -rtf -text -html CDCM/2017-02-Baseline/%-clean227.yml Blank.yml

Other code in the repository
----------------------------

The Stata folder contains Stata tools to test workbooks implementing aspects of the
EDCM charging methodology. See "How to use.txt" in the "Stata" folder for details.

The VBA folder contains VBA code and an Excel add-in with some spreadsheet analysis
tools. These are currently undocumented.

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

Franck Latrémolière, 13 November 2017.
