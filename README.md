github.com/f20/power-models
===========================

This repository contains an open source Perl 5 system to construct
Microsoft Excel spreadsheet models to calculate electricity distribution
use of system charges.

The repository also contains some historical data used by regional electricity
distribution companies in England, Scotland and Wales in models to set their
use of system charges, in a format suitable for use with the Perl 5 code.

Go to http://dcmf.co.uk/models/ to download some workbooks built using this code.

To get started with building spreadsheet models on your own computer using
this code, follow the instructions below.

Step 1. Set-up a Perl 5 development environment.
------------------------------------------------

You need a terminal interface, and Perl 5 (preferably v5.10 or later,
but most of the code is compatible with v5.8.8).

This is normally easy to set-up on desktop and server computing platforms:
* On Apple macOS, the built-in Terminal.app and Perl 5 installations are good.
* On Microsoft Windows, the built-in Command Prompt and the Strawberry Perl
package available from strawberryperl.com are good.
* On FreeBSD and many Linux distributions, console applications and Perl 5 are
either pre-installed or available from the ports/packages system.

Step 2. Download the code.
---------------------------

Either download https://github.com/f20/power-models/archive/master.zip and
extract all the files from it, or use a git client to clone the repository.

Step 3. Install any missing modules.
-------------------------------------

Change directory to the root of the extraction folder or repository, and
try these sample commands:

    perl -Icpan -Ilib -MSpreadsheetModel::Book::Manufacturing -e "SpreadsheetModel::Book::Manufacturing->factory->runAllWithFiles('models/Other/Formats sampler.yml')"

    perl -Icpan -Ilib -MSpreadsheetModel::Book::Manufacturing -e "SpreadsheetModel::Book::Manufacturing->factory(validate=>['lib'])->runAllWithFiles('models/CDCM/2017-02-Baseline/%-extras227.yml','models/CDCM/2017-02/SPEN-SPM.yml')"

This should create Microsoft Excel workbooks in the current working directory.
If not, examine the error messages. Sometimes the problem is a missing
module which can be installed from CPAN (www.cpan.org).

Other code in the repository
----------------------------

The "EDCM Stata validator" and "EDCM Stata generator" folders under models
contains Stata tools to test aspects of implementations of the EDCM charging
methodology. See "How to use.txt" under "EDCM Stata validator" for more
information.  This code is deprecated and unmaintained.

Licensing
---------

Components of this software are licensed under open source licences; see
the source code for details.

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

Franck Latrémolière, 23 February 2025.
