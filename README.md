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

To create models on your own computer using this code, clone this repository, check that you
have Perl 5.8.8 or higher installed and working, change to the root of the repository, and try
some of the following sample commands.  You might need to install some modules from CPAN; error
messages should say which.  Here are a few sample commands:

    perl run/make.pl CDCM/Current/%-wfl132.yml CDCM/Data-2014-02/NPG-Northeast.yml
    perl run/make.pl -xlsx CDCM/Current/%-clean132.yml CDCM/Current/Blank1001.yml
    perl run/make.pl ModelM/Current/%-postDCP118.yml ModelM/Data-2014-02/SSEPD-SEPD.yml
    perl run/make.pl -xlsx EDCM/Issue-70/%-i70-FCP.yml EDCM/Issue-70/%-i70-LRIC.yml EDCM/Data-2013-02/ENWL.yml EDCM/Data-2013-02/SPEN-SPM.yml Other/Blank.yml


This software is licensed under open source licences. Check the source code for details.

THIS SOFTWARE AND DATA ARE PROVIDED "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL ANY AUTHOR OR CONTRIBUTOR BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT
LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

Franck Latrémolière, 23 March 2014.
