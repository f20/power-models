===========================
github.com/f20/power-models
===========================

This repository contains:

* Perl code to create Microsoft Excel spreadsheet models that implement
parts of the methods used by the regional electricity distribution companies
in England, Scotland and Wales to set their use of system charges.

* Extracts from data published by these companies to populate these models,
in a YAML format designed for use with the Perl code above.

To see this code in action, go to http://dcmf.co.uk/models/.

To create models on your own computer using this code, clone this
repository, check that you have a reasonably model version of Perl 5
installed (5.8.8 or high should be fine), change to the root of the
repository, and try the following sample commands:

    perl run/make.pl CDCM/Previous/%-model100.yml CDCM/Data-2012-02/UKPN-LPN.yml
    perl run/make.pl -xlsx CDCM/Previous/%-clean130.yml CDCM/Data-2012-12/*
    perl run/make.pl -xlsx CDCM/Current/%-clean132.yml Blank.yml
    perl run/make.pl ModelM/Current/%-postDCP096.yml ModelM/Data-2012-02/*
    perl run/other/mkedcm2.pl -xlsx -small
    perl run/make.pl -comparedata ModelM/Current/%-postDCP096.yml ModelM/Data-2013-02/*LPN*

You will probably need to install a few modules from CPAN first. Hopefully
the error messages about missing dependencies are explicit enough. I have
not tried to run this on Microsoft Windows, but it should work.

The code is licensed under open source licences. Code that I wrote is
licensed under a two-clause BSD licence. A Perl Artistic Licence or GNU
Public Licence applies to the main packages from CPAN included in this
repository. Check the source code for details.

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

Franck Latrémolière, 19 March 2013.
