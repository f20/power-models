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
repository, check that you have Perl 5 installed, and try things like the
following sample commands (from the root of the repository):

    perl perl5/make.pl CDCM/%-model100.yml CDCM/2012-02/UKPN-LPN.yml
    perl perl5/make.pl -xlsx CDCM/%-clean130.yml CDCM/2012-12/*
    perl perl5/make.pl -xlsx CDCM/%-bloat132.yml Blank.yml
    perl perl5/make.pl -xlsx MM/%-MM-postDCP096.yml MM/2012-02/*
    perl perl5/EDCM/mkedcm.pl -small
    perl perl5/EDCM2/mkedcm2.pl -xlsx -small

You will probably need to install a few modules from CPAN first. Hopefully
the error messages about missing dependencies are explicit enough.

The code is licensed under open source licences. Code that I wrote is
licensed under the two-clause BSD licence. A Perl Artistic Licence or GNU
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

Franck Latrémolière, 23 February 2013.
