===========================
github.com/f20/power-models
===========================

This repository provides some code used to create Microsoft Excel
spreadsheet models that implement parts of the methods used by the regional
electricity distribution companies in England, Scotland and Wales to set
their use of system charges.  It also contains extracts from data published
by these companies to populate these models.

To see this code in action, go to http://dcmf.co.uk/models/.

To create models on your own computer using this code, try any of the
following sample commands (from the root of the repository):

    perl perl5/make.pl templates/'CDCM base'/%-clean100.yml data/model100-2012-02/*
    perl perl5/make.pl -xlsx templates/Method\ M/M-%-postDCP096.yml data/modelm-2012-02/*
    perl perl5/EDCM/mkedcm.pl -small
    perl perl5/EDCM2/mkedcm2.pl -small

The code is licensed under a mixture of permissive open source licences, mainly
two-clause BSD (for the code created specifically for this work) and the Perl
Artistic Licence (for the main packages from CPAN included in this repository).
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

Franck Latrémolière, 12 November 2012.
