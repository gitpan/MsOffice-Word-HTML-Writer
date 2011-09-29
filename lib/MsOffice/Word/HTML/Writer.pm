package MsOffice::Word::HTML::Writer;

use warnings;
use strict;
use MIME::QuotedPrint qw/encode_qp/;
use MIME::Base64      qw/encode_base64/;
use MIME::Types;
use Carp;
use Params::Validate qw/validate SCALAR HASHREF/;


our $VERSION = '1.01';


sub new {
  my $class = shift;

  # validate named parameters
  my $param_spec = {
    title        => {type => SCALAR, optional => 1},
    head         => {type => SCALAR, optional => 1},
    hf_head      => {type => SCALAR, optional => 1},
    WordDocument => {type => HASHREF, optional => 1},
   };
  my %params = validate(@_, $param_spec);

  # create instance
  my $self = {
    MIME_parts   => [],
    sections     => [{}],
    title        => $params{title}
                 || "Document generated by MsOffice::Word::HTML::Writer",
    head         => $params{head}      || "",
    hf_head      => $params{hf_head}   || "",
    WordDocument => $params{WordDocument},
   };

  bless $self, $class;
}


sub create_section {
  my $self = shift;

  # validate named parameters
  my $param_spec = {page => {type => HASHREF, optional => 1}};
  $param_spec->{$_} = {type => SCALAR, optional => 1}
    for qw/header footer first_header first_footer new_page/;
  my %params = validate(@_, $param_spec);

  # if first automatic section is empty, delete it
  $self->{sections} = []
    if scalar(@{$self->{sections}}) == 1 && !$self->{sections}[0]{content};

  # add the new section
  push @{$self->{sections}}, \%params;
}


sub write {
  my $self = shift;

  # add html arguments to current section content
  $self->{sections}[-1]{content} .= join ("", @_);
}



sub save_as {
  my ($self, $filename) = @_;

  # default extension is ".doc"
  $filename .= ".doc" unless $filename =~ /\.\w{1,5}$/;

  # open the file
  open my $fh, ">:crlf", $filename
    or croak "could not open >$filename: $!";

  # write content and close
  print $fh $self->content;
  close $fh;
}


sub attach {
  my ($self, $name, $open1, $open2, @other) = @_;

  # open a handle to the attachment (need to dispatch according to number
  # of args, because perlfunc/open() has complex prototyping behaviour)
  my $fh;
  if (@other) { 
    open $fh, $open1, $open2, @other
      or croak "open $open1, $open2, @other : $!"; 
  }
  elsif ($open2) {
    open $fh, $open1, $open2
      or croak "open $open1, $open2 : $!"; 
  }
  else {
    open $fh, $open1
      or croak "open $open1 : $!"; 
  }

  # slurp the content
  binmode($fh) unless $name =~ /\.(html?|css|te?xt|rtf)$/i;
  local $/;
  my $attachment = <$fh>;

  # add the attachment (filename and content)
  push @{$self->{MIME_parts}}, ["files/$name", $attachment];
}


sub page_break {
  my ($self, $break) = @_;
  $break ||= 'always';
  return qq{<br clear='all' style='page-break-before:$break'>\n};
}


sub tab {
  my ($self, $n_tabs) = @_;
  $n_tabs ||= 1;
  return qq{<span style='mso-tab-count:$n_tabs'></span>};
}

sub field {
  my ($self, $fieldname, $args, $content) = @_;

  for ($args, $content) {
    $_ ||= "";                              # undef replaced by empty string
    s/&/&amp;/g,  s/</&lt;/g, s/>/&gt;/g;   # replace HTML entities
  }

  my $field;

  # when args : long form of field encoding
  if ($args) {
    my $space = qq{<span style='mso-spacerun:yes'>�</span>};
    $field = qq{<span style='mso-element:field-begin'></span>}
           . $space . $fieldname . $space . $args
           . qq{<span style='mso-element:field-separator'></span>}
           . $content
           . qq{<span style='mso-element:field-end'></span>};
  }
  # otherwise : short form of field encoding
  else {
    $field = qq{<span style='mso-field-code:"$fieldname"'>$content</span>};
  }

  return $field;
}

sub quote {
  my ($self, $text) = @_;
  my $args = $text;
  $args =~ s/"/\\"/g;
  $args = qq{"$args"};
  $args =~ s/"/&quot;/g;
  return $self->field('QUOTE', $args, $text);
}



sub content {
  my ($self) = @_;

  # separator for parts in MIME document
  my $boundary = qw/__NEXT_PART__/;

  # MIME multipart header
  my $mime = qq{MIME-Version: 1.0\n}
           . qq{Content-Type: multipart/related; boundary="$boundary"\n\n}
           . qq{MIME document generated by MsOffice::Word::HTML::Writer\n\n};

  # generate each part (main document must be first)
  my @parts = $self->_MIME_parts;
  my $filelist = $self->_filelist(@parts);
  for my $pair ($self->_main, @parts, $filelist) {
    my ($filename, $content) = @$pair;
    my $mime_type = MIME::Types->new->mimeTypeOf($filename);
    my ($encoding, $encoded);
    if ($mime_type =~ /^text|xml$/) {
      $encoding = 'quoted-printable';
      $content  =~ s/\r\n/\n/g;
      $encoded  = encode_qp($content, ''); # '': no "soft line breaks"
    }
    else {
      $encoding = 'base64';
      $encoded  = encode_base64($content);
    }

    $mime .= qq{--$boundary\n}
          .  qq{Content-Location: file:///C:/foo/$filename\n}
          .  qq{Content-Transfer-Encoding: $encoding\n}
          .  qq{Content-Type: $mime_type\n\n}
          .  $encoded
          . "\n";
  }

  # close last MIME part
  $mime .= "--$boundary--\n";

  return $mime;
}


#======================================================================
# PRIVATE METHODS
#======================================================================

sub _main {
  my ($self) = @_;

  # body : concatenate content from all sections
  my $body = "";
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    # section break
    if ($i > 1) {
      # type of break
      my $break = $section->{new_page};
      $break = 'always' if $break && $break !~ /\w/; # if true but not a word
      $break ||= 'auto';                             # if false
      # otherwise, type of break will just be the word given in {new_page}

      # insert into body
      my $style = qq{page-break-before:$break;mso-break-type:section-break};
      $body .= qq{<br clear=all style='$style'>\n};
    }

    # section content
    $body .= qq{<div class="Section$i">\n$section->{content}\n</div>\n};

    $i += 1;
  }

  # assemble head and body into a full document
  my $html
    = qq{<html xmlns:v="urn:schemas-microsoft-com:vml"\n}
    . qq{      xmlns:o="urn:schemas-microsoft-com:office:office"\n}
    . qq{      xmlns:w="urn:schemas-microsoft-com:office:word"\n}
    . qq{      xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"\n}
    . qq{      xmlns="http://www.w3.org/TR/REC-html40">\n}
    . $self->_head
    . qq{<body>\n$body</body>\n}
    . qq{</html>\n};
  return ["main.htm", $html];
}


sub _head {
  my ($self) = @_;

  # HTML head : link to filelist, title, view format and styles
  my $head 
    = qq{<head>\n}
    . qq{<link rel=File-List href="files/filelist.xml">\n}
    . qq{<title>$self->{title}</title>\n}
    . $self->_xml_WordDocument
    . qq{<style>\n} . $self->_section_styles . qq{</style>\n}
    . $self->{head}
    . qq{</head>\n};
  return $head;
}



sub _xml_WordDocument {
  my ($self) = @_;
  my $xml_root = $self->{WordDocument} or return "";
  return "<xml><w:WordDocument>\n" 
       . _w_xml($xml_root)
       . "</w:WordDocument></xml>\n";
}


sub _w_xml {
  my $node = shift;
  my $xml = "";
  while (my ($k, $v) = each %$node) {
    $xml .= $v ? (               # �l�ment avec contenu
                   "<w:$k>"
                  . (ref $v ? _w_xml($v) : $v)
                  . "</w:$k>\n" )
               : "<w:$k />\n";     # �l�ment sans contenu
  }
  return $xml;
}


sub _section_styles {
  my ($self) = @_;

  my $styles = "";
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    my $properties = "";

    # page properties (size and margin)
    foreach my $prop (qw/size margin/) {
      my $val = $section->{page}{$prop} or next;
      $properties .= qq{  $prop:$val;\n};
    }

    # headers and footers 
    my $has_first_page;
    foreach my $prop (qw/header_margin footer_margin/) {
      my $val = $section->{page}{$prop} or next;
      (my $property = $prop) =~ s/_/-/g;
      $properties .= qq{  mso-$property:$val;\n};
    }
    foreach my $hf (qw/header footer first_header first_footer/) {
      $section->{$hf} or next;
      $has_first_page = 1 if $hf =~ /^first/;
      (my $property = $hf) =~ s/_/-/;
      $properties 
        .= qq{  mso-$property:url("files/header_footer.htm") $hf$i;\n};
    }
    $properties .= qq{  mso-title-page:yes;\n} if $has_first_page;

    # style definitions for this section
    $styles .= qq[\@page Section$i {\n$properties}\n]
            .  qq[div.Section$i {page:Section$i}\n];
    $i += 1;
  }

  return $styles;
}


sub _MIME_parts {
  my ($self) = @_;

  # attachments supplied by user
  my @parts = @{$self->{MIME_parts}};

  # additional attachment : computed file with headers and footers
  my $hf_content = $self->_header_footer;
  unshift @parts, ["files/header_footer.htm", $hf_content] if $hf_content;

  return @parts;
}


sub _header_footer {
  my ($self) = @_;

  # create a div for each header/footer in each section
  my $hf_divs = "";
  my $i = 1;
  foreach my $section (@{$self->{sections}}) {

    # deal with headers/footers defined in that section
    foreach my $hf (qw/header footer first_header first_footer/) {
      $section->{$hf} or next;
      (my $style = $hf) =~ s/^first_//;
      $hf_divs .= qq{<div style='mso-element:$style' id='$hf$i'>\n}
               .  $section->{$hf} . "\n"
               .  qq{</div>\n};
    }

    $i += 1;
  }

  # if at least one such div, need to create an attached file
  my $header_footer = !$hf_divs ? "" :
        qq{<html>\n}
      . qq{<head>\n}
      . qq{<link id=Main-File rel=Main-File href="../main.htm">\n}
      . $self->{hf_head}
      . qq{</head>\n}
      . qq{<body>\n} . $hf_divs . qq{</body>\n}
      . qq{</html>\n};

  return $header_footer;
}



sub _filelist {
  my ($self, @parts) = @_;

  # xml header
  my $xml = qq{<xml xmlns:o="urn:schemas-microsoft-com:office:office">\n}
          . qq{ <o:MainFile HRef="../main.htm"/>\n};

  # refer to each attached file
  foreach my $part (@parts) {
    $xml .= qq{ <o:File HRef="$part->[0]"/>\n};
  }

  # the filelist is itself an attached file
  $xml .= qq{ <o:File HRef="filelist.xml"/>\n};

  # closing tag;
  $xml .=  qq{</xml>\n};

  return ["files/filelist.xml", $xml];
}



1;

__END__

=head1 NAME

MsOffice::Word::HTML::Writer - Writing documents for MsWord in HTML format

=head1 SYNOPSIS

  use MsOffice::Word::HTML::Writer;
  my $doc = MsOffice::Word::HTML::Writer->new(
    title        => "My new doc",
    WordDocument => {View => 'Print'},
  );
  
  $doc->write("<p>hello, world</p>", 
              $doc->page_break, 
              "<p>hello from another page</p>");
  
  $doc->create_section(
    page => {size   => "21.0cm 29.7cm",
             margin => "1.2cm 2.4cm 2.3cm 2.4cm"},
    header => sprintf("Section 2, page %s of %s", 
                                  $doc->field('PAGE'), 
                                  $doc->field('NUMPAGES')),
    footer => sprintf("printed at %s", 
                                  $doc->field('PRINTDATE')),
    new_page => 1, # or 'left', or 'right'
  );
  $doc->write("this is the second section, look at header/footer");
  
  $doc->attach("my_image.gif", $path_to_my_image);
  $doc->write("<img src='files/my_image.gif'>");
  
  $doc->save_as("/path/to/some/file");

=head1 DESCRIPTION

=head2 Goal

The present module is one way to programatically generate documents
targeted for Microsoft Word (MsWord). It doesn't need
MsWord to be installed, and doesn't even require a Win32 machine
(which is why it is not in the C<Win32> namespace).

=head2 MsWord and HTML

MsWord can read documents encoded in old native binary format, in Rich
Text Format (RTF), in XML (either ODF or OOXML), or -- maybe this is
less known -- in HTML, with some special markup for pagination and
other MsWord-specific features. Such HTML documents are often in
several parts, because attachments like images or headers/footers need
to be in separate files; however, since it is more convenient to carry
all data in a single file, MsWord also supports the "MHTML" format (or
"MHT" for short), i.e. an encapsulation of a whole HTML tree into a
single file encoded in MIME multipart format. This format can be
generated interactively from MsWord by calling the "SaveAs" menu and
choosing the F<.mht> extension.

Documents saved with a F<.mht> extension will not directly 
reopen in MsWord : when clicking on such documents, Windows
chooses Internet Explorer as the default display program.
However, these documents can be simply renamed with a
F<.doc> extension, and will then open directly in MsWord.
By the way, the same can be done with XML or RTF documents.
That is to say, MsWord is able to recognize the internal
format of a file, without any dependency on the filename.

=head2 Features of the module

C<MsOffice::Word::HTML::Writer> helps you to programatically generate
MsWord documents in MHT format. The advantage of this technique is
that one can rely on standard HTML mechanisms for layout control, such
as styles, tables, divs, etc. Of course this markup can be produced
using your favorite HTML templating module; the added value
of C<MsOffice::Word::HTML::Writer> is to help building the 
MIME multipart file, and provide some abstractions for 
representing MsWord-specific features (headers, footers, fields, etc.).

=head2 Advantages of MHT format

The MHT format is probably the most convenient
way for programmatic document generation, because

=over

=item *

unlike Excel, MsWord native binary format (used in versions up to 2003)
is unpublished and therefore cannot be generated without the MsWord executable.

=item *

remote control of the MsWord program through an OLE connection,
as in L<Win32::Word::Writer|Win32::Word::Writer>, requires a
local installation of Microsoft Office, and is not well
suited for server-side generation because the MsWord program might hang
or might open dialog boxes that require user input.

=item *

generation of documents in RTF is possible, but 
authoring the models requires deep knowledge of the RTF structure
--- see L<RTF::Writer>.

=item *

authoring models in XML also requires
deep knowledge of the XML structure.

Instead of working directly at the XML level, one could use the
L<OpenOffice::OODoc> distribution on CPAN, which provides programmatic
access to the "ODF" XML format used by OpenOffice. MsWord is able to
read and produce such ODF files, but this is not fully satisfactory
because in that mode many MsWord features are disabled or restricted.

The XML format used by MsWord is called "OOXML"; to
my knowledge, there is no CPAN module providing an API to 
this format.


=back

By contrast, C<MsOffice::Word::HTML::Writer> allows you to 
produce documents even with little knowledge of MsWord.
Besides, since the content is in HTML, it can be assembled
with any HTML tool, and therefore also requires little knowledge
of Perl.

One word of warning, however : opening MHT documents in MsWord is
a bit slower than native binary or RTF documents, because MsWord needs to
parse the HTML, compute the layout and convert it into its internal
representation.  Therefore MHT format is not recommended for very
large documents.

=head2 Usage

C<MsOffice::Word::HTML::Writer> is used in production
at Geneva courts of law, for generating thousands of documents
per day, from hundreds of models, with an architecture of 
reusable document parts implemented by Template Toolkit mechanisms
(macros, blocks and views).


=head1 METHODS

B<General convention> : method names that start
with a I<verb> may change the internal state of the 
writer object (for example L</write>, L</create_section>);
method names that are I<nouns> return data without modifying
the internal state (for example L</field>, L</content>, L<page_break>).



=head2 new

    my $doc = MsOffice::Word::HTML::Writer->new(%params);

Creates a new writer object. Optional parameters are :

=over

=item title

document title

=item head

any HTML declarations you may want to include in the
C<head> part of the generated document (for example
inline CSS styles or links to attached stylesheets).

=item hf_head

any HTML declarations you may want to include in the
C<head> part of the I<headers and footers> HTML document
(MsWord requires headers and footers to be 
specified as C<div>s in a separate HTML document).

=item WordDocument

a hashref of options to include as an XML island in the 
HTML C<head>, corresponding to various options in the 
MsWord "Tools/Options" panel. These will be included
in a XML element named C<< <w:WordDocument> >>, and
all children elements will be automatically prefixed
by C<w:>. The hashref may contain nested hashrefs, such as

  WordDocument => { View => 'Print',
                    Compatibility => {DoNotExpandShiftReturn => "",
                                      BreakWrappedTables     => ""} }

Names and values of options
must be found from the Microsoft documentation, or from
reverse engineering of HTML files generated by MsWord.

=back

Parameters may also be passed as a hashref instead of a hash.


=head2 write

  $doc->write("<p>hello, world</p>");

Adds some HTML into the document body.


=head2 attach

  $doc->attach($localname, $filename);
  $doc->attach($localname, "<", \$content);
  $doc->attach($localname, "<&", $filehandle);

Adds an attachment into the document; the attachment will be encoded
as a MIME part and will be accessible under C<files/$localname>.

The remaining arguments to C<attach> specify the source of the attachment;
they are directly passed to L<perlfunc/open> and therefore have the same
API flexibility : you can specify a filename, a reference to a memory
variable, a reference to another filehandle, etc.



=head2 create_section

  $doc->create_section(
    page => {size   => "21.0cm 29.7cm",
             margin => "1.2cm 2.4cm 2.3cm 2.4cm"},
    header => sprintf("Section 2, page %s of %s", 
                                  $doc->field('PAGE'), 
                                  $doc->field('NUMPAGES')),
    footer => sprintf("printed at %s", 
                                  $doc->field('PRINTDATE')),
    new_page => 1, # or 'left', or 'right'
  );

Opens a new section within the document
(or, if this is called before any L</write>, 
setups pagination parameters for the first section).
Subsequent calls to the L</write> method will add content to
that section, until the next L</create_section> call.

Pagination parameters are all optional and may be given
either as a hash or as a hashref; accepted parameters are :

=over

=item page

Hashref of CSS page styles, such as :

=over

=item size

Paper size (for example C<21cm 29.7cm>)

=item margin

Margins (top right bottom left).

=item header_margin

Margin for header

=item footer_margin

Margin for footer

=back


=item header

Header content (in HTML)

=item first_header

Header content for the first page of that section.

=item footer

Footer content (in HTML).

=item first_footer

Footer content for the first page.

=item new_page

If true, a page break will be inserted before the new section.
If the argument is the word C<'left'> or C<'right'>, one or two
page breaks will be inserted so that the next page is formatted
as a left (right) page.

=back



=head2 save_as

  $doc->save_as("/path/to/some/file");

Generates the MIME document and saves it at the given location.
If no extension is present, file extension F<.doc> will be added
by default to the filename.


=head2 content

Returns the whole MIME-encoded document as a single string; this is
used internally by the L</save_as> method.  Direct call is useful if
you don't want to save the document into a file, but want to do
something else like embedding it in a message or a ZIP file, or
returning it as an HTTP response.


=head2 page_break

  my $html = $doc->page_break;
  my $html = $doc->page_break('left');
  my $html = $doc->page_break('right');

Returns HTML markup for encoding a page break.
If an argument C<'left'> or C<'right'> is given, one or two
page breaks will be inserted so that the next page is formatted
as a left (right) page.

=head2 tab

  my $html = $doc->tab($n_tabs);

Returns HTML markup for encoding one or several tabs. If C<$n_tab> is
omitted, it defaults to 1.


=head2 field

  my $html = $doc->field($fieldname, $args, $content);

Returns HTML markup for a MsWord field.

Optional C<$args> is a string with arguments or flags for
the field. See MsWord help documentation for the list of
field names and their associated arguments or flags.

Optional C<$content> is the initial displayed content for the
field (because unfortunately MsWord does not immediately compute
the field content when opening the document; users will have
to explicitly request to update all fields, by selecting the whole
document and then hitting the F9 key).

Here are some examples :

  my $header = sprintf "%s of %s", $doc->field('PAGE'), 
                                   $doc->field('NUMPAGES');
  my $footer = sprintf "created at %s, printed at %s", 
                 doc->field(CREATEDATE => '\\@ "d MM yyyy"'),
                 doc->field(PRINTDATE  => '\\@ "dddd d MMMM yyyy" \\* Upper');
  my $quoted = $doc->field('QUOTE', '"hello, world"', 'hello, world');

=head2 quote

  my $html = $doc->quote($text);

Shortcut to produce a QUOTE field (see last field example just above).


=head1 AUTHORING MHT DOCUMENTS

=head2 HTML for MsWord

MsWord does not support the full HTML and CSS standard,
so authoring MHT documents requires some trial and error.
Basic divs, spans, paragraphs and tables,
are reasonably supported, together with their common CSS
properties; but fancier features  like floats, absolute 
positioning, etc. may yield some surprises.

To specify widths and heights, you will get better results
by using CSS properties rather than attributes of the 
HTML table model.

In case of difficulties for implementing specific features, 
try to see what MsWord does with that feature when saving
a document in HTML format (plain HTM, not MHT!). 
The generated HTML is quite verbose, but after eliminating
unnecessary tags one can sometimes figure out which are 
the key tags (they start with C<o:>  or C<w:>) or the
key attributes (they start with C<mso->) which correspond
to the desired functionality.

=head2 Collaboration with the Template Toolkit

The L<Template Toolkit|Template> (TT for short) 
is a very helpful tool for generating the HTML.
Below are some hints about collaboration between
the two modules.

=head3 Client code calls both TT and Word::HTML::Writer

The first mode is to use the Template Toolkit for
generating various document parts, and then assemble
them into C<MsOffice::Word::HTML::Writer>.

  use Template;
  my $tmpl_app = Template->new(%options);
  $tmpl_app->process("doctmpl/html_head.tt", \%data, \my $html_head);
  $tmpl_app->process("doctmpl/body.tt",      \%data, \my $body);
  $tmpl_app->process("doctmpl/header.tt",    \%data, \my $header);
  $tmpl_app->process("doctmpl/footer.tt",    \%data, \my $footer);
  
  use MsOffice::Word::HTML::Writer;
  my $doc = MsOffice::Word::HTML::Writer->new(
    title  => $data{title},
    head   => $html_head,
  );
  $doc->create_section(
    header => $header,
    footer => $footer,
  );
  $doc->write($body);
  $doc->save_as("/path/to/some/file");

This architecture is straightforward, but various document parts 
are split into several templates, which might be inconvenient
when maintaining a large body of document templates.

=head3 HTML parts as blocks in a single template

Document parts might also be encoded as blocks within one
single template : 

  [% BLOCK html_head %]
  <style>...CSS...</style>
  [% END; # BLOCK html_head %]
  
  [% BLOCK body %]
    Hello, world
  [% END; # BLOCK body %]
  
  etc.

Then the client code calls each block in turn to gather
the various parts :

  use Template::Context;
  my $tmpl_ctxt = Template::Context->new(%options);
  my $tmpl      = $tmpl_ctxt->template("doctmpl/all_blocks.tt");
  my $html_head = $tmpl_ctxt->process($tmpl->blocks->{html_head}, \%data);
  my $body      = $tmpl_ctxt->process($tmpl->blocks->{body},      \%data);
  my $header    = $tmpl_ctxt->process($tmpl->blocks->{header},    \%data);
  my $footer    = $tmpl_ctxt->process($tmpl->blocks->{footer},    \%data);
  
  # assemble into MsOffice::Word::HTML::Writer, same as before


=head3 Template toolkit calls MsOffice::Word::HTML::Writer

Now let's look at a different architecture: the client code
calls the Template toolkit, which in turn calls
C<MsOffice::Word::HTML::Writer>. 

The most common way to call modules from TT is to use
a I<TT plugin>; but since there is currently 
no TT plugin for C<MsOffice::Word::HTML::Writer>,
we will just tell TT that templates can load regular
Perl modules, by turning on the C<LOAD_PERL> option.

The client code looks like any other TT application; but the output of
the L<process|Template/process> method is a fully-fledged MHT
document, instead of plain HTML.

  use Template;
  my $tmpl_app = Template->new(LOAD_PERL => 1, %other_options);
  $tmpl_app->process("doc_template.tt", \%data, \my $msword_doc);

Within C<doc_template.tt>, we have

  [% # main entry point
  
     # gather various parts
     SET html_head = PROCESS html_head;
     SET header    = PROCESS header;
     SET footer    = PROCESS footer;
     SET body      = PROCESS body;
  
     # create Word::HTML::Writer object
     USE msword = MsOffice.Word.HTML.Writer(head=html_head);
  
     # setup section format
     CALL msword.create_section(
        page => {size          => "21.0cm 29.7cm",
                 margin        => "1cm 2.5cm 1cm 2.5cm",
                 header_margin => "1cm",
                 footer_margin => "0cm",},
        header => header,
        footer => footer
      );
  
      # write the body
     CALL msword.write(body);
  
     # return the MIME-encoded MsWord document
     msword.content();  %]
  
  [% BLOCK html_head %]
  ...

=head3 Inheritance through TT views

The above architecture can be refined one step further,
by using L<TT views|Template::Manual::Views> to 
encapsulate documents. Views have an inheritance mechanism,
so it becomes possible to define families of document
templates, that inherit properties or methods from common
ancestors. Let us start with F<generic_letter.tt2>, 
a generic letter template :

  [% VIEW generic_letter
        title="Generic letter template";
  
       BLOCK main;
         USE msword = MsOffice.Word.HTML.Writer(
            title => view.title,
            head  => view.html_head(),
         );
         view.write_body();
         msword.content();
       END; # BLOCK main
    
       BLOCK write_body;
         CALL msword.create_section(
            page   => {size          => "21.0cm 29.7cm",
                       margin        => "1cm 2.5cm 1cm 2.5cm"},
            header => view.header(),
            footer => view.footer()
         );
         CALL msword.write(view.body());
       END; # BLOCK write_body
    
       BLOCK body;
         view.letter_head();
         view.letter_body();
       END; # BLOCK body
    
       BLOCK letter_body; %]
        Generic letter body; please override BLOCK letter_body in subviews
    [% END; # BLOCK letter_body;
  
       # ... other blocks for header, footer, letter_head, etc.
  
     END; # VIEW generic_letter
  
  [% # call main() method if this templated was loaded directly
     letter.main() UNLESS component.caller %]

This is quite similar to an object-oriented class : assignments
within the view are like object attributes (i.e. the C<title>
variable), and blocks within the view are like methods.

After the end of the view, we call the C<main> method, but 
only if that view was called directly from client code.
If the view is inherited, as displayed below, then the 
call to C<main> will be from the subview.

Now we can define a specific letter template that inherits
from the generic letter and overrides the C<letter_body> block :

  [% PROCESS generic_letter.tt2; # loads the parent view 
  
     VIEW advertisement;
  
       BLOCK letter_body; %]
  
         <p>Dear [% receiver.name %],</p>
         <p>You have won a wonderful [% article %].
            Just call us at [% sender.phone %].</p>
         <p>Best regards,</p>
         [% view.signature(name => sender.name ) %]
  
  [%   END; # BLOCK letter_body
     END; # VIEW advertisement
  
     advertisement.main() UNLESS component.caller %]


=head1 TO DO

Many features could be added; for example:

  - odd/even pages
  - link same header/footers across several sections
  - multiple columns
  - watermarks (I tried hard to reverse engineer MsWord behaviour, 
    but it still doesn't work ... couldn't figure out all details 
    of VML markup)

Contributions welcome!


=head1 AUTHOR

Laurent Dami, C<< <laurent DOT dami AT etat DOT geneve DOT ch> >>

=head1 BUGS

Please report any bugs or feature requests to
C<bug-win32-word-html-writer at rt.cpan.org>, or through the web interface at
L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=MsOffice-Word-HTML-Writer>.
I will be notified, and then you'll automatically be notified of progress on
your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc MsOffice::Word::HTML::Writer

You can also look for information at:

=over 4

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/MsOffice-Word-HTML-Writer>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/MsOffice-Word-HTML-Writer>

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=MsOffice-Word-HTML-Writer>

=item * Search CPAN

L<http://search.cpan.org/dist/MsOffice-Word-HTML-Writer>

=back

=head1 SEE ALSO

L<Win32::Word::Writer>, L<RTF::Writer>, L<Spreadsheet::WriteExcel>,
L<OpenOffice::OODoc>.


=head1 COPYRIGHT & LICENSE

Copyright 2009 Laurent Dami, all rights reserved.

This program is free software; you can redistribute it and/or modify it
under the same terms as Perl itself.

