#!perl -T

use Test::More tests => 2;

use MsOffice::Word::HTML::Writer;

my $doc = MsOffice::Word::HTML::Writer->new(
  title => "Demo",
  WordDocument => {View => 'Print',
                   Compatibility => {DoNotExpandShiftReturn => ""} },
 );
$doc->write("hello, world");

my $content = $doc->content;

like($content, qr(<w:View>Print</w:View>), "View => print");
like($content, qr(<w:Compatibility><w:DoNotExpandShiftReturn />),
                    "Compatibility => DoNotExpandShiftReturn");
