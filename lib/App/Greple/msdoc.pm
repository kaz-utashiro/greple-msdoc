=head1 NAME

msdoc - Greple module for access MS office documents

=head1 VERSION

Version 0.03

=head1 SYNOPSIS

greple -Mmsdoc

=head1 DESCRIPTION

This module makes it possible to search Microsoft docx/xlsx/pptx file.

Microsoft document consists of multiple files archived in zip format.
Document data is stored in "word/document.xml", "xl/worksheets/*.xml"
or "ppt/slides/*.xml".  This module extracts the content of these
files and replaces the search target data.

=head1 OPTIONS

=over 7

=item B<--indent>

Indent XML document before search.

=item B<--text>

Remove XML markups and extract document text.

=item B<--text-double>

Append double newlines after each sentence.

=item B<--dump>

Simply print all converted data.

=back

=head1 LICENSE

Copyright (C) Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=head1 SEE ALSO

L<https://github.com/kaz-utashiro/greple-msdoc>

=head1 AUTHOR

Kazumasa Utashiro

=cut

package App::Greple::msdoc;

our $VERSION = '0.03';

use strict;
use warnings;
use v5.14;
use Carp;
use utf8;

use Exporter 'import';
our @EXPORT      = ();
our %EXPORT_TAGS = ();
our @EXPORT_OK   = ();

use App::Greple::Common;
use Data::Dumper;

our $newline = 1;

push @EXPORT, '&extract_text';
sub extract_text {
    my @s;
    for (m{<[apw]:p\b[^>]*>(.*?)</[apw]:p>}g) {
	s{<[apw]:(delText)>.*?</[apw]:\1>}{}g;
	s{<(wp:posOffset)>.*?</\1>}{}g;
	s{<(w:drawing)>.*?</\1>}{}sg;
	s{<v:numPr>}{・}g;
	s{</?(?:[apvw]|wp|pic)\d*:.*?>}{}g;
	s{</?(?:mc|wpg|wps|ma14|o):[^>]*>}{}g;
	push @s, $_ if $_ ne "";
    }
    my $separator = "\n" x $newline;
    $_ = join $separator, @s, "" if @s;
}

push @EXPORT, '&indent_xml';
sub indent_xml {
    my %arg = @_;
    my $file = delete $arg{&FILELABEL} or die;

    my %nonewline = map { $_ => 1 } do {
	map  { @{$_->[1]} }
	grep { $file =~ $_->[0] } (
	    [ qr/\.docx$/, [ qw(w:t w:delText w:instrText wp:posOffset) ] ],
	    [ qr/\.pptx$/, [ qw(a:t) ] ],
	    [ qr/\.xlsx$/, [ qw(v f formula1) ] ],
	);
    };

    my $level = 0;
    my $indent_mark = "| ";

    s{
	(?<mark>
	  (?<single>
	    < (?<tag>[\w:]+) [^>]* />
	  )
	  |
	  (?<open>
	    < (?<tag>[\w:]+) [^>]* (?<!/) >
	  )
	  |
	  (?<close>
	    < / (?<tag>[\w:]+) >
	  )
	)
    }{
	if (not $+{single} and $nonewline{$+{tag}}) {
	    join("", $+{open} ? $indent_mark x $level : "",
		 $+{mark},
		 $+{close} ? "\n" : "");
	}
	else {
	    $+{close} and $level--;
	    ($indent_mark x ($+{open} ? $level++ : $level)) . $+{mark} . "\n";
	}
    }gex;
}

1;

__DATA__

option default \
	--if '/\.docx$/:unzip -p /dev/stdin word/document.xml' \
	--if '/\.xlsx$/:unzip -p /dev/stdin xl/worksheets/*.xml' \
	--if '/\.pptx$/:unzip -p /dev/stdin ppt/slides/*.xml'

option --text --begin extract_text
help   --text Extract text

expand --double --begin 'sub{$__PACKAGE__::newline=2}'
option --text-double --double --text
help   --text-double Extract text with double space

define (#delText) <w:delText>.*?</w:delText>
option --indent --begin indent_xml --exclude (#delText)
help   --indent Indent XML data

option --dump --le &sub{} --need 0 --all
help   --dump Print entire data
