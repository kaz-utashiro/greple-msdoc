=head1 NAME

msdoc - Greple module for access MS office documents

=head1 VERSION

Version 0.02

=head1 SYNOPSIS

greple -Mmsdoc

=head1 DESCRIPTION

This module makes it possible to search Microsoft docx/xlsx/pptx file.

Microsoft document is consists of multiple files archived in zip
format.  Document data is stored in "word/document.xml",
"xl/worksheets/*.xml" or "ppt/slides/*.xml".  This module extracts the
content of these files and replaces the search target data.

=head1 OPTIONS

=over 7

=item B<--indent>

Indent XML document before search.

=item B<--text>

Remove XML markups and extract document text.

=item B<--text-double>

Insert double space between sentence in text format.

=item B<--dump>

Simply print all converted data.

=back

=head1 LICENSE

Copyright (C) Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=head1 AUTHOR

Kazumasa Utashiro

=cut

package App::Greple::msdoc;

our $VERSION = '0.02';

use strict;
use warnings;

use v5.14;

use Carp;

use utf8;
use Encode;

use Exporter 'import';
our @EXPORT      = ();
our %EXPORT_TAGS = ();
our @EXPORT_OK   = ();

use App::Greple::Common;
use List::Util qw( min max first sum );
use Data::Dumper;

our $document = "word/document.xml";
my $separator = "\n";

push @EXPORT, '&double_space';
sub double_space {
    $separator = "\n\n";
}

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
    $_ = join($separator, @s) . "\n" if @s;
}

push @EXPORT, '&indent_xml';
sub indent_xml {
    my %arg = @_;
    my $file = delete $arg{&FILELABEL} or die;

    my %nonewline = map { $_ => 1 } do {
	if ($file =~ /\.docx$/) {
	    qw(w:t w:delText w:instrText wp:posOffset);
	}
	elsif ($file =~ /\.pptx$/) {
	    qw(a:t);
	}
	elsif ($file =~ /\.xlsx$/) {
	    qw(v f formula1);
	}
	else {
	    return;
	}
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
	my $s = "";
	if (not $+{single} and $nonewline{$+{tag}}) {
	    $s = join("", $+{open} ? $indent_mark x $level : "",
			  $+{mark},
			  $+{close} ? "\n" : "");
	}
	else {
	    $+{close} and --$level;
	    $s = ($indent_mark x $level) . $+{mark} . "\n";
	    $+{open}  and ++$level;
	}
	$s;
    }gex;
}

1;

__DATA__

option default \
	--if '/\.docx$/:unzip -p /dev/stdin word/document.xml' \
	--if '/\.xlsx$/:unzip -p /dev/stdin xl/worksheets/*.xml' \
	--if '/\.pptx$/:unzip -p /dev/stdin ppt/slides/*.xml'

option --text --begin extract_text

option --double --begin double_space

option --text-double --double --text

define (#delText) <w:delText>.*?</w:delText>
option --indent --begin indent_xml --exclude (#delText)

option --xls --begin xlsx_xml --exclude (#delText)

option --dump -e '(?=never)__match' --need 0 --all
