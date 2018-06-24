=head1 NAME

msdoc - Greple module for access MS office docx/pptx/xlsx documents

=head1 VERSION

Version 0.04

=head1 SYNOPSIS

greple -Mmsdoc

=head1 DESCRIPTION

This module makes it possible to search string in Microsoft
docx/pptx/xlsx file.

Microsoft document consists of multiple files archived in zip format.
String information is stored in "word/document.xml",
"ppt/slides/*.xml" or "xl/sharedStrings.xml".  This module extracts
these data and replaces the search target.

=head1 OPTIONS

=over 7

=item B<--indent>

Indent XML document before search.

=item B<--indent-mark>=I<string>

Set indentation string.  Default is C<| >.

=item B<--text>

Extract text part from XML data.  This process is done by very simple
method and may include redundant data.

After every paragraph, single newline is inserted for I<.pptx> and
I<.xlsx> file, and double newlines for I<.docx> file.  Use
B<--space> option to change this behavior.

=item B<--space>=I<n>

Specify number of newlines inserted after every paragraph.  Any
non-negative integer is allowed including zero.

=item B<-1>, B<-2>

Shorthand for B<--space> I<1> and I<2>.

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

our $VERSION = '0.04';

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

our $indent_mark = "| ";
our $opt_space = undef;
my $newline;

push @EXPORT, '&extract_text';
sub extract_text {
    my %arg = @_;
    my $file = delete $arg{&FILELABEL} or die;
    $newline = "\n" x do {
	if    (defined $opt_space) { $opt_space }
	elsif ($file =~ /\.docx$/) { 2 }
	else                       { 1 }
    };

    my @xml = grep { length } split /<\?xml\b[^>]*\?>\s*/;
    my @text = map { extract_xml($_) } @xml;
    $_ = join "\n", @text;
}
sub extract_xml {
    local $_ = shift;
    my @p;
    while (m{<(?<tag>[apw]:p|si)\b[^>]*>(?<para>.*?)</\g{tag}>}sg) {
	my $p = $+{para};
	my @s;
	while ($p =~ m{<(?<tag>(?:[apw]:)?t)\b[^>]*>(?<text>[^<]*?)</\g{tag}>}sg) {
	    push @s, $+{text} if $+{text} ne '';
	}
	push @p, join('', @s, $newline) if @s;
    }
    join '', @p;
}

push @EXPORT, '&indent_xml';
sub indent_xml {
    my %arg = @_;
    my $file = delete $arg{&FILELABEL} or die;

    my %nonewline = do {
	map  { $_ => 1 }
	map  { @{$_->[1]} }
	grep { $file =~ $_->[0] } (
	    [ qr/\.docx$/, [ qw(w:t w:delText w:instrText wp:posOffset) ] ],
	    [ qr/\.pptx$/, [ qw(a:t) ] ],
	    [ qr/\.xlsx$/, [ qw(t v f formula1) ] ],
	);
    };

    my $level = 0;

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
	--if '/\.xlsx$/:unzip -p /dev/stdin xl/sharedStrings.xml' \
	--if '/\.pptx$/:unzip -p /dev/stdin ppt/slides/*.xml'

builtin space=i $opt_space
builtin indent-mark=s $indent_mark

option --text --begin extract_text
help   --text Extract text

option -1 --space 1
option -2 --space 2

define (#delText) <w:delText>.*?</w:delText>
option --indent --begin indent_xml --exclude (#delText)
help   --indent Indent XML data

option --dump --le &sub{} --need 0 --all
help   --dump Print entire data

#  LocalWords:  msdoc Greple greple Mmsdoc docx ppt xml pptx xlsx xl
