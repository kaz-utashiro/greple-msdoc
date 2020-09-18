=head1 NAME

msdoc - Greple module for access MS office docx/pptx/xlsx documents

=head1 VERSION

Version 1.03

=head1 SYNOPSIS

greple -Mmsdoc

=head1 DESCRIPTION

This module makes it possible to search string in Microsoft
docx/pptx/xlsx file.

Microsoft document consists of multiple files archived in zip format.
String information is stored in "word/document.xml",
"ppt/slides/*.xml" or "xl/sharedStrings.xml".  This module extracts
these data and replaces the search target.

By default, text part from XML data is extracted.  This process is
done by very simple method and may include redundant information.

Strings are simply connected into paragraph for I<.docx> and I<.pptx>
document.  For I<.xlsx> document, single space is inserted between
them.  Use B<--separator> option to change this behavior.

After every paragraph, single newline is inserted for I<.pptx> and
I<.xlsx> file, and double newlines for I<.docx> file.  Use
B<--space> option to change.

=head1 OPTIONS

=over 7

=item B<--dump>

Simply print all converted data.  Additional pattern can be specified,
and they will be highlighted inside whole text.

    $ greple -Mmsdoc --dump -e foo -e bar buz.docx

=item B<--space>=I<n>

Specify number of newlines inserted after every paragraph.  Any
non-negative integer is allowed including zero.

=item B<--separator>=I<string>

Specify the separator string placed between each component strings.

=item B<--indent>

Extract indented XML document, not a plain text.

=item B<--indent-mark>=I<string>

Set indentation string.  Default is C<| >.

=back

=head1 INSTALL

cpanm App::Greple::msdoc

=head1 SEE ALSO

L<https://github.com/kaz-utashiro/greple-msdoc>

=head1 AUTHOR

Kazumasa Utashiro

=head1 LICENSE

Copyright 2018-2020 Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

package App::Greple::msdoc;

our $VERSION = '1.03';

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
use Data::Dumper;

our $indent_mark = "| ";
our $opt_space = undef;
our $opt_separator = undef;
our $opt_type;
our $default_format = 'text';

sub separate_xml {
    s{ (?<=>) ([^<]*) }{ $1 ? "\n$1\n" : "\n" }gex;
}

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

use App::optex::textconv::Zip;
use App::optex::textconv::msdoc;
*to_text = \&App::optex::textconv::msdoc::to_text;

my %formatter = (
    'indent-xml'   => \&indent_xml,
    'separate-xml' => \&separate_xml,
    );

sub extract_content {
    my %arg = @_;
    my $file = $arg{&FILELABEL} or die;
    my $pid = open(STDIN, '-|') // croak "process fork failed: $!";
    binmode STDIN, ':encoding(utf8)';
    if ($pid) {
	return $pid;
    }
    my $format = $arg{format} // $default_format;
    if ($format eq 'text') {
	print decode 'utf8', to_text($file);
	exit;
    } elsif ($format =~ /xml$/) {
	my $zip = App::optex::textconv::Zip->new($file);
	my $xml = decode 'utf8', join "\n", map { $zip->extract($_) } $zip->list;
	if (my $sub = $formatter{$format}) {
	    $sub->(&FILELABEL => $file) for $xml;
	}
	print $xml;
	exit;
    }
    die;
}

1;

__DATA__

help	default		ignore
help	--space		Number of newlines after paragraph
help	--separator	Separator between each strings
help	--indent	Indent XML data
help	--indent-mark	Specify text for indentation
help	--type		Specify document type (docx, pptx, xlsx)
help	--dump		Print entire data
help	--msdoc-format	ignore

# --text option is deprecated, and to be removed
option	--text		$<move(0,0)>
help	--text		ignore

option default \
	--if '/\.(docx|pptx|xlsx)$/:&__PACKAGE__::extract_content'

builtin space=i $opt_space
builtin separator=s $opt_separator
builtin type=s $opt_type
builtin msdoc-format=s $default_format

define (#delText) <w:delText>.*?</w:delText>

##
## --indent, --indent-mark
##
option --indent --msdoc-format=indent-xml
builtin indent-mark=s $indent_mark

##
## --dump
##
option --dump --le &sub{} --need 0 --all --epilogue 'sub{exit(0)}'

#  LocalWords:  msdoc Greple greple Mmsdoc docx ppt xml pptx xlsx xl
