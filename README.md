# NAME

msdoc - Greple module for access MS office documents

# VERSION

Version 0.03

# SYNOPSIS

greple -Mmsdoc

# DESCRIPTION

This module makes it possible to search Microsoft docx/xlsx/pptx file.

Microsoft document consists of multiple files archived in zip format.
Document data is stored in "word/document.xml", "xl/worksheets/\*.xml"
or "ppt/slides/\*.xml".  This module extracts the content of these
files and replaces the search target data.

# OPTIONS

- **--indent**

    Indent XML document before search.

- **--text**

    Remove XML markups and extract document text.

- **--text-double**

    Append double newlines after each sentence.

- **--dump**

    Simply print all converted data.

# LICENSE

Copyright (C) Kazumasa Utashiro.

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

# SEE ALSO

[https://github.com/kaz-utashiro/greple-msdoc](https://github.com/kaz-utashiro/greple-msdoc)

# AUTHOR

Kazumasa Utashiro
