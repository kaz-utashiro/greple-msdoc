# NAME

msdoc - Greple module for access MS office docx/pptx/xlsx documents

# VERSION

Version 0.04

# SYNOPSIS

greple -Mmsdoc

# DESCRIPTION

This module makes it possible to search string in Microsoft
docx/pptx/xlsx file.

Microsoft document consists of multiple files archived in zip format.
String information is stored in "word/document.xml",
"ppt/slides/\*.xml" or "xl/sharedStrings.xml".  This module extracts
these data and replaces the search target.

# OPTIONS

- **--indent**

    Indent XML document before search.

- **--indent-mark**=_string_

    Set indentation string.  Default is `| `.

- **--text**

    Extract text part from XML data.  This process is done by very simple
    method and may include redundant data.

    After every paragraph, single newline is inserted for _.pptx_ and
    _.xlsx_ file, and double newlines for _.docx_ file.  Use
    **--space** option to change this behavior.

- **--space**=_n_

    Specify number of newlines inserted after every paragraph.  Any
    non-negative integer is allowed including zero.

- **-1**, **-2**

    Shorthand for **--space** _1_ and _2_.

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
