# XPlus

## Description

[XPlus](http://x-vba.com/xplus) is an Excel Function library that adds over 100 functions to Excel, making it much
easier to do common tasks within Excel. XPlus is written purely in VBA and has no 
external dependencies, making it very easy to install and incorporate into an Excel
Workbook. Additionally, since XPlus is MIT Licensed, and thus has very little restrictions
on use, it can be used commercially and personally free and charge, and can be shipped
to others within an Excel Workbook free of charge. This means that even if another user
doesn't have XPlus installed, they can still use XPlus when you send your workbook to
them. 

## Documentation

Official documentation for XPlus can be found here. The official documentation gives
in-depth explanations and examples of the functions and XPlus, and the function list
below should be used primarily as a reference only.

## Function List

Below is a list of all functions within XPlus:

 * Validators
   - IS\_EMAIL
   - IS\_PHONE
   - IS\_CREDIT\_CARD
   - IS\_URL
   - IS\_IP\_FOUR
   - CREDIT\_CARD\_NAME
   - FORMAT\_FRACTION
   - FORMAT\_PHONE
   - FORMAT\_CREDIT\_CARD

 * Utilities
   - DISPLAY\_TEXT
   - JSONIFY
   - UUID\_FOUR
   - HIDDEN
   - ISERRORALL
   - JAVASCRIPT
   - SUMSHEET
   - AVERAGESHEET
   - MAXSHEET
   - MINSHEET
   - HTML\_TABLEIFY
   - HTML\_ESCAPE
   - HTML\_UNESCAPE
   - SPEAK\_TEXT

 * StringMetrics
   - HAMMING
   - LEVENSHTEIN
   - LEV\_STR
   - LEV\_STR\_OPT
   - DAMERAU
   - DAM\_STR
   - DAM\_STR\_OPT
   - PARTIAL\_LOOKUP

 * StringManipulation
   - CAPITALIZE
   - LEFT\_FIND
   - RIGHT\_FIND
   - LEFT\_SEARCH
   - RIGHT\_SEARCH
   - SUBSTR
   - SUBSTR\_SEARCH
   - REPEAT
   - FORMATTER
   - ZFILL
   - TEXT\_JOIN
   - SPLIT\_TEXT
   - COUNT\_WORDS
   - CAMEL\_CASE
   - KEBAB\_CASE
   - REMOVE\_CHARACTERS
   - COMPANY\_CASE

 * Regex
   - REGEX\_SEARCH
   - REGEX\_TEST
   - REGEX\_REPLACE

 * Range
   - FIRST\_UNIQUE
   - SORT\_RANGE
   - REVERSE\_RANGE
   - COLUMNIFY
   - ROWIFY
   - SUMN
   - AVERAGEN
   - MAXN
   - MINN
   - SUMHIGH
   - SUMLOW
   - AVERAGEHIGH
   - AVERAGELOW
   - INRANGE

 * Random
   - RANDOM\_SAMPLE
   - RANDOM\_RANGE
   - RANDOM\_SAMPLE\_PERCENT

 * Properties
   - RANGE\_COMMENT
   - RANGE\_HYPERLINK
   - RANGE\_NUMBER\_FORMAT
   - RANGE\_FONT
   - RANGE\_NAME
   - RANGE\_WIDTH
   - RANGE\_HEIGHT
   - SHEET\_NAME
   - SHEET\_CODE\_NAME
   - SHEET\_TYPE
   - WORKBOOK\_TITLE
   - WORKBOOK\_SUBJECT
   - WORKBOOK\_AUTHOR
   - WORKBOOK\_MANAGER
   - WORKBOOK\_COMPANY
   - WORKBOOK\_CATEGORY
   - WORKBOOK\_KEYWORDS
   - WORKBOOK\_COMMENTS
   - WORKBOOK\_HYPERLINK\_BASE
   - WORKBOOK\_REVISION\_NUMBER
   - WORKBOOK\_CREATION\_DATE
   - WORKBOOK\_LAST\_SAVE\_TIME
   - WORKBOOK\_LAST\_AUTHOR

 * Network
   - USER\_NAME
   - USER\_DOMAIN
   - COMPUTER\_NAME

 * Math
   - INTERPOLATE\_NUMBER
   - INTERPOLATE\_PERCENT

 * File
   - FILE\_CREATION\_TIME
   - FILE\_LAST\_MODIFIED\_TIME
   - FILE\_DRIVE
   - FILE\_NAME
   - FILE\_FOLDER
   - FILE\_PATH
   - FILE\_SIZE
   - FILE\_TYPE
   - FILE\_EXTENSION
   - READ\_FILE
   - WRITE\_FILE
   - PATH\_JOIN

 * Environment
   - ENVIRONMENT

 * DateTime
   - WEEKDAY\_NAME
   - MONTH\_NAME
   - QUARTER
   - TIME\_CONVERTER
   - DAYS\_OF\_MONTH

 * Color
   - RGB2HEX
   - HEX2RGB
   - RGB2HSL
   - HEX2HSL
   - HSL2RGB
   - HSL2HEX
   - RGB2HSV

## License

The MIT License (MIT)

Copyright © 2020 Anthony Mancini

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
