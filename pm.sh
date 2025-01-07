#!/bin/bash

tmp_txt=$(mktemp)
tmp_pdf=/tmp/EDITION.pdf
echo "// Text output in $tmp_txt"
echo "// PDF output in $tmp_pdf"
tmp_ps=$(mktemp)
python3 /Users/apetitbianco/Documents/GitHub/sca/results.py > $tmp_txt && \
    enscript --header='Classement au %D{%d/%m/%y} %C||Page $% of $=' \
	     --margin=10:10:10:10 -1 -f Courrier8 -e --word-wrap --media=A4 \
	     $tmp_txt -o $tmp_ps && \
    ps2pdf $tmp_ps $tmp_pdf
