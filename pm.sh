#!/bin/bash
#
# Driver script for results.py. Convert the text output to a PDF

RESULTS=/Users/apetitbianco/Documents/GitHub/sca/results.py
INPUT=/USers/apetitbianco/Downloads/EDITION.txt

tmp_txt=$(mktemp)
tmp_pdf=/tmp/EDITION.pdf
echo "// Text output in $tmp_txt"
echo "// PDF output should be in $tmp_pdf"
tmp_ps=$(mktemp)
python3 $RESULTS $INPUT > $tmp_txt && \
    enscript --header='Classement au %D{%d/%m/%y} %C||Page $% of $=' \
	     --margin=10:10:10:10 -1 -f Courrier8.5 -e --word-wrap --media=A4 \
	     $tmp_txt -o $tmp_ps >/dev/null 2>&1
if [ $? -eq 0 ];
then
    ps2pdf $tmp_ps $tmp_pdf
    open -a Preview $tmp_pdf
else
    echo "** No output generated!"
fi    
