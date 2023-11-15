#Early Inspiration, Replaced with Python


#!/bin/bash
EDITOR=nvim
BASEPATH=~/Documents/CV/cover_letters

STARTDIR=$PWD
JOB_TITLE=$1
COMPANY_NAME=$2
REQUISITION_ID=$3
SUBJECT="$COMPANY_NAME - $JOB_TITLE - $REQUISITION_ID"

cd $BASEPATH
cp generic.md src/$JOB_TITLE.md
cd src
sed -i "s/JOBTITLE/$JOB_TITLE/g" $JOB_TITLE.md
sed -i "s/COMPANYNAME/$COMPANY_NAME/g" $JOB_TITLE.md
sed -i "s/REQUISITIONID/$REQUISITION_ID/g" $JOB_TITLE.md
$EDITOR $JOB_TITLE.md
xclip -selection clipboard $JOB_TITLE.md
pandoc $JOB_TITLE.md -o $SUBJECT.docx
soffice --convert-to pdf $SUBJECT.docx --outdir ~/Documents/CV/cover_letters 2> /dev/null

cd $STARTDIR
