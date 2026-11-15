#!/bin/bash
# Regenerates a member tracking catalogue: gen-mbm-all.sh <form> <catalogue file name>
set -e
cd /c/Users/Naresh/AppData/Local/Temp
export TRK_FORM="$1" TRK_CAT="$2"
node analyse-mbmtrk.js > /dev/null
node analyse-mbmtrk-headers.js > /dev/null
node gen-mbmtrk-catalogue.js
node gen-mbmtrk-spec.js
node analyse-mbmtrk-detail.js > /dev/null
node gen-mbmtrk-detail-spec.js | head -1
