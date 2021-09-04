#!/bin/bash

export NETRC='/private/.ssh/netrc'
export PYEXCH_REGEX_JSON='{ 
"IAMSICK" : "(sick|doctor|dr. appt)",
"VACA" : "(vacation|OOTO|OOO|out of the office|out of office|PTO|Paid)" 
}'

python atest.py
