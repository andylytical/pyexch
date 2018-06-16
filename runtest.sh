#!/bin/bash

export NETRC='anetrc'
export PYEXCH_REGEX_JSON='{ 
"IAMSICK" : "(sick|doctor|dr. appt)",
"VACA" : "(vacation|OOTO|OOO|out of the office|out of office|PTO|Paid)" 
}'

./env/bin/python atest.py
