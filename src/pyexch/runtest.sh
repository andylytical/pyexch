#!/bin/bash

[[ -f /tmp/firstrun ]] || {
  pip install --upgrade pip
  pip install -r ../../requirements.txt
  touch /tmp/firstrun
  apt update && apt -y install vim less
}

export NETRC='/home/.ssh/netrc'
export PYEXCH_OAUTH_CONFIG='/home/.ssh/exchange_oauth.yaml'
export PYEXCH_TOKEN_FILE='/home/.ssh/exchange_token'
export PYEXCH_REGEX_JSON='{ 
"IAMSICK" : "(sick|doctor|dr. appt)",
"VACA" : "(vacation|OOTO|OOO|out of the office|out of office|PTO|Paid)",
"TRIAGE" : "^Triage",
}'
export PYEXCH_REGEX_JSON='{"TRIAGE":"^Triage: ", "SHIFTCHANGE":"^Triage Shift Change: "}'

python atest.py
