#!/bin/bash

export PYEXCH_USER='aloftus'
export PYEXCH_AD_DOMAIN='UOFI'
export PYEXCH_EMAIL_DOMAIN='illinois.edu'
export PYEXCH_PWD_FILE=$HOME/.ssh/imap_illinois_edu

python atest.py
