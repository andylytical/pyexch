#!/bin/bash

export PYEXCH_AD_DOMAIN='UOFI'
export PYEXCH_EMAIL_DOMAIN='illinois.edu'
export PYEXCH_PWD_FILE=$HOME/.ssh/imap_illinois_edu

ENV/bin/python atest.py
