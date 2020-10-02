#!/bin/bash

git pull
git add .
echo
echo
echo "You entered commit message: $1"
echo
echo
git commit -m "$1"
git push
