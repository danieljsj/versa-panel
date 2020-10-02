#!/bin/bash

echo "message: $1"

git pull
git add .
git commit -m "$1"
git push
