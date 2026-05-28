#!/bin/bash

find ./src -type f \( -name "*.ts" -o -name "*.vue" -o -name "*.js" \) -exec wc -l {} \; | sort -rn | head -20