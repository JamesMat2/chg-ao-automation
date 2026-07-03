#!/bin/bash
# Called by n8n with: bash /root/run_pipeline.sh FILE_ID FILE_NAME
FILE_ID=$1
FILE_NAME=$2
python3 /root/pipeline.py "$FILE_ID" "$FILE_NAME"
