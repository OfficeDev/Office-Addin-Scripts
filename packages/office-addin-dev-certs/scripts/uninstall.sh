#!/bin/bash
hashes=$(security find-certificate -c "$1" -a -Z | grep SHA-1 | awk '{ print $NF }')
for hash in $hashes; do
    security delete-certificate -Z $hash
done