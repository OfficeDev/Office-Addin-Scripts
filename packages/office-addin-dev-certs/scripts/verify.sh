#!/bin/bash
certs=$(security find-certificate -a -c $1 -p)
while read line; do
    if [[ "$line" == *"--BEGIN"* ]]; then
        cert=$line
    else
        cert="$cert"$'\n'"$line"
        if [[ "$line" == *"--END"* ]]; then
            # testValue=$(echo "$cert" | openssl x509 -checkend 86400 | grep "will not expire" -c)
            # echo "$testValue"
            if [ 0 -lt $(echo "$cert" | openssl x509 -checkend 86400 | grep -c "will not expire") ]; then
                echo "$cert"
            fi
        fi
    fi
done <<< "$certs"