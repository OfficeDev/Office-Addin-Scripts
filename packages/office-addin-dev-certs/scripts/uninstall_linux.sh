#!/bin/bash

if [ -f "/usr/sbin/update-ca-certificates" ]; then
    sudo rm -r /usr/local/share/ca-certificates/office-addin-dev-certs/$1 && sudo /usr/sbin/update-ca-certificates --fresh
elif [ -f "/usr/sbin/update-ca-trust" ]; then
    sudo rm -r /etc/ca-certificates/trust-source/anchors/office-addin-dev-certs-ca.crt && sudo /usr/sbin/update-ca-trust
fi
