#!/bin/bash

if [ -f "/usr/sbin/update-ca-certificates" ]; then
     sudo mkdir -p /usr/local/share/ca-certificates/office-addin-dev-certs && sudo cp $1 /usr/local/share/ca-certificates/office-addin-dev-certs && sudo /usr/sbin/update-ca-certificates
elif [ -f "/usr/sbin/update-ca-trust" ]; then
    sudo cp $1 /etc/ca-certificates/trust-source/anchors/office-addin-dev-certs-ca.crt && sudo /usr/sbin/update-ca-trust
fi
