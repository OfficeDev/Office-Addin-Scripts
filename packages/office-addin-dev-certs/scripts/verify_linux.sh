#!/bin/bash

if [ -f "/usr/sbin/update-ca-certificates" ]; then
    echo [ -f /usr/local/share/ca-certificates/office-addin-dev-certs/$1 ] && openssl x509 -in /usr/local/share/ca-certificates/office-addin-dev-certs/$1 -checkend 86400 -noout
elif [ -f "/usr/sbin/update-ca-trust" ]; then
    echo [ -f /etc/ca-certificates/trust-source/anchors/office-addin-dev-certs-ca.crt ] && openssl x509 -in /etc/ca-certificates/trust-source/anchors/office-addin-dev-certs-ca.crt -checkend 86400 -noout
fi
