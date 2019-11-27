#!/bin/sh
function download() {
    curl $1 | tar -zxv
    if [ $? -ne 0 ]
    then
        echo "Failed downloading and extracting: $1"
        exit 1
    fi
}

download https://common-lisp.net/project/armedbear/releases/1.6.0/abcl-bin-1.6.0.tar.gz
download http://mirrors.ibiblio.org/apache/poi/release/bin/poi-bin-4.1.1-20191023.tar.gz 
