#!/bin/bash

PRODUCT_VERSION="5.2.0"
BUILD_NUMBER="111"

echo "----------------------------------------"
echo "Building for mobile"
echo "----------------------------------------"

BASEDIR="$(cd "$(dirname "$0")" && pwd)"
cd $BASEDIR

if [ -z "$1" ]
  then
    echo "You must set argument with path for sdkjs copy..."
    exit 1
fi

if [[ -d "$1" ]]
	then
		echo "Folder already exist: $1"
else
	mkdir "$1"
fi

mkdir "$1/documents"
mkdir "$1/spreadsheets"
mkdir "$1/presentations"

npm install

PRODUCT_VERSION=$PRODUCT_VERSION BUILD_NUMBER=$BUILD_NUMBER grunt --level=WHITESPACE_ONLY --mobile=true --noclosure=true --formatting=PRETTY_PRINT
echo grunt --level=ADVANCED --mobile=true  --noclosure=true

printf $'\r' > temp.txt

cat "../../web-apps-pro/vendor/xregexp/xregexp-all-min.js" "temp.txt" "../../web-apps/vendor/underscore/underscore-min.js" "temp.txt" "../common/Native/Wrappers/common.js" "temp.txt" "../common/Native/jquery_native.js" "temp.txt" > "banners.js"

cat "banners.js" "../word/sdk-all-min.js" "../word/sdk-all.js" > "$1/documents/script.bin"

rm -f -r "banners.js"

cat "../../web-apps-pro/vendor/xregexp/xregexp-all-min.js" "temp.txt" "../../web-apps/vendor/underscore/underscore-min.js" "temp.txt" "../cell/native/common.js" "temp.txt" "../common/Native/jquery_native.js" "temp.txt" > "banners.js"

cat "banners.js" "../cell/sdk-all-min.js" "../cell/sdk-all.js" > "$1/spreadsheets/script.bin"

rm -f -r "banners.js"

cat "../../web-apps-pro/vendor/xregexp/xregexp-all-min.js" "temp.txt" "../../web-apps/vendor/underscore/underscore-min.js" "temp.txt" "../common/Native/Wrappers/common.js" "temp.txt" "../common/Native/jquery_native.js" "temp.txt" > "banners.js"

cat "banners.js" "../slide/sdk-all-min.js" "../slide/sdk-all.js" > "$1/presentations/script.bin"

rm -f -r "banners.js"
rm -f -r "temp.txt"

printf $PRODUCT_VERSION.$BUILD_NUMBER > "$1/documents/sdk.version"

printf $PRODUCT_VERSION.$BUILD_NUMBER > "$1/spreadsheets/sdk.version"

printf $PRODUCT_VERSION.$BUILD_NUMBER > "$1/presentations/sdk.version"
