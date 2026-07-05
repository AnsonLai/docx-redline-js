#!/usr/bin/env bash
# Validates generated validation fixtures against the ECMA-376 transitional
# wordprocessingml XSD using xmllint. Used by the nightly validation workflow
# and runnable locally on any machine with curl, unzip, and xmllint.
#
# Usage: scripts/validate-fixtures-xsd.sh [fixtures-dir]
set -euo pipefail

FIXTURES_DIR="${1:-tmp/validation-docx}"
CACHE_DIR="${OOXML_SCHEMA_DIR:-.cache/ooxml-schemas}"
ECMA_ZIP_URL="https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip"
XML_XSD_URL="https://www.w3.org/2001/xml.xsd"

command -v xmllint >/dev/null || { echo "xmllint not found (install libxml2-utils)"; exit 1; }

shopt -s nullglob
fixtures=("$FIXTURES_DIR"/*.document.xml)
if [ ${#fixtures[@]} -eq 0 ]; then
  echo "No *.document.xml fixtures in $FIXTURES_DIR — run: node scripts/export-validation-fixtures.mjs"
  exit 1
fi

if [ ! -f "$CACHE_DIR/wml.xsd" ]; then
  echo "Downloading ECMA-376 Part 4 transitional schemas..."
  mkdir -p "$CACHE_DIR"
  curl -sSL --retry 3 -o "$CACHE_DIR/ecma376-4.zip" "$ECMA_ZIP_URL"
  unzip -o -q "$CACHE_DIR/ecma376-4.zip" "OfficeOpenXML-XMLSchema-Transitional.zip" -d "$CACHE_DIR"
  unzip -o -q "$CACHE_DIR/OfficeOpenXML-XMLSchema-Transitional.zip" -d "$CACHE_DIR"
  curl -sSL --retry 3 -o "$CACHE_DIR/xml.xsd" "$XML_XSD_URL"
  # ECMA's published XSDs import the xml namespace without a schemaLocation;
  # point them at the local copy so xmllint can compile offline.
  sed -i.bak 's|<xsd:import namespace="http://www.w3.org/XML/1998/namespace"/>|<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>|' "$CACHE_DIR"/*.xsd
  rm -f "$CACHE_DIR"/*.xsd.bak "$CACHE_DIR/ecma376-4.zip"
fi

xmllint --noout --schema "$CACHE_DIR/wml.xsd" "${fixtures[@]}"
echo "XSD validation passed for ${#fixtures[@]} fixture(s)."
