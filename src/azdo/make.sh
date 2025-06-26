jq -r '.version |= "0.0." + (((split(".")[2] | tonumber) + 1) | tostring)' vss-extension.json > vss-extension.json.tmp
mv vss-extension.json.tmp vss-extension.json
npm ci
npm run build
tfx extension create --json --root . --manifest-globs vss-extension.json --loc-root ../../ --output-path ../../bin/pingmint.vsix
