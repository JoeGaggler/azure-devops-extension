# az login --allow-no-subscriptions
jq -r '.version |= "0.0." + (((split(".")[2] | tonumber) + 1) | tostring)' vss-extension.json > vss-extension.json.tmp
mv vss-extension.json.tmp vss-extension.json
npm ci
npm run build

# HACK: add charset to CSS files
for file in dist/assets/*.css; do
  if ! grep -q '^@charset "UTF-8";' "$file"; then
    echo '@charset "UTF-8";' | cat - "$file" > "$file.tmp" && mv "$file.tmp" "$file"
  fi
done

tfx extension create --json --root . --manifest-globs vss-extension.json --loc-root ../../ --output-path ../../bin/pingmint.vsix
TOKEN=$(az account get-access-token --query accessToken --output tsv)
tfx extension publish --publisher pingmint --vsix ../../bin/pingmint.vsix --auth-type pat -t "$TOKEN"
