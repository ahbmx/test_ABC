
CERT_HASH=$(openssl x509 -noout -modulus -in server.crt | openssl md5)

find . -type f \( -name "*.key" -o -name "*.pem" \) 2>/dev/null | while read file; do
  KEY_HASH=$(openssl rsa -noout -modulus -in "$file" 2>/dev/null | openssl md5)
  if [ "$CERT_HASH" = "$KEY_HASH" ]; then
    echo "MATCH FOUND: $file"
  fi
done
