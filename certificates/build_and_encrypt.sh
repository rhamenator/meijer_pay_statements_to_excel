#!/bin/bash

# Set variables with paths
PROGRAM_NAME="pay_statements_to_excel"
PROJECT_DIRECTORY="/mnt/d/meijer_${PROGRAM_NAME}"
PYTHON_SCRIPT_DIR="${PROJECT_DIRECTORY}"  # Replace with the directory of your .py file
CERT_DIR="${PROJECT_DIRECTORY}/certificates"  # Replace with the directory of your certificates
OUTPUT_DIR="${PROJECT_DIRECTORY}/dist"
PYTHON_SCRIPT="${PROGRAM_NAME}.py"
EXE_NAME="${PROGRAM_NAME}.exe"
CERT_FILE="${PROGRAM_NAME}.pfx"
ENCRYPTED_CERT="${PROGRAM_NAME}.pfx.enc"
BASE64_CERT="${PROGRAM_NAME}.pfx.enc.b64"
GITHUB_REPO="rhamenator/meijer_pay_statements_to_excel"  # Replace with your GitHub repo
SECRET_CERT_NAME="CERT_PFX"
SECRET_PASS_NAME="CERT_PASSWORD"
CERT_PASSWORD="aNd9^^iWB@puToBe"  # Replace with your actual certificate password
SIGNTOOL="/mnt/c/Program Files (x86)/Windows Kits/10/App Certification Kit/signtool.exe"

# Step 1: Navigate to the directory containing the Python script
cd "$PYTHON_SCRIPT_DIR" || { echo "Failed to navigate to $PYTHON_SCRIPT_DIR"; exit 1; }

# Step 2: Compile the Python script into an executable
pyinstaller --onefile "$PYTHON_SCRIPT"

if [ $? -ne 0 ]; then
    echo "An error occurred creating the executable."
    exit 1
fi

# Step 3: Navigate to the directory containing the certificates
cd "$CERT_DIR" || { echo "Failed to navigate to $CERT_DIR"; exit 1; }

# Step 4: Sign the executable
"$SIGNTOOL" sign /fd SHA256 /f "$CERT_FILE" /p "$CERT_PASSWORD" /t http://timestamp.digicert.com /v "$OUTPUT_DIR/$EXE_NAME"

if [ $? -ne 0 ]; then
    echo "An error occurred signing the executable."
    exit 1
fi

# Step 5: Encrypt the certificate
openssl enc -aes-256-cbc -pbkdf2 -in "$CERT_FILE" -out "$ENCRYPTED_CERT" -pass pass:"$CERT_PASSWORD"

if [ $? -ne 0 ]; then
    echo "An error occurred encrypting the certificate."
    exit 1
fi

# Step 6: Encode the encrypted certificate in base64
base64 "$ENCRYPTED_CERT" > "$BASE64_CERT"

if [ $? -ne 0 ]; then
    echo "An error occurred encoding the certificate."
    exit 1
fi

# Step 7: Use GitHub CLI to add the encoded certificate to GitHub Secrets
echo "Adding certificate to GitHub Secrets..."
gh secret set "$SECRET_CERT_NAME" -b"$(cat "$BASE64_CERT")" -R "$GITHUB_REPO"

if [ $? -ne 0 ]; then
    echo "An error occurred adding the certificate to GitHub SECRET_CERT_NAME of GITHUB_REPO."
    exit 1
fi

gh secret set "$SECRET_PASS_NAME" -b"$CERT_PASSWORD" -R "$GITHUB_REPO"

if [ $? -ne 0 ]; then
    echo "An error occurred adding the certificate to GitHub CERT_NAME_PASSWORD of GITHUB_REPO."
    exit 1
fi

# Cleanup
# rm "$ENCRYPTED_CERT"
# rm "$BASE64_CERT"

echo "Script execution completed."
