@echo off

REM Set variables with paths
SET "PROGRAM_NAME=pay_statements_to_excel"
SET "PROJECT_DIRECTORY=D:\meijer_%PROGRAM_NAME%"
SET "PYTHON_SCRIPT_DIR=%PROJECT_DIRECTORY%"  REM Replace with the directory of your .py file
SET "CERT_DIR=%PROJECT_DIRECTORY%\certificates"  REM Replace with the directory of your certificates
SET "OUTPUT_DIR=%PROJECT_DIRECTORY%\dist"
SET "PYTHON_SCRIPT=%PROGRAM_NAME%.py"
SET "EXE_NAME=%PROGRAM_NAME%.exe"
SET "CERT_FILE=%PROGRAM_NAME%.pfx"
SET "ENCRYPTED_CERT=%PROGRAM_NAME%.pfx.enc"
SET "BASE64_CERT=%PROGRAM_NAME%.pfx.enc.b64"
SET "GITHUB_REPO=rhamenator/meijer_pay_statements_to_excel"  REM Replace with your GitHub repo
SET "SECRET_CERT_NAME=CERT_PFX"
SET "SECRET_PASS_NAME=CERT_PASSWORD"
SET "CERT_PASSWORD=aNd9^^iWB@puToBe"
SET "SIGNTOOL=C:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe"

REM Step 1: Navigate to the directory containing the Python script
cd /d %PYTHON_SCRIPT_DIR%

REM Step 2: Compile the Python script into an executable
pyinstaller --onefile %PYTHON_SCRIPT%

if errorlevel 1 (
    echo An error occurred creating the executable.
    exit /b 1
)
REM Step 3: Navigate to the directory containing the certificates
cd /d %CERT_DIR%

REM Step 4: Sign the executable
REM echo %CERT_PASSWORD%

"%SIGNTOOL%" sign /fd SHA256 /f %CERT_FILE% /p %CERT_PASSWORD% /t http://timestamp.digicert.com /v %OUTPUT_DIR%\%EXE_NAME%

if errorlevel 1 (
    echo An error occurred signing the executable.
    exit /b 1
)

REM Step 5: Encrypt the certificate
openssl enc -aes-256-cbc -pbkdf2 -in %CERT_FILE% -out %ENCRYPTED_CERT% -pass pass:%CERT_PASSWORD%

if errorlevel 1 (
    echo An error occurred encrypting the certificate.
    exit /b 1
)
REM Step 6: Encode the encrypted certificate in base64
certutil -encode %ENCRYPTED_CERT% %BASE64_CERT%

if errorlevel 1 (
    echo An error occurred encoding the certificate.
    exit /b 1
)

REM Step 7: Use GitHub CLI to add the encoded certificate to GitHub Secrets
echo Adding certificate to GitHub Secrets...
gh secret set %SECRET_CERT_NAME% -b"%BASE64_CERT%" -R %GITHUB_REPO%

if errorlevel 1 (
    echo An error occurred adding the certificate to GitHub SECRET_CERT_NAME of GITHUB_REPO.
    exit /b 1
)

gh secret set %SECRET_PASS_NAME% -b"%CERT_PASSWORD%" -R %GITHUB_REPO%

if errorlevel 1 (
    echo An error occurred adding the certificate to GitHub CERT_NAME_PASSWORD of GITHUB_REPO.
    exit /b 1
)

REM REM Cleanup
REM del %ENCRYPTED_CERT%
REM del %BASE64_CERT%

echo Script execution completed.
