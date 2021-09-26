. .\crypto.ps1

# ˆÃ†‰»
Encrypt-Plain -Plain "TestPassword1234" -Path ".\secret_sample.enc"

# •œ†
$plain = Decrypt-Secure -Path ".\secret_sample.enc"

Write-Host $plain

