. .\crypto.ps1

# �Í���
Encrypt-Plain -Plain "TestPassword1234" -Path ".\secret_sample.enc"

# ����
$plain = Decrypt-Secure -Path ".\secret_sample.enc"

Write-Host $plain

