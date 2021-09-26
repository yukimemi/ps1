
' ˆÃ†‰»
cCode = CreateObject("WScript.Shell").Run("crypto.cmd -Enc -Plain TestPassword1234 -Path .\secret_sample.enc", 1, True)

' •œ†
Set exec = CreateObject("WScript.Shell").Exec("crypto.cmd -Path .\secret_sample.enc")
plain = exec.StdOut.ReadAll
plain = Left(plain, Len(plain) - 2)
WScript.Echo "•œ†Œ‹‰Ê: " & plain


