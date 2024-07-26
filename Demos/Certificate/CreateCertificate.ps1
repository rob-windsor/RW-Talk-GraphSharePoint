$certname = "M365NYC" 
$cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256

Export-Certificate -Cert $cert -FilePath "$certname.cer"

$mypwd = ConvertTo-SecureString -String "pass@word1" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "$certname.pfx" -Password $mypwd