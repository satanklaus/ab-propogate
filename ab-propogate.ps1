$profile = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\"
$cit_ldaptype = "f209613c13644332b2183613005fb209\"
$cit_ldaptype_hex = [byte[]](0xF2, 0x09, 0x61, 0x3C, 0x13, 0x64, 0x43, 0x32, 0xB2, 0x18, 0x36, 0x13, 0x00, 0x5F, 0xB2, 0x09)
$cit_ldapconfig = "c494fa4a2cf941859f2a7b45fe416326\"
$cit_ldapconfig_hex = [byte[]](0xC4,0x94,0xFA,0x4A,0x2C,0xF9,0x41,0x85,0x9F,0x2A,0x7B,0x45,0xFE,0x41,0x63,0x26)
$backup = "9207f3e0a3b11019908b08002b2a56c2\"
$LDAPdisplayname = "Адресная книга РКОМИ"
$LDAPserver = "192.168.1.60"
$LDAPport = "389"
$LDAPsearchbase = "ou=resources,dc=rk,dc=local"
$ablist = "9375CFF0413111d3B88A00104B2A6676"
$abkey = "{ED475419-B0D6-11D2-8C3B-00104B2A6676}"

$backup_prop = $(Get-ItemProperty -path $profile$backup -Name "01023d01")."01023d01"
$backup_sets = $(Get-ItemProperty -path $profile$backup -Name "01023d0e")."01023d0e"
$count_prop = [math]::floor($(Get-ItemProperty -path $profile$backup -Name "01023d01")."01023d01".Length/16)
$count_sets = [math]::floor($(Get-ItemProperty -path $profile$backup -Name "01023d0e")."01023d0e".Length/16)

$cit_ldaptype_present = $false
for ($i = 0; $i -lt $count_prop; $i += 1)
{ 
	if ("$($backup_prop[(16*$i)..(16*($i+1)-1)])" -like "$cit_ldaptype_hex") 
	{
		$cit_ldaptype_present = $true
		echo "found type"
		break
	}
}
$cit_ldapconfig_present = $false
for ($i = 0; $i -lt $count_sets; $i += 1)
{ 
	if ("$($backup_sets[(16*$i)..(16*($i+1)-1)])" -like "$cit_ldapconfig_hex") 
	{
		$cit_ldapconfig_present = $true
		break
	}
}

New-Item -Path $profile$cit_ldaptype -Force
Set-ItemProperty -Path $profile$cit_ldaptype -Name '00033009' -Value ([byte[]](0x0,0x0,0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldaptype -Name '00033e03' -Value ([byte[]](0x23,0x0,0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldaptype -Name '001e3001' -Value "Microsoft LDAP Directory"
Set-ItemProperty -Path $profile$cit_ldaptype -Name '001e3006' -Value "Microsoft LDAP Directory"
Set-ItemProperty -Path $profile$cit_ldaptype -Name '001e300a' -Value "EMABLT.DLL"
Set-ItemProperty -Path $profile$cit_ldaptype -Name '001e3d09' -Value "EMABLT"
Set-ItemProperty -Path $profile$cit_ldaptype -Name '001e3d13' -Value "{6485D268-C2AC-11D1-AD3E-10A0C911C9C0}"
Set-ItemProperty -Path $profile$cit_ldaptype -Name '01023d0c' -Value $cit_ldapconfig_hex

New-Item -Path $profile$cit_ldapconfig -Force
Set-ItemProperty -Path $profile$cit_ldapconfig -Name '00033009' -Value ([byte[]](0x20,0x0,0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "00036623" -Value ([byte[]](0x00,0x0,0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "000b6613" -Value ([byte[]](0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "000b6615" -Value ([byte[]](0x00,0x0))
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "000b6622" -Value ([byte[]](0x0,0x0))
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e3001" -Value $LDAPdisplayname
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e3d09" -Value "EMABLT"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e3d0a" -Value "BJABLR.DLL"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e3d0b" -Value "ServiceEntry"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e3d13" -Value "{6485D268-C2AC-11D1-AD3E-10A0C911C9C0}"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6600" -Value $LDAPserver
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6601" -Value $LDAPport
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6602" -Value ""
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6603" -Value $LDAPsearchbase
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6604" -Value "(&(mail=*)(|(mail=%s*)(|(cn=%s*)(|(sn=%s*)(givenName=%s*)))))"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6605" -Value "SMTP"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6606" -Value "mail"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6607" -Value "60"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6608" -Value "100"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6609" -Value "120"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660a" -Value "15"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660b" -Value ""
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660c" -Value "OFF"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660d" -Value "OFF"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660e" -Value "NONE"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e660f" -Value "OFF"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6610" -Value "postalAddress"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6611" -Value "cn"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "001e6612" -Value "1"
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "01023615" -Value ([byte[]](0x50, 0xA7, 0x0A, 0x61, 0x55, 0xDE, 0xD3, 0x11, 0x9D, 0x60, 0x0, 0xC0, 0x4F, 0x4C, 0x8E, 0xFA)) 
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "01023d01" -Value $cit_ldaptype_hex

Set-ItemProperty -Path $profile$cit_ldapconfig -Name "01026631" -Value ([byte[]](0x98, 0x17, 0x82, 0x92, 0x5B, 0x43, 0x03, 0x4B, 0x99, 0x5D, 0x5C, 0xC6, 0x74, 0x88, 0x7B, 0x34 )) 
Set-ItemProperty -Path $profile$cit_ldapconfig -Name "101e3d0f" -Value ([byte[]](0x2, 0x0,0x0,0x0,0xC,0x0,0x0,0x0,0x17,0x0,0x0,0x0,0x45,0x4D,0x41,0x42,0x4C,0x54,0x2E,0x44,0x4C,0x0,0x42,0x4A,0x41,0x42,0x4C,0x52,0x2E,0x44,0xC,0x4C,0x0))

New-ItemProperty -Path $profile$cit_ldapconfig -Name "01026617" -PropertyType Binary -Value ([byte[]]@()) -Force
New-ItemProperty -Path $profile$cit_ldapconfig -Name "S001e67f1" -PropertyType Binary -Value ([byte[]]@()) -Force

if (-not $cit_ldaptype_present)  { Set-ItemProperty -Path $profile$backup -Name "01023d01" -Value $($backup_prop + $cit_ldaptype_hex) }
if (-not $cit_ldapconfig_present){ Set-ItemProperty -Path $profile$backup -Name "01023d0e" -Value $($backup_sets + $cit_ldapconfig_hex)}

Remove-ItemProperty -Path $profile$ablist -Name $abkey
