#################################
# Domain Inventory Script v1.2	#
# igroykt (c)29.09.14		#
#################################

###### SETTINGS ######
$MAILBOT=""
$ADMIN=""
$SMTPSERVER=""
$MAILENCODING=[System.Text.Encoding]::UTF8
$SEND_MAIL="1"
$PATH="c:\inetpub\report\"
$INDEX="index.html"
$ERROR_LOG="error.log"

###### DO NOT MODIFY ######
Import-Module ActiveDirectory
$LIST=Get-ADComputer -Filter 'ObjectClass -eq "Computer"'|Select -Expand DNSHostName
$PC_COUNT=([adsisearcher]"(ObjectClass=computer)").FindAll().count
$DATE=Get-Date
$STYLE=@'
<style>
body { background-color:#dddddd;
       font-family:Tahoma;
       font-size:12pt; }
td, th { border:1px solid black; 
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
table { margin-left:50px; }
a:link, a:visited {
color: #004F8B;
text-decoration: underline;
font-weight: bolder;
}
a:hover {
text-decoration: none;
}
</style>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Expires" CONTENT="-1">
'@

function make_clean
{
	Write-Host "==== MAKE CLEAN" -ForeGroundColor Green
	Remove-Item -Path $PATH\*.html -ErrorAction SilentlyContinue
	Remove-Item -Path $PATH\$ERROR_LOG -ErrorAction SilentlyContinue
}

function make_index
{
	Write-Host "==== MAKE INDEX" -ForeGroundColor Green
	$TITLE="<h1>�������� ������������: $DATE</h1>"
	ConvertTo-HTML -head $STYLE -Body "$TITLE"|Out-File "$PATH\$INDEX"
	ConvertTo-HTML -head $STYLE -Body "<strong>���������� ����������� � ������:</strong> $PC_COUNT"|Out-File "$PATH\$INDEX" -Append
	ConvertTo-HTML -head $STYLE -Body "<strong>������:</strong><br>"|Out-File "$PATH\$INDEX" -Append
	Get-ChildItem -Path $PATH\*.html|Select-Object Name|Where-Object {$_.Name -NotMatch "$INDEX"}|
	%{"<a href=$($_.Name) target=_blank>$($_.Name)</a><br>"}|Out-File "$PATH\$INDEX" -Append
}

function make_log
{
	Write-Host "==== MAKE LOG" -ForeGroundColor Green
	New-Item $PATH\$ERROR_LOG -Type File|Out-Null
}

function make_report
{
	Write-Host "==== MAKE REPORT" -ForeGroundColor Green
	foreach($FQDN in $LIST)
	{
		if(test-connection -computername $FQDN -quiet)
		{
			if(Get-Content $PATH\$FQDN.html | Where-Object { $_.Contains("false") -eq "<embed hidden=false></embed>"})
			{
				$ErrorActionPreference="SilentlyContinue"
				Try{
					Write-Host "#### FOUND HOST: $FQDN" -ForeGroundColor Green
					ConvertTo-HTML -head $STYLE -Body "<h2>[ <u>$FQDN</u> ]</h2>"|Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>������������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING USERNAME ]"|
					Get-WmiObject Win32_ComputerSystem -computername $FQDN|
					Select-Object @{expression={$_.username}; label='������������'}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>������������ �������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING OPERATING SYSTEM ]"|
					Get-WmiObject Win32_OperatingSystem -computername $FQDN|
					Select-Object @{expression={$_.CSname}; label='������� ���'}, @{expression={$_.Caption}; label='�����������'}, @{expression={$_.Serialnumber}; label='�������� �����'}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>����������� ���������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING CENTRAL PROCESSOR UNIT ]"|Get-WmiObject CIM_Processor -computername $FQDN|
					Select-Object @{expression={$_.Name}; label='������'}, @{expression={$_.SocketDesignation}; label='�����'}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>����������� �����:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING MOTHERBOARD AND RAM CAPACITY ]"|Get-WmiObject Win32_ComputerSystem -computername $FQDN| 
					Select-Object @{expression={$_.Manufacturer}; label='�������������'}, @{expression={$_.Model}; label='������'}, @{label='��� (��)'; expression={"{0:N0}" -f ($_.TotalPhysicalMemory/1GB)}}, @{expression={$_.SystemType}; label='�����������'}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>���������������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING VIDEOCONTROLLER ]"|Get-WmiObject CIM_VideoController -computername $FQDN|
					Select-Object @{expression={$_.Caption}; label='���������������'}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>�������� ����������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING HARD DRIVE AND CAPACITY ]"|Get-WmiObject CIM_DiskDrive -computername $FQDN|
					Select-Object @{expression={$_.Model}; label='����������'}, @{label='����� (��)'; expression={"{0:N0}" -f ($_.Size/1GB)}}|
					ConvertTo-HTML -head $STYLE -As LIST|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>����������� ������������ ����������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ DETECTING MEMORY TYPE ]"
					$MEM_TYPE = "DDR-3", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM","EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM", "FEPROM","EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM", "DDR-1", "DDR-2"
					$COL1=@{Name='��� ������'; Expression={$MEM_TYPE[$_.MemoryType]}}
					Get-WmiObject Win32_PhysicalMemory -computername $FQDN| Select-Object @{expression={$_.BankLabel}; label='����'},$COL1|
					ConvertTo-HTML -head $STYLE|
					Out-File "$PATH\$FQDN.html" -Append
					
					ConvertTo-HTML -head $STYLE -Body "<strong>������ ������������� ��������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
					Write-Host "[ PERFORMING LIST OF APPLICATIONS ]"|Get-WmiObject Win32_Product -computername $FQDN|Sort-Object Name|
					Select-Object @{expression={$_.Name}; label='������������'}, @{expression={$_.Version}; label='������'}|
					ConvertTo-HTML -head $STYLE|
					Out-File "$PATH\$FQDN.html" -Append
				}
				Catch [System.Runtime.InteropServices.COMException]
				{
					Write-Host "[ RPC SERVER UNAVAILABLE OR UNIX-LIKE OPERATING SYSTEM ]" -ForeGroundColor Yellow
					ConvertTo-HTML -head $STYLE -Body "<h2>[ <u>$FQDN</u> ]</h2>"|Out-File "$PATH\$FQDN.html"
					ConvertTo-HTML -head $STYLE -Body "<strong>������ RPC �� �������� ��� UNIX-like ������������ �������.<br><br>������ �������� � <a href=mailto:$ADMIN>���������� ��������������<a>.</strong>"|
					Out-File "$PATH\$FQDN.html" -Append
				}
			}
			else
			{
				Write-Host "#### SKIP: $FQDN" -ForeGroundColor Yellow
			}
		}
		else
		{
			Write-Host "#### OFFLINE: $FQDN" -ForeGroundColor Red
		}
	}
}

function make_force_report
{
	make_clean
	make_log
	ConvertTo-HTML -head $STYLE -Body "<center><strong><p>���� ��������� ������!</p><p>������� �����...</p></strong></center>"|Out-File "$PATH\$INDEX"
	Write-Host "==== MAKE FORCE REPORT" -ForeGroundColor Green
	foreach($FQDN in $LIST)
	{
		if(test-connection -computername $FQDN -quiet)
		{
			$ErrorActionPreference="Stop"
			Try{
				Write-Host "#### ONLINE: $FQDN" -ForeGroundColor Green
				ConvertTo-HTML -head $STYLE -Body "<h2>[ <u>$FQDN</u> ]</h2>"|Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>������������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING USERNAME ]"|
				Get-WmiObject Win32_ComputerSystem -computername $FQDN|
				Select-Object @{expression={$_.username}; label='������������'}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>������������ �������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING OPERATING SYSTEM ]"|
				Get-WmiObject Win32_OperatingSystem -computername $FQDN|
				Select-Object @{expression={$_.CSname}; label='������� ���'}, @{expression={$_.Caption}; label='�����������'}, @{expression={$_.Serialnumber}; label='�������� �����'}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>����������� ���������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING CENTRAL PROCESSOR UNIT ]"|Get-WmiObject CIM_Processor -computername $FQDN|
				Select-Object @{expression={$_.Name}; label='������'}, @{expression={$_.SocketDesignation}; label='�����'}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>����������� �����:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING MOTHERBOARD AND RAM CAPACITY ]"|Get-WmiObject Win32_ComputerSystem -computername $FQDN| 
				Select-Object @{expression={$_.Manufacturer}; label='�������������'}, @{expression={$_.Model}; label='������'}, @{label='��� (��)'; expression={"{0:N0}" -f ($_.TotalPhysicalMemory/1GB)}}, @{expression={$_.SystemType}; label='�����������'}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>���������������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING VIDEOCONTROLLER ]"|Get-WmiObject CIM_VideoController -computername $FQDN|
				Select-Object @{expression={$_.Caption}; label='���������������'}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>�������� ����������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING HARD DRIVE AND CAPACITY ]"|Get-WmiObject CIM_DiskDrive -computername $FQDN|
				Select-Object @{expression={$_.Model}; label='����������'}, @{label='����� (��)'; expression={"{0:N0}" -f ($_.Size/1GB)}}|
				ConvertTo-HTML -head $STYLE -As LIST|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>����������� ������������ ����������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ DETECTING MEMORY TYPE ]"
				$MEM_TYPE = "DDR-3", "Other", "DRAM", "Synchronous DRAM", "Cache DRAM","EDO", "EDRAM", "VRAM", "SRAM", "RAM", "ROM", "Flash", "EEPROM", "FEPROM","EPROM", "CDRAM", "3DRAM", "SDRAM", "SGRAM", "RDRAM", "DDR-1", "DDR-2"
				$COL1=@{Name='��� ������'; Expression={$MEM_TYPE[$_.MemoryType]}}
				Get-WmiObject Win32_PhysicalMemory -computername $FQDN| Select-Object @{expression={$_.BankLabel}; label='����'},$COL1|
				ConvertTo-HTML -head $STYLE|
				Out-File "$PATH\$FQDN.html" -Append
				
				ConvertTo-HTML -head $STYLE -Body "<strong>������ ������������� ��������:</strong>"|Out-File "$PATH\$FQDN.html" -Append
				Write-Host "[ PERFORMING LIST OF APPLICATIONS ]"|Get-WmiObject Win32_Product -computername $FQDN|Sort-Object Name|
				Select-Object @{expression={$_.Name}; label='������������'}, @{expression={$_.Version}; label='������'}|
				ConvertTo-HTML -head $STYLE|
				Out-File "$PATH\$FQDN.html" -Append
			}
			Catch [System.Runtime.InteropServices.COMException]
			{
				Write-Host "[ RPC SERVER UNAVAILABLE OR UNIX-LIKE OPERATING SYSTEM ]" -ForeGroundColor Yellow
				ConvertTo-HTML -head $STYLE -Body "<h2>[ <u>$FQDN</u> ]</h2>"|Out-File "$PATH\$FQDN.html"
				ConvertTo-HTML -head $STYLE -Body "<strong>������ RPC �� �������� ��� UNIX-like ������������ �������.<br><br>������ �������� � <a href=mailto:$ADMIN>���������� ��������������<a>.</strong>"|
				Out-File "$PATH\$FQDN.html" -Append
			}
			Catch
			{
				Write "Error: $_.Exception" >>$PATH\$ERROR_LOG
				if($SEND_MAIL -eq "1"){Send-MailMessage -From $MAILBOT -To $ADMIN -Subject "������ ��� ��������������" -SmtpServer $SMTPSERVER -Body "$_.Exception" -Encoding $MAILENCODING}
			}
		}
		else
		{
			Write-Host "#### OFFLINE: $FQDN" -ForeGroundColor Red
			ConvertTo-HTML -head $STYLE -Body "<embed hidden=false></embed>"|Out-File "$PATH\$FQDN.html"
			ConvertTo-HTML -head $STYLE -Body "<h2>[ <u>$FQDN</u> ]</h2>"|Out-File "$PATH\$FQDN.html"
			ConvertTo-HTML -head $STYLE -Body "<strong>�� ������� ������������ � ����������.<br><br>��������� �������:<br>-��������<br>-ICMP-������ ����������� firewall'��<br>-��������� ������ COM<br>-��������� �������� �� ���<br>-��������� ����� �� ����������, �� ������������ � ������ ����������� ������<br><br>������ �������� � <a href=mailto:$ADMIN>���������� ��������������<a>.</strong>"|
			Out-File "$PATH\$FQDN.html" -Append
		}
	}
	make_index
}

switch ($args[0])
{
	rotate { make_clean }
	report { make_report }
	force-report { make_force_report }
	default { write-output "Usage: ./dis.ps1 {rotate|report|force-report}" }
}