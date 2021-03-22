# A script to monitor the status of Vizrt's Viz graphics rendering engine
# Supported Viz commands can be found at https://documentation.vizrt.com/viz-engine-guide-3.14.pdf
# To run this script in Windows, %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NoProfile -file thisFile.ps1
# By Jun Ye - 18 March 2021
# Email: jellun@hotmail.com

#cmdkey /generic:"ip address" /user:"domain\xxx" /pass:"xxxxxxxx"
$credential=(get-credential domain\xxx)

function Send-Email {
	param(
		[string]$toAddress = "jellun@hotmail.com", 
		[string]$fromAddress = "Default",
		[string]$subject = "Automated Email", 
		[string]$attachment = "Default", 
		[string]$emailbody = "Default"
		)

	#set basic parameters
	$smtpServer = "xxx.net"
	$defaultDomain = "nle.xxx.net"
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)

	#apply the from address
	if($fromAddress -eq "Default"){
		$fromAddress = "{0}@{1}" -f $env:computername, $defaultDomain
	}
	$msg.From = $fromAddress

	#apply each to address
	foreach($address in $toAddress -split ",") {
		$msg.To.Add($address)
	}
	$msg.Subject = $subject

	#check if the attachment is valid
	if (($attachment -ne "Default") -and (Test-Path $attachment)){
		$att = new-object Net.Mail.Attachment($attachment)
		$msg.Attachments.Add($att)
		#update the body text
		if($emailbody -eq "Default"){
			$emailbody = "Attached is the file {0}" -f $attachment
		}
		$msg.Body = $emailbody 
		$smtp.Send($msg)
		$att.Dispose()
	} else {
		#update the body text
		if($emailbody -eq "Default"){
			$emailbody = "Automated email: no body specified by sender"
		}
		$msg.Body = $emailbody
		$smtp.Send($msg)
	}
}

[int] $port = 6100
$engine = "127.0.0.1" # Hostname or IP address
#$engine = "localhost" # Hostname or IP address: [127.0.0.1] [localhost] [MachineName] [MachineName.DomainName]

$saddrf = [System.Net.Sockets.AddressFamily]::InterNetwork
$stype = [System.Net.Sockets.SocketType]::Stream
$ptype = [System.Net.Sockets.ProtocolType]::TCP
$script:sock = New-Object System.Net.Sockets.Socket $saddrf, $stype, $ptype
$script:sock.TTL = 26

$Enc = [System.Text.Encoding]::ASCII

if ($engine.Split('.').Length -ge 4) # An IP address
{
	$script:IPaddr = [system.net.IPAddress]::Parse($engine)
	$script:endPoint = New-Object System.Net.IPEndPoint $script:IPaddr, $port
	try
	{
		$script:sock.Connect($script:endPoint)
		if ($script:sock.Connected)
		{
			$script:endPoint
		}
	}
	catch
	{
		$Error
		exit
	}
}
else # A Hostname
{
	$script:hostEntry = [System.Net.Dns]::GetHostEntry($engine)
	foreach ($aHostAddr in $script:hostEntry.AddressList)
    {
        $script:endPoint = New-Object System.Net.IPEndPoint $aHostAddr, $port
        if ($script:endPoint.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
        {
			try
			{
				$script:sock.Connect($script:endPoint)
				if ($script:sock.Connected)
				{
					$script:endPoint
					break
				}
			}
			catch
			{
				$Error
				exit
			}
        }
    }
}

if (!($script:sock.Connected))
{
    "Cannot connect to the engine! Exit!"
    exit
}

try
{
	$Commands = "0 DEBUG_CONTROL*RENDERINFO*CONSOLE GET`0"
    $BufferSend = $Enc.GetBytes($Commands)
    "Sending command: {$Commands} to Viz Engine"
    $SentBytes = $script:sock.Send($BufferSend)
    "{0} characters sent to: {1} " -f $SentBytes,$engine
    [byte[]]$BufferRevd = New-Object Byte[] 33554432
    $ReceivedBytes = $script:sock.Receive($BufferRevd)
    $ReceivedMsg = $Enc.GetString($BufferRevd, 0, $ReceivedBytes)
    "Received message: {$ReceivedMsg} from Viz Engine"
    "Received message length: $($ReceivedMsg.Length*2)"
    if ($ReceivedMsg.Contains("get suitable video decompressor"))
    {
        Send-Email "email address 1,email address 2" "Default" "Alert - video decompressor error occured" "Default" "OMG! Human intervention is required! You better restart me now!"
    }

	(Get-Process | where {$_.MainWindowTitle -eq "X64 Viz Engine [1]"}) | kill -Force
	Start-Sleep -Seconds 5
	Start-Process -FilePath "C:\Users\Public\Desktop\Viz Engine 3.14.4 #1 x64.lnk"
}
catch
{
    $Error
}

try
{
    $script:sock.Close()
    $script:sock.Dispose()
}
catch
{
    $Error
}
