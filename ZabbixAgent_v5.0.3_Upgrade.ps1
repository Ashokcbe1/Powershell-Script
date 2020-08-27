#To Stop Exsiting Agent Service
stop-service "Zabbix Agent"

#To Delete Exsiting Agent Service
sc.exe delete "Zabbix Agent"

#Removing the Exsiting Agent Folder
Remove-Item –path 'C:\Program Files\Zabbix-Agent' -Recurse

#To delay time between upper and lower commands
Start-Sleep -s 10

#Gets the server host name
$serverHostname =  Invoke-Command -ScriptBlock {hostname}

# Creates Zabbix DIR
mkdir "C:\Program Files\Zabbix-Agent"

# Unzipping file to c:\zabbix
Expand-Archive "\\4D-VM05\Software\zabbix_agent-5.0.3-windows-amd64-openssl.zip" "c:\Program Files\Zabbix-Agent"  


# Replaces 127.0.0.1 with your Zabbix server IP in the config file
(Get-Content -Path "c:\Program Files\Zabbix-Agent\conf\zabbix_agentd.conf") | ForEach-Object {$_ -Replace '127.0.0.1', "172.16.1.76"} | Set-Content -Path "c:\Program Files\Zabbix-Agent\conf\zabbix_agentd.conf"

# Replaces hostname in the config file
(Get-Content -Path "c:\Program Files\Zabbix-Agent\conf\zabbix_agentd.conf") | ForEach-Object {$_ -Replace 'Windows host', "$ServerHostname"} | Set-Content -Path "c:\Program Files\Zabbix-Agent\conf\zabbix_agentd.conf"

# Attempts to install the agent with the config in c:\zabbix
& 'C:\Program Files\Zabbix-Agent\bin\zabbix_agentd.exe' --config 'C:\Program Files\Zabbix-Agent\conf\zabbix_agentd.conf' --install

# Attempts to start the agent
Start-Service "Zabbix Agent"
