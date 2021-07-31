##################################################################################################

$computer_list = Import-Csv -Path "C:\computerList.csv" -Header computer
$password = "123456?Ab" | ConvertTo-SecureString -asPlainText -Force;       
$username = "gameloft\truong_tranduy"                                              
$credential = New-object System.Management.Automation.PSCredential($username,$password)

##################################################################################################

function Execute_regedit{
    param($computername,$credential)

    foreach ($pc in $computername) {
        
        #Test connection first, if true return computer is up
        $test_connection = Test-Connection -ComputerName $pc.computer -Count 2 -Quiet

        if ($test_connection){
            
            #Check the username and password before running.
            try 
            {
                $new_session = New-PSSession -ComputerName $pc.computer -Credential $credential -ErrorAction Stop
            }
            catch
            {
                throw "Please check:
                    1. The username or password is incorrect!
                    2. Computer name $pc lack header computer
                "
            }

            #Create errormachine to create log after run completely.

            $script_block = {
                param($credential)
            
                #Get username from group administrators
                $get_localmember = Get-LocalGroupMember Administrators | Where-Object {$_.ObjectClass -eq "User"} | Where-Object {$_.PrincipalSource -eq "ActiveDirectory"};
                $username = $get_localmember.Name;
                
                #Check username in group administrators
                #if there are more 2 users in group administrators, they will return error
                if ($username.count -gt 1){
                    Write-Host "`n############################################`n"
                    $Host.UI.WriteErrorLine("Error: Cannot identify user on this machine ${env:Computername} `nReason: There are more than 2 users in the group administrator`n$username")
                    Write-Host "`n############################################`n"
                    break
                }
                    
                #Convert Username into SID
                $user = New-Object System.Security.Principal.NTAccount($username)
                $sid_temp = $user.Translate([System.Security.Principal.SecurityIdentifier])
                $sid = $sid_temp.Value

                #Create file reg
                $regedit_path = "C:\safe_sender.reg"
                New-Item -Path $regedit_path -ItemType "file" -Force | Out-Null
                $line1 = "Windows Registry Editor Version 5.00"
                Add-Content -Path $regedit_path -Value "$line1`n";
                for ($i=10;$i -le 16;$i++){
                    $l1 = "[HKEY_USERS\"+$sid+"\SOFTWARE\Policies\Microsoft\office\"+$i+".0\outlook\options\mail]"
                    $l2 = '"unblockspecificsenders"=dword:00000001'
                    $l3 = '"junkmailtrustcontacts"=dword:00000001'
                    $l4 = '"trustedzone"=dword:00000000'
                    $l5 = '"intranet"=dword:00000001'
                    $l6 = '"UnblockSafeZone"=dword:00000001'
                    Add-Content -Path $regedit_path -Value "$l1`n$l2`n$l3`n$l4`n$l5`n$l6`n";
                }
                $line3="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\gameloft.com]"
                $line4='"*"=dword:00000002'
                $line5='"https"=dword:00000001'

                $line7="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\gameloft.org]"
                $line8='"*"=dword:00000002'
                $line9='"https"=dword:00000001'

                $line11="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\vivendi.com]"
                $line12='"*"=dword:00000002'
                $line13='"https"=dword:00000001'
                Add-Content -Path $regedit_path -Value "$line3`n$line4`n$line5`n`n$line7`n$line8`n$line9`n`n$line11`n$line12`n$line13";

                #Get content in file that just created
                Get-Content -Path $regedit_path 
                
                #Check file reg
                $test_reg = Test-Path -Path "$regedit_path"
                if ($test_reg){
                    reg import "$regedit_path" | Out-Null
                    Write-Host "Import regedit successfully in $env:Computername" -ForegroundColor Green
                    Remove-Item -Path "$regedit_path" -Force
                }

            }

            #Execute the command in remote machine.
            Invoke-Command -Session $new_session -ScriptBlock $script_block -ArgumentList $credential

            #Close the connection to release the resource that PSSession was using.
            Remove-PSSession -Session $new_session
        
        }else{
            #return computer is offline!
            Write-Host "Computer $pc.computer is offline" -ForegroundColor Red
            }
        }
}

Execute_regedit -computername $computer_list -credential $credential