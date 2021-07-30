##################################################################################################

$host_computer = $env:COMPUTERNAME
$computer_list = Import-Csv -Path "C:\computerList.csv" -Header computer
$password = "123456?Ab" | ConvertTo-SecureString -asPlainText -Force;       
$username = "gameloft\truong_tranduy"                                              
$credential = New-object System.Management.Automation.PSCredential($username,$password)

##################################################################################################

function Execute_regedit{
    param($computername,$host_computer,$credential)

    foreach ($pc in $computername) {
        
        #Test connection first, if true return computer is up
        $test_connection = Test-Connection -ComputerName $pc.computer -Count 2 -Quiet

        if ($test_connection){
            
            try 
            {
                $new_session = New-PSSession -ComputerName $pc.computer -Credential $credential -ErrorAction Stop
                $new_session
            }
            catch
            {
                throw "Please check:
                    1. The username or password is incorrect!
                    2. Computer name $pc lack header computer
                "
            }

            $script_block = {
                param($host_computer, $credential)
            
                #Get username from group administrators
                $get_localmember = Get-LocalGroupMember Administrators | Where-Object {$_.ObjectClass -eq "User"} | Where-Object {$_.PrincipalSource -eq "ActiveDirectory"};
                $username = $get_localmember.Name;

                #Convert Username into SID
                $user = New-Object System.Security.Principal.NTAccount($username)
                $sid_temp = $user.Translate([System.Security.Principal.SecurityIdentifier])
                $sid = $sid_temp.Value

                #Create file reg
                $regedit_path = "C:\safe_sender.reg"
                New-Item -Path $regedit_path -ItemType "file" -Force;
                $line1 = "Windows Registry Editor Version 5.00"
                $line2=""
                Add-Content -Path $regedit_path -Value "$line1";
                Add-Content -Path $regedit_path -Value "$line2";
                for ($i=10;$i -le 16;$i++){
                    $l1 = "[HKEY_USERS\"+$sid+"\SOFTWARE\Policies\Microsoft\office\"+$i+".0\outlook\options\mail]"
                    $l2 = '"unblockspecificsenders"=dword:00000001'
                    $l3 = '"junkmailtrustcontacts"=dword:00000001'
                    $l4 = '"trustedzone"=dword:00000000'
                    $l5 = '"intranet"=dword:00000001'
                    $l6 = '"UnblockSafeZone"=dword:00000001'
                    $l7 = ""
                    Add-Content -Path $regedit_path -Value "$l1";
                    Add-Content -Path $regedit_path -Value "$l2";
                    Add-Content -Path $regedit_path -Value "$l3";
                    Add-Content -Path $regedit_path -Value "$l4";
                    Add-Content -Path $regedit_path -Value "$l5";
                    Add-Content -Path $regedit_path -Value "$l6";
                    Add-Content -Path $regedit_path -Value "$l7";
                }
                $line3="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\gameloft.com]"
                $line4='"*"=dword:00000002'
                $line5='"https"=dword:00000001'
                $line6 = ""
                $line7="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\gameloft.org]"
                $line8='"*"=dword:00000002'
                $line9='"https"=dword:00000001'
                $line10 =""
                $line11="[HKEY_USERS\"+$sid+"\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\vivendi.com]"
                $line12='"*"=dword:00000002'
                $line13='"https"=dword:00000001'
                Add-Content -Path $regedit_path -Value "$line3";
                Add-Content -Path $regedit_path -Value "$line4";
                Add-Content -Path $regedit_path -Value "$line5";
                Add-Content -Path $regedit_path -Value "$line6";
                Add-Content -Path $regedit_path -Value "$line7";
                Add-Content -Path $regedit_path -Value "$line8";
                Add-Content -Path $regedit_path -Value "$line9";
                Add-Content -Path $regedit_path -Value "$line10";
                Add-Content -Path $regedit_path -Value "$line11";
                Add-Content -Path $regedit_path -Value "$line12";
                Add-Content -Path $regedit_path -Value "$line13";
                
                Get-Content -Path $regedit_path 
                
                #Check file reg
                $test_reg = Test-Path -Path "$regedit_path"
                if ($test_reg){
                    reg import "$regedit_path"
                    Remove-Item -Path "$regedit_path" -Force
                }

            }

            #Execute the command in remote machine.
            Invoke-Command -Session $new_session -ScriptBlock $script_block -ArgumentList $credential -ErrorAction SilentlyContinue

            #Close the connection to release the resource that PSSession was using.
            Remove-PSSession -Session $new_session
        
        }else{
            #return computer is offline!
            Write-Host "Computer $pc.computer is offline" -ForegroundColor Red
            }
        }
}

Execute_regedit -computername $computer_list -host_computer $host_computer -credential $credential
