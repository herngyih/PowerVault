function Get-SpaceRequired
{
	param($OS,$freeSpace)
	
	if($OS -eq 'W2K16'){
	
		$holding_space_required = 15 - $freeSpace
		
		if($holding_space_required.GetType().Name -eq 'double')
        {
            $space_required = [Math]::Truncate($holding_space_required) + 1

            return $space_required
        }

        else{

            $space_required = $holding_space_required

            return $space_required
        }
	
	}
	
	else{

            $holding_space_required = 7 - $freeSpace

            if($holding_space_required.GetType().Name -eq 'double')
            {
                $space_required = [Math]::Truncate($holding_space_required) + 1

                return $space_required
            }

            else{

                $space_required = $holding_space_required

                return $space_required
            }
	
	}
}

function Get-ObjectList
{
    param($key1)
    
    $id= whoami
    $trimmed_id = $id.toString().split('\')[1]

    if($key1 -eq "computer"){$content = Get-Content C:\Users\$trimmed_id\Desktop\computers.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "group"){$content = Get-Content C:\Users\$trimmed_id\Desktop\groups.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "LANID"){$content = Get-Content C:\Users\$trimmed_id\Desktop\lanIDs.txt|Where-Object{$_.trim() -ne ''}}
    elseif($key1 -eq "country"){$content = Get-Content C:\Users\$trimmed_id\Desktop\countries.txt|Where-Object{$_.trim() -ne ''}}

    return $content

}

function Get-C_FreeSpace
{
    param($ComputerName)
    try{
            $Computer_LocalDisks = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $ComputerName -ErrorAction Stop| Where-Object {$_.DriveType -eq 3}
    
            $disk_enumerator = -1

            foreach($drive in $Computer_LocalDisks.DeviceID)
            {
                $disk_enumerator++
                if($drive -eq "C:")
                {
                    return [math]::Round(($Computer_LocalDisks[$disk_enumerator].FreeSpace)/1GB,2)
                }
            }
    }

    catch{
            return "UTC"
    }

    
}

function Get-OS
{
    param($ComputerName)
    
    try{
        $Computer_OS = Get-WmiObject -Class Win32_OperatingSystem -comp $ComputerName -ErrorAction Stop
    
        if($Computer_OS.caption -match "Windows" -and $Computer_OS.caption -match "2008"){return "W2K8"}
        elseif($Computer_OS.caption -match "Windows" -and $Computer_OS.caption -match "2012"){return "W2K12"}
        elseif($Computer_OS.caption -match "Windows" -and $Computer_OS.caption -match "2016"){return "W2K16"}
        elseif($Computer_OS.caption -match "Windows" -and $Computer_OS.caption -match "2019"){return "W2K19"}
        elseif($Computer_OS.caption -match "Windows" -and $Computer_OS.caption -match "2022"){return "W2K22"}
        else{return $Computer_OS.caption}
    }
    catch{
        return "UTC"
    }

    
    
}

function Get-Compliant
{
    param($OS,$freeSpace)

    if(($OS -eq "W2K16") -and ($freeSpace -lt 15)){return "NO"}
	elseif(($OS -eq "W2K8") -or ($OS -eq "W2K12") -or ($OS -eq "W2K19") -or ($OS -eq "W2K22")  -and ($freeSpace -lt 7)){return "NO"}
    elseif(($OS -eq "UTC") -or ($freeSpace -eq "UTC")){return "UTC"}
    else{return "YES"}
}

function Delete-Folder_Files
{
    param($computer)
    #Get-ChildItem -Path \\$computer\c$\Windows\WinSxS\ManifestCache -ErrorAction SilentlyContinue |ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\Logs\CBS -ErrorAction SilentlyContinue |ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\Logs\DISM -ErrorAction SilentlyContinue |ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\Web -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\'Downloaded Program Files' -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\ProgramData\Microsoft\Windows\WER\ReportQueue -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\SoftwareDistribution\Download -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\inetpub\logs\LogFiles -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Windows\ccmcache -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\Recycler -ErrorAction SilentlyContinue |ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}
    Get-ChildItem -Path \\$computer\c$\'$Recycle.Bin' -ErrorAction SilentlyContinue|ForEach-Object{Remove-Item $_.FullName -Force -Recurse -ErrorAction SilentlyContinue}

}

function Get-Hash
{
    param($code,$computer,$bFree_space,$OS,$aFree_space,$space_required)

    #param($code,$computer,$bFree_space,$OS,$aFree_space,$space_required)

    if($code -eq 1)
    {
          $hash=@{
                       "Computer" = $computer
                       "OS" = $OS
                       "Remedied"  = "NR"
                       "Before(GB)" = $bFree_space
                       "After(GB)"= $aFree_space
                       "Required(GB)" = ""
                 }

          return $hash
    }

    elseif($code -eq 2)
    {
          $hash=@{
                       "Computer" = $computer
                       "OS" = $OS
                       "Remedied"  = "YES"
                       "Before(GB)" = $bFree_space
                       "After(GB)"= $aFree_space
                       "Required(GB)" = ""
                 }

          return $hash
    }

    elseif($code -eq 3)
    {
          $hash=@{
                       "Computer" = $computer
                       "OS" = $OS
                       "Remedied"  = "NO"
                       "Before(GB)" = $bFree_space
                       "After(GB)"= $aFree_space
                       "Required(GB)" = $space_required
                 }

          return $hash
    }
    
    elseif($code -eq 4)
    {
          $hash=@{
                       "Computer" = $computer
                       "OS" = $OS
                       "Remedied"  = "UTC"
                       "Before(GB)" = $bFree_space
                       "After(GB)"= $aFree_space
                       "Required(GB)" = ""
                 }

          return $hash
    }
    
}



    $stamp = get-date -Format "dd-MMM-yyyy HHmm"
    $my_name = whoami | Out-String
    $user_name = $my_name.Split("\")[1].Trim()
    $current_directory = pwd
    
    
    $obj =@()
    $computers = Get-ObjectList computer
    $i = 0
    ForEach($computer in $computers)
    {
            $i++
            $remaining = $computers.count - $i
            Write-Host -NoNewline "Checking and cleaning C:\ in $computer."$remaining "remains"
            
            $bFree_space = Get-C_FreeSpace $computer
            $OS = Get-OS $computer
            $compliant = Get-Compliant $OS $bFree_space

            if($compliant -eq "YES")      {$obj += New-Object psobject -Property (Get-Hash 1 $computer $bFree_space $OS $bFree_space)}
            elseif($compliant -eq "UTC")  {$obj += New-Object psobject -Property (Get-Hash 4 $computer $bFree_space $OS $bFree_space)}
            
            elseif($compliant -eq "NO")   
            {
              
                Delete-Folder_Files $computer
                $aFree_space = Get-C_FreeSpace $computer
                $compliant = Get-Compliant $OS $aFree_space

                if($compliant -eq "YES")  {$obj += New-Object psobject -Property (Get-Hash 2 $computer $bFree_space $OS $aFree_space)}
                
                elseif($compliant -eq "NO"){
                
                    $space_required = Get-SpaceRequired $os $aFree_space
                    
                    $obj += New-Object psobject -Property (Get-Hash 3 $computer $bFree_space $OS $aFree_space $space_required)
                    
                }
            }




            
            Write-Host "...Done"
        
    }

    $obj|Select-Object Computer,OS,Remedied,'Before(GB)','After(GB)','Required(GB)'| FT
    $obj|Select-Object Computer,OS,Remedied,'Before(GB)','After(GB)','Required(GB)'|Export-Csv c:\users\$user_name\desktop\cResults_$stamp"_"$user_name.csv -NoTypeInformation
    Write-Host "Exported to a CSV file in desktop " 
    








  
