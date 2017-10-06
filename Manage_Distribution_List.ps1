$Cred = Get-Credential "Edit with admin email"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Cred -Authentication Basic -AllowRedirection
Import-PSSession $Session



function Disconnect-O365 {
    Remove-PSSession $Session
}



Do{

#Creare  o Eliminare GRuppo
Write-Host " "
Write-Host "Selezionare l'attività sulle liste di distribuzione"
Write-Host "Crea lista"
Write-Host "Edita il manager di un gruppo"
Write-Host "Rimuovi lsita"
Write-Host "Aggiungere utente a lista"
Write-Host "rImuove utente da lista"
Write-Host "cerca lista di diStribuzione"
Write-Host "Quit"
Write-Host ""
$action = Read-Host -Prompt 'Digitare la lettera maiuscola'



switch ($action){
"c"{
    $retry = ""
    $i = 1
    do{
    $name = Read-Host -Prompt 'Nome della lista da creare'    
    if (-not (Get-distributionGroup -identity $name)){
    $members = Read-Host -Prompt 'Partecipanti alla lista, separati da virgola'
    foreach($member in $members.Split(",")){
        if ($i = 1){
        New-DistributionGroup -Name $name -PrimarySmtpAddress "$name@example.com" -Members $member -ManagedBy tsw\andrea.feltrin -type Distribution -RequireSenderAuthenticationEnabled $false         
        Set-DistributionGroup -Identity $name -HiddenFromAddressListsEnabled $true
        $i++
        }else{
            Add-DistributionGroupMember -Identity $name -Member "$member@example.com"
        }
        }        
        $retry = "n"
    }else{
        Write-Host "Lista esistente"
    }
    }while($retry -ne "n")
    }
    
"r"{
    $names = Read-Host -Prompt 'Nome della lista da rimuovere, indicare più liste separate da virgola'
    foreach($name in $names.Split(",")){
        Remove-DistributionGroup -Identity $name 
        }
    }
"e"{
$dl = @("list of distribution list")
$groupname = Read-Host -Prompt 'Nome della lista da editare, indicare più liste separate da virgola'
    foreach($gpname in $groupname.Split(",")){
        $group = Get-DistributionGroup -Identity $gpname | select ManagedBy

        $user = $dl.ForEach({Get-DistributionGroupMember -Identity $_ | select name})
        $newmgr = $user.name + $group.ManagedBy | sort | Get-Unique
        Set-DistributionGroup -Identity $gpname -managedby $newmgr

}
    }
"a"{
    $name = Read-Host -Prompt 'Nome della lista da editare'
    $members = Read-Host -Prompt 'Email da aggiungere alla lista, indicare più caselle separate da virgola'
    foreach($member in $members.Split(",")){
        Add-DistributionGroupMember -Identity $name -Member $member 
        }
    }
"i"{
    $name = Read-Host -Prompt 'Nome della lista da editare'
    $members = Read-Host -Prompt 'Email da rimuovere dalla lista, indicare più caselle separate da virgola'
    foreach($member in $members.Split(",")){
        Remove-DistributionGroupMember -Identity $name -Member $member 
        }
      }

"s"{
    $reload = ""
    do{
    $name = Read-Host -Prompt 'Nome della lista da cercare'
    $distribution = Get-distributionGroup -identity $name
        if ($distribution){
            Write-Host "lista $distribution identificata"
            $getname = get-distributiongroupmember -identity $name
            Write-Host $getname
            $reload = "y"
        }else{
            Write-Host "Lista di distribuzione non trovata"
            
        }
        }while($reload -ne "y")
    }

"q"{ $exit = "n" }

}


}
While ($exit -ne "n"){
    Disconnect-O365
}

