$RedirectRuleName="Redirect to Personal Email"

function New-O365Membership.Member{
    [CmdletBinding()]
   Param(
        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$False,ParameterSetName="MemberUPN")]
        [Switch]$ObjectOnly, 

        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$True,ParameterSetName="CSV")]
        [Parameter(Mandatory=$True,ParameterSetName="Parts")]
            [ValidateNotNullOrEmpty()]
            [string]$DomainName,

        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$True,ParameterSetName="Parts")]
            [ValidateNotNullOrEmpty()]
            [string]$FirstName,

        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$True,ParameterSetName="Parts")]
            [ValidateNotNullOrEmpty()]
            [string]$LastName,

        [Parameter(Mandatory=$false,ParameterSetName="ObjectOnly")]      
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
            [Boolean]$IsActive=$false,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
            [string]$MobileNumber=$null,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
            [string]$JobTitle,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]        
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
            [string]$Company,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
            [string]$EmailAddress,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
        [Parameter(Mandatory=$false,ParameterSetName="Parts")]
        [string[]]$ProgramGroupNames,  #Array fo strings

        [Parameter(Mandatory=$True,ParameterSetName="CSV")]
        [ValidateScript({Test-Path $_ -PathType Leaf})] 
        [string] $CSVPath,

        [Parameter(Mandatory=$True,ParameterSetName="MemberUPN")]
        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
        [ValidateNotNullOrEmpty()]
        [string] $MemberUPN,

        [Parameter(Mandatory=$True,ParameterSetName="Member")]
        [ValidateNotNullOrEmpty()]
        [object] $Member


    )

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
    Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"BoundParameters follow")
    Write-Debug  -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")

    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))

    switch ($PsCmdlet.ParameterSetName) { 
        "ObjectOnly"  {                 
            Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Creating Object"))
            $member = New-Object psobject -Property @{
                FirstName = $FirstName
                LastName = $LastName
                DomainName=$DomainName
                IsActive=$IsActive
                Company = $Company
                JobTitle = $JobTitle
                MobileNumber = $MobileNumber
                EmailAddress= $EmailAddress           
                DisplayName=$null
                LastSyncTime=$null                     
		        History=@()
                _ProgramGroupNames=$null
                RedirectRule=$null
            }
            # Add Custom Type Name
            $member.pstypenames.insert(0,"O365Membership.Member")
            Add-Member -InputObject $member -MemberType ScriptProperty -Name ProgramGroupNames -Value {$this._ProgramGroupNames} -SecondValue {
                [string[]]$this._ProgramGroupNames = @()
                $args[0] | ForEach-Object {
                    if(-not [string]::IsNullOrEmpty($_)){
                        $this._ProgramGroupNames+=$_
                    }
                }
            }
            $member.ProgramGroupNames = $ProgramGroupNames

            Add-Member -InputObject $member -MemberType ScriptProperty -Name FullName -Value {("{0} {1}" -f $this.FirstName, $this.LastName)}
            Add-Member -InputObject $member -MemberType ScriptProperty -Name FullNameDotted -Value {("{0}.{1}" -f ($this.FirstName.Trim() -replace " ", "_"   ) ,($this.LastName.Trim() -replace " ", "_" ) )}
            Add-Member -InputObject $member -MemberType ScriptProperty -Name UPN -Value {"{0}@{1}" -f $this.FullNameDotted, $this.DomainName}
            Add-Member -InputObject $member -MemberType ScriptProperty -Name Summary -Value {
                ("Firstname='{5}', Lastname='{6}', UPN='{0}', IsActive={1}, EmailAddress='{2}', MobileNumber ='{4}', ProgramGroupNames= '{3}'" `
                    -f $this.UPN, $this.IsActive, $this.EmailAddress, $this.ProgramGroupNames, $this.mobilenumber, $this.firstname, $this.lastname) 
            }

            #   ****************  PropertyWarnings
            Add-Member -InputObject $member -MemberType ScriptProperty -Name PropertyWarnings -Value {
                $PropertyWarnings=@{}
                if($this.IsActive){
                    if([string]::IsNullOrWhiteSpace($this.MobileNumber)){
                        $PropertyWarnings.Add("MobileNumber","Appears to be blank - Usually we have a phone number for each contact")
                    }                 
                    if($this.ProgramGroupNames.count -eq 0){
                        $PropertyWarnings.Add("ProgramGroupNames","ProgramGroupNames is empty - Member will not be part of any groups")
                    }  

                }
                return $PropertyWarnings
            }

            #   ****************  PropertyErrors
            Add-Member -InputObject $member -MemberType ScriptProperty -Name PropertyErrors -Value {
                $PropertyErrors=@{}
                if([string]::IsNullOrWhiteSpace($this.FirstName)){
                    $PropertyErrors.Add("FirstName","Appears to be blank - Must have First Name")
                }
                if([string]::IsNullOrWhiteSpace($this.LastName)){
                    $PropertyErrors.Add("LastName","Appears to be blank - Must have Last Name")
                }
                if($this.IsActive){
                    if([string]::IsNullOrWhiteSpace($this.EmailAddress)){
                        $PropertyErrors.Add("EmailAddress","Appears to be blank - Email Address is needed for Active Member")
                    }
                    if($this.LastSyncTime -ne $null){
                        If($this.RedirectRule -eq $null){
                            $PropertyErrors.Add("RedirectRule","Appears to be blank - RedirectRule SHould Exist")                            
                        }
                        if($this.RedirectRule.RedirectTo -ne ("""{0}"" [SMTP:{0}]" -f $this.emailaddress )){
                            $PropertyErrors.Add("RedirectRule.RedirectTo",("RedirectTo has email {0}, Expected {1}" -f (($this.RedirectTo -split " ")[0] -split '"')[1], $this.emailaddress))                            
                        }                       
                    }

                }
               $PropertyErrors
            } 
            Add-Member -InputObject $member -MemberType ScriptProperty -Name HasPropertyWarnings -Value {$this.PropertyWarnings.count -gt 0}
            Add-Member -InputObject $member -MemberType ScriptProperty -Name HasPropertyErrors -Value {$this.PropertyErrors.count -gt 0}
            Add-Member -InputObject $member -MemberType ScriptProperty -Name ProgramNames -Value {
                $ProgramNames=@()
                foreach ($ProgramGroupName in $this.ProgramGroupNames){
                    $ProgramName=$ProgramGroupName.Split("-")[0]
                    if(-not [string]::IsNullOrEmpty($ProgramName)){
                        if(($ProgramNames | Where-Object {$_ -eq $ProgramName}) -eq $null){
                            $ProgramNames += $ProgramName
                        }
                    }
                } 
                $ProgramNames
                }
            Add-Member -InputObject $member -MemberType ScriptProperty -Name GroupNames -Value {
                $GroupNames=@()
                foreach ($ProgramGroupName in $this.ProgramGroupNames){
                    $GroupName=$ProgramGroupName.Split("-")[1]
                    if(-not [string]::IsNullOrEmpty($GroupName)){
                        if(($GroupNames | Where-Object {$_ -eq $GroupName}) -eq $null){
                            $GroupNames += $GroupName
                        }
                    }
                } 
                $GroupNames
                }
            Add-Member -InputObject $member -MemberType ScriptProperty -Name InferredDisplayName -Value {
                if($this.GroupNames -ne $null){
                    ('{0} ({1})' -f $this.FullName, ($this.GroupNames -join ","))
                }
                else{
                     ("{0} (No Groups)" -f $this.FullName)
               }

            }
            Add-Member -InputObject $Member -MemberType ScriptMethod -Name Equals -Force -Value {
                param(
                    [Parameter(Mandatory=$True)]
                    [object]$memberToCompare
                )
                $isEqual=$True
                $isEqual = ($isEqual -and ($this.firstname -eq $memberToCompare.firstname))
                $isEqual = ($isEqual -and ($this.lastname -eq $memberToCompare.lastname))
                $isEqual = ($isEqual -and ($this.DomainName -eq $memberToCompare.DomainName))
                $isEqual = ($isEqual -and ($this.MobileNumber -eq $memberToCompare.MobileNumber))
                $isEqual
           }
            Add-Member -InputObject $Member -MemberType ScriptMethod -Name Compare -Force -Value {
                param(
                    [Parameter(Mandatory=$True)]
                    [object]$memberToCompare
                )
                $compareResults=@{}
                $propertyNames=@("firstname","lastname","domainname","mobilenumber","IsActive","emailaddress")
                foreach($propertyName in $propertyNames){
                    if(($this | select -expand $propertyName).ToString() -eq ($memberToCompare | select -expand $propertyName).ToString()){
                        $compareResults.Add($propertyName,"Same") 
                    }
                    else{
                         $compareResults.Add($propertyName,"Different")
                    }
                }
                $groupSame=$True
                foreach($programGroupName in $this.ProgramGroupNames){
                    $groupSame = $groupSame -and (($memberToCompare.ProgramGroupNames | where {$_ -eq $programGroupName}) -ne $null)
                }
                if($groupSame){
                    $compareResults.Add("ProgramGroupNames","Same") 
                }
                else{
                     $compareResults.Add("ProgramGroupNames","Different") 
                }
                $compareResults
            }         
            Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member Created Follows"))
            Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),($member | fl | Out-String))

            Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("-------------- Member"))

            $member
        }
        "CSV"{
            Import-CSV -Path $CSVPath | ForEach-Object {
                Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("CSV Row = {0}" -f $_))
                New-O365Membership.Member -ObjectOnly -DomainName $DomainName -FirstName $_.'First Name'.trim() -LastName $_.'Last Name'.trim() -IsActive ([System.Convert]::ToBoolean($_.'Is Active')) `
                    -MobileNumber $_.'Mobile Number'.trim() -EmailAddress $_.'Email Address (Personal)'.trim() -ProgramGroupNames (((($_.'Contact Groups') -replace ";#[0-9]+","") -replace ";#",",") -split ",")
                
            }
        }
        "MemberUPN" {
            try{
                
                New-O365Membership.Member -ObjectOnly -Domain ($MemberUPN -split "@")[1] -FirstName ((($MemberUPN -split "@")[0] -split "\.")[0] -replace "_", " ")  -LastName ((($MemberUPN -split "@")[0] -split "\.")[1] -replace "_", " ")
            }
            catch{
                Write-Error -Message ("Something happend when trying to split MemberUPN name") -RecommendedAction "UPN probably isn't in a format that is needed for a Member"
                # Return null to indicate problem
                $null
            }
        }
        "Parts"{
            $toNewMember = New-O365Membership.Member -ObjectOnly -DomainName $DomainName -FirstName $FirstName -LastName $LastName -IsActive $IsActive -MobileNumber $MobileNumber -ProgramGroupNames $ProgramGroupNames -EmailAddress $EmailAddress
            $newMember = New-O365Membership.Member -Member $toNewMember                        
        }
        "Member"{
            if(-not $member.HasPropertyErrors){
                try{
                    $msolUser=New-MsolUser -UserPrincipalName $member.UPN -LicenseAssignment "etly:STANDARDWOFFPACK" -UsageLocation "US"  `
                        -DisplayName $member.InferredDisplayName -FirstName $member.FirstName -LastName $member.LastName -ErrorAction Stop -ErrorVariable $newMSOluserError
                }
                catch{
                    Write-Error -Message ("Unable to add mailbox for member {0} with Exception {1}" -f $member.UPN, $_.Exception)
                    return
                }

                $LoopCount=0
                $SleepSeconds=15
                $WarningDelayStartSeconds=60
                $waitStart = (Get-Date)
                Write-Verbose -Message ("Waiting for Mailbox to get created for Member {0} - This can take a few minutes. Will wait {1} seconds before warnings. Starting at {2} {3}" -f $member.UPN,$WarningDelayStartSeconds, $waitStart,$waitStart.Kind  )
                while((Get-Mailbox -Identity $member.UPN -ErrorAction SilentlyContinue) -eq $null){
                    $waitedTime = New-TimeSpan -Start $waitStart -End (Get-Date)
                    if($waitedTime.Seconds -gt $WarningDelayStartSeconds){
                        Write-Warning -Message ("  Waited {0} seconds so far for mailbox to create for new member {1}, Sleeping {2} seconds." -f $waitedTime.Seconds, $member.UPN, $SleepSeconds)
                    }
                    Start-Sleep -Seconds $SleepSeconds              
                }               
               Write-Verbose -Message ("Waited {0} seconds total" -f (New-TimeSpan -Start $waitStart -End (Get-Date)).seconds )

               Write-Verbose -Message ("Set Properties for new user {0}" -f $member.UPN  )
               Set-O365Membership.Member -Member $member              
            }
            else{
                Write-Error -Message ("Unable to Create New Member for {0} - Has PropertyErrors {1}" -f $member.upn, ($member.PropertyErrors | fl | Out-String))
            }
            
        }
    }
    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Leaving"))

}

function Get-O365Membership.Member{
    [CmdletBinding(DefaultParameterSetName="Default")]
   Param(
        [Parameter(Mandatory=$True,ParameterSetName="ByDomain")]
            [ValidateNotNullOrEmpty()]
            [string]$DomainName,

        [Parameter(Mandatory=$True,ParameterSetName="ByMemberUPN")]
            [ValidateNotNullOrEmpty()]
            [string]$MemberUPN

    )
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("Entering"))
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,"----------------------BoundParameters follow")
    Write-Debug -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,"--------------------- BoundParameters End")

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))

    $msolusers=@()
    $dlMembers=@()
    switch ($PsCmdlet.ParameterSetName) { 
        "ByDomain"  {
            $DistributionGoups=Get-DistributionGroup -Filter "{(WindowsEmailAddress -like '*@$DomainName') -and (WindowsEmailAddress -like '*-*') } "  
            foreach($DistributionGoup in $DistributionGoups){
                $dlMembers = Get-DistributionGroupMember -Identity $DistributionGoup.PrimarySMTPAddress
            }      
        }
        "ByMemberUPN" {
            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("Getting MSOLuser for {0}" -f $MemberUPN))
            $msolusers+= Get-MsolUser -UserPrincipalName $MemberUPN -ErrorAction SilentlyContinue
        }
        default {
            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("Getting All MSOLusers"))
            $msolusers+= Get-MsolUser -ErrorAction SilentlyContinue
        }        
    }                 

    foreach($DLmember in $DLmembers){
        if($DLmember.RecipientType -eq "UserMailbox"){
            if(($msolusers | where {$_.UserPrincipalName -eq $DLmember.PrimarySMTPAddress}) -eq $null){
                $msolusers += Get-MsolUser -UserPrincipalName $DLmember.PrimarySMTPAddress
            }
        }
    }


    foreach($msoluser in $msolusers){
        Write-Debug  -Message ("Function {0} called  from{1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("Processing MSOLuser {0}" -f $MSOlUser.UserPrincipalName))

        $member = New-O365Membership.Member -MemberUPN $msoluser.UserPrincipalName -ErrorAction SilentlyContinue
        if($member -ne $null){
            $member.LastSyncTime = Get-Date
            $member.DomainName = $MsolUser.UserPrincipalName.Split("@")[1]
            $member.IsActive = (-not $MSOLUser.BlockCredential)
            $member.MobileNumber = [string]$MSOLUser.MobilePhone
            $member.JobTitle = $MSOLUser.Title
            if(-not [string]::IsNullOrWhiteSpace($MSOLUser.AlternateEmailAddresses[0])){
                $member.EmailAddress = $MSOLUser.AlternateEmailAddresses[0]
            }
            else{
                $member.EmailAddress=$null
            }
            $member.ProgramGroupNames = (Get-O365Membership.Group -MemberUPN $member.UPN | select -expand ProgramGroupName)
            $member.RedirectRule = Get-InboxRule $RedirectRuleName -Mailbox $member.UPN -ErrorAction SilentlyContinue
            Write-Verbose -Message ("Got Member - {0}" -f $member.Summary)
        }
        else{
            Write-Error -Message ("Unable to parse {0} with New-O365Membership.Member - Assume this MSOLuser is NOT a Member" -f $msoluser.UserPrincipalName )
        }
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,"---------Member to return follows")
        Write-Debug -Message (($member) | Format-list | Out-String)
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,"---------------------")

        $member 
    }

    Write-Debug  -Message ("Function {0} called  from{1} - {2}" -f $MyInvocation.InvocationName, (gcs)[1].Location,("Leaving"))

}

function Set-O365Membership.Member{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ParameterSetName="MemberUPN")]
        [ValidateNotNullOrEmpty()]
        [string] $MemberUPN, 

        [Parameter(Mandatory=$False,ParameterSetName="MemberUPN")]
        [string] $MobileNumber,

        [Parameter(Mandatory=$False,ParameterSetName="MemberUPN")]
        [string] $EmailAddress,

        [Parameter(Mandatory=$False,ParameterSetName="MemberUPN")]
        [string[]] $ProgramGroupNames,

        [Parameter(Mandatory=$False,ParameterSetName="MemberUPN")]
        [boolean] $IsActive,

        [Parameter(Mandatory=$True,ParameterSetName="Member")]
        [ValidateNotNullOrEmpty()]
        [object] $Member
    )
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
    Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"BoundParameters follow")
    Write-Debug  -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")

    Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))

    switch ($PsCmdlet.ParameterSetName) { 
        "MemberUPN"  {
            $member = New-O365Membership.Member -ObjectOnly -MemberUPN $MemberUPN

            $argLine="-UserPrincipalName {0}" -f $MemberUPN
            if($PSBoundParameters.ContainsKey("MobileNumber")){
                $Member.MobileNumber=$MobileNumber
                if(-not [string]::IsNullOrEmpty($MobileNumber)){
                    $argLine+=" -MobilePhone ""{0}""" -f $MobileNumber
                }
                else{
                    $argLine+=" -MobilePhone {0}" -f '$null'

                }
            }
            if($PSBoundParameters.ContainsKey("EmailAddress")){
                $Member.EmailAddress = $EmailAddress
                if(-not [string]::IsNullOrEmpty($emailaddress)){
                    $argLine+=" -AlternateEmailAddresses {0}" -f [string[]]$EmailAddress
                }
                else{
                    # Need to send Null if no email address because this expects an array of string
                    $argLine+=" -AlternateEmailAddresses {0}" -f '$null'
                }
            }                        
            if($PSBoundParameters.ContainsKey("IsActive")){
                $Member.IsActive = -not $PSBoundParameters["IsActive"]
                $argLine+=" -BlockCredential {0}" -f ([int](-not $PSBoundParameters["IsActive"]))
            }

            if($PSBoundParameters.ContainsKey("ProgramGroupNames")){
                $Member.ProgramGroupNames = $PSBoundParameters["ProgramGroupNames"]
            }

            $argLine+=(" -DisplayName ""{0}""" -f $Member.InferredDisplayName)

            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Calling MSOLuser with args {0}" -f $argLine ))
            Invoke-Expression -Command ("Set-MSOLuser {0}" -f $argLine)
            $gottonMSOluser = Get-MsolUser -UserPrincipalName $MemberUPN
            $gottonGroups = Get-O365Membership.Group -MemberUPN $MemberUPN
            $GottonProgramGroupNames = (Get-O365Membership.Group -MemberUPN $MemberUPN | Select -Expand ProgramGroupName)
            if(-not $gottonMSOluser.BlockCredential){
                Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member {0} IS Active" -f $MemberUPN ))
                #Handle Email
                $currentEmail = $gottonMSOluser.AlternateEmailAddresses[0]
                if($currentEmail -ne $null){
                    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member {0} has alternate email {1}" -f $MemberUPN, $currentEmail ))
                    $redirectrule=Get-InboxRule $RedirectRuleName -Mailbox $MemberUPN -ErrorAction SilentlyContinue
                    if($redirectrule -ne $null){
                        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Found Existing InboxRule with Identity {0}" -f $redirectrule.Identity ))
                        $currentRuleEmail=(($redirectrule.RedirectTo -split " ")[0] -replace '"',"")
                        if($currentRuleEmail -ne $currentEmail){
                            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("InboxRule has Email {0}. Change to {1}" -f $currentRuleEmail,$currentEmail  ))
                            Set-InboxRule -Identity $redirectrule.identity -RedirectTo $currentEmail -Confirm:$false
                        }
                    }
                    else{
                        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Did not find Existing InboxRule so create one"))
                        $inboxrule=New-InboxRule -Mailbox $MemberUPN -RedirectTo $currentEmail -Name $RedirectRuleName
                        if($inboxrule -ne $null){
                            Write-Debug  -Message ("Function {0} called {1} - {2}" -f (gcs)[0].Location, (gcs)[1].Location ,("Rule ID {0} Created" -f $inboxrule.Identity))
                        }
                        else{
                            Write-Error -Message ("RedirectRUle for user {0} did not create" -f $MemberUPN)
                        }
                    }

                }
                else{
                    Write-Error  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member{0} does NOT have alternate email - This isn't expected" -f $MemberUPN))
                }

                #Handle Groups            
                if($PSBoundParameters.ContainsKey("ProgramGroupNames")){
                    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member{0} Has Groups '{1}'. Needs to have Groups '{2}'" -f $MemberUPN, ($GottonProgramGroupNames -join ","), ($programGroupNames -join ",") ))
                    $ToRemoveProgramGroupNames=@()
                    $ToAddProgramGroupNames=@()
                    foreach($programGroupName in $programGroupNames){
                        if(($gottonGroups | Where-Object {$_.ProgramGroupName -eq $programGroupName}) -eq $null){
                            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("No Current Membership in Group {0} - Need to add" -f $programGroupName ))
                            $ToAddProgramGroupNames+=$programGroupName
                        }
                    }
                    foreach($gottenGroup in $gottonGroups){
                        if (($ProgramGroupNames | Where-Object {$_ -eq $gottenGroup.ProgramGroupName} ) -eq $null){
                            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Need to Remove Group {0}" -f $gottenGroup.ProgramGroupName ))
                            $ToRemoveProgramGroupNames+=$gottenGroup.ProgramGroupName
                        }

                    }
                    foreach($toAddprogramGroupName in $ToAddProgramGroupNames){
                        $groupUPN="{0}@{1}" -f $toAddprogramGroupName, ($MemberUPN -split "@")[1]
                        Write-Debug -Message ($MyInvocation.InvocationName + " - "+ ("Adding {0} to group '{1}'" -f $MemberUPN,$groupUPN) )
                        Add-DistributionGroupMember -Identity $groupUPN -member $MemberUPN -Confirm:$false                    
                    }

                    foreach($toRemoveprogramGroupName in $ToRemoveProgramGroupNames){
                        $groupUPN="{0}@{1}" -f $toRemoveprogramGroupName, ($MemberUPN -split "@")[1]
                        Write-Debug -Message ($MyInvocation.InvocationName + " - "+ ("Removing {0} from group '{1}'" -f $MemberUPN,$groupUPN) )
                        Remove-DistributionGroupMember -Identity $groupUPN -member $MemberUPN -Confirm:$false                    
                    }

                }
                Get-O365Membership.Member -MemberUPN $MemberUPN     
            }
            else{
                Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member{0} is NOT Active" -f $MemberUPN ))
                Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member {0} is Not Active -  Remove RedirectRule if it exist" -f $MemberUPN ))
                Remove-InboxRule $RedirectRuleName -Confirm:$false -Mailbox $MemberUPN -Force:$True -ErrorAction SilentlyContinue
                Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Member {0} is NOT Active - Remove from Groups" -f $MemberUPN ))
                $gottonGroups | ForEach-Object {Remove-DistributionGroupMember -Identity $_.UPN -member $MemberUPN -Confirm:$false -ErrorAction SilentlyContinue}

            }
           

        }
        "Member" {
            (Set-O365Membership.Member -MemberUPN $Member.UPN -IsActive $Member.IsActive -EmailAddress $Member.emailaddress -ProgramGroupNames $Member.ProgramGroupNames -MobileNumber $Member.MobileNumber)
        }
    }

    
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Leaving")

}

function Sync-O365Membership.Member{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$DomainName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$FirstName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$LastName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$MobileNumber,
        
        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$EmailAddress,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$ProgramGroupNames,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [boolean]$IsActive

        
    )
    BEGIN {
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
        Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters Start")
        Write-Debug  -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")

        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))

        switch ($PsCmdlet.ParameterSetName) {
            "Parts"{
            } 
        }
    }
    PROCESS{
        $member = New-O365Membership.Member -objectonly -DomainName $DomainName -FirstName $FirstName -LastName $LastName -MobileNumber $MobileNumber -IsActive $IsActive -EmailAddress $EmailAddress -ProgramGroupNames $ProgramGroupNames 
        if((Get-O365Membership.Member -MemberUPN $member.UPN -Verbose:$False) -ne $null){
            Write-Verbose -Message ("Setting Member {0} - {1}" -f $member.UPN, $member.Summary)
            Set-O365Membership.Member -Member $member -Verbose:$False > $null
            if(-not $member.IsActive){
                Write-Warning -Message ("****  Member {0} NOT active - Can be removed from list" -f $member.UPN)
            }

        }
        else{
            if($member.IsActive){            
                Write-Verbose -Message ("Newing Member {0} - {1}" -f $member.UPN,  $member.Summary)
                New-O365Membership.Member -Member $member -Verbose:$False > $null
            }
            else{
                 Write-Warning -Message ("****  Do Nothing with {0} - Not Active and Not Online so could be removed from member list - {1} " -f $member.UPN, $member.Summary)               
            }

        }

    }
    End{
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Leaving")    
    }

}

function Compare-O365Membership.Member{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$DomainName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$FirstName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$LastName,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$MobileNumber,
        
        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$EmailAddress,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$ProgramGroupNames,

        [Parameter(Mandatory=$False,ParameterSetName="Parts")]
        [Parameter(ValueFromPipelineByPropertyName=$True)]
        [boolean]$IsActive
    )
    BEGIN {
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
        Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters Start")
        Write-Debug  -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")

        Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))

        switch ($PsCmdlet.ParameterSetName) {
            "Parts"{
            } 
        }
    }
    PROCESS{
        $member = New-O365Membership.Member -objectonly -DomainName $DomainName -FirstName $FirstName -LastName $LastName -MobileNumber $MobileNumber -IsActive $IsActive -EmailAddress $EmailAddress -ProgramGroupNames $ProgramGroupNames 
        $memberGotten = Get-O365Membership.Member -MemberUPN $member.UPN -Verbose:$False
        if($memberGotten -ne $null){
            $compareResults = $member.Compare($memberGotten)
            $differences=@()
            foreach($key in $compareResults.keys)
            {
                if($compareResults[$key] -ne "Same"){
                   $differences+=$key
                }
            }
            if($differences -ne $null){
                Write-Warning -Message ("Members are different for {0} with differences {1} - Member Online ({2}), Member Local ({3})" -f $member.UPN, ($differences -join ","), $memberGotten.summary, $member.summary)
            }
            else{
                Write-Verbose -Message ("Members are same for Member Online = {0}" -f $memberGotten.summary)
            }
        }
        else{
            Write-Warning -Message ("Unable to find a member to compare with {0}" -f $member.upn)
        }

    }
    End{
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Leaving")    
    }

}


function Get-O365Membership.Group{
    [CmdletBinding(DefaultParameterSetName="ObjectOnly")]
   Param(
        [Parameter(Mandatory=$True,ParameterSetName="MemberUPN")]
            [ValidateNotNullOrEmpty()]
            [string]$MemberUPN
    )
    

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
    Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"BoundParameters follow")
    Write-Debug -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")

    switch ($PsCmdlet.ParameterSetName) { 
        "MemberUPN"  { 
            $DomainName = ($MemberUPN -split "@")[1]
            $DistributionGroups=Get-DistributionGroup -Filter "{(WindowsEmailAddress -like '*@$DomainName') -and (WindowsEmailAddress -like '*-*') }" `
                | Where-Object { (Get-DistributionGroupMember -Identity $_.PrimarySMTPAddress | Where-Object { $_.PrimarySMTPAddress -eq $MemberUPN} ) -ne $null  }
    
        }
    }           
    foreach($DistributionGroup in $DistributionGroups){
        Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Processing DistributionGroup {0}"-f $DistributionGroup.PrimarySMTPAddress ) )

        $MemberUPNs=Get-DistributionGroupMember -Identity $DistributionGroup.PrimarySMTPAddress | select -expand PrimarySMTPAddress
        $ProgramGroupName= ($DistributionGroup.WindowsEmailAddress -split "@")[0]

        New-O365Membership.Group -DomainName ($DistributionGroup.PrimarySmtpAddress -split "@")[1] -ProgramName ($ProgramGroupName -split "-")[0] -GroupName ($ProgramGroupName -split "-")[1] -MemberUPNs $MemberUPNs
    }

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Leaving")

}

function New-O365Membership.Group{
    [CmdletBinding(DefaultParameterSetName="ObjectOnly")]
   Param(
        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
            [ValidateNotNullOrEmpty()]
            [string]$DomainName,

        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
            [ValidateNotNullOrEmpty()]
            [string]$ProgramName,

        [Parameter(Mandatory=$True,ParameterSetName="ObjectOnly")]
            [ValidateNotNullOrEmpty()]
            [string]$GroupName,

        [Parameter(Mandatory=$False,ParameterSetName="ObjectOnly")]
            [string[]]$MemberUPNs=@()


    )
    

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Entering")
    Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"BoundParameters follow")
    Write-Debug -Message (($MyInvocation.BoundParameters) | Format-list | Out-String)
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"--------------------- BoundParameters End")
    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("ParameterSetName= {0}" -f $PsCmdlet.ParameterSetName ))
    switch ($PsCmdlet.ParameterSetName) { 
        "ObjectOnly"  {                 
            Write-Debug  -Message ("Function {0} called {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),("Creating Object"))
            $group = New-Object psobject -Property @{
                DomainName = $DomainName
                ProgramName = $ProgramName
                GroupName=$GroupName
		        MemberUPNs=$MemberUPNs
             }
            Add-Member -InputObject $group -MemberType ScriptProperty -Name ProgramGroupName -Value {("{0}-{1}" -f $this.ProgramName.Trim(), $this.GroupName.Trim())}
            Add-Member -InputObject $group -MemberType ScriptProperty -Name UPN -Value {("{0}@{1}" -f $this.ProgramGroupName,$this.DomainName.Trim())}
            Write-Debug  -Message ("Function {0} called bh {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Group Created follow")
            Write-Debug -Message (($group) | Format-list | Out-String)
            Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"---------------------End")
            $group
        }

    }

    Write-Debug  -Message ("Function {0} called by {1} - {2}" -f $MyInvocation.InvocationName, ((gcs)[1] | select Location),"Leaving")

}
