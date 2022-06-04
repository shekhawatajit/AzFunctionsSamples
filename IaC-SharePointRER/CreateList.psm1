#function to add new list 
function New-RequestList {
    Param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$HubSiteUrl,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$ClientID,
    
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Certificate,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$TenantID
    )        
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
    
    Write-Information "Creating request list in Hub site if does not exist."
    $ListTitle = "Project Request"
    $ListDescription = "List for requesting new Project Site."
    #Connecting to site
    Connect-PnPOnline -Url $HubSiteUrl -ClientId $ClientID -CertificateBase64Encoded $Certificate -Tenant $TenantID
    
    # Creating list
    $List = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (! $List) {
        $List = New-PnPList -Title $ListTitle -Url "lists/ProjectRequest" -Template GenericList -EnableVersioning -OnQuickLaunch:$false
        $UpdatedList = Set-PnPList -Identity $ListTitle -Description $ListDescription 
    }
    # Creating fields if does not exist
    $Fld1 = Get-PnPField -List $ListTitle -Identity 'Owners' -ErrorAction SilentlyContinue
    if (! $Fld1) {
        $Fld1 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='UserMulti' DisplayName='Owners' List='UserInfo' Required='TRUE' EnforceUniqueValues='FALSE' ShowField='ImnName' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Mult='TRUE' Sortable='FALSE' ID='{51d3d5ca-08d5-4248-a65e-65889da08cb3}' StaticName='Owners' Name='Owners' Description='Owners of this project site and group.'/>"
    }
    $Fld2 = Get-PnPField -List $ListTitle -Identity 'Members' -ErrorAction SilentlyContinue
    if (! $Fld2) {
        $Fld2 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='UserMulti' DisplayName='Members' List='UserInfo' Required='FALSE' EnforceUniqueValues='FALSE' ShowField='ImnName' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Mult='TRUE' Sortable='FALSE' ID='{aa2dfe68-d2aa-483f-b96c-5eef95cb0982}' StaticName='Members' Name='Members' Description='Members of this project site and group.'/>"
    }
    $Fld3 = Get-PnPField -List $ListTitle -Identity 'Visitors' -ErrorAction SilentlyContinue
    if (! $Fld3) {
        $Fld3 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='UserMulti' DisplayName='Visitors' List='UserInfo' Required='FALSE' EnforceUniqueValues='FALSE' ShowField='ImnName' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Mult='TRUE' Sortable='FALSE' ID='{1a3dfe68-d2aa-483f-b96c-5eef95cb0982}' StaticName='Visitors' Name='Visitors' Description='Visitors of this project site and group.'/>"
    }
    $Fld4 = Get-PnPField -List $ListTitle -Identity 'Description' -ErrorAction SilentlyContinue
    if (! $Fld4) {
        $Fld4 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='Note' DisplayName='Description' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='FALSE' RichTextMode='Compatible' IsolateStyles='FALSE' Sortable='FALSE' ID='{c86c4ad5-0d6b-437d-8dfa-d61b6bead418}' StaticName='Description' Name='Description' Description='Project Description' RestrictedMode='TRUE' AppendOnly='FALSE'/>"
    }
    # rename Title field
    $Fld = Set-PnPField -List $List -Identity "Title" -Values @{Title = "Project Title"; Description='Title of the project, a SharePoint site will be created using this title.'}

    # updating default view
    $Views = Get-PnPView -List $List
    $Views = Set-PnPView -List $List -Identity $Views[0].Id -Fields "Title", "Owners", "Members", "Visitors"
    $RetVal = $List.Id.Guid.ToString();
    return $RetVal
}
function New-MailList {
    Param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$HubSiteUrl,
        
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$ClientID,
    
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$Certificate,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$TenantID
    )        
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
    
    Write-Information "Creating mail list in Hub site if does not exist."
    $ListTitle = "Project Request EMails"
    $ListDescription = "List for system email for reporting status on new Project Site request."
    #Connecting to site
    Connect-PnPOnline -Url $HubSiteUrl -ClientId $ClientID -CertificateBase64Encoded $Certificate -Tenant $TenantID
    
    # Creating list
    $List = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (! $List) {
        $List = New-PnPList -Title $ListTitle -Url "lists/ProjectStatusEmails" -Template GenericList -EnableVersioning -OnQuickLaunch:$false
        $UpdatedList = Set-PnPList -Identity $ListTitle -Description $ListDescription 
    }
    # Creating fields if does not exist
    $Fld1 = Get-PnPField -List $ListTitle -Identity 'Receiver' -ErrorAction SilentlyContinue
    if (! $Fld1) {
        $Fld1 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='User' DisplayName='Receiver' List='UserInfo' Required='TRUE' EnforceUniqueValues='FALSE' ShowField='ImnName' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Mult='TRUE' Sortable='FALSE' ID='{56a3d5ca-08d5-4248-a65e-65889da08cb3}' StaticName='Receiver' Name='Receiver' Description='Receiver of Project status.'/>"
    }
    $Fld2 = Get-PnPField -List $ListTitle -Identity 'Status' -ErrorAction SilentlyContinue
    if (! $Fld4) {
        $Fld4 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='Note' DisplayName='Status' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='FALSE' RichTextMode='Compatible' IsolateStyles='FALSE' Sortable='FALSE' ID='{c73c4ad5-0d6b-437d-8dfa-d61b6bead418}' StaticName='Status' Name='Status' Description='Project Status' RestrictedMode='TRUE' AppendOnly='FALSE'/>"
    }
    # rename Title field
    $Fld = Set-PnPField -List $List -Identity "Title" -Values @{Title = "EMail Subject"; Description='EMail Subject.'}

    # updating default view
    $Views = Get-PnPView -List $List
    $Views = Set-PnPView -List $List -Identity $Views[0].Id -Fields "Title", "Receiver", "Status"
    $RetVal = $List.Id.Guid.ToString();
    return $RetVal
}
Export-ModuleMember -Function New-RequestList, New-MailList 