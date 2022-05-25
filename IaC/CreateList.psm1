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
        [String]$Secret,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$TenantID

    )        
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
    
    Write-Information "Creating request list in Hub site if does not exist."
    $ListTitle = "OIP Request"
    $ListDescription = "List for requesting new OIP Site."
    #Connecting to site
    Connect-PnPOnline -Url $GovernanceSiteUrl -ClientId $ClientID -Tenant $TenantID -CertificateBase64Encoded $Secret
    
    # Creating list
    $List = Get-PnPList -Identity $ListTitle -ErrorAction SilentlyContinue
    if (! $List) {
        $List = New-PnPList -Title $ListTitle -Url "lists/OIPRequest" -Template GenericList -EnableVersioning -OnQuickLaunch:$false
        $UpdatedList = Set-PnPList -Identity $ListTitle -Description $ListDescription 
    }
    # Creating fields if does not exist
    $Fld1 = Get-PnPField -List $ListTitle -Identity 'LibraryTitle' -ErrorAction SilentlyContinue
    if (! $Fld1) {
        $Fld1 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='Text' DisplayName='Library/List Title' Required='TRUE' MaxLength='255' ID='{63e74288-d95c-4569-8461-6e2b5aa8e0fe}' Name='LibraryTitle' Description='Title of SharePoint List or Library where policy will be enforced.'/>"
    }
    $Fld2 = Get-PnPField -List $ListTitle -Identity 'Policy' -ErrorAction SilentlyContinue
    if (! $Fld2) {
        $Fld2 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Name='Policy' Type='Choice' DisplayName='Policy Name' Description='Name of policy which will be enforced.' Required='TRUE' Format='Dropdown' FillInChoice='FALSE' ID='{4e992580-6539-41c4-8830-f5b002741e90}' ><Default>ContentOwnership</Default><CHOICES><CHOICE>ContentOwnership</CHOICE><CHOICE>ContentValidity</CHOICE><CHOICE>ContentArchaival</CHOICE><CHOICE>ContentDelete</CHOICE></CHOICES></Field>"
    }
    $Fld3 = Get-PnPField -List $ListTitle -Identity 'FieldDefinitions' -ErrorAction SilentlyContinue
    if (! $Fld3) {
        $Fld3 = Add-PnPFieldFromXml -List $List -FieldXml "<Field Type='Note' DisplayName='Field Definitions' Required='FALSE' NumLines='6' RichText='FALSE' ID='{eeb4c696-2f1f-4a8b-ae42-38b49d6d0493}' Name='FieldDefinitions' RichTextMode='Compatible' Description='XML Schema of fields which will be added in List/Library. These field will be added by enable policy enforcement.' />"
    }

    # rename Title field
    $Fld = Set-PnPField -List $List -Identity "Title" -Values @{Title = "Site Url"; Description='SharePoint site URL where policy will be enforced.'}

    # updating default view
    $Views = Get-PnPView -List $List
    $Views = Set-PnPView -List $List -Identity $Views[0].Id -Fields "Title", "Library/List Title", "Policy", "Field Definitions"
    $RetVal = $List.Id.Guid.ToString();
    return $RetVal
}
Export-ModuleMember -Function New-RequestList