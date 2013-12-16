#################################################################################
#           Microsoft SharePoint Server 2013 Managed Metadata Service           #
#                     Taxonomy Backup Script (version 0.1)                      #
#                           by Roydon Gyles-Bedford                             #
#                      based on Bedrich Chaloupka's script                      #
#-------------------------------------------------------------------------------#
#                                                                               #
# This script restores taxonomy data from XML backup file generated by          #
# Taxonomy Backup Script (version 0.1).                                         #
#                                                                               #
# Release Notes:                                                                #
# The script restores complete taxonomy into desired term store.                #
# The items in the [System], [Search Dictionaries], and [People] term groups    #
# are not restored.                                                             #
#                                                                               #
# The script restores the items in the structure with following attributes:     #
#    * TermSet - Name, Contact, Owner, Description, Stakeholders,               #
#      IsAvailableForTagging, IsOpenForTermCreation                             #
#    * Term - Name, Owner, Descriptions and Labels for all available languages, #
#      IsAvailableForTagging, IsDeprecated                                      #
#                                                                               #
# The script does not restore merged ids for a term.                            #
#                                                                               #
#################################################################################

param(
    [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$ClientContext,
    [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$LiteralPath
)
process {

    ##########################################################
    ###  loadXMLFile: checks and load the backup XML file  ###
    ##########################################################

    function loadXMLFile ([string]$fPath) {
        if (Test-Path $fPath) {
            [xml]$xmlDoc = Get-Content $fPath
            $file = Get-ChildItem $fPath
            Write-Host “The backup file: ” $file.Name " has been loaded successfully."
        } else {
            Write-Host -ForegroundColor Red "ERROR: The specified file path does not exist!"
            Write-Host "Please run the script with valid file path or enter valid file path when prompted."
            Write-Host "Example: ./TaxonomyRestore.ps1 d:\backup\2012-03-14_15-27-59_mms_bak.xml"
            Break
        }
        Return $xmlDoc
    }

    ###########################################################################
    ###  restore-TermGroup: restores a termgroup and calls restore-TermSet  ###
    ###########################################################################
    function restore-TermGroup([System.Xml.XmlElement] $xmlTermGroup) {
        $groupName = $xmlTermGroup.GetAttribute("Name")
        $groupId = [guid] $xmlTermGroup.GetAttribute("Id")
        $group = $termStore.CreateGroup($groupName,$groupId)
        $group.Description = $xmlTermGroup.GetAttribute("Description")
        $termStore.CommitAll()
        $ClientContext.load($group)
        $ClientContext.ExecuteQuery()
        Write-Host "The group: " $group.Name " has been restored."

        foreach ($xmlTermSet in $xmlTermGroup.TermSet) {
            restore-TermSet $xmlTermSet
        }
    }

    ####################################################################
    ###  restore-TermSet: restores a termset and calls restore-Term  ###
    ####################################################################
    function restore-TermSet([System.Xml.XmlElement] $xmlTermSet) {
        $termSetName = $xmlTermSet.GetAttribute("Name")
        $termSetId = [guid] $xmlTermSet.GetAttribute("Id")
        $termSet = $group.CreateTermSet($termSetName, $termSetId,$defaultLanguage)
        $termStore.CommitAll()
        $ClientContext.load($termSet)
        $ClientContext.ExecuteQuery()
        if ($xmlTermSet.GetAttribute("IsOpenForTermCreation") -eq "True") {$termSet.IsOpenForTermCreation = $True}
        if ($xmlTermSet.GetAttribute("IsAvailableForTagging") -eq "False") {$termSet.IsAvailableForTagging = $False}
        if ($xmlTermSet.GetAttribute("Contact")) {$termSet.Contact = $xmlTermSet.GetAttribute("Contact")}
        foreach ($stakeholders in $xmlTermSet.Stakeholders) {
            $stakeholders = $stakeholders.Split()
            foreach ($stakeholder in $stakeholders) {
                $termSet.AddStakeholder($stakeholder)}
        }
        foreach ($CustomProperty in $xmlTermSet.CustomProperty) {
            $termSet.SetCustomProperty($CustomProperty.Key,$CustomProperty.InnerText)
        }
        $termSet.Description = $xmlTermSet.GetAttribute("Description")
        $termSet.Owner = $xmlTermSet.GetAttribute("Owner")
        $termStore.CommitAll()
        $ClientContext.ExecuteQuery()
        Write-Host "The term set: " $termSet.Name " has been restored."

        foreach($xmlTerm in $xmlTermSet.Term) {
            restore-Term $xmlTerm
        }
        if ($xmlTermSet.CustomSortOrder) {
            $termSet.CustomSortOrder = $xmlTermSet.CustomSortOrder.InnerText
            $termStore.CommitAll()
            $ClientContext.ExecuteQuery()
        }
    }

    ######################################################################################
    ###  restore-Term: restores terms in the root level and calls restore-ChildTerm    ###
    ######################################################################################

    function restore-Term ([System.Xml.XmlElement] $xmlTerm) {
        foreach ($label in $xmlTerm.Label) {
            $labelLanguage = $label.GetAttribute("Language")
            if (($label.GetAttribute("IsDefaultForLanguage") -eq "True") -and ($labelLanguage -eq $defaultLanguage)) {
                $labelName = $label.InnerText
                $termId = [guid]$xmlTerm.GetAttribute("Id")
                $term = $termSet.CreateTerm($labelName, $labelLanguage, $termId)
                $termStore.CommitAll()
                $ClientContext.load($term)
                $ClientContext.ExecuteQuery()
            }    
        }

        if ($xmlTerm.GetAttribute("IsAvailableForTagging") -eq "False") {$term.IsAvailableForTagging = $False}
        $term.Owner = $xmlTerm.GetAttribute("Owner")
        foreach ($label in $xmlTerm.Label) {
            $labelLanguage = $label.GetAttribute("Language")
            $labelName = $label.InnerText
            $isDefaultLabel = $False
            if ($label.GetAttribute("IsDefaultForLanguage") -eq "True") {$isDefaultLabel = $True}
            if (!(($isDefaultLabel -eq $True) -and ($labelLanguage -eq $defaultLanguage))) {
                $dummyLabel = $term.CreateLabel($labelName, $labelLanguage, $isDefaultLabel) 
            }    
        }
        if ($xmlTerm.Description) {
            foreach ($description in $xmlTerm.Description) {
                $descriptionLanguage = $description.GetAttribute("Language")
                $descriptionText = $description.InnerText
                $term.SetDescription($descriptionText, $descriptionLanguage)
            }
        }    
        if ($xmlTerm.GetAttribute("IsDeprecated") -eq "True") {$term.Deprecate($True)}
        foreach ($CustomProperty in $xmlTerm.CustomProperty) {
            $term.SetCustomProperty($CustomProperty.Key,$CustomProperty.InnerText)
        }
        foreach ($LocalCustomProperty in $xmlTerm.LocalCustomProperty) {
            $term.SetLocalCustomProperty($LocalCustomProperty.Key,$LocalCustomProperty.InnerText)
        }
        Write-Host "The term " $term.Name " has been restored as root level item"
        $termStore.CommitAll()
        $ClientContext.ExecuteQuery()

        foreach($xmlChildTerm in $xmlTerm.ChildTerms.Term) {
            restore-ChildTerm -xmlTerm $xmlChildTerm -parentTerm $term
        }
        if ($xmlTerm.CustomSortOrder) {
            $term.CustomSortOrder = $xmlTerm.CustomSortOrder.InnerText
            $termStore.CommitAll()
            $ClientContext.ExecuteQuery()
        }
    }
    #########################################################################################
    ###  restore-ChildTerm: restores child terms and recusivly calls restore-ChildTerm    ###
    #########################################################################################

    function restore-ChildTerm ([System.Xml.XmlElement] $xmlTerm, [Microsoft.SharePoint.Client.Taxonomy.Term] $parentTerm) {
        foreach ($label in $xmlTerm.Label) {
            $labelLanguage = $label.GetAttribute("Language")
            if (($label.GetAttribute("IsDefaultForLanguage") -eq "True") -and ($labelLanguage -eq $defaultLanguage)) {
                $labelName = $label.InnerText
                $termId = [guid]$xmlTerm.GetAttribute("Id")
                $term = $parentTerm.CreateTerm($labelName, $labelLanguage, $termId)
                $termStore.CommitAll()
                $ClientContext.load($term)
                $ClientContext.ExecuteQuery()
            }    
        }

        if ($xmlTerm.GetAttribute("IsAvailableForTagging") -eq "False") {$term.IsAvailableForTagging = $False}
        $term.Owner = $xmlTerm.GetAttribute("Owner")
        foreach ($label in $xmlTerm.Label) {
            $labelLanguage = $label.GetAttribute("Language")
            $labelName = $label.InnerText
            $isDefaultLabel = $False
            if ($label.GetAttribute("IsDefaultForLanguage") -eq "True") {$isDefaultLabel = $True}
            if (!(($isDefaultLabel -eq $True) -and ($labelLanguage -eq $defaultLanguage))) {
                $dummyLabel = $term.CreateLabel($labelName, $labelLanguage, $isDefaultLabel) 
            }    
        }
        if ($xmlTerm.Description) {
            foreach ($description in $xmlTerm.Description) {
                $descriptionLanguage = $description.GetAttribute("Language")
                $descriptionText = $description.InnerText
                $term.SetDescription($descriptionText, $descriptionLanguage)
            }
        }
        if ($xmlTerm.GetAttribute("IsDeprecated") -eq "True") {$term.Deprecate($True)}
        foreach ($CustomProperty in $xmlTerm.CustomProperty) {
            $term.SetCustomProperty($CustomProperty.Key,$CustomProperty.InnerText)
        }
        foreach ($LocalCustomProperty in $xmlTerm.LocalCustomProperty) {
            $term.SetLocalCustomProperty($LocalCustomProperty.Key,$LocalCustomProperty.InnerText)
        }
        Write-Host "The term " $term.Name " has been restored as a child"
        $termStore.CommitAll()
        $ClientContext.ExecuteQuery()

        foreach($xmlChildTerm in $xmlTerm.ChildTerms.Term) {
            restore-ChildTerm -xmlTerm $xmlChildTerm -parentTerm $term
        }
        if ($xmlTerm.CustomSortOrder) {
            $term.CustomSortOrder = $xmlTerm.CustomSortOrder.InnerText
            $termStore.CommitAll()
            $ClientContext.ExecuteQuery()
        }
    }
   

    ###  Checks if exists file path argument  ###
    $xmlData = loadXMLFile $LiteralPath

    ###  Gets the term store object  ###
    $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ClientContext)


    ###  Gets term store object based on the name from backup file (might need to be rewritten to get it from sesssion!)  ###
    foreach($xmltermStore in $xmlData.TermStores.TermStore) {
        $termStoreName = $xmltermStore.GetAttribute("Name")
        $termStore = $session.TermStores.GetByName($termStoreName);
        $ClientContext.Load($termStore)
        $ClientContext.ExecuteQuery()

        $defaultLanguage = $termStore.DefaultLanguage
        foreach($xmlTermGroup in $xmltermStore.Group) {
            $groupName = $xmlTermGroup.GetAttribute("Name")
            if($groupName -eq "System" -or $groupName -eq "Search Dictionaries" -or $groupName -eq "People") {
                Write-Host "Skipping Term Group: $groupName"
            } else {
                restore-TermGroup $xmlTermGroup
            }
        }
    }

}
end {}
####################################
###  TaxonomyRestore script END  ###
####################################