#################################################################################
#           Microsoft SharePoint Server 2013 Managed Metadata Service           #
#                     Taxonomy Backup Script (version 0.1)                      #
#                           by Roydon Gyles-Bedford                             #
#                      based on Bedrich Chaloupka's script                      #
#-------------------------------------------------------------------------------#
#                                                                               #
# This script backs-up taxonomy data to a XML backup file.                      #
#                                                                               #
# Release Notes:                                                                #
# The script backs-up complete taxonomy of the term store including items       #
# in the [System] term group. The script backs-up the items and following       #
# attributes:                                                                   #
#    * TermSet - Name, Contact, Owner, Description, Stakeholders,               #
#      IsAvailableForTagging, IsOpenForTermCreation, LastModifiedDate,          #
#      CreatedDate, ID, CustomSortOrder                                         # 
#    * Term - Name, Owner, Descriptions and Labels for all available languages, #
#      IsAvailableForTagging, IsDeprecated, LastModifiedDate, CreatedDate, ID,  #
#      IsReused, ParentTermId                                                   #
#                                                                               #
# The script does not back-up merged ids for a term.                            #
#                                                                               #  
#                                                                               #
#################################################################################  

param(
    [parameter(Mandatory=$true, ValueFromPipeline=$true)][Microsoft.SharePoint.Client.ClientContext]$ClientContext,
    [parameter(Mandatory=$true, ValueFromPipeline=$true)][string]$LiteralPath
)

process {

    ########################################################
    ###  checkPath: checks and load the backup XML file  ###
    ########################################################


    function check-Path ([string]$fPath) {
        if (!(Test-Path $fPath)) {
            Write-Host -ForegroundColor Red "ERROR: The specified file path does not exist!"
            Write-Host "Please run the script with valid file path or enter valid file path when prompted."
            Write-Host "The file name will be generated automatically with this pattern: yyyy-M-d_H-m-s_mms_bak.xml"
            Write-Host "Example: ./TaxonomyBackup.ps1 -Path d:\backup\"
            Break
        }
    }

    #######################################################################################
    ###  saveChildTerms: walks through the term structure under a term in root level    ###
    ###  and generates <Term> elements in the back-up XML file                          ###
    #######################################################################################
    function loadChildTerms ([Microsoft.SharePoint.Client.Taxonomy.Term] $parentTerm) {
        Write-Host "`t`t`t`tLoading Child Terms for $($parentTerm.Name) ..."
        foreach ($term in $parentTerm.Terms) {
            if ($term.TermsCount -gt 0) {
                $ClientContext.Load($term.Terms)
            }
        }
        $ClientContext.ExecuteQuery()
        foreach ($term in $parentTerm.Terms) {
            if ($term.TermsCount -gt 0) {
                loadChildTerms -parentTerm $term
            }
        }
    }
    function loadChildTermsExtraInfo ([Microsoft.SharePoint.Client.Taxonomy.Term] $parentTerm) {
        Write-Host "`t`t`t`tLoading Child Terms Extra Info for $($parentTerm.Name) ..."
        $count = 0
        foreach ($term in $parentTerm.Terms) {
            $ClientContext.Load($term.SourceTerm)
            $ClientContext.Load($term.Labels)
            $count += 1
            # Reduce the number of round trips needed
            if($count % 50 -eq 0) {
                $ClientContext.ExecuteQuery()
            }
            if ($term.TermsCount -gt 0) {loadChildTermsExtraInfo -parentTerm $term}
        }
        $ClientContext.ExecuteQuery()
    }
    function saveChildTerms ([Microsoft.SharePoint.Client.Taxonomy.Term] $parentTerm, [System.Xml.XmlElement] $xmlParentTerm) {

        Write-Host "`t`t`t`tSaving ChildTerms for $($parentTerm.Name) ..."
        $xmlChildTerms = $xmlDoc.CreateElement("ChildTerms")
        $xmlParentTerm.AppendChild($xmlChildTerms) | Out-Null

        #$childTerms = $parentTerm.GetTerms($childs)
        foreach ($term in $parentTerm.Terms){ # $childTerms) {
            $xmlTerm = $xmlDoc.CreateElement("Term")
            $xmlTerm.SetAttribute("Name",$term.Name)
            $xmlTerm.SetAttribute("Id",$term.Id.toString())
            $xmlTerm.SetAttribute("Owner",$term.Owner)
            $xmlTerm.SetAttribute("CreatedDate",$term.CreatedDate)
            $xmlTerm.SetAttribute("LastModifiedDate",$term.LastModifiedDate)
            $xmlTerm.SetAttribute("IsAvailableForTagging",$term.IsAvailableForTagging.toString())
            $xmlTerm.SetAttribute("IsDeprecated",$term.IsDeprecated.toString())
            $xmlTerm.SetAttribute("ParentTermId",$parentTerm.Id.toString())
            $xmlTerm.SetAttribute("IsReused",$term.IsReused.toString())
            if ($term.IsReused -eq $true) {
                $xmlTerm.SetAttribute("IsReused",$term.IsReused.toString())
                $xmlTerm.SetAttribute("SourceTermId",$term.SourceTerm.Id.toString())
            }
            if ([string]$term.CustomSortOrder -ne "") {
                $xmlSortOrder = $xmlDoc.CreateElement("CustomSortOrder")
                $xmlSortOrder.InnerText = $term.CustomSortOrder
                $xmlTerm.AppendChild($xmlSortOrder) | Out-Null
            }
            foreach($key in $term.LocalCustomProperties.Keys){
                $xmlCustomProperty = $xmlDoc.CreateElement("LocalCustomProperty")
                $xmlCustomProperty.InnerText = $term.LocalCustomProperties[$key]
                $xmlCustomProperty.SetAttribute("Key",$key)
                $xmlTerm.AppendChild($xmlCustomProperty) | Out-Null
            }
            foreach($key in $term.CustomProperties.Keys){
                $xmlCustomProperty = $xmlDoc.CreateElement("CustomProperty")
                $xmlCustomProperty.InnerText = $term.CustomProperties[$key]
                $xmlCustomProperty.SetAttribute("Key",$key)
                $xmlTerm.AppendChild($xmlCustomProperty) | Out-Null
            }
            foreach ($mergedTermId in $term.MergedTermIds) {
                $xmlMergedTermId = $xmlDoc.CreateElement("MergedTermId")
                $xmlMergedTermId.InnerText = $mergedTermId
                $xmlTerm.AppendChild($xmlMergedTermId) | Out-Null
            }
            foreach ($label in $term.Labels) {
                $xmlLabel = $xmlDoc.CreateElement("Label")
                $xmlLabel.InnerText = $label.Value
                $xmlLabel.SetAttribute("Language",$label.Language)
                $xmlLabel.SetAttribute("IsDefaultForLanguage",$label.IsDefaultForLanguage.toString())
                $xmlTerm.AppendChild($xmlLabel) | Out-Null
            }
            foreach ($language in $termStore.Languages) {
                if ($term.GetDescription($language).Value) { 
                    $xmlDescription = $xmlDoc.CreateElement("Description")
                    $xmlDescription.SetAttribute("Language",$language)
                    $xmlDescription.InnerText = $term.GetDescription($language).Value
                    $xmlTerm.AppendChild($xmlDescription) | Out-Null
                }                  
            }
            $xmlChildTerms.AppendChild($xmlTerm) | Out-Null
            if ($term.TermsCount -gt 0) {saveChildTerms -parentTerm $term -xmlparentTerm $xmlTerm}
        
        }
    
    }

    ###  Checks if exists file path argument  ###
    check-Path $LiteralPath


    ###  Initiates the XML file structure  ###
    [xml]$xmlDoc = "<?xml version=""1.0"" encoding=""utf-8""?><TermStores></TermStores>"

    ###  Gets the term store object  ###
    $session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ClientContext)
    Write-Host "Loading TermStores ..."
    $ClientContext.Load($session.TermStores)
    $ClientContext.ExecuteQuery()

    foreach ($termStore in $session.TermStores) {
        Write-Host "Saving TermStore $($termStore.Name) ..."
        $xmlStore = $xmlDoc.CreateElement("TermStore")
        $xmlStore.SetAttribute("Name",$termStore.Name)
        $xmlStore.SetAttribute("ContentTypePublishingHub",$termStore.ContentTypePublishingHub)
        $xmlStore.SetAttribute("DefaultLanguage",$termStore.DefaultLanguage)
        $xmlStore.SetAttribute("Id",$termStore.Id)
        $xmlStore.SetAttribute("WorkingLanguage",$termStore.WorkingLanguage)
        $xmlDoc.DocumentElement.AppendChild($xmlStore) | Out-Null

        Write-Host "`tLoading Groups for $($termStore.Name) ..."
        $ClientContext.Load($termStore.Groups)
        $ClientContext.ExecuteQuery()
        foreach ($group in $termStore.Groups) {
            Write-Host "`t`tLoading TermSets for $($group.Name) ..."
            $ClientContext.Load($group.TermSets)
            $ClientContext.ExecuteQuery()
            foreach ($termSet in $group.TermSets) {
                Write-Host "`t`t`tLoading Terms for $($termSet.Name) ..."
                $ClientContext.Load($termSet.Terms)
            }
            $ClientContext.ExecuteQuery()
            
            foreach ($termSet in $group.TermSets) {
                foreach ($term in $termSet.Terms) {
                    if ($term.TermsCount -gt 0) {
                        $ClientContext.Load($term.Terms)
                    }
                }
            }
            $ClientContext.ExecuteQuery()
            foreach ($termSet in $group.TermSets) {
                foreach ($term in $termSet.Terms) {
                    if ($term.TermsCount -gt 0) {
                        loadChildTerms -parentTerm $term
                    }
                }
            }
            $count = 0
            foreach ($termSet in $group.TermSets) {
                foreach ($term in $termSet.Terms) {
                    $ClientContext.Load($term.SourceTerm)
                    $ClientContext.Load($term.Labels)
                    $count += 1
                    # Reduce the number of round trips needed
                    if($count % 50 -eq 0) {
                        $ClientContext.ExecuteQuery()
                    }
                    if ($term.TermsCount -gt 0) {loadChildTermsExtraInfo -parentTerm $term}
                }
            }
            $ClientContext.ExecuteQuery()
        }
        foreach ($group in $termStore.Groups) {
            Write-Host "`tSaving Group $($group.Name) ..."
            $xmlGroup = $xmlDoc.CreateElement("Group")
            $xmlGroup.SetAttribute("Name",$group.Name)
            $xmlGroup.SetAttribute("Id",$group.Id.toString())
            $xmlGroup.SetAttribute("Description",$group.Description)
            $xmlGroup.SetAttribute("CreatedDate",$group.CreatedDate)
            $xmlGroup.SetAttribute("LastModifiedDate",$group.LastModifiedDate)
            $xmlGroup.SetAttribute("IsSiteCollectionGroup",$group.IsSiteCollectionGroup.toString())
            $xmlGroup.SetAttribute("IsSystemGroup",$group.IsSystemGroup.toString())
            $xmlStore.AppendChild($xmlGroup) | Out-Null

            foreach ($termSet in $group.TermSets) {
                    Write-Host "`t`tSaving TermSet $($termSet.Name) ..."
                    $xmlTermSet = $xmlDoc.CreateElement("TermSet")
                    $xmlTermSet.SetAttribute("Name",$termSet.Name)
                    $xmlTermSet.SetAttribute("Id",$termSet.Id.toString())
                    if ($termSet.Owner -ne "") {$xmlTermSet.SetAttribute("Owner",$termSet.Owner)}
                    if ($termSet.Contact -ne "") {$xmlTermSet.SetAttribute("Contact",$termSet.Contact)}
                    $xmlTermSet.SetAttribute("Description",$termSet.Description)
                    $xmlTermSet.SetAttribute("CreatedDate",$termSet.CreatedDate)
                    $xmlTermSet.SetAttribute("LastModifiedDate",$termSet.LastModifiedDate)
                    $xmlTermSet.SetAttribute("IsAvailableForTagging",$termSet.IsAvailableForTagging.toString())
                    $xmlTermSet.SetAttribute("IsOpenForTermCreation",$termSet.IsOpenForTermCreation.toString())  
                    if ($termSet.Stakeholders -ne "") { 
                        $xmlStakeholders = $xmlDoc.CreateElement("Stakeholders")
                        $xmlStakeholders.InnerText = $termSet.Stakeholders
                        $xmlTermSet.AppendChild($xmlStakeholders) | Out-Null
                    }
                    if ([string]$termSet.CustomSortOrder -ne "") {
                        $xmlSortOrder = $xmlDoc.CreateElement("CustomSortOrder")
                        $xmlSortOrder.InnerText = $termSet.CustomSortOrder
                        $xmlTermSet.AppendChild($xmlSortOrder) | Out-Null
                    }
                    foreach($key in $termSet.CustomProperties.Keys){
                        $xmlCustomProperty = $xmlDoc.CreateElement("CustomProperty")
                        $xmlCustomProperty.InnerText = $termSet.CustomProperties[$key]
                        $xmlCustomProperty.SetAttribute("Key",$key)
                        $xmlTermSet.AppendChild($xmlCustomProperty) | Out-Null
                    }
                    $xmlGroup.AppendChild($xmlTermSet) | Out-Null


                    Write-Host "`t`t`tSaving Terms for $($termSet.Name) ..."
                    foreach ($term in $termSet.Terms) {
                        $xmlTerm = $xmlDoc.CreateElement("Term")
                        $xmlTerm.SetAttribute("Name",$term.Name)
                        $xmlTerm.SetAttribute("Id",$term.Id.toString())
                        $xmlTerm.SetAttribute("Owner",$term.Owner)
                        $xmlTerm.SetAttribute("CreatedDate",$term.CreatedDate)
                        $xmlTerm.SetAttribute("LastModifiedDate",$term.LastModifiedDate)
                        $xmlTerm.SetAttribute("IsAvailableForTagging",$term.IsAvailableForTagging.toString())
                        $xmlTerm.SetAttribute("IsDeprecated",$term.IsDeprecated.toString())
                        if ($term.IsReused -eq $true) {
                            $xmlTerm.SetAttribute("IsReused",$term.IsReused.toString())
                            $xmlTerm.SetAttribute("SourceTermId",$term.SourceTerm.Id.toString())
                        }
                        if ($term.IsKeyword -eq $true) {
                            $xmlTerm.SetAttribute("IsKeyword",$term.IsKeyword.toString())
                        }
                        if ([string]$term.CustomSortOrder -ne "") {
                            $xmlSortOrder = $xmlDoc.CreateElement("CustomSortOrder")
                            $xmlSortOrder.InnerText = $term.CustomSortOrder
                            $xmlTerm.AppendChild($xmlSortOrder) | Out-Null
                        }
                        foreach($key in $term.CustomProperties.Keys){
                            $xmlCustomProperty = $xmlDoc.CreateElement("CustomProperty")
                            $xmlCustomProperty.InnerText = $term.CustomProperties[$key]
                            $xmlCustomProperty.SetAttribute("Key",$key)
                            $xmlTerm.AppendChild($xmlCustomProperty) | Out-Null
                        }
                        foreach($key in $term.LocalCustomProperties.Keys){
                            $xmlCustomProperty = $xmlDoc.CreateElement("LocalCustomProperty")
                            $xmlCustomProperty.InnerText = $term.LocalCustomProperties[$key]
                            $xmlCustomProperty.SetAttribute("Key",$key)
                            $xmlTerm.AppendChild($xmlCustomProperty) | Out-Null
                        }
                        foreach ($mergedTermId in $term.MergedTermIds) {
                            $xmlMergedTermId = $xmlDoc.CreateElement("MergedTermId")
                            $xmlMergedTermId.InnerText = $mergedTermId
                            $xmlTerm.AppendChild($xmlMergedTermId) | Out-Null
                        }

                        foreach ($label in $term.Labels) {
                            $xmlLabel = $xmlDoc.CreateElement("Label")
                            $xmlLabel.InnerText = $label.Value
                            $xmlLabel.SetAttribute("Language",$label.Language)
                            $xmlLabel.SetAttribute("IsDefaultForLanguage",$label.IsDefaultForLanguage.toString())
                            $xmlTerm.AppendChild($xmlLabel) | Out-Null
                        }

                        foreach ($language in $termStore.Languages) {
                            if ($term.GetDescription($language).Value) { 
                                $xmlDescription = $xmlDoc.CreateElement("Description")
                                $xmlDescription.SetAttribute("Language",$language)
                                $xmlDescription.InnerText = $term.GetDescription($language).Value
                                $xmlTerm.AppendChild($xmlDescription) | Out-Null
                            }                  
                        }
                    
                        $xmlTermSet.AppendChild($xmlTerm) | Out-Null
                        if ($term.TermsCount -gt 0) {saveChildTerms -parentTerm $term -xmlparentTerm $xmlTerm}

                    }
                }
        }
    }


    ###  Generates the file name and saves the XML backup file to the specified path  ###
    if (!($LiteralPath.EndsWith("\"))) {
        $LiteralPath = $LiteralPath + "\"
    }
    $fileDateTime = Get-Date -Format yyyy-M-d_H-m-s
    $outputFileName = $fileDateTime + "_mms_bak.xml"
    $xmlDoc.Save($LiteralPath + $outputFileName)

}
end {}
####################################
###  TaxonomyBackup script END  ###
####################################