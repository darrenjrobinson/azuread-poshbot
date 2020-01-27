@{
    RootModule = 'PoshBot.AzureAD.psm1'
    ModuleVersion = '1.0.0'
    
    Description = 'PoshBot module Azure Active Directory'
    Author = 'Darren J Robinson'
    CompanyName = 'Community'
    Copyright = '(c) 2020 Darren J Robinson. All rights reserved.'
    PowerShellVersion = '5.0.0'
    
    GUID = 'd964e35b-4815-433a-a86c-a50ca5e09a84'
    
    RequiredModules = @('PoshBot','MSAL.PS')
    FunctionsToExport = '*'
    
    PrivateData = @{
        # These are permissions we'll expose in our poshbot module even though version 1.0.0 only provides Read Functions
        Permissions = @(
            @{
                Name = 'read'
                Description = 'Run commands that have Read Only access to the MIM Service'
            }
            @{
                Name = 'write'
                Description = 'Run commands that have Write access to the MIM Service'
            }
        )
    } # End of PrivateData hashtable
}
