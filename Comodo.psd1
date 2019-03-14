@{
    # Point to your module psm1 file...
    RootModule = 'Comodo.psm1'

    # Be sure to specify a version
    ModuleVersion = '0.1.0'

    Description = 'Comodo Service Desk Plugin for PoshBot'
    Author = 'BeverlyHuff'
    CompanyName = 'KAKNets'
    Copyright = 'KAKNets'
    PowerShellVersion = '5.0.0'

    # Generate your own GUID
    GUID = '3d0d33cd-bee5-4c20-99e3-fff9aab222b8'

    # We require poshbot...
    RequiredModules = @('PoshBot')

    # Ideally, define these!
    FunctionsToExport = '*'

    PrivateData = @{
        # These are permissions we'll expose in our poshbot module
        Permissions = @(
            @{
                Name = 'status'
                Description = 'Get Comodo Ticket Status'
            }
            @{
                Name = 'reply'
                Description = 'Reply to Comodo Tickets'
            }
            @{
                Name = 'close'
                Description = 'Close Comodo Tickets'
            }

        )
    } # End of PrivateData hashtable
}