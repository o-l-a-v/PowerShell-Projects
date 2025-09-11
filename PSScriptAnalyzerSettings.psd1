@{
    IncludeDefaultRules = $true
    ExcludeRules        = @(
        'PSPossibleIncorrectUsageOfAssignmentOperator',
        'PSUseBOMForUnicodeEncodedFile'
    )
    Rules               = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.4')
        }
    }
}
