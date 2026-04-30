# Tests/MetadataOps.Tests.ps1
# Parameter-validation tests for MetadataOps functions (no COM required)

BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\ExcelPOSH\ExcelPOSH.psd1'
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module (Resolve-Path $modulePath).Path -Force -ErrorAction Stop
}

AfterAll {
    Get-Module ExcelPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'Get-ExcelDocumentProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelDocumentProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelDocumentProperty).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has PropertyName parameter (optional)' {
        (Get-Command Get-ExcelDocumentProperty).Parameters['PropertyName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Custom switch' {
        (Get-Command Get-ExcelDocumentProperty).Parameters['Custom'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelDocumentProperty' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelDocumentProperty).CmdletBinding | Should -BeTrue
    }
    It 'Has PropertyName parameter' {
        (Get-Command Set-ExcelDocumentProperty).Parameters['PropertyName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Value parameter' {
        (Get-Command Set-ExcelDocumentProperty).Parameters['Value'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Custom switch' {
        (Get-Command Set-ExcelDocumentProperty).Parameters['Custom'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelConnection' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelConnection).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelConnection).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelProtection' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelProtection).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Get-ExcelProtection).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter (optional)' {
        (Get-Command Get-ExcelProtection).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelProtection' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelProtection).CmdletBinding | Should -BeTrue
    }
    It 'Has WorkbookPath parameter' {
        (Get-Command Set-ExcelProtection).Parameters['WorkbookPath'] | Should -Not -BeNullOrEmpty
    }
    It 'Has SheetName parameter (optional)' {
        (Get-Command Set-ExcelProtection).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Password parameter' {
        (Get-Command Set-ExcelProtection).Parameters['Password'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Unprotect switch' {
        (Get-Command Set-ExcelProtection).Parameters['Unprotect'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelComment' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelComment).CmdletBinding | Should -BeTrue
    }
    It 'Has SheetName parameter' {
        (Get-Command Get-ExcelComment).Parameters['SheetName'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Range parameter (optional)' {
        (Get-Command Get-ExcelComment).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Set-ExcelComment' {
    It 'Has CmdletBinding' {
        (Get-Command Set-ExcelComment).CmdletBinding | Should -BeTrue
    }
    It 'Has Range parameter' {
        (Get-Command Set-ExcelComment).Parameters['Range'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Text parameter' {
        (Get-Command Set-ExcelComment).Parameters['Text'] | Should -Not -BeNullOrEmpty
    }
    It 'Has Remove switch' {
        (Get-Command Set-ExcelComment).Parameters['Remove'] | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-ExcelTip' {
    It 'Has CmdletBinding' {
        (Get-Command Get-ExcelTip).CmdletBinding | Should -BeTrue
    }
    It 'Has AsJson switch' {
        (Get-Command Get-ExcelTip).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
    }
    It 'Returns a tip without COM' {
        $result = Get-ExcelTip
        $result.tip | Should -Not -BeNullOrEmpty
    }
    It 'Returns JSON when AsJson is specified' {
        $result = Get-ExcelTip -AsJson
        $result | Should -BeOfType [string]
        { $result | ConvertFrom-Json } | Should -Not -Throw
    }
}
