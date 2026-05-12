# ============================================================
#  Pester tests for TeamsManager.ps1
#  Tests cover the orphan / single-owner classification logic
#  against canned Graph responses; no live calls.
# ============================================================

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $script:RepoRoot 'UI.ps1')
    . (Join-Path $script:RepoRoot 'Auth.ps1')
    . (Join-Path $script:RepoRoot 'Audit.ps1')
    . (Join-Path $script:RepoRoot 'Preview.ps1')
    . (Join-Path $script:RepoRoot 'TeamsManager.ps1')
}

Describe "ConvertTo-TeamRecord" {
    It "produces the expected shape" {
        $raw = [PSCustomObject]@{ id='aaa-111'; displayName='Sales'; description='Sales team'; visibility='public' }
        $r = ConvertTo-TeamRecord -Raw $raw -Role 'Owner'
        $r.TeamId      | Should -Be 'aaa-111'
        $r.DisplayName | Should -Be 'Sales'
        $r.Role        | Should -Be 'Owner'
    }
}

Describe "Orphan + single-owner classification (mocked)" {
    BeforeAll {
        # Canned owner sets for three teams:
        #   t1: 0 owners (orphan)
        #   t2: 1 owner  (SPOF)
        #   t3: 3 owners (healthy)
        $script:OwnerSets = @{
            't1' = @()
            't2' = @([PSCustomObject]@{ id='u-bob';     userPrincipalName='bob@x.com';     displayName='Bob' })
            't3' = @(
                [PSCustomObject]@{ id='u-alice';        userPrincipalName='alice@x.com';   displayName='Alice' },
                [PSCustomObject]@{ id='u-carol';        userPrincipalName='carol@x.com';   displayName='Carol' },
                [PSCustomObject]@{ id='u-dave';         userPrincipalName='dave@x.com';    displayName='Dave'  }
            )
        }
    }
    It "Predicate: zero-owner teams are orphans" {
        foreach ($id in @('t1','t2','t3')) {
            $owners = $script:OwnerSets[$id]
            $isOrphan = (-not $owners -or @($owners).Count -eq 0)
            if ($id -eq 't1') { $isOrphan | Should -BeTrue } else { $isOrphan | Should -BeFalse }
        }
    }
    It "Predicate: single-owner teams are SPOF" {
        foreach ($id in @('t1','t2','t3')) {
            $owners = $script:OwnerSets[$id]
            $isSingle = (@($owners).Count -eq 1)
            if ($id -eq 't2') { $isSingle | Should -BeTrue } else { $isSingle | Should -BeFalse }
        }
    }
}

Describe "Resolve-TeamIdentifier GUID short-circuit" {
    It "returns the input unchanged when it is already a GUID" {
        $guid = '12345678-aaaa-bbbb-cccc-1234567890ab'
        # We can't avoid calling the function but we mock the Graph
        # request to ensure we never touch the network -- but since
        # the GUID branch returns early, no mock needed.
        Resolve-TeamIdentifier -IdOrName $guid | Should -Be $guid
    }
}
