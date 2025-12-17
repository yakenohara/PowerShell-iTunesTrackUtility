# <CAUTION!>
# このファイルは UTF-8 (BOM 付き) で保存すること。
# </CAUTION!>

# Private
Get-ChildItem $PSScriptRoot\Private\*.ps1 -Recurse |
    ForEach-Object { . $_ }

# Public
Get-ChildItem $PSScriptRoot\Public\*.ps1 -Recurse |
    ForEach-Object { . $_ }

Export-ModuleMember -Function (Get-ChildItem $PSScriptRoot\Public\*.ps1).BaseName
