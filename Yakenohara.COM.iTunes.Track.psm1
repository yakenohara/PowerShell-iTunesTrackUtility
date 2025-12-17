# <CAUTION!>
# このファイルは UTF-8 (BOM 付き) で保存すること。
# </CAUTION!>

# Private
# Get-ChildItem $PSScriptRoot\Private\*.ps1 -Recurse |
#     ForEach-Object { . $_.FullName }

# Public
Get-ChildItem $PSScriptRoot\Public\*.ps1 -Recurse |
    ForEach-Object { . $_.FullName }

