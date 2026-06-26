#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputDirectory = "manual"
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$RepositoryRoot = Split-Path -Parent $PSScriptRoot
$ManualDirectory = Join-Path -Path $RepositoryRoot -ChildPath $OutputDirectory
$ManualMarkdown = Join-Path -Path $ManualDirectory -ChildPath "AdminTools-Manual.md"
$ManualHtml = Join-Path -Path $ManualDirectory -ChildPath "AdminTools-Manual.html"
$ManualDocx = Join-Path -Path $ManualDirectory -ChildPath "AdminTools-Manual.docx"
$TemporaryDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("admintools-docx-" + [guid]::NewGuid().ToString("N"))

$DocumentFiles = @(
    @{ Title = "README"; Path = "README.md" },
    @{ Title = "Overview"; Path = "docs/overview.md" },
    @{ Title = "Usage Guide"; Path = "docs/usage.md" },
    @{ Title = "Parameters Reference"; Path = "docs/parameters.md" },
    @{ Title = "Outputs and Report Files"; Path = "docs/outputs.md" },
    @{ Title = "Examples"; Path = "docs/examples.md" },
    @{ Title = "Scan-ADComputers Manual"; Path = "docs/scan-adcomputers.md" },
    @{ Title = "AD Excel Reporting"; Path = "docs/ad-excel-reporting.md" },
    @{ Title = "Get-ADAdminActivity Manual"; Path = "docs/ad-admin-activity.md" },
    @{ Title = "User Account Management"; Path = "docs/user-account-management.md" },
    @{ Title = "Troubleshooting"; Path = "docs/troubleshooting.md" },
    @{ Title = "Security Policy"; Path = "SECURITY.md" },
    @{ Title = "Contributing Guide"; Path = "CONTRIBUTING.md" },
    @{ Title = "Changelog"; Path = "CHANGELOG.md" }
)

$ScriptManuals = @(
    @{
        Title = "Scan-ADComputers Manual"
        BaseName = "Scan-ADComputers-Manual"
        Subtitle = "Computer inventory, operational validation, JSON config, reporting, and troubleshooting reference."
        Documents = @(
            @{ Title = "Scan-ADComputers Manual"; Path = "docs/scan-adcomputers.md" },
            @{ Title = "Security Policy"; Path = "SECURITY.md" }
        )
    },
    @{
        Title = "Get-ADAdminActivity Manual"
        BaseName = "Get-ADAdminActivity-Manual"
        Subtitle = "Domain Controller Security log audit reporting, privileged-admin filtering, output, and troubleshooting reference."
        Documents = @(
            @{ Title = "Get-ADAdminActivity Manual"; Path = "docs/ad-admin-activity.md" },
            @{ Title = "Security Policy"; Path = "SECURITY.md" }
        )
    },
    @{
        Title = "Manage-ADUserAccounts Manual"
        BaseName = "Manage-ADUserAccounts-Manual"
        Subtitle = "User account reports, lockout analysis, single-user audit lookup, reset actions, and troubleshooting reference."
        Documents = @(
            @{ Title = "Manage-ADUserAccounts Manual"; Path = "docs/user-account-management.md" },
            @{ Title = "Security Policy"; Path = "SECURITY.md" }
        )
    }
)

function ConvertTo-XmlText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape($Text)
}

function New-WordRun {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "This pure helper constructs and returns WordprocessingML text; it does not change system state."
    )]
    param(
        [string]$Text,
        [switch]$Code
    )

    $EscapedText = ConvertTo-XmlText -Text $Text
    $SpaceAttribute = if ($Text -match "^\s|\s$") { " xml:space=`"preserve`"" } else { "" }

    if ($Code) {
        return "<w:r><w:rPr><w:rFonts w:ascii=`"Courier New`" w:hAnsi=`"Courier New`"/><w:sz w:val=`"19`"/></w:rPr><w:t$SpaceAttribute>$EscapedText</w:t></w:r>"
    }

    return "<w:r><w:t$SpaceAttribute>$EscapedText</w:t></w:r>"
}

function New-WordParagraph {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "This pure helper constructs and returns WordprocessingML text; it does not change system state."
    )]
    param(
        [string]$Text,
        [string]$Style,
        [switch]$Code
    )

    $StyleXml = if ([string]::IsNullOrWhiteSpace($Style)) {
        ""
    }
    else {
        "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>"
    }

    return "<w:p>$StyleXml$(New-WordRun -Text $Text -Code:$Code)</w:p>"
}

function ConvertTo-HtmlText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function ConvertTo-AnchorId {
    param(
        [string]$Text
    )

    $Anchor = $Text.ToLowerInvariant()
    $Anchor = [regex]::Replace($Anchor, "[^a-z0-9]+", "-")
    $Anchor = $Anchor.Trim("-")

    if ([string]::IsNullOrWhiteSpace($Anchor)) {
        return "section"
    }

    return $Anchor
}

function ConvertTo-HtmlInline {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ($null -eq $Text) {
        return ""
    }

    $Result = New-Object System.Text.StringBuilder
    $CurrentIndex = 0
    $LinkPattern = [regex]'\[([^\]]+)\]\((#[A-Za-z0-9_-]+|https?://[^)\s]+)\)'

    foreach ($Match in $LinkPattern.Matches($Text)) {
        [void]$Result.Append((ConvertTo-HtmlText -Text $Text.Substring($CurrentIndex, $Match.Index - $CurrentIndex)))
        $LinkText = ConvertTo-HtmlText -Text $Match.Groups[1].Value
        $LinkTarget = ConvertTo-HtmlText -Text $Match.Groups[2].Value
        [void]$Result.Append(("<a href=`"{0}`">{1}</a>" -f $LinkTarget, $LinkText))
        $CurrentIndex = $Match.Index + $Match.Length
    }

    if ($CurrentIndex -lt $Text.Length) {
        [void]$Result.Append((ConvertTo-HtmlText -Text $Text.Substring($CurrentIndex)))
    }

    return $Result.ToString()
}

function ConvertFrom-MarkdownToManualHtml {
    param(
        [string[]]$MarkdownLines,
        [string]$GeneratedDate,
        [string]$ManualTitle = "AdminTools Manual",
        [string]$ManualSubtitle = "Active Directory administration scripts, security controls, reporting workflows, and troubleshooting reference."
    )

    $HtmlBody = New-Object System.Collections.Generic.List[string]
    $InCodeBlock = $false
    $InList = $false
    $InPreformatted = $false
    $AnchorCounts = @{}

    $HtmlBody.Add("<div class=`"cover`">")
    $HtmlBody.Add(("<p class=`"cover-title`">{0}</p>" -f (ConvertTo-HtmlText -Text $ManualTitle)))
    $HtmlBody.Add(("<p class=`"cover-subtitle`">{0}</p>" -f (ConvertTo-HtmlText -Text $ManualSubtitle)))
    $HtmlBody.Add(("<p class=`"cover-meta`">Generated: {0}</p>" -f (ConvertTo-HtmlText -Text $GeneratedDate)))
    $HtmlBody.Add("<p class=`"cover-meta`">Format: Microsoft Word DOCX manual with linked navigation.</p>")
    $HtmlBody.Add("</div>")

    foreach ($Line in $MarkdownLines) {
        if ($Line -eq "# AdminTools Manual" -or $Line -like "Generated: *") {
            continue
        }

        if ($Line -match '^```') {
            if ($InList) {
                $HtmlBody.Add("</ul>")
                $InList = $false
            }

            if ($InPreformatted) {
                $HtmlBody.Add("</code></pre>")
                $InPreformatted = $false
            }

            if ($InCodeBlock) {
                $HtmlBody.Add("</code></pre>")
                $InCodeBlock = $false
            }
            else {
                $HtmlBody.Add("<pre><code>")
                $InCodeBlock = $true
            }

            continue
        }

        if ($InCodeBlock) {
            $HtmlBody.Add((ConvertTo-HtmlText -Text $Line))
            continue
        }

        if ($Line -match '^\|.*\|$') {
            if ($InList) {
                $HtmlBody.Add("</ul>")
                $InList = $false
            }

            if (-not $InPreformatted) {
                $HtmlBody.Add("<pre><code>")
                $InPreformatted = $true
            }

            $HtmlBody.Add((ConvertTo-HtmlText -Text $Line))
            continue
        }
        elseif ($InPreformatted) {
            $HtmlBody.Add("</code></pre>")
            $InPreformatted = $false
        }

        if ([string]::IsNullOrWhiteSpace($Line)) {
            if ($InList) {
                $HtmlBody.Add("</ul>")
                $InList = $false
            }

            continue
        }

        if ($Line -match '^(#{1,4})\s+(.+)$') {
            if ($InList) {
                $HtmlBody.Add("</ul>")
                $InList = $false
            }

            $Level = $Matches[1].Length
            $HeadingText = $Matches[2]
            $AnchorId = ConvertTo-AnchorId -Text $HeadingText
            if ($AnchorCounts.ContainsKey($AnchorId)) {
                $AnchorCounts[$AnchorId] += 1
                $AnchorId = "{0}-{1}" -f $AnchorId, $AnchorCounts[$AnchorId]
            }
            else {
                $AnchorCounts[$AnchorId] = 1
            }

            $ClassAttribute = if ($Level -eq 1) { " class=`"chapter`"" } else { "" }
            $HtmlBody.Add(("<h{0} id=`"{1}`"{2}>{3}</h{0}>" -f $Level, $AnchorId, $ClassAttribute, (ConvertTo-HtmlText -Text $HeadingText)))
            continue
        }

        if ($Line -match '^[-*]\s+(.+)$') {
            if (-not $InList) {
                $HtmlBody.Add("<ul>")
                $InList = $true
            }

            $HtmlBody.Add(("<li>{0}</li>" -f (ConvertTo-HtmlInline -Text $Matches[1])))
            continue
        }

        if ($InList) {
            $HtmlBody.Add("</ul>")
            $InList = $false
        }

        $HtmlBody.Add(("<p>{0}</p>" -f (ConvertTo-HtmlInline -Text $Line)))
    }

    if ($InCodeBlock -or $InPreformatted) {
        $HtmlBody.Add("</code></pre>")
    }

    if ($InList) {
        $HtmlBody.Add("</ul>")
    }

    return @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>AdminTools Manual</title>
  <style>
    @page { margin: 0.75in; }
    body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.35; color: #111827; }
    a { color: #0f62fe; text-decoration: none; }
    h1 { font-size: 22pt; margin: 24pt 0 8pt; page-break-after: avoid; color: #111827; }
    h2 { font-size: 17pt; margin: 18pt 0 6pt; page-break-after: avoid; color: #1f2937; }
    h3 { font-size: 14pt; margin: 14pt 0 5pt; page-break-after: avoid; color: #374151; }
    h4 { font-size: 12pt; margin: 12pt 0 4pt; page-break-after: avoid; color: #4b5563; }
    p { margin: 0 0 7pt; }
    ul { margin: 0 0 8pt 18pt; padding: 0; }
    li { margin: 0 0 3pt; }
    pre { font-family: Consolas, "Courier New", monospace; font-size: 9pt; background: #f3f4f6; border: 1px solid #d1d5db; padding: 8pt; white-space: pre-wrap; }
    .cover { page-break-after: always; min-height: 9in; padding-top: 2in; }
    .cover-title { font-size: 32pt; font-weight: 700; margin: 0 0 14pt; color: #111827; }
    .cover-subtitle { font-size: 15pt; margin: 0 0 30pt; color: #374151; }
    .cover-meta { font-size: 11pt; color: #4b5563; margin: 0 0 5pt; }
    .reference-page { page-break-after: always; }
    .toc { margin-left: 0; }
    .toc li { list-style-type: none; margin-bottom: 5pt; }
    .chapter { page-break-before: always; }
  </style>
</head>
<body>
$($HtmlBody -join "`n")
</body>
</html>
"@
}

function New-ManualMarkdown {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        "PSUseShouldProcessForStateChangingFunctions",
        "",
        Justification = "This pure helper constructs and returns Markdown content; file writes occur outside the function."
    )]
    param(
        [string]$Title,
        [object[]]$Documents,
        [string]$GeneratedDate
    )

    $ExistingDocuments = @()
    foreach ($Document in $Documents) {
        $FullPath = Join-Path -Path $RepositoryRoot -ChildPath $Document.Path
        if (Test-Path -LiteralPath $FullPath) {
            $ExistingDocuments += $Document
        }
    }

    $Content = New-Object System.Collections.Generic.List[string]
    $Content.Add(("# {0}" -f $Title))
    $Content.Add("")
    $Content.Add(("Generated: {0}" -f $GeneratedDate))
    $Content.Add("")
    $Content.Add("## Reference Index")
    $Content.Add("")
    $Content.Add("Use this section to jump to the major areas of the manual.")
    $Content.Add("")

    foreach ($Document in $ExistingDocuments) {
        $AnchorId = ConvertTo-AnchorId -Text $Document.Title
        $Content.Add(("- [{0}](#{1}) - {2}" -f $Document.Title, $AnchorId, $Document.Path))
    }

    $Content.Add("")

    foreach ($Document in $ExistingDocuments) {
        $FullPath = Join-Path -Path $RepositoryRoot -ChildPath $Document.Path
        $Content.Add("")
        $Content.Add(("# {0}" -f $Document.Title))
        $Content.Add("")

        foreach ($Line in [string[]](Get-Content -LiteralPath $FullPath)) {
            if ($Line -match '^(#{1,5})\s+(.+)$') {
                $Content.Add(("#" + $Line))
                continue
            }

            $Content.Add($Line)
        }
    }

    return $Content
}

function Resolve-LibreOfficeCommand {
    $Command = Get-Command -Name "libreoffice" -ErrorAction SilentlyContinue
    if ($Command) {
        return $Command.Source
    }

    $Command = Get-Command -Name "soffice" -ErrorAction SilentlyContinue
    if ($Command) {
        return $Command.Source
    }

    $CandidatePaths = @(
        $env:LIBREOFFICE_PATH,
        (Join-Path -Path $env:ProgramFiles -ChildPath "LibreOffice\program\soffice.exe"),
        (Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath "LibreOffice\program\soffice.exe"),
        (Join-Path -Path $env:LOCALAPPDATA -ChildPath "Programs\LibreOffice\program\soffice.exe")
    )

    foreach ($CandidatePath in $CandidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($CandidatePath) -and (Test-Path -LiteralPath $CandidatePath -PathType Leaf)) {
            return $CandidatePath
        }
    }

    return $null
}

function Convert-ManualHtmlToDocxWithWord {
    param(
        [string]$HtmlPath,
        [string]$DocxPath
    )

    if (-not $IsWindows -and $PSVersionTable.PSEdition -eq "Core") {
        return $false
    }

    $Word = $null
    $Document = $null

    try {
        $Word = New-Object -ComObject Word.Application -ErrorAction Stop
        $Word.Visible = $false
        $Document = $Word.Documents.Open((Resolve-Path -LiteralPath $HtmlPath).Path)

        if (Test-Path -LiteralPath $DocxPath) {
            Remove-Item -LiteralPath $DocxPath -Force
        }

        $Document.SaveAs([ref]$DocxPath, [ref]16)
        return (Test-Path -LiteralPath $DocxPath)
    }
    catch {
        return $false
    }
    finally {
        if ($Document) {
            $Document.Close([ref]$false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
        }

        if ($Word) {
            $Word.Quit() | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null
        }
    }
}
function Convert-ManualHtmlToDocx {
    param(
        [string]$HtmlPath,
        [string]$DocxPath,
        [string]$OutputDirectory
    )

    $LibreOffice = Resolve-LibreOfficeCommand
    if (-not $LibreOffice) {
        return Convert-ManualHtmlToDocxWithWord -HtmlPath $HtmlPath -DocxPath $DocxPath
    }

    $LibreOfficeProfile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("admintools-lo-profile-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $LibreOfficeProfile -Force | Out-Null

    if (Test-Path -LiteralPath $DocxPath) {
        Remove-Item -LiteralPath $DocxPath -Force
    }

    $LibreOfficeArguments = @(
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        ("-env:UserInstallation=file://{0}" -f $LibreOfficeProfile),
        "--convert-to",
        "docx:Office Open XML Text",
        "--outdir",
        $OutputDirectory,
        $HtmlPath
    )

    try {
        & $LibreOffice @LibreOfficeArguments | Out-Host
        return (Test-Path -LiteralPath $DocxPath)
    }
    finally {
        if (Test-Path -LiteralPath $LibreOfficeProfile) {
            Remove-Item -LiteralPath $LibreOfficeProfile -Recurse -Force
        }
    }
}

New-Item -ItemType Directory -Path $ManualDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $TemporaryDirectory -Force | Out-Null

$GeneratedDate = Get-Date -Format "yyyy-MM-dd"
$ExistingDocumentFiles = @()
foreach ($DocumentFile in $DocumentFiles) {
    $FullPath = Join-Path -Path $RepositoryRoot -ChildPath $DocumentFile.Path
    if (Test-Path -LiteralPath $FullPath) {
        $ExistingDocumentFiles += $DocumentFile
    }
}

$Markdown = New-Object System.Collections.Generic.List[string]
$Markdown.Add("# AdminTools Manual")
$Markdown.Add("")
$Markdown.Add(("Generated: {0}" -f $GeneratedDate))
$Markdown.Add("")
$Markdown.Add("## Reference Index")
$Markdown.Add("")
$Markdown.Add("Use this section to jump to the major areas of the manual.")
$Markdown.Add("")

foreach ($DocumentFile in $ExistingDocumentFiles) {
    $AnchorId = ConvertTo-AnchorId -Text $DocumentFile.Title
    $Markdown.Add(("- [{0}](#{1}) - {2}" -f $DocumentFile.Title, $AnchorId, $DocumentFile.Path))
}

$Markdown.Add("")

foreach ($DocumentFile in $ExistingDocumentFiles) {
    $FullPath = Join-Path -Path $RepositoryRoot -ChildPath $DocumentFile.Path
    $Markdown.Add("")
    $Markdown.Add(("# {0}" -f $DocumentFile.Title))
    $Markdown.Add("")

    foreach ($Line in [string[]](Get-Content -LiteralPath $FullPath)) {
        if ($Line -match '^(#{1,5})\s+(.+)$') {
            $Markdown.Add(("#" + $Line))
            continue
        }

        $Markdown.Add($Line)
    }
}

Set-Content -LiteralPath $ManualMarkdown -Value $Markdown -Encoding UTF8
Set-Content -LiteralPath $ManualHtml -Value (ConvertFrom-MarkdownToManualHtml -MarkdownLines $Markdown -GeneratedDate $GeneratedDate) -Encoding UTF8

$CreatedDocxWithLibreOffice = $false
$LibreOffice = Resolve-LibreOfficeCommand
if ($LibreOffice) {
    $LibreOfficeProfile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("admintools-lo-profile-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $LibreOfficeProfile -Force | Out-Null

    if (Test-Path -LiteralPath $ManualDocx) {
        Remove-Item -LiteralPath $ManualDocx -Force
    }

    $LibreOfficeArguments = @(
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        ("-env:UserInstallation=file://{0}" -f $LibreOfficeProfile),
        "--convert-to",
        "docx:Office Open XML Text",
        "--outdir",
        $ManualDirectory,
        $ManualHtml
    )

    & $LibreOffice @LibreOfficeArguments | Out-Host
    $CreatedDocxWithLibreOffice = Test-Path -LiteralPath $ManualDocx

    if (Test-Path -LiteralPath $LibreOfficeProfile) {
        Remove-Item -LiteralPath $LibreOfficeProfile -Recurse -Force
    }
}

if (-not $CreatedDocxWithLibreOffice) {

$Body = New-Object System.Collections.Generic.List[string]
$InCodeBlock = $false

foreach ($Line in $Markdown) {
    if ($Line -match '^```') {
        $InCodeBlock = -not $InCodeBlock
        continue
    }

    if ($InCodeBlock) {
        $Body.Add((New-WordParagraph -Text $Line -Code))
        continue
    }

    if ([string]::IsNullOrWhiteSpace($Line)) {
        $Body.Add("<w:p/>")
        continue
    }

    if ($Line -match "^# (.+)$") {
        $Body.Add((New-WordParagraph -Text $Matches[1] -Style "Heading1"))
        continue
    }

    if ($Line -match "^## (.+)$") {
        $Body.Add((New-WordParagraph -Text $Matches[1] -Style "Heading2"))
        continue
    }

    if ($Line -match "^### (.+)$") {
        $Body.Add((New-WordParagraph -Text $Matches[1] -Style "Heading3"))
        continue
    }

    if ($Line -match "^#### (.+)$") {
        $Body.Add((New-WordParagraph -Text $Matches[1] -Style "Heading4"))
        continue
    }

    if ($Line -match "^[-*]\s+(.+)$") {
        $Body.Add((New-WordParagraph -Text ("- " + $Matches[1])))
        continue
    }

    if ($Line -match "^\|.*\|$") {
        $Body.Add((New-WordParagraph -Text $Line -Code))
        continue
    }

    $Body.Add((New-WordParagraph -Text $Line))
}

$DocumentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
$($Body -join "`n")
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>
"@

$StylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="220" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="180" w:after="60"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style></w:styles>
"@

New-Item -ItemType Directory -Path (Join-Path $TemporaryDirectory "_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $TemporaryDirectory "word/_rels") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $TemporaryDirectory "docProps") -Force | Out-Null

Set-Content -LiteralPath (Join-Path $TemporaryDirectory "[Content_Types].xml") -Encoding UTF8 -Value "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/word/document.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml`"/><Override PartName=`"/word/styles.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml`"/><Override PartName=`"/docProps/core.xml`" ContentType=`"application/vnd.openxmlformats-package.core-properties+xml`"/><Override PartName=`"/docProps/app.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.extended-properties+xml`"/></Types>"
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "_rels/.rels") -Encoding UTF8 -Value "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument`" Target=`"word/document.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties`" Target=`"docProps/core.xml`"/><Relationship Id=`"rId3`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties`" Target=`"docProps/app.xml`"/></Relationships>"
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "word/_rels/document.xml.rels") -Encoding UTF8 -Value "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "word/document.xml") -Encoding UTF8 -Value $DocumentXml
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "word/styles.xml") -Encoding UTF8 -Value $StylesXml
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "docProps/core.xml") -Encoding UTF8 -Value "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:dcmitype=`"http://purl.org/dc/dcmitype/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>AdminTools Manual</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy></cp:coreProperties>"
Set-Content -LiteralPath (Join-Path $TemporaryDirectory "docProps/app.xml") -Encoding UTF8 -Value "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Properties xmlns=`"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties`" xmlns:vt=`"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes`"><Application>Codex</Application></Properties>"

if (Test-Path -LiteralPath $ManualDocx) {
    Remove-Item -LiteralPath $ManualDocx -Force
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($TemporaryDirectory, $ManualDocx)
Remove-Item -LiteralPath $TemporaryDirectory -Recurse -Force
}
elseif (Test-Path -LiteralPath $TemporaryDirectory) {
    Remove-Item -LiteralPath $TemporaryDirectory -Recurse -Force
}

Write-Information "Created $ManualMarkdown"
Write-Information "Created $ManualHtml"
Write-Information "Created $ManualDocx"

foreach ($ScriptManual in $ScriptManuals) {
    $ScriptManualMarkdown = Join-Path -Path $ManualDirectory -ChildPath ("{0}.md" -f $ScriptManual.BaseName)
    $ScriptManualHtml = Join-Path -Path $ManualDirectory -ChildPath ("{0}.html" -f $ScriptManual.BaseName)
    $ScriptManualDocx = Join-Path -Path $ManualDirectory -ChildPath ("{0}.docx" -f $ScriptManual.BaseName)

    $ScriptMarkdown = New-ManualMarkdown -Title $ScriptManual.Title -Documents $ScriptManual.Documents -GeneratedDate $GeneratedDate
    $ScriptHtml = ConvertFrom-MarkdownToManualHtml `
        -MarkdownLines $ScriptMarkdown `
        -GeneratedDate $GeneratedDate `
        -ManualTitle $ScriptManual.Title `
        -ManualSubtitle $ScriptManual.Subtitle

    Set-Content -LiteralPath $ScriptManualMarkdown -Value $ScriptMarkdown -Encoding UTF8
    Set-Content -LiteralPath $ScriptManualHtml -Value $ScriptHtml -Encoding UTF8

    $CreatedScriptDocx = Convert-ManualHtmlToDocx `
        -HtmlPath $ScriptManualHtml `
        -DocxPath $ScriptManualDocx `
        -OutputDirectory $ManualDirectory

    if (-not $CreatedScriptDocx) {
        Write-Warning "Could not create $ScriptManualDocx because LibreOffice Writer conversion was unavailable."
    }
    else {
        Write-Information "Created $ScriptManualMarkdown"
        Write-Information "Created $ScriptManualHtml"
        Write-Information "Created $ScriptManualDocx"
    }
}
