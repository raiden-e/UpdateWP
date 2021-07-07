[CmdletBinding()]
param (
    [string]$ExportPath,
    [switch]$ExportAll
)
# settings

if (!($ExportPath)) { $ExportPath = Resolve-Path "$home\OneDrive*ruhr-uni-bochum.de\_Studium" }
# Create OneNote application
$oneLog = "$home\one_lastedit.log"
try{
    if(Get-Content $oneLog -ge Get-Date -Format "%y%M%d"){
        Write-Host "Daily Export Already Done"
        exit(0)
    }
}
catch{
    Write-Warning "Couldn't find '$oneLog', proceeding"
}
$OneNote = New-Object -ComObject OneNote.Application

# Set note hierarchy
[xml]$Hierarchy = ""
$OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)
function Export-All {
    param (
        [string]$Notebook
    )

    # Get info about each section
    $Notebook.ChildNodes | Where-Object { !$_.isRecycleBin } | ForEach-Object {
        $Section = $_
        $SectionIndex = [array]::indexof($Notebook.ChildNodes, $_)
        $SectionName = "$($SectionIndex)_$($Section.name)"
        $SectionPath = Join-Path -Path $NotebookPath -ChildPath $SectionName

        Write-Host "Processing Section: $SectionName"
        New-Item -Force -Path $SectionPath -ItemType directory | Out-Null

        # Process pages
        $Section.ChildNodes | ForEach-Object {
            $Page = $_
            $PageIndex = [array]::indexof($Section.ChildNodes, $_)
            $PageName = "$($PageIndex)_$($Page.name)".Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
            $PagePath = Join-Path -Path $SectionPath -ChildPath $PageName

            Write-Host "Processing Page: $PageName"
            New-Item -Force -Path $PagePath -ItemType directory | Out-Null

            $PageHtmPath = Join-Path -Path $PagePath -ChildPath 'index.htm'
            if (!(Test-Path -Path $PageHtmPath)) {
                Write-Host "Export Page as Htm: $PageHtmPath"
                $OneNote.Publish($Page.ID, $PageHtmPath, 7, "")
            }

            $PageDocxPath = Join-Path -Path $PagePath -ChildPath 'index.docx'
            if (!(Test-Path -Path $PageDocxPath)) {
                Write-Host "Export Page as Docx: $PageDocxPath"
                $OneNote.Publish($Page.ID, $PageDocxPath, 5, "")
            }

            $PageMdPath = Join-Path -Path $PagePath -ChildPath 'index.md'
            if (!(Test-Path -Path $PageMdPath)) {
                Write-Host "Convert Docx to Md: $PageMdPath"
                Set-Location $PagePath
                pandoc.exe --extract-media=./ .\index.docx -o index.md -t gfm
            }

            Export Attachments
            $xml = ''
            $schema = @{one = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
            $onenote.GetPageContent($Page.ID, [ref]$xml)
            $xml | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | ForEach-Object {
                $AttachmentPath = Join-Path -Path $PagePath -ChildPath $_.Node.preferredName
                Write-Host "Export Attachment: $($AttachmentPath)"
                Copy-Item -Force $_.Node.pathCache -Destination $AttachmentPath

            }
        }
    }
}

# Get info foreach notebook
$Hierarchy.Notebooks.Notebook | Where-Object { $_.name -eq 'Studium' } | ForEach-Object {
    $Notebook = $_
    $Name = $Notebook.name
    $NotebookPath = Join-Path -Path $ExportPath -ChildPath $Name

    Write-Host "Trying to export Notebook: $Name"


    $NotebookOnepkgPath = Join-Path -Path $ExportPath -ChildPath "$($Name).onepkg"
    if (Test-Path -Path $NotebookOnepkgPath) {
        try {
            $NotebookOnepkgPath | Remove-Item -Force
        }
        catch {
            Write-Error "Couldn't delete $NotebookOnepkgPath!"
            return 1
        }
    }
    Write-Host "Export Notebook as Onepkg: $NotebookOnepkgPath"

    if (!(Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory
    }
    $OneNote.Publish($Notebook.ID, $NotebookOnepkgPath, 1, "")

    if ($ExportAll) {
        New-Item -Force -Path $NotebookPath -ItemType directory | Out-Null
        Export-All($Notebook)
    }
    Get-Date -Format "%y%M%d" | Out-File $oneLog
}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNote)
