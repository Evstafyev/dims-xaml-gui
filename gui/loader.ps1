[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$XamlPath
)

[xml]$Global:xmlWPF = Get-Content -Path $XamlPath

try{
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
} catch {
    Throw "Failed to load Windows Presentation Framework assemblies."
}

$Global:xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))

$xmlWPF.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
    Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
}
