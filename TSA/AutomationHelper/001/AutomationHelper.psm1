function Get-Greeting {
    [CmdletBinding()]
    param (
        [string]$Name
    )
    return "Hello, $Name!"
}

function Get-Assembly2017 {
    [CmdletBinding()]
    param (
        [string]$dllPath
    )

    Write-Warning "Please make sure the SDL Trados Studio 2017 (Studio5) version is installed to execute this script!"
    
    $dllPaths = @(
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.Core.Globalization.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.Core.Settings.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.ProjectApi.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.ProjectAutomation.Core.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.ProjectAutomation.FileBased.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.ProjectAutomation.Settings.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.LanguagePlatform.Core.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.LanguagePlatform.TranslationMemory.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio5\Sdl.LanguagePlatform.TranslationMemoryApi.dll"
    )

    foreach ($dllPath in $dllPaths) {
        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath
        }
        else {
            Write-Host "DLL not found: $dllPath" -ForegroundColor Red
        }
    }

    $url = "https://github.com/vegetaz/UnsafeDLL/raw/master/4.0.30319/System.Runtime.CompilerServices.Unsafe.dll"
    $webClient = New-Object System.Net.WebClient
    $dllBytes = $webClient.DownloadData($url)
    [System.Reflection.Assembly]::Load($dllBytes)

    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Management.Automation.Runspaces;

public class RunspacedDelegateFactory
{
    public static Delegate NewRunspacedDelegate(Delegate _delegate, Runspace runspace)
    {
        Action setRunspace = () => Runspace.DefaultRunspace = runspace;

        return ConcatActionToDelegate(setRunspace, _delegate);
    }

    private static Expression ExpressionInvoke(Delegate _delegate, params Expression[] arguments)
    {
        var invokeMethod = _delegate.GetType().GetMethod("Invoke");

        return Expression.Call(Expression.Constant(_delegate), invokeMethod, arguments);
    }

    public static Delegate ConcatActionToDelegate(Action a, Delegate d)
    {
        var parameters =
            d.GetType().GetMethod("Invoke").GetParameters()
            .Select(p => Expression.Parameter(p.ParameterType, p.Name))
            .ToArray();

        Expression body = Expression.Block(ExpressionInvoke(a), ExpressionInvoke(d, parameters));

        var lambda = Expression.Lambda(d.GetType(), body, parameters);

        var compiled = lambda.Compile();

        return compiled;
    }
}
'@
}

function Get-Assembly2021 {
    [CmdletBinding()]
    param (
        [string]$dllPath
    )

    Write-Warning "Please make sure the SDL Trados Studio 2021 (Studio16) version is installed to execute this script!"
    
    $dllPaths = @(
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.Core.Globalization.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.Core.Settings.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.ProjectApi.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.ProjectAutomation.Core.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.ProjectAutomation.FileBased.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.ProjectAutomation.Settings.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.LanguagePlatform.Core.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.LanguagePlatform.TranslationMemory.dll",
        "C:\Program Files (x86)\SDL\SDL Trados Studio\Studio16\Sdl.LanguagePlatform.TranslationMemoryApi.dll"
    )

    foreach ($dllPath in $dllPaths) {
        if (Test-Path $dllPath) {
            Add-Type -Path $dllPath
        }
        else {
            Write-Host "DLL not found: $dllPath" -ForegroundColor Red
        }
    }

    $url = "https://github.com/vegetaz/UnsafeDLL/raw/master/4.0.30319/System.Runtime.CompilerServices.Unsafe.dll"
    $webClient = New-Object System.Net.WebClient
    $dllBytes = $webClient.DownloadData($url)
    [System.Reflection.Assembly]::Load($dllBytes)

    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Management.Automation.Runspaces;

public class RunspacedDelegateFactory
{
    public static Delegate NewRunspacedDelegate(Delegate _delegate, Runspace runspace)
    {
        Action setRunspace = () => Runspace.DefaultRunspace = runspace;

        return ConcatActionToDelegate(setRunspace, _delegate);
    }

    private static Expression ExpressionInvoke(Delegate _delegate, params Expression[] arguments)
    {
        var invokeMethod = _delegate.GetType().GetMethod("Invoke");

        return Expression.Call(Expression.Constant(_delegate), invokeMethod, arguments);
    }

    public static Delegate ConcatActionToDelegate(Action a, Delegate d)
    {
        var parameters =
            d.GetType().GetMethod("Invoke").GetParameters()
            .Select(p => Expression.Parameter(p.ParameterType, p.Name))
            .ToArray();

        Expression body = Expression.Block(ExpressionInvoke(a), ExpressionInvoke(d, parameters));

        var lambda = Expression.Lambda(d.GetType(), body, parameters);

        var compiled = lambda.Compile();

        return compiled;
    }
}
'@
}

function Get-Guids {
    [CmdletBinding()]
    param (
        [Sdl.ProjectAutomation.Core.ProjectFile[]] $files
    )
    
    $guids = $files | Select-Object -ExpandProperty Id
    return [System.Guid[]]$guids
}

function Get-Project {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [Alias("Location", "PrjLoc")]
        [String] $ProjectLocation
    )

    # Tìm kiếm tệp dự án SDL Trados
    $ProjectFile = Get-ChildItem -Path $ProjectLocation -Filter *.sdlproj -File | Select-Object -First 1

    if (-not $ProjectFile) {
        Throw "The SDL Trados Studio File (.sdlproj) not found in '$ProjectLocation' folder."
    }

    # Tạo đối tượng dự án
    $Project = New-Object Sdl.ProjectAutomation.FileBased.FileBasedProject ($ProjectFile.FullName)
    return $Project
}

function Write-TaskProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Object] $Caller,
        [Parameter(Mandatory = $true)]
        [Object] $ProgressEventArgs
    )

    $Percent = $ProgressEventArgs.PercentComplete
    $Status = $ProgressEventArgs.Status

    if ($host.name -eq 'ConsoleHost') {
        Write-Host "$($Percent.ToString().PadLeft(5))%	$Status	$StatusMessage`r" -NoNewLine
        if ($Percent -eq 100 -and $Status -eq "Completed") {
            Write-Host
        }
    }
    else {
        Write-Progress -Activity "Processing task" -PercentComplete $Percent -Status $Status
        if ($Percent -eq 100 -and $Status -eq "Completed") {
            Write-Progress -Activity "Processing task" -Completed
        }
    }
}

function Write-TaskMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Object] $Caller,
        
        [Parameter(Mandatory = $true)]
        [Sdl.ProjectAutomation.Core.TaskMessageEventArgs] $MessageEventArgs
    )

    $Message = $MessageEventArgs.Message

    if ($null -ne $Message) {
        Write-Host "`n$($Message.Level): $($Message.Message)" -ForegroundColor DarkYellow

        if ($null -ne $Message.Exception) {
            Write-Host "$($Message.Exception)" -ForegroundColor Magenta
        }
    }
}

function Test-Task {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Sdl.ProjectAutomation.Core.AutomaticTask] $taskToValidate
    )

    $taskStatus = $taskToValidate.Status
    $taskName = $taskToValidate.Name

    if ($taskStatus -eq [Sdl.ProjectAutomation.Core.TaskStatus]::Completed) {
        Write-Host "Task $taskName successfully completed." -ForegroundColor Green
    }
    else {
        switch ($taskStatus) {
            [Sdl.ProjectAutomation.Core.TaskStatus]::Failed {
                Write-Host "Task $taskName failed." -ForegroundColor Red
            }
            [Sdl.ProjectAutomation.Core.TaskStatus]::Invalid {
                Write-Host "Task $taskName not valid." -ForegroundColor Red
            }
            [Sdl.ProjectAutomation.Core.TaskStatus]::Rejected {
                Write-Host "Task $taskName rejected." -ForegroundColor Red
            }
            [Sdl.ProjectAutomation.Core.TaskStatus]::Cancelled {
                Write-Host "Task $taskName cancelled." -ForegroundColor Red
            }
            Default {
                Write-Host "Task $taskName status: $taskStatus" -ForegroundColor Cyan
            }
        }

        foreach ($message in $taskToValidate.Messages) {
            if ($null -ne $message.ProjectFileId) {
                $AffectedFile = $Project.GetTargetLanguageFiles() | Where-Object { $_.Id -eq $message.ProjectFileId }
                if ($AffectedFile) {
                    Write-Host "$($AffectedFile.Language)`t$($AffectedFile.Folder)$($AffectedFile.Name)"
                }
            }
            Write-Host ($message.Message -Replace '(`n|`r)+$', '') -ForegroundColor Red
        }
    }
}

function Update-MainTMs {
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $true)]
        [Alias("Location", "PrjLoc")]
        [String] $ProjectLocation,        
        [Alias("TrgLng")]
        [String] $TargetLanguages
    )

    try {
        $ResolvedProjectPath = (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath
        $Project = Get-Project -ProjectLocation $ResolvedProjectPath
        # $ProjectSettings = $Project.GetSettings() # Tạm thời không sử dụng $ProjectSettings

        $TargetLanguagesList = if ($TargetLanguages) {
            $TargetLanguages -Split "\s+|;\s*|,\s*"
        }
        else {
            @($Project.GetProjectInfo().TargetLanguages.IsoAbbreviation)
        }

        Write-Host "`nUpdating main translation memories..." -ForegroundColor White

        $TargetFilesGuids = @()
        foreach ($TargetLanguage in $TargetLanguagesList) {
            $TargetFiles = $Project.GetTargetLanguageFiles($TargetLanguage)
            $TargetFilesGuids += Get-Guids $TargetFiles
        }

        $OnTaskProgress = ${function:Write-TaskProgress} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskStatusEventArgs]]
        $OnTaskProgressDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskProgress, [Runspace]::DefaultRunspace)

        $OnTaskMessage = ${function:Write-TaskMessage} -as [System.EventHandler[Sdl.ProjectAutomation.Core.TaskMessageEventArgs]]
        $OnTaskMessageDelegate = [RunspacedDelegateFactory]::NewRunspacedDelegate($OnTaskMessage, [Runspace]::DefaultRunspace)

        $Task = $Project.RunAutomaticTask($TargetFilesGuids, [Sdl.ProjectAutomation.Core.AutomaticTaskTemplateIds]::UpdateMainTranslationMemories, $OnTaskProgressDelegate, $OnTaskMessageDelegate)
        Test-Task $Task

        Write-Host "Done"
    }
    catch {
        Write-Host "An error occurred while updating the main TMs: $_" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    }
}


function Export-Package {
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $true)]
        [Alias("Location", "PrjLoc")]
        [String] $ProjectLocation,

        [Parameter (Mandatory = $true)]
        [Alias("PkgLoc")]
        [String] $PackageLocation,

        [Alias("TrgLng")]
        [String] $TargetLanguages,

        [ValidateSet("Translate", "Review")]
        [String] $Task = "Translate",

        [ValidateSet("None", "UseExisting", "CreateNew")]
        [Alias("PrjTM", "ProjectTMs")]
        [String] $ProjectTM = "CreateNew",

        [Alias("PkgCmt")]
        [String] $PackageComment = "",

        [Alias("IncAS", "IncludeAutoSuggest")]
        [Switch] $IncludeAutoSuggestDictionaries,

        [Alias("IncTM", "IncludeMainTM")]
        [Switch] $IncludeMainTMs,

        [Alias("IncTB", "IncludeTermbase")]
        [Switch] $IncludeTermbases,

        [Alias("RecAna", "Recompute", "RecomputeAnalyse", "RecomputeAnalyze")]
        [Switch] $RecomputeAnalysis,

        [Alias("IncRep", "IncludeExisting", "IncludeExistingReport", "IncludeReports", "IncludeReport")]
        [Switch] $IncludeExistingReports,

        [Alias("KeepAT", "KeepATProviders", "KeepATProvider")]
        [Switch] $KeepAutomatedTranslationProviders,

        [Alias("RemoveAT", "RemoveATProviders", "RemoveATProvider")]
        [Switch] $RemoveAutomatedTranslationProviders,

        [Alias("RmvSrvTM", "RemoveServerTMs", "RemoveServerTM")]
        [Switch] $RemoveServerBasedTMs
    )

    if (!(Test-Path $PackageLocation)) {
        New-Item -Path $PackageLocation -Force -ItemType Directory | Out-Null
    }

    $PackageOptions = [Sdl.ProjectAutomation.Core.ProjectPackageCreationOptions]::new()

    $PackageOptions.IncludeAutoSuggestDictionaries = $IncludeAutoSuggestDictionaries.IsPresent
    $PackageOptions.IncludeMainTranslationMemories = $IncludeMainTMs.IsPresent
    $PackageOptions.IncludeTermbases = $IncludeTermbases.IsPresent
    $PackageOptions.ProjectTranslationMemoryOptions = [Sdl.ProjectAutomation.Core.ProjectTranslationMemoryPackageOptions] $ProjectTM
    $PackageOptions.RemoveAutomatedTranslationProviders = !$KeepAutomatedTranslationProviders.IsPresent
    $PackageOptions.RemoveServerBasedTranslationMemories = $RemoveServerBasedTMs.IsPresent
    $PackageOptions.RecomputeAnalysisStatistics = $RecomputeAnalysis.IsPresent

    if ($IncludeExistingReports -or $ProjectTM -eq "CreateNew" -or $RecomputeAnalysis) {
        if ($PackageOptions.PSObject.Properties.Match("IncludeReports").Count -gt 0) {
            $PackageOptions.IncludeReports = $true
        }
        if ($IncludeExistingReports -and $PackageOptions.PSObject.Properties.Match("IncludeExistingReports").Count -gt 0) {
            $PackageOptions.IncludeExistingReports = $true
        }
    }

    $PackageDueDate = [DateTime]::MaxValue

    $Project = Get-Project (Resolve-Path -LiteralPath $ProjectLocation).ProviderPath

    $TargetLanguagesList = if ($TargetLanguages) {
        $TargetLanguages -split "\s+|;\s*|,\s*"
    }
    else {
        $Project.GetProjectInfo().TargetLanguages.IsoAbbreviation
    }

    Write-Host "`nCreating packages..." -ForegroundColor White

    foreach ($Language in $TargetLanguagesList) {
        $User = "$($Language) translator"
        $PackageName = "$($Project.GetProjectInfo().Name)_$($Language)"
        
        Write-Host "$PackageName$ProjectPackageExtension"

        $TaskFiles = Get-TaskFileInfoFiles $Project $Language
        $ManualTask = $Project.CreateManualTask($Task, $User, $PackageDueDate, $TaskFiles)

        $Package = $Project.CreateProjectPackage($ManualTask.Id, $PackageName, $PackageComment, $PackageOptions, ${function:Write-PackageProgress}, ${function:Write-PackageMessage})

        if ($Package.Status -eq [Sdl.ProjectAutomation.Core.PackageStatus]::Completed) {
            $Project.SavePackageAs($Package.PackageId, "$PackageLocation\$PackageName$ProjectPackageExtension")
        }
        else {
            Write-Host "Package creation failed, cannot save it!" -ForegroundColor Red
        }

        Remove-Variable TaskFiles, ManualTask, Package
    }
}

function Wait-ForExit {
    [CmdletBinding()]
    param ()

    Write-Host "Press 'Q' key to exit..."
    try {
        do {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } until ($key.Character -eq 'q' -or $key.Character -eq 'Q')
    }
    finally {
        Write-Host "Exited."
    }
}

Export-ModuleMember -Function Get-Greeting
Export-ModuleMember -Function Get-Assembly2017
Export-ModuleMember -Function Get-Assembly2021
Export-ModuleMember -Function Get-Guids
Export-ModuleMember -Function Get-Project
Export-ModuleMember -Function Write-TaskProgress
Export-ModuleMember -Function Write-TaskMessage
Export-ModuleMember -Function Test-Task
Export-ModuleMember -Function Export-Package
Export-ModuleMember -Function Update-MainTMs
Export-ModuleMember -Function Wait-ForExit