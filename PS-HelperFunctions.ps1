# PowerShell & Azure & O365 Helper Functions

function require-Module {
    param ([string]$Name,[boolean]$isInteractive=$true,[boolean]$AlwaysUpdate=$false)

    if((Get-Module $Name) -and $AlwaysUpdate) {
        write-host "Module found. Forcing Update of $Name and importing as current user." -ForegroundColor Cyan
        update-module $Name -ErrorAction SilentlyContinue
        import-module $Name -ErrorAction SilentlyContinue

        if (!(Get-Module $Name)) {
            write-host "Failed to install or import module $Name." -ForegroundColor Red
            if(!$isInteractive) { return $false }
        } else {
            write-host "Successfully imported module $Name." -ForegroundColor Green
            if(!$isInteractive) { return $true }
        }
    } elseif (!(Get-Module $Name)) {
        write-host "Module $Name not imported." -ForegroundColor Cyan
        import-module $Name -ErrorAction SilentlyContinue
        if (!(Get-Module $Name)) {
            write-host "Module not found. Installing $Name and importing as current user." -ForegroundColor Yellow
            install-module $Name -Scope CurrentUser -ErrorAction SilentlyContinue
            import-module $Name -ErrorAction SilentlyContinue
            if (!(Get-Module $Name)) {
                write-host "Failed to install or import module $Name." -ForegroundColor Red
                if(!$isInteractive) { return $false }
            } else {
                write-host "Successfully imported module $Name." -ForegroundColor Green
                if(!$isInteractive) { return $true }
            }
        } else {
                write-host "Successfully imported module $Name." -ForegroundColor Green
                if(!$isInteractive) { return $true }
        }
    } else {
        write-host "Module $Name already imported." -ForegroundColor Green       
    }
}

function errorAndExit([string]$message)
{
    #logError $message
    if ($Host.Name -eq 'Windows PowerShell ISE Host') {
        throw $message
    } else {
        exit 1
    }
}

function require-AzConnect {
    # Connect to Azure manually as an appropriate Administrative user
    $AzConnection = Get-AzContext -ErrorAction SilentlyContinue
    if(-not ($AzConnection.Tenant.Id)) {
        AzConnection = Connect-AzAccount
    } else {
        write-host "Already logged into Azure as $($AzConnection.Account.Id)" -ForegroundColor Green
    }
}

function require-MSOLConnect {
    # Connect to MSOL manually as an appropriate Administrative user
    if(-not (Get-MsolDomain -ErrorAction SilentlyContinue)) {
        Connect-MsolService
        if(-not (Get-MsolDomain -ErrorAction SilentlyContinue)) {
            write-host "Failed to connect to MSOnline" -ForegroundColor Red
        } else {
            write-host "Connect to MSOnline successfully" -ForegroundColor Green
        }
    } else {
        write-host "Already logged into MSOnline" -ForegroundColor Green
    }
}

function Set-Subscription {
    param ([string]$Subscription)
    write-host "Selecting $Subscription as default Azure subscription" -ForegroundColor cyan
    $sub = Select-AzSubscription $Subscription -ErrorAction SilentlyContinue
    if($sub.Subscription.Name -ne $Subscription) {
            Write-host "Failed to select subscription $Subscription" -Verbose -ForegroundColor red
            errorAndExit -message "Failed to select subscription $Subscription"
    }
}

Function Use-CorpProxy {
    # Set to use default proxy creds when making internet calls. Relies on current user IE Proxy settings
    [System.Net.WebRequest]::DefaultWebProxy.Credentials =  [System.Net.CredentialCache]::DefaultCredentials

    # Ignore Proxy Self-Signed Cert if required (for HTTPS Inspection Proxies)
    if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type)
    {
$certCallback = @"
    using System;
    using System.Net;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    public class ServerCertificateValidationCallback
    {
        public static void Ignore()
        {
            if(ServicePointManager.ServerCertificateValidationCallback ==null)
            {
                ServicePointManager.ServerCertificateValidationCallback += 
                    delegate
                    (
                        Object obj, 
                        X509Certificate certificate, 
                        X509Chain chain, 
                        SslPolicyErrors errors
                    )
                    {
                        return true;
                    };
            }
        }
    }
"@
        Add-Type $certCallback
    }
    [ServerCertificateValidationCallback]::Ignore()
}

# Helper Function (global scope) to create Array of Named PSO Objects with specific types
Function Create-PSOFunction([string]$Name,[string[]]$Fields,[boolean]$Validate=$false) {
    # Example Usage:
    # $Name = 'MyAppTemplate'
    # $Fields = @('[string]AppCode','[string]Name','isEnabled','[string[]]Links')
    # $FunctionCode  = Create-PSOFunction -Name $Name -Fields $Fields -Validate $True
    # $AppTemplate = @() # Configure array of PS Objects
    # $LinksArray  = @('https://bit.ly/blah','https://bit.ly/deblah')
    # $AppTemplate += Add-PSOMyAppTemplate -AppCode 'SuperApp' -Name 'Super-App' -isEnabled $true -Links $LinksArray
    # $AppTemplate
    # NOTE: -Validate param will annonouce status of creation or why it's failed

    #Security First...Check $Name just includes alphanumeric
    if($Name -match "\W+") {
        if($Validate) { write-host 'Bad name. Must be alphanumeric only!' -ForegroundColor Red }
        return $false
    }

    $NL = "`n"
    
    $FunctionName = "Add-PSO$Name"
    $FunctionCode = "function global:$FunctionName("

    # Validate fields and amend for function param
    $FieldList = $Fields -join ','

    #Security First...Check $Fields just includes alphanumeric or allowed chars
    if($FieldList -match ':' -or $FieldList -match '\$' -or $FieldList -match ';' -or $FieldList -match '\\' -or $FieldList -match "'" -or $FieldList -match '\/') {
        if($Validate) { write-host 'Bad list of Fields. Must ONLY contain alphanumeric characters or square brackets!' -ForegroundColor Red }
        return $false
    } else {
        $Params = @()
        [PSCustomObject]$ParamsPSOArray = @()

        ForEach($Field in $Fields) {
            [PSCustomObject]$SetItem = @{}

            if($Field -match '\]\]') {
                $SplitParams = $Field -split ']]'
                $FieldType = $SplitParams[0] + ']]'
                $FieldName = $SplitParams[1]
            } elseif($Field -match '\[') {
                $SplitParams = $Field -split ']'
                $FieldType = $SplitParams[0] + ']'
                $FieldName = $SplitParams[1]
            } else {
                $FieldType = ''
                $FieldName = $SplitParams[1]
            }
            $Params += $FieldType + '$' + $FieldName
            $SetItem.Add('FieldType',[string]$FieldType)
            $SetItem.Add('FieldName',[string]$FieldName)
            $ParamsPSOArray += New-Object -TypeName psobject -Property $SetItem
        }
    }

    $ParamsList = $Params -join ','
    $FunctionCode += $ParamsList + ")$NL{$NL"

    $FunctionCode += '[PSCustomObject]$SetItem = @{}' + $NL
    ForEach($ParamsPSO in $ParamsPSOArray) {
            $FunctionCode += '$SetItem.Add(' + "'" + $ParamsPSO.FieldName + "'" + ','+ $ParamsPSO.FieldType + '$' + $ParamsPSO.FieldName + ")" + $NL
    }

    $FunctionCode += 'return New-Object -TypeName psobject -Property $SetItem' + $NL + '}'
    Try {
        Invoke-Expression "$FunctionCode"
        # -ErrorAction SilentlyContinue
    } Catch {
        if($Validate) { 
            write-host "Helper Function $FunctionName failed to be created." -ForegroundColor Red
            return $false
        }        
    }

    if($Validate) { 
        write-host "Helper Function $FunctionName created." -ForegroundColor Green
    }
    
    return $FunctionCode
}
