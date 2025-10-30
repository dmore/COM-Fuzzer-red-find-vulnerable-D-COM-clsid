#  Copyright 2025 Remco van der Meer. All Rights Reserved.
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#  http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

# Check if OleViewDotNet module is installed
$module = Get-Module -ListAvailable -Name OleViewDotNet

if (-not $module) {
    Write-Host "OleViewDotNet module is not installed."
    $answer = Read-Host "Do you want to install it now? (Y/N)"
    
    if ($answer -match '^[Yy]$') {
        try {
            Install-Module -Name OleViewDotNet -Scope CurrentUser -Force
            Write-Host "OleViewDotNet module installed successfully."
        }
        catch {
            Write-Host "Failed to install OleViewDotNet module. Error: $_"
            exit
        }
    }
    else {
        Write-Host "Exiting script because OleViewDotNet is required."
        exit
    }
}

# Import OleViewDotNet
Import-Module OleViewDotNet

# Source the external scripts into this module.
. "$PSScriptRoot\source\PrepareFunctions.ps1"
. "$PSScriptRoot\source\FuzzerFunctions.ps1"
. "$PSScriptRoot\source\DataExporter.ps1"
. "$PSScriptRoot\source\Neo4jDataMapper.ps1"
. "$PSScriptRoot\source\Neo4jImporter.ps1"
. "$PSScriptRoot\source\Neo4jWrapper.ps1"
. "$PSScriptRoot\source\PML-Importer.ps1"
. "$PSScriptRoot\source\ParameterFunctions.ps1"