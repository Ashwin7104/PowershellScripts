<#
	{
		"info": {
			"Statement": "Code is poetry",
			"Author": "Joerg Hochwald",
			"Contact": "joerg.hochwald@outlook.com",
			"Link": "http://hochwald.net",
			"Support": "https://github.com/jhochwald/MyPowerShellStuff/issues"
		},
		"Copyright": "(c) 2012-2015 by Joerg Hochwald. All rights reserved."
	}
	Redistribution and use in source and binary forms, with or without modification,
	are permitted provided that the following conditions are met:
	1. Redistributions of source code must retain the above copyright notice, this list of
	   conditions and the following disclaimer.
	2. Redistributions in binary form must reproduce the above copyright notice,
	   this list of conditions and the following disclaimer in the documentation and/or
	   other materials provided with the distribution.
	3. Neither the name of the copyright holder nor the names of its contributors may
	   be used to endorse or promote products derived from this software without
	   specific prior written permission.
	THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
	IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
	AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
	CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
	CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
	SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
	OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
	POSSIBILITY OF SUCH DAMAGE.
	By using the Software, you agree to the License, Terms and Conditions above!
#>

# Base Directory
# This must match with the UpdateService/LocalePath entry ($Config.UpdateService.LocalePath)
# in the JSON configuration file if you want to use the automated update/Distribution features!
#$global:BaseDirectory = "C:\Users\ashwinkumars_s\Desktop\SyscoScripts\SyscoScripts\"

$global:BaseDirectory = split-path -parent $MyInvocation.MyCommand.Definition

# JSON configuration filename to use
$global:BaseConfig = "\config.json"

# Load and parse the JSON configuration file
try {
	$global:Config = Get-Content "$BaseDirectory$BaseConfig" -Raw -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue | ConvertFrom-Json -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
} catch {
	Write-PoshError -Message "The Base configuration file is missing!" -Stop
}

# Check the configuration
if (!($Config)) {
	Write-PoshError -Message "The Base configuration file is missing!" -Stop
}

<#
	Any further Script here
#>


<#
	Now we read the JSON File and use it
#>
# Internal Version information (For future use)
$global:ConfigVersion = ($Config.basic.ConfigVersion)

# Customer Info (For future use)
$global:Company = ($Config.basic.Customer)

# Environment (Production, Leaduser, Testing, Development)
$global:environment = ($Config.basic.environment)

<#
	Any further Script here
#>