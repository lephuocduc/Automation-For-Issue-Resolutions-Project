name: Package PowerShell App

on:
  push:
    branches:
      - master

jobs:
  build:
    runs-on: windows-latest
    steps:
      - name: Checkout code
        uses: actions/checkout@v2

      - name: PowerShell script
        uses: Amadevus/pwsh-script@v2.0.3
        with:
          script: |
            Install-Module -Name PowerShellProTools -Force
            Import-Module PowerShellProTools
            Merge-Script -ConfigFile 'package.psd1' -Verbose

      - name: Commit and Push Executable
        run: |
          # Navigate to the output directory where the .exe is located
          cd out
          
          # Check if the .exe file exists
          if (Test-Path 'AutomationProject.exe') {
              # Configure Git user details
              git config --global user.name "GitHub Action"
              git config --global user.email "action@github.com"
              
              # Stage the .exe file for commit
              git add AutomationProject.exe
              
              # Commit the changes
              git commit -m "Add AutomationProject.exe"
              
              # Push the changes back to the remote repository
              git push origin master
          } else {
              Write-Host "Executable not found in output directory."
          }
