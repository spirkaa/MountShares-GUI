<#

# Для создания использяемого файла (.exe) используется скрипт
# https://github.com/MScholtes/PS2EXE

./PS2EXE.ps1 `
-InputFile '.\MountShares-GUI.ps1' `
-OutputFile '.\MountShares-GUI.exe' `
-iconFile '.\favicon.ico' `
-title 'Подключение сетевых папок МЭШ' `
-product 'MountShares-GUI' `
-copyright 'pavlov@vkapotne.ru' `
-version 0.1.0.0 `
-x86 `
-verbose `
-noConfigfile `
-noConsole

#>

#########

$DomainName = "uvao"
$SchoolName = "SCH-KAPOT"

$FQDN = "$DomainName.obr.mos.ru"
$RootPath = "\\$FQDN\$SchoolName"

#########

Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

$ConnectForm = New-Object system.Windows.Forms.Form
$ConnectForm.text = "Подключение сетевых папок МЭШ"
$ConnectForm.ClientSize = "320,210"
$ConnectForm.BackColor = "#00acc1"
$ConnectForm.TopMost = $true
$ConnectForm.AutoScalemode = "Dpi"
$ConnectForm.AutoSize = $true
$ConnectForm.AutoSizeMode = "GrowOnly"
$ConnectForm.Font = "Microsoft Sans Serif, 10"

$HeaderLabel = New-Object system.Windows.Forms.Label
$HeaderLabel.text = "Подключение сетевых папок МЭШ. Введите данные вашей учетной записи ЭЖД"
$HeaderLabel.TextAlign = "MiddleCenter"
$HeaderLabel.width = 300
$HeaderLabel.height = 65
$HeaderLabel.Top = 10
$HeaderLabel.Left = 10

$LoginLabel = New-Object system.Windows.Forms.Label
$LoginLabel.text = "Логин"
$LoginLabel.TextAlign = "MiddleLeft"
$LoginLabel.AutoSize = $true
$LoginLabel.Top = $HeaderLabel.height + 15
$LoginLabel.Left = 10

$LoginText = New-Object system.Windows.Forms.TextBox
$LoginText.multiline = $false
$LoginText.AutoSize = $true
$LoginText.width = 200
$LoginText.Top = $HeaderLabel.height + 15
$LoginText.Left = $LoginLabel.width + 10
$LoginText.Font = 'Microsoft Sans Serif,12'

$PasswordLabel = New-Object system.Windows.Forms.Label
$PasswordLabel.text = "Пароль"
$PasswordLabel.TextAlign = "MiddleLeft"
$PasswordLabel.AutoSize = $true
$PasswordLabel.Top = $LoginLabel.Top + 30
$PasswordLabel.Left = 10

$PasswordText = New-Object system.Windows.Forms.MaskedTextBox
$PasswordText.multiline = $false
$PasswordText.AutoSize = $true
$PasswordText.width = 200
$PasswordText.Top = $LoginText.Top + 30
$PasswordText.Left = $PasswordLabel.width + 10
$PasswordText.Font = 'Microsoft Sans Serif,12'
$PasswordText.PasswordChar = '*'
$PasswordText.Add_KeyDown( { maskedTextBoxKeyDown })

$connectButton = New-Object system.Windows.Forms.Button
$connectButton.text = "Подключить"
$connectButton.BackColor = "#ffffff"
$connectButton.AutoSize = $true
$connectButton.width = 200
$connectButton.Top = $PasswordText.Top + 30
$connectButton.Left = $LoginLabel.width + 10
$connectButton.Add_Click( { connectSharedFolders })

$statusBar = New-Object System.Windows.Forms.StatusBar
$statusBar.Name = "statusBar"
$statusBar.Text = ""

# https://blog.netnerds.net/2015/10/use-base64-for-notifyicon-in-powershell/
$base64 = "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAACXBIWXMAAC4jAAAuIwF4pT92AAAD20lEQVR42u1ZWWgTQRiOWhG8+lBRER/
           qgYJQEBRRxFuw6IsP+lQU9CEmO7Pm2GSTVtDUA22y2dTbIAh9NHZnW4WCUGhR9MWXgkXxlla88KCCglfjP6LJ7JizuaruwA/JzjD/9+1+/7
           GzFsuvESYoPsrsnCWfYRIYTQQUIpxVdLy13Aa+nxSFgEqwzVKBAb77TAI8AZvTX4tdfls5bP9J5+CBM844tSbVdS3dOoQaa3ImgCV5K3L74
           qPJBKd38X9NoK9U5mjyvnfu836ittcnv2Dm+otGoBLBLXi9M00CJgGTQAkIwP9usJelMOz2ecpBoK9UuR67/QGTQDYCgiRvK11vJC81g9gk
           YBL4RwiIojgBueUN2OWrw7I8KxAIVP1VBAp6pcyYRt3yOjpfqNkl/5KKEChCIfuCJDno8XgmpcMV1MV59MhF1QQhTITDCkEXIrqttvIEXL4
           eUWpcxOOIxbaPUzS0GoAGAcudVIdeEV2oaDP3ACp4A+8/GrWOhzvtAP8D2U7tikagGAMyylgAbgW/zzKA/gT2ECR0DZ7MxRZNnD0qCLR2Om
           aAv+6UoDX0OKyjiErsa6isShLEBSaNBWBPeeBwh2+GCF5fliw00qF24Drw85oDfyek2beUNY2OZFD9hgl+zoEnp2PC5D/ioydQRSWk6Eih+
           leJcB/WDlHpVYQA1TJo+wbjaxgk0xyPW8aw60JXPZOA5D6Yf5sqPoIxYWZFCNBCxOm9mV8DNWAHzL3KlEaLRiBTIcOSTzbsr9mXw95fmSzT
           zt55+nTAr5oG9Dd4IrfB2uD33hNd4oRSE/hg9fmqDYFLhC7GR3/0inWiQVoEdaQA/ijcjjaza/P+RnbwvCOeLwFoe1V2XYSgFYZ9ARRHrpX
           zOwzF7TSNhYI/8mUikK4btUrSNIOuoXoye17n4mIn5/NlujrQ0rl7iqqLy7JKKFcCuYxTBNXAPp9/7xfS8drfczR1UsBG+Qr1fOYCSe+CuX
           sZgzjdl8NGRTpaCAGFCA1M1nlDe59kYOOAAbyGogZpdQrzU3WkKQmkG4WmUZWgI8m7iy6xRQquvWN7H7aYUfCwfjCFKr6HiG16GQlgLQlS2
           JMoVqBzQ02Ap81lpXsc8JjSjjeB/qfmBYAnkPE0TfJbU8RWQgIRTdzIXD/OZp1jMWs1V8xY8GG+Wo+YQBYb4nsZcP4lIaEOcVWSAO5lAN7l
           SPcn5nR0q6DmKz8C8oCx37cvNAQppEAm999PzuE29s2M6jwhLQ0dKuxQyeWvz/1M399rKGCafe7PTPPL2OwB4D4y5FAiNi7jOYbXRw2vzIT
           vB+fCwDlS6QmBAAAAAElFTkSuQmCC"

# Create a streaming image by streaming the base64 string to a bitmap streamsource
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()
$bitmap.Freeze()

# Convert the bitmap into an icon
$image = [System.Drawing.Bitmap][System.Drawing.Image]::FromStream($bitmap.StreamSource)
$icon = [System.Drawing.Icon]::FromHandle($image.GetHicon())

$ConnectForm.controls.AddRange(@($HeaderLabel, $LoginLabel, $PasswordLabel, $LoginText, $PasswordText, $connectButton, $statusBar))
$ConnectForm.FormBorderStyle = 'Fixed3D'
$ConnectForm.MaximizeBox = $false
$ConnectForm.MaximumSize = $ConnectForm.Size
$ConnectForm.MinimumSize = $ConnectForm.Size
$ConnectForm.StartPosition = "CenterScreen"
$ConnectForm.Icon = $icon

function connectSharedFolders {
    # https://blogs.technet.microsoft.com/dsheehan/2018/06/23/confirmingvalidating-powershell-get-credential-input-before-use/

    $ValidAccount = $False
    
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    if ($LoginText.Text -and $PasswordText.Text) {
        $UserName = $LoginText.Text
        $Password = ConvertTo-SecureString $PasswordText.Text -AsPlainText -Force

        $FailureMessage = $Null
        $credentials = New-Object System.Management.Automation.PsCredential($UserName, $Password)

        # Check domain connection. Check login and password against domain.
        $ContextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
        Try {
            $statusBar.Text = "Подключение к серверу $FQDN..."
            $PrincipalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ContextType, $FQDN
            Start-Sleep -Milliseconds 100
        }
        Catch {
            If ($_.Exception.InnerException -like "*Не удалось установить связь с сервером*") {
                $FailureMessage = "Не удалось установить связь с сервером $FQDN"
                $statusBar.Text = $FailureMessage

                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageIcon = [System.Windows.MessageBoxImage]::Error
                $MessageBody = $FailureMessage
                $MessageTitle = "Ошибка"
        
                $Result = [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
                $Result
            }
            Else {
                $FailureMessage = "Неожиданная ошибка: `"$($_.Exception.Message)`""
                $statusBar.Text = "Неожиданная ошибка"
            }
        }

        # If there wasn't a failure talking to the domain test the validation of the credentials, and if it fails record a failure message.
        If (!($FailureMessage)) {
            $ValidAccount = $PrincipalContext.ValidateCredentials($UserName, $Credentials.GetNetworkCredential().Password)
            If ($ValidAccount) {
                $statusBar.Text = "Введены правильные данные, породолжаем..."
            }
            else {
                $FailureMessage = "Неправильный логин или пароль"
                $statusBar.Text = $FailureMessage

                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageIcon = [System.Windows.MessageBoxImage]::Error
                $MessageBody = $FailureMessage
                $MessageTitle = "Ошибка"
        
                $Result = [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
                $Result
            }
        }

        if ($ValidAccount) {
            $CredentialsWithDomain = New-Object System.Management.Automation.PsCredential("$DomainName\$UserName", $Password)

            $statusBar.Text = "Освобождаем буквы H, S, T..."
            Remove-PSDrive -Name H, S, T -Scope Global -ErrorAction SilentlyContinue
            $psdrive = Get-PSDrive | Where-Object { $_.DisplayRoot -like '\\*' }
            foreach ($mapdrive in $psdrive) {
                if ($mapdrive.DisplayRoot -Like "$RootPath*") {
                    $driveLetter = ($mapdrive.Root) -replace "\\", ""
                    Start-Process -FilePath "net" -ArgumentList "use $driveLetter /delete" -WindowStyle "Hidden"
                }
            }
            Start-Sleep -Milliseconds 100

            $statusBar.Text = "Подключаем HOME..."
            New-PSDrive -Name "H" -Root "$RootPath\HOME\$UserName" -PSProvider "FileSystem" -Credential $CredentialsWithDomain -Persist -Scope Global -ErrorAction SilentlyContinue > $null
            Start-Sleep -Milliseconds 100

            $statusBar.Text = "Подключаем Share..."
            New-PSDrive -Name "S" -Root "$RootPath\Share" -PSProvider "FileSystem" -Credential $CredentialsWithDomain -Persist -Scope Global -ErrorAction SilentlyContinue > $null
            Start-Sleep -Milliseconds 100

            $statusBar.Text = "Подключаем Data..."
            New-PSDrive -Name "T" -Root "$RootPath\Data" -PSProvider "FileSystem" -Credential $CredentialsWithDomain -Persist -Scope Global -ErrorAction SilentlyContinue > $null
            Start-Sleep -Milliseconds 100

            if ((Test-Path H:, S:, T:) -contains $false) {
                $FailureMessage = "Не все диски были подключены, попробуйте еще раз!"
                $statusBar.Text = $FailureMessage

                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageIcon = [System.Windows.MessageBoxImage]::Error
                $MessageBody = $FailureMessage
                $MessageTitle = "Ошибка"
        
                $Result = [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
                $Result
            }
            else {
                $statusBar.Text = "Диски подключены"

                $connectButton.Enabled = $false
                $LoginText.Enabled = $false
                $PasswordText.Enabled = $false

                # Open explorer window
                explorer.exe

                $ConnectForm.Close()
            }
        }
    }
    else {
        $statusBar.Text = "Заполните все поля"
    }
}

function maskedTextBoxKeyDown {
    if ($_.KeyCode -eq "Enter")
    { connectSharedFolders }
}

[void]$ConnectForm.ShowDialog()
