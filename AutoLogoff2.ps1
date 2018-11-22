# Masa untuk logoff (dalam saat)
$Global:MasaUntukLogoff = 30*60
#$Global:MasaUntukLogoff = 60
# Masa tambahan sebelum benar-benar logoff (dalam saat)
$Global:MasaTambahan = 5*60
#$Global:MasaTambahan = 5
# Command untuk logoff komputer. Jangan usik jika tiada urusan
$Global:LogoffPath = "shutdown"
$Global:LogoffArgument = "/l", "/f"
#$Global:LogoffPath = "cmd"
#$Global:LogoffArgument = "/c", "pause"


Add-Type -AssemblyName System.Windows.Forms


function PaparMesej {
    param( [string]$tajuk, [string]$mesej, [string]$butangmsj = "Ok", [string]$jenismsj = "Information" )

    # Mulakan thread baru agar program boleh terus berjalan
    $job_paparmesej = Start-Job -Name KotakMesej -ScriptBlock {
        param( [string]$title, [string]$message, [string]$msgbutton, [string]$msgtype )
    
#        Add-Type -AssemblyName System.Windows.Forms
#        [System.Windows.Forms.MessageBox]::Show($message, $title, $msgbutton, $msgtype)
        Add-Type -AssemblyName Microsoft.VisualBasic
        [Microsoft.VisualBasic.Interaction]::MsgBox($message, "SystemModal", $title)
        # TODO: Fix this, this workaround is so dirty
    } -ArgumentList $tajuk, $mesej, $butangmsj, $jenismsj

    return $job_paparmesej
}

function TungguMesej {
    param( [System.Management.Automation.Job]$job_paparmesej, [long]$timeout = -1 )

    # Tunggu kotak mesej ditutup
    if ($timeout -lt 0) {
        Wait-Job -Job $job_paparmesej | Out-Null
    } else {
        Wait-Job -Job $job_paparmesej -Timeout $timeout | Out-Null
    }

    # Tentukan sama ada kotak mesej telah tertutup atau tidak
    if ($job_paparmesej.State -eq "Running" -or $job_paparmesej.HasMoreData -eq $false) {
        return $false
    } else {
        return $true
    }
}

function KeputusanMesej {
    param( [System.Management.Automation.Job]$job_paparmesej )

    # Tentukan sama ada kotak mesej telah ditutup
    if ($job_paparmesej.State -eq "Running") {
        return $null
    }

    # Dapatkan keputusan
    $keputusan = (Receive-Job -Job $job_paparmesej).Value

    # Pulangkan keputusan
    return $keputusan
}

function BersihkanMesej {
    Remove-Job -Name KotakMesej -Force
}


$Global:Ikon = [System.Windows.Forms.NotifyIcon]$null

function SetupIkon {
    # Diambil dari: https://social.technet.microsoft.com/Forums/en-US/16444c7a-ad61-44a7-8c6f-b8d619381a27/using-icons-in-powershell-scripts?forum=winserverpowershell
    $code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
    
    Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

    $Global:Ikon = New-Object System.Windows.Forms.NotifyIcon
    $Global:Ikon.Text = "Auto-Logoff"
    $Global:Ikon.Icon = [System.IconExtractor]::Extract("imageres.dll", 202, $true)
    $Global:Ikon.Visible = $true
    
    Register-ObjectEvent -InputObject $Global:Ikon -EventName MouseClick -SourceIdentifier IkonDiklik -Action {TriggerKlikIkon} | Out-Null
    Register-ObjectEvent -InputObject $Global:Ikon -EventName MouseDoubleClick -SourceIdentifier IkonDidwiklik -Action {TriggerDwiklikIkon} | Out-Null
}

function TriggerKlikIkon {
    PaparMesejIkon "$(FormatMasa $Global:MasaTinggal)" "Komputer boleh digunakan selama $(FormatMasa $Global:MasaTinggal)"
}

function TriggerDwiklikIkon {
    MintaMasaTambahan
}

function PaparMesejIkon {
    param( [string]$title, [string]$mesej )

    # Setup tajuk dan mesej
    $Global:Ikon.BalloonTipTitle = $title
    $Global:Ikon.BalloonTipText = $mesej
    $Global:Ikon.ShowBalloonTip(10000)
}

function BersihkanIkon {
    $Global:Ikon.Dispose()
    Unregister-Event -SourceIdentifier IkonDiklik
    Remove-Job -Name IkonDiklik
    Unregister-Event -SourceIdentifier IkonDidwiklik
    Remove-Job -Name IkonDidwiklik
}


function MintaMasaTambahan {
#    Start-Job -Name MintaMasaTambahan -ScriptBlock {
        $maklumat = Get-Credential -Message "Input katalaluan untuk masa tambahan:" -UserName "Pelajar"
    
        # Abaikan jika Cancel
        if ($maklumat -eq $null) {
            return
        }
    
        $username = $maklumat.UserName
        $katalaluan = $maklumat.GetNetworkCredential().Password

        Write-Host "[D] Cubaan permintaan: Username=$($username), Katalaluan=$($katalaluan)"
    
        # Diambil dari: https://stackoverflow.com/questions/10431964/powershell-to-check-local-admin-credentials
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $ds = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::Machine)
        $boleh = $ds.ValidateCredentials($username, $katalaluan)
    
        if ($boleh) {
            $Global:MasaTinggal = $Global:MasaUntukLogoff
        }
#    }
}

function BersihkanMintaMasaTambahan {
#    Remove-Job -Name MintaMasaTambahan
}


function FormatMasa {
    param( [long]$saat )

    $minit = [System.Math]::Floor($saat / 60)
    if ($minit -le 0) {
        # Gunakan unit saat
        return "$($saat) saat"
    } else {
        # Gunakan unit minit
        return "$($minit) minit"
    }
}


$Global:MasaTinggal = [long]$null

function TikPemasa {
    Write-Host "[D] Tik pemasa berjalan untuk $(FormatMasa $Global:MasaUntukLogoff)"
    $Global:MasaTinggal = $Global:MasaUntukLogoff

    for (; $Global:MasaTinggal -gt 0; $Global:MasaTinggal--) {
        Start-Sleep -Seconds 1
        if ($Global:MasaTinggal -eq [System.Math]::Floor($Global:MasaUntukLogoff * (15/100))) {
            PaparMesejIkon "Amaran" "Anda mempunyai masa sebanyak $(FormatMasa $Global:MasaTinggal) sebelum logoff"
            Write-Host "[I] Amaran 85% masa dipaparkan."
        }
    }

#    $mesej = PaparMesej "Masa tamat" "Anda diberi $(FormatMasa $Global:MasaTambahan) masa tambahan untuk simpan semua kerja anda. Jangan simpan kerja anda dalam komputer ini!" "OkCancel" "Exclamation"
    $mesej = PaparMesej "Masa tamat" "Anda diberi $(FormatMasa $Global:MasaTambahan) masa tambahan untuk simpan semua kerja anda. Jangan simpan kerja anda dalam komputer ini!" "Ok" "SystemModal"
    Write-Host "[I] Amaran 100% masa dipaparkan."

    for ($masa = $Global:MasaTambahan*10; $masa -ge 0; $masa--) {
        Start-Sleep -Milliseconds 100

#        # Kesan jika operator mengklik Cancel, kemudian minta masa tambahan
#        If (TungguMesej $mesej 0) {
#            $keputusan = KeputusanMesej $mesej
#            switch ($keputusan) {
#                "Cancel" {
#                    MintaMasaTambahan
#                }
#            }
#        }

        # Kesan jika masa tambahan diberi
        If ($Global:MasaTinggal -gt 0) {
            Write-Host "[W] Masa tinggal berubah. Memulakan semula tik pemasa."
            return TikPemasa
        }
    }

    # Jika masa tambahan habis, logout komputer
    Write-Host "[I] Menjalankan command: $($Global:LogoffPath) $($Global:LogoffArgument)"
    Start-Process -FilePath $Global:LogoffPath -ArgumentList $Global:LogoffArgument
}


try {

    Write-Host "Auto-Logoff v2.5 BETA
"
    SetupIkon
    
    PaparMesejIkon "Peringatan" "Komputer akan logoff selepas $(FormatMasa $Global:MasaUntukLogoff). Pastikan anda mempunyai tempat untuk simpan semua kerja anda."
    TikPemasa

} finally {

    Write-Host "[W] Membersih..."
    BersihkanIkon
    BersihkanMesej
    BersihkanMintaMasaTambahan
    Write-Host "[W] Program tamat. Terima kasih."

}