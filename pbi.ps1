# Importa os módulos necessários
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class User32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
}
"@

function Update-PowerBI {
    Add-Type -AssemblyName System.Windows.Forms
    $hWnd = [User32]::FindWindow($null, "Power BI Desktop")
    if ($hWnd -eq [IntPtr]::Zero) {
        # Tenta encontrar a janela pelo título parcial
        $hWnd = Get-Process -Name "PBIDesktop" | ForEach-Object {
            $sb = New-Object System.Text.StringBuilder 256
            [User32]::GetWindowText($_.MainWindowHandle, $sb, $sb.Capacity) | Out-Null
            if ($sb.ToString() -like "*Power BI*") {
                return $_.MainWindowHandle
            }
        }
    }
    if ($hWnd -ne [IntPtr]::Zero) {
        [User32]::SetForegroundWindow([IntPtr]$hWnd)
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("%")
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("c")
        Start-Sleep -Seconds 1
        [System.Windows.Forms.SendKeys]::SendWait("r")
    } else {
        Write-Output "Power BI Desktop não encontrado."
    }
}

# Chama a função para atualizar o Power BI
Update-PowerBI
