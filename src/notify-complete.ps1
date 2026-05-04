param(
    [string]$Text = "NEXT",
    [ValidateRange(0, 100)]
    [int]$Volume = 55
)

try {
    $voice = New-Object -ComObject SAPI.SpVoice
    $maleVoice = $voice.GetVoices() | Where-Object { $_.GetAttribute('Gender') -eq 'Male' } | Select-Object -First 1

    if ($maleVoice) {
        $voice.Voice = $maleVoice
    }

    $voice.Volume = $Volume
    [void]$voice.Speak($Text)
}
catch {
    Write-Error "Не удалось воспроизвести голосовое уведомление: $($_.Exception.Message)"
    exit 1
}
