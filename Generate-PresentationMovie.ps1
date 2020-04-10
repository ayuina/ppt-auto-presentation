param(
    [Parameter(Mandatory=$True)]
    [string]$pptxFile
)

begin {
    $workDir = "{0}\{1:yyyyMMdd-HHmmss}" -f [System.IO.Path]::GetDirectoryName($pptxFile), [DateTime]::Now
    $workFile = "{0}\{1}" -f $workDir, [System.IO.Path]::GetFileName($pptxFile)
    $d = [System.IO.Directory]::CreateDirectory($workDir)
    [System.IO.File]::Copy($pptxFile, $workFile)
    
    $app = New-Object -ComObject "PowerPoint.Application"
    $ppPlaceholderBody = 2
    $msoAnimEffectMediaPlay = 83
    $msoAnimTriggerAfterPrevious = 3
    $ppMediaTaskStatusInProgress = 1
    $ppMediaTaskStatusDone = 3

    $ppSaveAsMP4 = 39
    $vbCr = [char]13

    Add-Type –AssemblyName System.Speech
    $synthesizer = New-Object –TypeName System.Speech.Synthesis.SpeechSynthesizer
}

process {
    
    $pres = $app.Presentations.Open($workFile)
    try {

        Write-Host "========== generating audio from note text ========"        
        $audioOutputs = @()
        $pres.Slides | where { $_.HasNotesPage } | foreach {
            
            $slide = $_
            $audioInfo = @{ page = $slide.SlideIndex; audioFiles = @() }

            $slide.NotesPage.Shapes | where { $_.PlaceholderFormat.Type -eq $ppPlaceholderBody } | where { $_.HasTextFrame } | where { $_.TextFrame.HasText } | foreach {

                $lines = $_.TextFrame.TextRange.Text
                $noteIndex = 0
                $lines.Split($vbCr) | where { ![string]::IsNullOrWhiteSpace($_)  } | foreach { 
                    $text = $_
                    $output = "{0}\{1:000000}-{2:000000}.wav" -f $workDir, $slide.SlideIndex, $noteIndex++
                    Write-Host ("page {0} : audio {1} : text {2} " -f $slide.SlideIndex, $output, $text)

                    $synthesizer.SetOutputToWaveFile($output)
                    $synthesizer.Speak($text)
                    $audioInfo.audioFiles += $output
                }                 
            }
            $audioOutputs += $audioInfo
        }
        $synthesizer.Dispose()

        Write-Host "========== adding audio as animation effect ========"        
        $audioOutputs | foreach {
            $slide = $pres.Slides[$_.page]
            for($i = 0; $i -lt $_.audioFiles.Length; $i++) {
                Write-Host ("processing page {0}, audio file {1}" -f $_.page, $_.audioFiles[$i] )
                $mo = $slide.Shapes.AddMediaObject2($_.audioFiles[$i], $true, $false, 10, 10)
                $eff = $slide.TimeLine.MainSequence.AddEffect($mo, $msoAnimEffectMediaPlay,0, $msoAnimTriggerAfterPrevious)
                $eff.MoveTo( $i + 1 )
                $eff.EffectInformation.PlaySettings.HideWhileNotPlaying = $true
            }
              
            $_.audioFiles | foreach {
            }
        }
        $pres.Save()

        Write-Host "========== save files and export video ========"        
        $videofile = "$($workFile).mp4"
        $pres.CreateVideo($videofile, $false, 5, 720, 30, 85)
        $videofile

        while( $pres.CreateVideoStatus -eq $ppMediaTaskStatusInProgress)
        {
            Write-Host "Waiting for media output"
            Start-Sleep -Seconds 5
        }
    }
    finally {
        $pres.Close()
    }


}

end {
    $app.Quit()
}

#https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
#http://www.pptfaq.com/FAQ00481_Export_the_notes_text_of_a_presentation.htm
#https://answers.microsoft.com/en-us/msoffice/forum/all/how-do-i-add-an-mp3-narration-file-with-timing/1a670b64-4b0c-4bf3-9bd2-4d162ce30500
#https://answers.microsoft.com/en-us/msoffice/forum/all/using-vba-to-insert-and-automatically-play-audio/a56ac636-2a83-4c37-88d1-068ef01b52b9