param(
    [Parameter(Mandatory=$true)]
    [string]$pptxFile,
    [Parameter(Mandatory=$true)]
    [string]$configFile
)

$InformationPreference = "Continue"
. $configFile

function Main()
{
    Write-Information "Creating working directory"
    $workDirPath = "./work/{0:yyyyMMdd-HHmmss}-{1}" -f [DateTime]::Now, [System.IO.Path]::GetFileNameWithoutExtension($pptxFile)
    $workDir = [System.IO.Directory]::CreateDirectory($workDirPath)

    $sourceFilePath = Join-Path -Path $workDir.FullName -ChildPath "source.pptx"
    $original = Get-Item $pptxFile
    $source = $original.CopyTo($sourceFilePath, $true)

    Generate-Movie $source
}

function Generate-Movie ([System.IO.FileInfo]$sourceFile)
{
    $workDir = $sourceFile.Directory
    Write-Information "processing on $workDir ..."
    $outputFile = Join-Path -Path $workDir.FullName -ChildPath "output.mp4"

    $ppapp = New-Object -ComObject "PowerPoint.Application"
    try 
    {
        $presentation = $ppapp.Presentations.Open($sourceFile.FullName)
        try
        {
            # $context = ExportNotes -presentation $presentation

            # $context | foreach {
            #     $_['audio'] = Generate-Audio -workdir $workDir -page $_.page -text $_.text
            # }

            # $context | foreach {
            #     AddAudioAsAnimationEffect -presentation $presentation -page $_.page -audio $_.audio
            # }

            ExportNotes -presentation $presentation `
            | foreach {
                $_['audio'] = Generate-Audio -workdir $workDir -page $_.page -text $_.text
                Write-Output $_
            } `
            | foreach {
                AddAudioAsAnimationEffect -presentation $presentation -page $_.page -audio $_.audio
            }

            ExportTo-Movie -presentation $presentation -outputFile $outputFile

        }
        finally
        {
            $presentation.Close()
        }
    }
    finally 
    {
        $ppapp.Quit()
    }
}

function ExportNotes($presentation)
{
    $presentation.Slides | where { $_.HasNotesPage } | foreach {
        ExportNoteFromSlide $_
    }
}

function ExportNoteFromSlide($slide)
{
    Write-Information "exporting note text from page $($slide.SlideIndex) ..."

    $ppPlaceholderBody = 2

    $slide.NotesPage.Shapes `
    | where { $_.PlaceholderFormat.Type -eq $ppPlaceholderBody } `
    | where { $_.HasTextFrame } `
    | where { $_.TextFrame.HasText } `
    | foreach {
        $lines = $_.TextFrame.TextRange.Text
        Write-Output @{ page = $slide.SlideIndex; text = $lines }
    }
}

function Generate-Audio([System.IO.DirectoryInfo]$workDir, [int]$page, [string]$text)
{
    Write-Information "generationg audio for page $($page) ..."

    $ssmlOutput = Join-Path -Path $workDir.FullName -ChildPath ("{0:0000}.ssml" -f $page)
    $ssml = [xml]$config.ssmlBase.Clone()
    $ssml.speak.voice.InnerText = $text
    $ssml.Save( $ssmlOutput )

    $audioOutput = Join-Path -Path $workDir.FullName -ChildPath ("{0:0000}.{1}" -f $page, $config.speech.audioExtension)
    Invoke-RestMethod -Uri $config.speech.endpoint -Method Post `
        -Headers $config.speech.headers `
        -Body $ssml.OuterXml `
        -OutFile $audioOutput
    
    return $audioOutput
}

function AddAudioAsAnimationEffect($presentation, $page, $audio)
{
    write-Information "adding audio as animation effect for page $($page) ..."

    $slide = $presentation.Slides[$page]

    #https://learn.microsoft.com/ja-jp/office/vba/api/powerpoint.shapes.addmediaobject2
    $linkToFile = $true
    $saveWithDocument = $false
    $left = 10
    $top = 10
    $mo = $slide.Shapes.AddMediaObject2($audio, $linkToFile, $saveWithDocument, $left, $top)

    #https://learn.microsoft.com/ja-jp/office/vba/api/powerpoint.sequence.addeffect
    $msoAnimEffectMediaPlay = 83
    $msoAnimLevelNone = 0
    $msoAnimTriggerAfterPrevious = 3
    $eff = $slide.TimeLine.MainSequence.AddEffect($mo, $msoAnimEffectMediaPlay, $msoAnimLevelNone, $msoAnimTriggerAfterPrevious)
    $eff.EffectInformation.PlaySettings.HideWhileNotPlaying = $true

    $presentation.Save()
}

function ExportTo-Movie($presentation, $outputFile)
{
    write-Information "output to video : $outputFile ..."

    #https://learn.microsoft.com/ja-jp/office/vba/api/powerpoint.presentation.createvideo
    $useTimeingAndNaration = $false
    $defaultSlideDuration = 5
    $verticalResolution = 720
    $framePerSecoond = 30
    $quality = 80
    $presentation.CreateVideo($outputFile, $useTimeingAndNaration, $defaultSlideDuration, $verticalResolution, $framePerSecoond, $quality)

    $ppMediaTaskStatusInProgress = 1
    $ppMediaTaskStatusDone = 3
    while( $presentation.CreateVideoStatus -ne $ppMediaTaskStatusDone)
    {
        Write-Host "Waiting for media output $($pres.CreateVideoStatus)"
        Start-Sleep -Seconds 5
    }

    $presentation.Save()
}

Main
