$config = @{

ssmlBase = @"
    <speak version='1.0' xml:lang='ja-JP'>
        <voice name='ja-JP-Nanami:DragonHDLatestNeural'>
            talk script goes here
        </voice>
    </speak>
"@

speech = @{
    endpoint = 'https://yourRegion.tts.speech.microsoft.com/cognitiveservices/v1'
    headers = @{
        'Ocp-Apim-Subscription-Key' = 'your-speech-service-key'
        'Content-Type' = 'application/ssml+xml'
        'X-Microsoft-OutputFormat' = 'riff-24khz-16bit-mono-pcm'
    }
    audioExtension = 'wav'
}

}