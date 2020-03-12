# This script will run a curl request to get the current index prices and read them out loud to you every 15 minutes

Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

do{
    $indexes = $(curl -uri https://financialmodelingprep.com/api/v3/majors-indexes).content | ConvertFrom-Json
    foreach($index in $indexes.majorIndexesList){
        if($index.ticker -like ".*"){
            if($index.changes -like "-*"){ $changeVerb = "down" }else{ $changeVerb = "up" }

            $speak.Speak("The $($index.indexName) is currently at $($index.price), which is $($changeVerb) $($index.changes) for the day.")
        }
    }
    sleep -Seconds 900
}while( 1 -eq 1 )
