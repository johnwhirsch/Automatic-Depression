Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

do{        
    $webrequest = (Invoke-WebRequest -Uri https://money.cnn.com/data/markets/dow/)
    $table = $webrequest.ParsedHtml.getElementById('wsod_quoteLeft')

    foreach($row in $($table.getElementsByTagName('tr'))){

        $cells = $row.getElementsByTagName('td')
        $innerTextFound = $false;

        foreach($cell in $cells){ if($cell.innerText -eq "Open"){ $innerTextFound = $true; } }
        if($innerTextFound){ $openPrice = [int]$($cells[1].innerText) }
    }


    $changeCell = $webrequest.ParsedHtml.getElementsByTagName('td') | ? { $_.getAttributeNode('class').value -eq "wsod_change" }
    $dailyChange = $($changeCell.children)[1].innerText

    $currentCell = $webrequest.ParsedHtml.getElementsByTagName('td') | ? { $_.getAttributeNode('class').value -eq "wsod_last wsod_lastIndex" }
    $currentPrice = $($currentCell.children)[0].innerText

    if($dailyChange -like "-*"){ $changeVerb = "down" }else{ $changeVerb = "up" }
    
    $speak.Speak("The DOW is currently at $($currentPrice), which is $($changeVerb) $($dailyChange) for the day.")
       
    sleep -Seconds 900

}while( 1 -eq 1 )
