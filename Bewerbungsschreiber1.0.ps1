#------------------------------------------------
#Programm um Ordner zur suche eines MasterOrdners
#------------------------------------------------


#Vergiss nicht Um DIeses Programm Zu Nutzen Müssen wir Ein Bestehenden BewerbungsOrdner haben, dieser kann leer sein darf aber auch diplome beinhalten


# MasterOrdner auswählen mit den Dateien drinnen
$shell = New-Object -ComObject Shell.Application
$folderpath = $null
$rootFolders = $shell.NameSpace("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}").Items() | Where-Object {$_.IsFolder -eq $true}
foreach ($rootFolder in $rootFolders) {
    $folderpath = $rootFolder.GetFolder.Path
    $folder = $shell.BrowseForFolder(0, "Bitte wählen Sie einen Ordner aus, der Ihre PDF und Word Dateien enthält", 0, $folderpath)
    if ($folder -ne $null) {
        $folderpath = $folder.Self.Path
        Write-Host "Der ausgewählte Ordner ist: $folderpath"
        # Hier können Sie Ihren Code einfügen, um auf die Dateien in dem ausgewählten Ordner zuzugreifen
        break
    }
}

if ($folderpath -eq $null) {
    Write-Host "Kein Ordner ausgewählt."
}

# Masterordner für Bewerbungen
$masterOrdner = $folderpath


#------------------------------------------------
#Programm um ein Bewerbungsschreiben herzustellen
#------------------------------------------------

#----------------------------------------------------------------
# Windows Forms GUI erstellen
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Eingabe der Daten für die Bewerbung" # Fenstertitel festlegen
$Form.ClientSize = New-Object System.Drawing.Size(550,750) # Fenstergröße festlegen

#----------------------------------------------------------------
#Chat GPT-------------------------------------********************
#----------------------------------------------------------------
# Erstellt das Kontrollkästchen
$checkBox69 = New-Object System.Windows.Forms.CheckBox
$checkBox69.Location = New-Object System.Drawing.Point(100,600)
$checkBox69.Size = New-Object System.Drawing.Size(250,20)
$checkBox69.Text = "Bewebung Automatisch ausfüllen"
$form.Controls.Add($checkBox69)
#****************************************************************


#----------------------------------------------------------------
#****************************************************************
#----------------------------------------------------------------
# Erstes Label eingabe Name des Benutzers

$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Point(100,20) # Position relativ zum Fenster
$Label.Size = New-Object System.Drawing.Size(200,20) # Größe des Labels
$Label.Text = "Vor/Nachname von Ihnen:"

# Textfeld erstellen, in dem der Benutzer seinen Namen eingeben kann
$InputBox = New-Object System.Windows.Forms.TextBox
$InputBox.Location = New-Object System.Drawing.Point(100,50) # Position relativ zum Fenster
$InputBox.Size = New-Object System.Drawing.Size(200,20) # Größe des Textfelds


#----------------------------------------------------------------
# Button erstellen, der beim Klicken den eingegebenen Namen speichert und das Fenster schließt

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(100,650) # Position relativ zum Fenster
$OKButton.Size = New-Object System.Drawing.Size(75,23) # Größe des Buttons
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

# Controls (Label, Textfeld, Button) zum Fenster hinzufügen
$Form.Controls.Add($Label)
$Form.Controls.Add($InputBox)
$Form.Controls.Add($OKButton)


#----------------------------------------------------------------
#Zweite label eingabe (Firma)

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Point(100,80)
$Label2.Size = New-Object System.Drawing.Size(200,20)
$Label2.Text = "Name der Firma:"

$InputBox2 = New-Object System.Windows.Forms.TextBox
$InputBox2.Location = New-Object System.Drawing.Point(100,110)
$InputBox2.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label2)
$Form.Controls.Add($InputBox2)


#---------------------------------------------------------------
#drittes label eingabe (Adresse Firma)

$Label3 = New-Object System.Windows.Forms.Label
$Label3.Location = New-Object System.Drawing.Point(100,140)
$Label3.Size = New-Object System.Drawing.Size(200,20)
$Label3.Text = "Strasse der Firma:"

$InputBox3 = New-Object System.Windows.Forms.TextBox
$InputBox3.Location = New-Object System.Drawing.Point(100,170)
$InputBox3.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label3)
$Form.Controls.Add($InputBox3)

#------------------------------------------------------------
#Viertes label eingabe (PLZ Firma)

$Label5 = New-Object System.Windows.Forms.Label
$Label5.Location = New-Object System.Drawing.Point(100,200)
$Label5.Size = New-Object System.Drawing.Size(200,20)
$Label5.Text = "PLZ der Firma:"

$InputBox5 = New-Object System.Windows.Forms.TextBox
$InputBox5.Location = New-Object System.Drawing.Point(100,230)
$InputBox5.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label5)
$Form.Controls.Add($InputBox5)


#-------------------------------------------------------------
#Fünftes label eingabe (Ort Firma)

$Label4 = New-Object System.Windows.Forms.Label
$Label4.Location = New-Object System.Drawing.Point(100,260)
$Label4.Size = New-Object System.Drawing.Size(200,20)
$Label4.Text = "Ort der Firma:"

$InputBox4 = New-Object System.Windows.Forms.TextBox
$InputBox4.Location = New-Object System.Drawing.Point(100,290)
$InputBox4.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label4)
$Form.Controls.Add($InputBox4)


#------------------------------------------------------------
#Herr Frau Keine (Auswahlfenster)
#------------------------------------------------------------
# Label für die Auswahlmöglichkeiten erstellen
$Label6 = New-Object System.Windows.Forms.Label
$Label6.Location = New-Object System.Drawing.Point(100,330)
$Label6.Size = New-Object System.Drawing.Size(200,20)
$Label6.Text = "Anrede der Person:"

# ComboBox mit den Auswahlmöglichkeiten erstellen
$ComboBox1 = New-Object System.Windows.Forms.ComboBox
$ComboBox1.Location = New-Object System.Drawing.Point(100,360)
$ComboBox1.Size = New-Object System.Drawing.Size(200,20)
$ComboBox1.Items.AddRange(@("Herr", "Frau", "Keine"))
$ComboBox1.SelectedIndex = 0

# Label und TextBox für den Namen erstellen (sind zunächst unsichtbar)
$Label7 = New-Object System.Windows.Forms.Label
$Label7.Location = New-Object System.Drawing.Point(100,410)
$Label7.Size = New-Object System.Drawing.Size(200,20)
$Label7.Text = "Nachname der Person:"
#$Label7.Visible = $true

$InputBox7 = New-Object System.Windows.Forms.TextBox
$InputBox7.Location = New-Object System.Drawing.Point(100,440)
$InputBox7.Size = New-Object System.Drawing.Size(200,20)
#$InputBox7.Visible = $true

$Form.Controls.Add($Label7)
$Form.Controls.Add($InputBox7)


#----------------------------------------------------------------
# Herr Frau Auswahl für die ComboBox hinzufügen
$ComboBox1.add_SelectedIndexChanged({
    if ($ComboBox1.SelectedItem -eq "Herr") {
        $Label6.Text = "Sehr geehrter Herr"
    } elseif ($ComboBox1.SelectedItem -eq "Frau") {
        $Label6.Text = "Sehr geehrte Frau"
    } else {
        $Label6.Text = "Sehr geehrte Damen und Herren"
    }


    # Die Labels und Textboxen für den Namen anzeigen/ausblenden
    if ($ComboBox1.SelectedItem -eq "Keine") {
        $Label7.Visible = $false
        $InputBox7.Visible = $false
    } else {
        $Label7.Visible = $true
        $InputBox7.Visible = $true
    }
})

# Steuerelemente zum Formular hinzufügen
$Form.Controls.Add($Label6)
$Form.Controls.Add($ComboBox1)
$Form.Controls.Add($Label6)
$Form.Controls.Add($InputBox7)


#----------------------------------------------------------------
#Achtes label eingabe (Email der Person)

$Label8 = New-Object System.Windows.Forms.Label
$Label8.Location = New-Object System.Drawing.Point(100,470)
$Label8.Size = New-Object System.Drawing.Size(200,20)
$Label8.Text = "Email der Person eingeben:"

$InputBox8 = New-Object System.Windows.Forms.TextBox
$InputBox8.Location = New-Object System.Drawing.Point(100,500)
$InputBox8.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label8)
$Form.Controls.Add($InputBox8)


#----------------------------------------------------------------
#Neuntes label eingabe (Stellenbezeichnung)

$Label9 = New-Object System.Windows.Forms.Label
$Label9.Location = New-Object System.Drawing.Point(100,530)
$Label9.Size = New-Object System.Drawing.Size(200,20)
$Label9.Text = "Stellenbezeichung:"

$InputBox9 = New-Object System.Windows.Forms.TextBox
$InputBox9.Location = New-Object System.Drawing.Point(100,560)
$InputBox9.Size = New-Object System.Drawing.Size(200,20)

$Form.Controls.Add($Label9)
$Form.Controls.Add($InputBox9)


#------------------------------------------------------------
# Button als Default-Button festlegen (wird bei Enter-Druck ausgeführt)
$Form.AcceptButton = $OKButton

# Fenster öffnen und auf Button-Klick warten
$Result = $Form.ShowDialog()

# Wenn OK-Button geklickt wurde, den eingegebenen Namen speichern
if ($Result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $name = $InputBox.Text
    $firma = $InputBox2.Text
    $strasse = $InputBox3.Text
    $ort = $InputBox4.Text
    $plz = $InputBox5.Text
    $anr = $Label6.Text
    $anrname = $InputBox7.Text
    $Stellenbezeichnung = $InputBox9.Text
    $emailadresse = $InputBox8.Text
}

#------------------------------------------------
#Programm Word-Datei erstellen
#------------------------------------------------
#----------------------------------------------------------------
# Word-Objekt erstellen
$word = New-Object -ComObject "Word.Application"

# Dokument erstellen
$document = $word.Documents.Add()

Add-Type -AssemblyName System.Windows.Forms


#------------CHAT GPT-----------------------------------
#API----------------------------------------------------
#Prompts------------------------------------------------
#Kopieren in Word.--------------------------------------


#-----=-=-=-schreibe-deine-gpt/3.5-=-=---um-gptbewebung-zu-Nutzen==
#API Schlüssel
$openaiKey = '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'

#Frage Beschreiben
$question = "Schreibe mir eine Bewerbung mit den Informationen, firma $fimra ansprechsperson $anrname"

# Funktion Definieren
function ask-OpenAI {
    param(
       # [string]$question,
        [int]$tokens = 500
    )

    $url = "https://api.openai.com/v1/completions"
    $body = @{
        'model' = 'text-davinci-003'
        'prompt' = $question
        'temperature' = 0.2
        'max_tokens' = $tokens
        'top_p' = 1
        'n' = 1
        'frequency_penalty' = 1
        'presence_penalty' = 1
    } | ConvertTo-Json -Compress
    $headers = @{
        'Authorization' = "Bearer $openaiKey"
        'Content-Type' = 'application/json'
    }

    try {
        $response = Invoke-RestMethod -Uri $url -Method POST -Headers $headers -Body $body
        $answer = $response.choices.text.Trim()
        return $answer
    } catch {
        Write-Error $_.Exception.Message
    }
}

# Define the OpenAI API key
$openaiKey = "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"

# Create a new Word document object
$word = New-Object -ComObject word.application
$document = $word.documents.Add()

# Prompt the user for input
if ($checkBox69.Checked) {
$paragraph = $document.Content.Paragraphs.Add()

    $answer = ask-OpenAI -question $question -tokens 500
    $paragraph.Range.Text = "$answer"
} 

#-------------------------------------Else Word DATEI SELBST GENERIEREN UND SCHREIBEN--
else {
    # Insert text into the Word document
    $paragraph = $document.Content.Paragraphs.Add()
    $document.Content.Paragraphs.SpaceAfter = 0
    $document.Content.Paragraphs.SpaceBefore = 0

# Ändern der Schriftart und Schriftgröße
$paragraph.Range.Font.Name = "Calibri"
$paragraph.Range.Font.Size = 12

$paragraph.Range.Text += "$firma`r`n$strasse`r`n$plz $ort`r`n`r`n"

$paragraph.Range.Text += "$(Get-Date -Format 'd. MMMM yyyy')`r`n`r`nBewerbungsschreiben für die Stelle als $Stellenbezeichnung`r`n`r`n"

# Set font and size for body text
$paragraph.Range.Font.Size = 12
$paragraph.Range.Text += "$anr $anrname`r`n`r`n"
$paragraph.Range.Text += "Ich habe Ihr Inserat auf Ihrer Website gefunden und nachdem ich es gründlich durchgelesen habe, finde ich, dass meine Eigenschaften und meine Erfahrung perfekt passen."
$paragraph.Range.Text += "Hiermit bewerbe ich mich auch ohne den Hochschulabschluss, der verlangt wird."
$paragraph.Range.Text += "Ich habe dieses Jahr meinen Abschluss an der KV Business School zum Handelsschuldiplom Edupool erfolgreich abgeschlossen. In dieser Weiterbildung konnte ich viel von der wirtschaftlichen Seite der Unternehmen lernen und meine Erfahrungen in der Buchhaltung machen.`r"
$paragraph.Range.Text += "Meine Quereinsteigerausbildung zum Informatiker habe ich Anfang September bei der Benedict Schule in Zürich Altstetten angefangen. Doch dies heißt jedoch nicht, dass es mir an IT-Erfahrung mangelt. In dieser kurzen Zeit habe ich schon ein paar Module abgeschlossen und konnte so weitere Erfahrungen dazu sammeln.`n"
$paragraph.Range.Text += "Seit ich klein war, war und bin ich sehr interessiert an der IT. Ich habe schon in der Oberstufe die Programmierkurse besucht und konnte mich selbstständig sehr weit entwickeln. So habe ich allein HTML gelernt zu gebrauchen und gerade danach auch CSS. Auch in der Schule lernen wir C++, so konnte ich dort mein Skillset erweitern. Privat übe ich täglich meine Skills aus, so habe ich kleine Projekte aufgebaut und auch schon meine eigene Website erstellt. .NET ist als nächstes auf meiner Checkliste, da ich das Framework sehr spannend finde.`r`n"
$paragraph.Range.Text += "Meine engsten Freunde würden mich als motivierte Persönlichkeit beschreiben und meine Teamkameraden im Basketballverein auch dazu als Teamplayer. Ich habe Referenzen in meinem Lebenslauf angegeben. Diese Personen sind bereit, Ihnen Auskunft über mich zu geben. `r`n"
$paragraph.Range.Text += "Ich freue mich darauf, Ihnen mehr über mich bei einem persönlichen Gespräch erzählen zu können. `r`n`r`n"
$paragraph.Range.Text += "Freundliche Grüßern`n`n"
$paragraph.Range.Text += "$name`n"

# Anpassen des Abstands zwischen den Absätzen
$document
}




#----------------------------------------------------------------
# Dokument formatieren
$document.PageSetup.TopMargin = $document.Application.CentimetersToPoints(1.5)
$document.PageSetup.BottomMargin = $document.Application.CentimetersToPoints(1.5)



#------------------------------------------------
#Programm um Ordner zu erstellen 2.0 ZielOrdner erstellen auf Desktop
#------------------------------------------------

# Zielordner für kopierte Bewerbungen
$neuerOrdnerName = $firma
$desktopPath = [Environment]::GetFolderPath("Desktop")
$zielOrdner = "$desktopPath\$neuerOrdnerName"


# Erstelle Zielordner, wenn er nicht existiert
if(!(Test-Path $zielOrdner))
{
    New-Item -ItemType Directory -Path $zielOrdner
    Write-Host "Der Ordner $zielOrdner wurde erstellt."
}

# Kopiere alle Word-Dateien aus dem Masterordner in den Zielordner
Get-ChildItem -Path $masterOrdner -Filter *.docx | ForEach-Object {
    $zielPfad = Join-Path $zielOrdner $_.Name
    Copy-Item $_.FullName -Destination $zielPfad -Force
    Write-Host "Die Word-Datei $($_.Name) wurde nach $zielPfad kopiert."
}

# Kopiere alle PDF-Dateien aus dem Masterordner in den Zielordner
Get-ChildItem -Path $masterOrdner -Filter *.pdf | ForEach-Object {
    $zielPfad = Join-Path $zielOrdner $_.Name
    Copy-Item $_.FullName -Destination $zielPfad -Force
    Write-Host "Die PDF-Datei $($_.Name) wurde nach $zielPfad kopiert."
}


#------------------------------------------------
#Programm Word-Datei in Ordner speichern
#------------------------------------------------

#----------------------------------------------------------------
# Dokument speichern und schließen
$document.SaveAs("$desktopPath\$neuerOrdnerName\Bewerbungsschreiben als $Stellenbezeichnung für $firma.docx", [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)#Dokument speichern
$document.Close()# Word schliessen


#----------------------umwandeln in PDF--------------------------

$WordApp = New-Object -ComObject Word.Application
$WordDoc = $WordApp.Documents.Open("$desktopPath\$neuerOrdnerName\Bewerbungsschreiben als $Stellenbezeichnung für $firma.docx")
$WordDoc.SaveAs("$desktopPath\$neuerOrdnerName\Bewerbungsschreiben als $Stellenbezeichnung für $firma.pdf", [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF)
$WordDoc.Close()
$WordApp.Quit()

#------------------Word Dokument Löschen um nurnoch PDF Ordner zu haben

# Setze den Pfad zur .docx-Datei
$docxPath = "$desktopPath\$neuerOrdnerName\Bewerbungsschreiben als $Stellenbezeichnung für $firma.docx"

# Lösche die Datei, wenn sie existiert
if (Test-Path $docxPath) {
    Remove-Item $docxPath -Force
}



#------------------------------------------------
#Ende Skript
#------------------------------------------------
