import re
import sqlite3
import pyautogui
import pyperclip

from time import sleep
from windows_toasts import Toast, WindowsToaster

db = r'C:\ProgramData\TK\Klientoversikt\Computers.SQLite'
intune_url = r'https://intune.microsoft.com/?l=en.en-gb#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId'

shortcuts = {
    "bh": 'Bare hyggelig! Ha en fin dag videre. :)',
    "pw": 'Ønsker du å endre passordet ditt, kan du gjøre dette ved å besøke følgende lenke:\nhttps://passwordreset.microsoftonline.com\nNB: Her må du velge "Glemt passord" når du får valget.\n\nPassordregler for Trondheim kommune:\nFor å opprette et sterkt og komplekst passord, så må det inneholde følgende:\n\n* Minst 12 tegn\n* Stor bokstav\n* Små bokstaver\n* Tall\n\nVær oppmerksom på følgende:\n*Ikke bruk Æ, Ø eller Å\n*Passordet må ikke ha vært brukt tidligere\n*Unngå å bruke ditt eget navn\n*Unngå bruk av spesialtegn\n*Passordet blir det samme for Windows-/ og Google-innlogging\n*Maks 15 tegn om du skal logge på TKsak',
    "ptk": 'Ønsker du å endre passordet ditt, kan du gjøre dette ved å besøke følgende lenke:\npassord.trondheim.komune.no\nLogg inn med BankID.\n\nPassordregler for Trondheim kommune:\nFor å opprette et sterkt og komplekst passord, så må det inneholde følgende:\n\n* Minst 12 tegn\n* Stor bokstav\n* Små bokstaver\n* Tall\n\nVær oppmerksom på følgende:\n*Ikke bruk Æ, Ø eller Å\n*Passordet må ikke ha vært brukt tidligere\n*Unngå å bruke ditt eget navn\n*Unngå bruk av spesialtegn\n*Passordet blir det samme for Windows-/ og Google-innlogging\n*Maks 15 tegn om du skal logge på TKsak',
    "ring": 'Hei,\n\nKan du ringe oss på 72540060 eller ta kontakt via chat nede til høyre i [code]<a href="https://trondheim.service-now.com/ap?id=home" target="_blank">Ansattportalen</a>[/code], så får vi sett nærmere på det.\n\nHører fra deg!',
    "lift": 'Hei,\n\nKan du forsøke følgende:\n- Åpne LIFT og fjern infokapsler. Du finner veiledning for dette her: [code]<a href="https://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0010761" target="_blank">KB0010761 - Google Chrome - Fjerne informasjonskapsler til nettsider</a>[/code]\n- Kryss ut fanen og åpne LIFT på nytt\n- Logg på med Single Sign-On når du får alternativet\n\nFungerte dette?\nHører fra deg!',
    "strøm": 'Hei,\n\nKan du forsøke følgende:\n- Ta ut alle ledninger tilkoblet PC\n- Hold inn av-knapp på PC i minimum 60 sekunder\n- Koble til ledninger igjen og slå på PC\n\nFungerte dette?\nHører fra deg!',
    "iot": 'Da skal enheten være lagt til på IoT-nettet.\n\nFår å få koblet til IoT-nettverket må man selv "Legge til nettverk" på enheten.\nDa skriver du inn følgende informasjon:\n\nNavn: Tk-IoT\n(En må skrive helt identisk som dette siden det skiller på store og små bokstaver.)\nSikkerhet: WPA2/WPA3\nOm det er mulig må en slå av valg for: Privat Wi-Fi adresse.\nPassord: ientynntrad!',
    "cbnullst": 'Er det forsøkt en powerwash/nullstilling av CB-en?\n\nNullstilling gjennomføres slik:\n1 - Skru på Chromebook\n2- Trykk og hold på knappene Esc + Last inn på nytt samtidig, og trykk deretter på av/på-knappen . Slipp av/på-knappen. Chromebook gjør nå en omstart.\nSlipp opp de andre tastene når du ser en melding på skjermen.\n3 - Når melding «Please insert a recovery USB stick or SD card.» kommer på skjermen, trykk CTRL + D\n4 - Følg instruksjoner på skjerm: Trykk Enter på første bilde, og deretter Enter igjen på neste bilde.\n5 - Koble på trådløst nett for innrullering med SSID TK-Gjestenett. Passord behøves ikke.\n6 - Godta vilkår\n7 - Logg på Chromebook med (elevensident)@skole.trondheim.kommune.no.\n8 - Sjekk om Chromebook kobler til TK-nett ca 1-2 min etter eleven har logget inn',
    "smartboard": 'Hei,\n\nKan du forsøke følgende:\n- Koble PC til smartboard\n- Høyreklikk på skrivebord (Legg ned evt programmer) og velg "Skjerminnstillinger"\n- Bla ned og klikk "Avansert skjerm"\n- Endre valgt skjerm til smartboard via nedtrekksmenyen oppe til høyre, og endre oppdateringsfrekvens til 30Hz\n\nFungerte dette?\nHører fra deg!',
    "olkweb": 'Opplæringskontoret kan hjelpe med passord og innlogging til Olkweb.\nhttps://olkweb.no/#ContactSection, skriv en melding til de under "Kontakt oss',
    "infokapsler": "Hei,\n\nKan du forsøke følgende:\n- Naviger til nettsiden der du får feilmelding.\n- Trykk på de to strekene med prikker til venstre for nettadressen.\n- Velg 'Informasjonskapsler og data fra nettsteder'\n- Velg 'Administrer informasjonskapsler og nettstedsdata'\n- Trykk på søppelbøtten på høyre side helt til det er tomt, deretter ferdig.\n- Last inn siden på nytt etterpå.",
    "powerwash": 'Det kan høres ut som noe har skjedd under innrullering/innhenting av sertifikater på Chromebooken din.\n\nDa kan du forsøke en såkalt powerwash for å tilbakestille, og så innrullere CB-en på nytt.\n- Logg ut av Chromebook\n- Trykk ned tastene "CTRL + SHIFT + ALT + R" samtidig når du står på innloggingsskjermen.\n- Velg "Restart"\n- Når Chromebook starter på nytt, velg "Powerwash" og deretter "Continue\n- Etter ny omstart, velg "Get Started"\n- Velg TK-gjestenett og trykk Next\n- Trykk på Accept and continue\nNår du har fått logget inn skal CB automatisk hopper over til TK-nett igjen.\n\nVeiledning med bilder:\nhttps://trondheim.service-now.com/ap?id=kb_article&sys_kb_id=3c8e5d881b0e1150170aa9b3b24bcb3d&table=kb_knowledge&searchTerm=powerwash',
    "iphone": 'Oppsett av e-post på mobil med iPhone (iOS)\n\nhttps://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0016313&table=kb_knowledge&searchTerm=Oppsett%20av%20e-post%20p%C3%A5%20mobil%20med%20iPhone%20(iOS)',
    "deakt": 'Merkantil eller leder på enhet må sende inn "Ny ansatt på enhet" skjema, så skal konto bli aktivert igjen.\nDe må også sørge for at dato og hovedarbeidsforhold i HR-portalen er riktig.',
    "pxe": 'Da PCen har falt ut av systemet, må en PXE Boot forsøkes.\nPXE boot veiledning:\n\nhttps://docs.google.com/document/d/1LpRE7uEKBZCP1QlzNDFlGHK8C4qoj5DIxrO2IAYe_zc/edit#heading=h.3r8kcvs2wp7b\n\nNB:\xa0\nViktig at PC er tilkobletstrøm og er tilkoblet nettverkskabel med kommunalt nett.\nNår nullstillingen er ferdig så kommer du til punktet etter du velger språk og tastatur, etter dette så må nettverket "Migreringsnett" velges hvis du ikke bruker nettverkskabel.\xa0\nPassordet til dette er "Velkommen26!',
    "oppdater": 'Hei, takk for din henvendelse!\n\nHelt ny PC mangler som regel mange oppdateringer,\xa0kan du søke på "Se etter oppdateringer"? Nederst på oppgavelinjen så er det et forstørrelsesglass som du kan trykke på å deretter søke på "Se etter oppdateringer", åpne denne.\xa0Her skal det stå "Du er oppdatert" om alt er riktig, hvis ikke så må du trykke på "Se etter oppdateringer" (dette kan ta 1-3 min før noe dukker opp), da begynner maskinen og laste ned og installere programmer, restart maskinen når det står at alle er lasted ned og klar for omstart.\n\nEtter dette så skal alt være i orden, ring oss om det fortsatt er problemer etter dette.',
    "data": "Du skal du ikke trenge å være på WIFI, 4g/5g (data roaming) er bra nok til dette. Og telefoner fra kommunen har ubegrenset data, det er også mer stabilt enn å være på kommunalt wifi.",
    "vpn": 'Hei, takk for din henvendelse!\nEnhetsleder/merkantil på enheten kan bestille tilgang til ekstern pålogging (VPN/Cisco Anyconnect) for de ansatte.\n\nBestillingen heter (Ansattilgang fagsystem) som er under bestillinger i Ansattportalen.',
    "hp": 'Du må melde en sak til Helsesupport via Ansattportalen.\nDirekte lenke til skjema: https://trondheim.service-now.com/ap?id=sc_cat_item&sys_id=45f7dc5c8779ed10900431973cbb3526\n\nI skjema så må du velge "Fagprogram" også "Helseplattformen". Da går saken direkte til de :-)\xa0\nVi har dessverre ikke tilganger eller kunnskap inne i selve Epic, så dette er noe de må sjekke opp.',
    "ikke": 'Hei,\nVi avslutter denne saken siden vi ikke har mottatt noen tilbakemelding fra deg.\n\nHvis problemet fortsatt vedvarer, er du velkommen til å gjenåpne saken eller kontakte oss på telefon 72540060, eller åpne en supportchat via følgende lenke: https://trondheim.service-now.com/ap?id=home',
    "chrome": 'Følg veiledning for tilbakestilling av Chrome:\nhttps://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0011607',
    "tknett": 'Veiledning på å koble til TK-nett mobil:\n\nTilkobling til TK-nett med iPhone -\xa0https://trondheim.service-now.com/ap?id=kb_article_view&sys_kb_id=1df4e2001b0d3590dfcd64e0b24bcb76\nTilkobling til TK-nett med Android -\xa0https://trondheim.service-now.com/ap?id=kb_article_view&sys_kb_id=297793051b5d3590170aa9b3b24bcb2c',
    "mic": 'Følg veiledning i dokumentet:\nhttps://docs.google.com/document/d/1m9Iyy1WMNPGomVRZaipHcsY47mssqua6AMZoeYbbPgQ/edit\nFungerer det etter dette?',
    "firm": 'Følg veiledning for å installere program i Firmaportalen:\nhttps://docs.google.com/document/d/1LQle405N91wBa3v9joA3BtO--KiKV_nyU9Le7MqDqLk/edit\n\n1. Åpne Firmaportal\n2. Søk på program du ønsker å installere\n3. Trykk på Installer/prøv på nytt',
    "id": 'Dette må du ta opp med ID-kontoret.\n\nE-post:\nid-kontor.postmottak@trondheim.kommune.no\n\nTlf:\n991 00 247\n\nSe her for mer informasjon om ID-kort, Yubikey og lignende:\nhttps://intranett.trondheim.kommune.no/idkontoret/',
    "clisup": 'Hei,\n\nFungerer printer ellers på PC?\nIsåfall, kan du sjekke om PC er registrert i EPIC?\nHer er veiledning for dette:\nhttps://trondheim.service-now.com/ap?id=kb_article&table=kb_knowledge&sys_kb_id=8f7a9bd31b0b5d14dfcd64e0b24bcb0e&searchTerm=clisup',
    "citrix": 'Følg veiledning for å installere Citrix i Firmaportalen:\nhttps://docs.google.com/document/d/1LQle405N91wBa3v9joA3BtO--KiKV_nyU9Le7MqDqLk/edit\n*Åpne Firmaportal\n*Søk på program du ønsker å installere\n*Trykk på Installer/prøv på nytt',
    "bestill": 'Takk for din henvedelse!\nProgrammet kan bestilles via "Ny programvare" skjema i Ansattportalen.\nMerkantil eller leder kan bestille dette for deg.',
    "ipad": 'Veiledning for å nullstille Ipad\nhttps://support.apple.com/no-no/108931\n\nHvis du har glemt passordet for Apple-kontoen din\nhttps://support.apple.com/no-no/102656',
    "auten": 'Gå til lenke via PC: aka.ms/mfasetup\nSkann Qr-kode i Microsoft autenticator appen',
    "fresh": 'Startet tilbakestilling nå, det kan ta opptil 30 minutter.\nSørg for at PC er påslått, er tilkoblet strøm og på wifi eller nettverkskabel.\n\nNår PC starter opp igjen så vil den be deg om å koble til et nettverk.\nVelg Migreringsnettet, passordet er: Velkommen26!',
    "sikker": "Er alt her prøvd:\n\n-Skal ligge et 2-sidig ark først i bunken.\n-Kodeark er brukt først i dokumentbunken.\n-Ikke flere kodeark rett etter hverandre.\n-Arkene snudd riktig vei.\n-Framsiden først ved bruk av 2-sidige QR-ark.\n-Kodearkene ikke er 'slitt', skriv i så fall ut nye ark.\n(IKKE kopier kodearkene da kvaliteten forringes. Skriv ut fra den originale PDF-filen det antallet du trenger).",
    "sysvak": 'Må logge på via sikkersone, logg på Helseplattformen og velg Kunnskapsbasen.\nDeretter skriv inn denne lenken: https://sysvaknett.fhi.no/',
    "android": 'Oppsett av e-post på mobil med Android\n\nhttps://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0016288',
    "utskrift": 'Hvordan legge til skriver til PC:\nhttps://trondheim.service-now.com/ap?id=kb_article&table=kb_knowledge&sys_kb_id=db0a90771bdd211011c85205604bcb30',
    "factory": 'Kan du prøve å nullstille den?\n\n1. Slå av telefonen\n2. Slå den på ved å holde i volum opp og slå av/på knappen samtidig!\n3. Hold på volum opp knappen men slipp slå av/på knappen når ""Samsung Galaxy"" Secured by Knox dukker opp på skjermen.\n\nNår du er inne i Android Recovery:\n\n4. Flytt deg til ""Wipe data/factory reset"" med volum opp/ned knappen\n5. Velg denne ved å trykke på slå av/på knappen\n6. Telefonen settes til Fabrikk innstillinger etter et par få sekunder\n7. Velg deretter ""Reboot system now""\n\nDu kan gjerne sjekke følgende video for å se hvordan dette gjøres:\nhttps://www.youtube.com/watch?v=sP77k1dsWKI&t=153s\n\n\nEtter at tilbakestillingen er utført må telefonen settes opp på nytt. Veiledning finner du her:\nhttps://trondheim.service-now.com/ap?id=kb_article_view&table=kb_knowledge&sysparm_article=KB0017022&searchTerm=fmu\n\n\nHåper at dette hjelper.\nHører fra deg!',
    "yubikey": '(registrere yubikey)\nLogg på med Pass/ID kort eller Bankid\n\nhttps://www.trk.buypass.no/user-portal',
    "batteri": 'Hei, takk for din henvendelse!\n\nVi kan kjøre en batteritest:\n\n1. Installer "HP Support Assistant" via Firmaportal. (Firmaportal er installert på PC-en din)\n2. Åpne programmet og trykk "Fortsett som gjest"\n3. På Hjem-skjermen, trykk på "Batteri"\n4. Koble fra laderen, og kjør batterisjekk\n5. Oppdater saken med hva resultatet er.\n\nHører fra deg. :)',
    "rammeavtale": 'https://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0015194',
    "ikonhp": 'function New-Fullpath{\nparam([string]$path)\n$tmppath = $path.Split("")\n$n = $tmppath.Count\n[string]$newpath = $null\nfor($i=0;$i -lt $n;$i++){\n$newpath += $tmppath[$i]+""\nif(-not(Test-Path $newpath)){\nNew-Item $newpath -ItemType Directory\n}\n}\n}\nNew-Fullpath -path \'C:\\ProgramData\\TK\\Icons\'\n\n$archive = \'C:\\ProgramData\\TK\\Icons\'\n$url = "https://tkafiles.blob.core.windows.net/tkicons/Helseplattformen_logo.ico"\nInvoke-WebRequest $url -OutFile \'C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico\' -UseBasicParsing\n\n#Remove old shortcuts\nif (Test-Path -Path "$env:Public\\Desktop\\HP - Prod.lnk") {\nRemove-Item -Path "$env:Public\\Desktop\\HP - Prod.lnk"\n}\nif (Test-Path -Path "$env:Public\\Desktop\\HP - NonProd Reserve.lnk") {\nRemove-Item -Path "$env:Public\\Desktop\\HP - NonProd Reserve.lnk"\n}\nif (Test-Path -Path "$env:Public\\Desktop\\HP - Non-Prod.lnk") {\nRemove-Item -Path "$env:Public\\Desktop\\HP - Non-Prod.lnk"\n}\nif (Test-Path -Path "$env:Public\\Desktop\\HP - Prod Reserve.lnk") {\nRemove-Item -Path "$env:Public\\Desktop\\HP - Prod Reserve.lnk"\n}\n\n\n$ShortcutFile = "$env:Public\\Desktop\\HP - N kkelbrikke.lnk"\xa0\n$WScriptShell = New-Object -ComObject WScript.Shell\xa0\n$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)\xa0\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\xa0\n$Shortcut.Arguments = "https://fidportaltrd.helseplattformen.no"\n$shortcut.IconLocation="C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\xa0\n$Shortcut.Save()\xa0\n\n$ShortcutFile = "$env:Public\\Desktop\\HP - BankID.lnk"\xa0\n$WScriptShell = New-Object -ComObject WScript.Shell\xa0\n$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)\xa0\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\xa0\n$Shortcut.Arguments = "https://idpportal.helseplattformen.no"\n$shortcut.IconLocation="C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\xa0\n$Shortcut.Save()\xa0\n\n$ShortcutFile = "$env:Public\\Desktop\\HP - Buypass ID.lnk"\xa0\n$WScriptShell = New-Object -ComObject WScript.Shell\xa0\n$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)\xa0\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\xa0\n$Shortcut.Arguments = "https://idpportal.helseplattformen.no"\n$shortcut.IconLocation="C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\xa0\n$Shortcut.Save()\xa0\n\n#####################################\n\n$Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs\\HP - N kkelbrikke.lnk")\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\n$Shortcut.Arguments = "https://fidportaltrd.helseplattformen.no"\n$Shortcut.IconLocation = "C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\n$Shortcut.Save()\n\n$Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs\\HP - Non-Prod.lnk")\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\n$Shortcut.Arguments = "https://fidportaltrdnp.helseplattformen.no"\n$Shortcut.IconLocation = "C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\n$Shortcut.Save()\n\n$Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs\\HP - BankID.lnk")\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\n$Shortcut.Arguments = "https://idpportal.helseplattformen.no"\n$Shortcut.IconLocation = "C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\n$Shortcut.Save()\n\n$Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs\\HP - Buypass ID.lnk")\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\n$Shortcut.Arguments = "https://idpportal.helseplattformen.no"\n$Shortcut.IconLocation = "C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\n$Shortcut.Save()\n\n$Shortcut = $WScriptShell.CreateShortcut("$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs\\HP - NonProd Reserve.lnk")\n$Shortcut.TargetPath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"\n$Shortcut.Arguments = "https://idpportalnp.helseplattformen.no"\n$Shortcut.IconLocation = "C:\\ProgramData\\TK\\Icons\\Helseplattformen_logo.ico, 0"\n$Shortcut.Save()\n',
    "googledisk": 'Kunnskapsartikkel: Google disk - hva skjer med filer når man bytter enhet?\n\nhttps://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0011154&table=kb_knowledge&searchTerm=min%20disk%20overgang%20ou%20tk',
    "adrl": 'Det er enhetsleder som må sende inn ønske om ny ADRL liste med dette skjemaet:\n\nhttps://trondheim.service-now.com/ap?id=sc_cat_item&table=sc_cat_item&sys_id=77183b83db88c490d8a881cc0b9619c6&recordUrl=com.glideapp.servicecatalog_cat_item_view.do%3Fv%3D1&sysparm_id=77183b83db88c490d8a881cc0b9619c6\n\nEnhetsleder har tilgang og vil kunne velge "opprette ADRL liste',
    "pin": 'Her er veiledning for å endre Pin Kode i Windows:\n\nhttps://trondheim.service-now.com/ap?id=kb_article_view&sysparm_article=KB0015796&table=kb_knowledge&searchTerm=pin%20kode%20pc',
}

toastShortcuts = {
    "pw": "Passordtekst (aka.ms).",
    "pw": "Passordtekst (Passord Trondheim)",
    "bh": "Sier farvel. ♥", 
    "ring": "Ber bruker ringe inn", 
    "lift": "Infokapsler LIFT.",
    "strøm": "Strømfiks!",
    "iot": "Enhet lagt til på IoT-nett.",
    "cbnullst": "Hard reset CB.",
    "smartboard": "Smartboard - Endre frekvens.",
    "olkweb": "Kontaktinfo opplæringskontoret.",
    "infokapsler": "Slette infokapsler",
    "powerwash": "Powerwash CB.",
    "iphone": "Veiledning - Oppsett av epost (iPhone)",
    "deakt": "Henvisning merkanil - Ny ansatt på enhet.",
    "pxe": "Veiledning PXE Boot.",
    "oppdater": "Se etter oppdateringer.",
    "data": "Mobilnett good. TK-Nett bad.",
    "vpn": "Henvisning merkanil - Bestilling VPN-tilgang.",
    "hp": "Henvisning brukerhjelp-skjema. HP.",
    "ikke": "Case closed. Ingen svar.",
    "chrome": "Tilbakestilling Chrome.",
    "tknett": "Veiledning TK-nett.",
    "mic": "Deaktivert mikrofon.",
    "firm": "Firmaportalen. Generell veiledning.",
    "id": "Henvisning ID-kontoret.",
    "clisup": "Er PC registrert i EPIC?",
    "citrix": "Veiledning - Installere Citrix",
    "bestill": "Henvisning merkantil - Bestilling av program.",
    "ipad": "Veiledning - Tilbakestille iPad.",
    "auten": "Veiledning - MFA setup.",
    "fresh": "Startet tilbakestilling av PC.",
    "sikker": "Veiledning - Sikker Scan feilsøking.",
    "sysvak": "Veiledning - Pålogging sysvaknett.",
    "android": "Veiledning - Oppsett av epost (Android)",
    "utskrift": "Veiledning - Legge til skriver.",
    "factory": "Veiledning - Hard reset (Android)",
    "yubikey": "Veiledning - Registrere yubikey",
    "batteri": "Veiledning - Batteritest",
    "rammeavtale": "Informasjon - Rammeavtale",
    "ikonhp": "Syke syke scriptet.",
    "googledisk": "Informasjon - Bytte av enhet.",
    "pin": "Veiledning - Endre PIN Windows."
}


# Display toast notification
def toast(message):
    toaster = WindowsToaster("Magic Button")
    newToast = Toast()
    newToast.text_fields = [message]
    toaster.show_toast(newToast)
    

# Format MAC-address from Intune's ugly string
def convertMac(macAddress):
    formattedMac = ":".join([macAddress[i:i+2] for i in range(0, 12, 2)])
    pyperclip.copy(formattedMac)
    toast("Converted string to MAC Address.")


# Get printerinfo ezpz
def compilePrinterInfo(printerInfo):
    if len(printerInfo.strip().replace("\r", "").split("\n")) > 1000:
        dump = printerInfo.strip().replace("\r", "").split("\n")[20:60]
    
    else:
        dump = printerInfo.strip().replace("\r", "").split("\n")

    lopenummer = dump[dump.index("Physical identifier")+1]
    lopenummer = re.findall(r"TKP\d+", lopenummer)[0]
    locationInfo = dump[dump.index("Location/Department")+1].split(",")
    rom = dump[dump.index("Location/Department")+1].split(",")[2:5] # Fucking comma
    tmp = ""
    
    for item in rom:
        item = item.lstrip()
        tmp += f"{item},"

    pyperclip.copy(f"Enhet: {locationInfo[0]}\nAdresse: {locationInfo[1]}\nEtg./Rom: {tmp[:-1].capitalize()}\nLøpenummer: {lopenummer}\nModell: {dump[dump.index('Type/Model')+1]}")
    toast("Compiled printer info.")


# Make text lowercase and capitalize after every period
def capitalizeString(text):
    string = re.sub(r'\. *([a-z])', lambda match: '. ' + match.group(1).upper(), text.lower())
    pyperclip.copy(string.capitalize())
    toast("De-Narinified text.")


def convertIDfromAD(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Computername_Intune, IntuneDeviceID FROM Computers WHERE Computername_AD = '{deviceID}'")
            info = info.fetchall()
            pyperclip.copy(info[0][0])
            
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Converted to new Device ID.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][1]}"

            toaster.show_toast(newToast)            
            return

    except Exception as e:
        print (e)
        toast ("Encountered error. Aborting.\nDevice might not exist yet.")
        return


def convertAndGetInfo(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Computername_Intune, IntuneDeviceID FROM Computers WHERE Computername_AD = '{deviceID}'")
            info = info.fetchall()
            info = cursor.execute(f"SELECT Model, Enrollment_date, Department, Aktiv, IntuneDeviceID FROM Computers WHERE Computername_Intune = '{deviceID}'")
            info = info.fetchall()
            pyperclip.copy (f"Løpenummer: {deviceID}\nEnhet: {info[0][2]}\nModell: {info[0][0]}\nInnrullert: {info[0][1]}\nAktiv: {info[0][3]}")

            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Converted from old ID and grabbed computer info.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][4]}"

            toaster.show_toast(newToast)
    except:
        toast("Encountered error. Aborting.\nDevice might not exist yet.")
        return
    

def getComputerInfo(deviceID):
    try:
        with sqlite3.connect(db) as connection:
            cursor = connection.cursor()
            info = cursor.execute(f"SELECT Model, Enrollment_date, Department, Aktiv, IntuneDeviceID FROM Computers WHERE Computername_Intune = '{deviceID}'")
            info = info.fetchall()
            pyperclip.copy (f"Løpenummer: {deviceID}\nEnhet: {info[0][2]}\nModell: {info[0][0]}\nInnrullert: {info[0][1]}\nAktiv: {info[0][3]}")

            print(f"{intune_url}/{info[0][4]}")
            # Custom toast for computer lookup
            toaster = WindowsToaster('MagicButton')
            newToast = Toast()
            newToast.duration
            newToast.text_fields = ["Grabbed computer info.\nClick here to open in Intune"]
            newToast.launch_action = f"{intune_url}/{info[0][4]}"

            toaster.show_toast(newToast)   
            return

    except:
        toast("Encountered error. Aborting.\nDevice might not exist yet.")
        return
    

def formatSheetADRL():
    # Move to leftmost cell
    pyautogui.leftClick()
    pyautogui.press('home')
    sleep(0.1)

    # Mark leftmost column
    pyautogui.keyDown("ctrl")
    pyautogui.press("space")
    pyautogui.keyUp("ctrl")
    sleep(0.1)
    

    # Split into columns
    pyautogui.keyDown("alt")
    pyautogui.press("d")
    pyautogui.keyUp("alt")
    pyautogui.press("e")
    sleep(0.1)

    # Alternating colors
    pyautogui.keyDown("alt")
    pyautogui.press("o")
    pyautogui.keyUp("alt")
    pyautogui.press("l")
    pyautogui.keyDown("shift")
    pyautogui.press("tab")
    pyautogui.press("tab")
    pyautogui.keyUp("shift")
    pyautogui.press("enter")
    sleep(0.1)

    # Create filter
    pyautogui.keyDown("alt")
    pyautogui.press("d")
    pyautogui.keyUp("alt")
    pyautogui.press("f")

    toast("Formatted sheet for ADRL")
    return