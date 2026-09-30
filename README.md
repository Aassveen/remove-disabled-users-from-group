# Remove Disabled Users From Entra ID Group

PowerShell-script som finner deaktiverte brukerkontoer blant de direkte medlemmene i en Microsoft Entra ID-gruppe, og tilbyr å slette disse brukerne fra tenant-en.

> **Viktig:** Dette scriptet sletter brukerkontoer fra Entra ID. Det fjerner dem ikke bare fra gruppen. Test alltid med `-WhatIf` og kontroller kandidatlisten før en reell kjøring.

## Krav

- PowerShell 7 anbefales.
- Microsoft Graph PowerShell-modulene `Microsoft.Graph.Authentication`, `Microsoft.Graph.Groups` og `Microsoft.Graph.Users`.
- Den innloggede identiteten trenger `GroupMember.Read.All` og `User.ReadWrite.All`, med administratortillatelse der tenant-en krever det.

Installer modulene ved behov:

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users -Scope CurrentUser
```

## Bruk

Kjør først en simulering med gruppe-ID:

```powershell
./Remove-DisabledUsersFromGroup.ps1 -GroupId "<group-object-id>" -WhatIf
```

Du kan også slå opp gruppen med eksakt visningsnavn. Navnet må matche nøyaktig én gruppe:

```powershell
./Remove-DisabledUsersFromGroup.ps1 -GroupDisplayName "Former Contractors" -WhatIf
```

Når kandidatlisten er kontrollert, kan scriptet kjøres interaktivt. PowerShell ber om bekreftelse for hver konto:

```powershell
./Remove-DisabledUsersFromGroup.ps1 -GroupId "<group-object-id>" -Confirm
```

Et eksempel som bare forhåndsviser kandidater finnes i [examples/Preview-DisabledUsers.ps1](examples/Preview-DisabledUsers.ps1).

## Sikkerhetsdetaljer

- Bare direkte gruppemedlemmer som er brukerobjekter behandles.
- Kontoens `accountEnabled`-status leses på nytt før brukeren blir kandidat, og kontrolleres igjen før sletting.
- Aktiverte brukere og brukere med ukjent status blir ikke slettet.
- Alle kandidater vises før sletting forsøkes.
- `-WhatIf` forhåndsviser operasjonen. `-Confirm` ber om bekreftelse per bruker.

Skriptet fjerner ikke brukere fra grupper via nestede grupper; det behandler bare gruppens direkte medlemmer.
