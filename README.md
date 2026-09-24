# netserv

Information Technology & Communication VB6 Network Enumerator (`NetServ.exe` / NetServer) that lists servers in a domain/workgroup by type via `NetServerEnum` (workstations, SQL, DCs, print, dial-in, AFP, etc.) and can send NetBIOS messages with `NetMessageBufferSend`. Open `NetServer.vbp` in the VB6 IDE.

**Source last updated:** 2026-08-27 · **Language:** VB6 · **Target:** VB6 Win32 · **Output:** WinForms exe

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `Project1` (`NetServer.vbp`) | VB6 | WinForms exe | Project1 |

## How to open

Open the `.vbp` in Visual Basic 6.0 IDE:
- `NetServer.vbp`

## Requirements

- Visual Basic 6.0 IDE
- Registered OCX/DLL dependencies referenced by the `.vbp` (may need to be installed separately):
  - `COMCTL32.OCX`

## Attribution and provenance

Working copy from my Historical Dev folder `VB/netserv`.
Company names in `.vbp` files: Information Technology & Communication.

## License

MIT © 2026 VaderConsulting for Dave Robinson's code. See `LICENSE`.
