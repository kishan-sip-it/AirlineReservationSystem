# Airline Reservation System (VB6)

A classic VB6 desktop application for managing airline flight reservations —
login, flight/fare/branch/aircraft management, booking, cancellation, and
printed reports (boarding info, cancellation slip, flight-fare summary).

This copy has been cleaned up from the original student project and had two
real bugs fixed so it runs without any extra configuration (see **What was
fixed** below).

## ⚠️ Correction from an earlier version of this guide
An earlier pass at this project incorrectly pointed at `ReservationInfo.vbp`
as the main project file. That `.vbp` only wires up 2 of the 13 forms and
isn't the real app — it was a leftover/incomplete side file from development
and has been removed from this repo.

**The actual, correct project file is `AirReservation.vbp`.**

## Requirements

- Windows (7/8.1/10/11, 32 or 64-bit — the VB6 IDE itself is 32-bit only and
  isn't officially supported past Windows 7, but generally still runs fine in
  compatibility mode on newer Windows)
- **Visual Basic 6.0 Professional or Enterprise edition** (the Learning
  edition doesn't include the Data Environment / Data Report designers this
  project uses). You'll need your own legitimately-owned install media —
  this repo doesn't include or link to an installer.
- No database engine to install separately — this uses the Microsoft Jet
  engine built into Windows via `Microsoft.Jet.OLEDB.4.0`, which ships with
  Windows/Office. No ODBC DSN setup is required (see below).

## What was fixed 

The original source had two categories of bugs that would have broken it
the moment it was copied to a new machine:

1. **Most forms and both Data Environment designers connected through an
   ODBC DSN** (`Provider=MSDASQL.1;...Data Source=Air` / `...=AirReservation`)
   that doesn't exist on a fresh machine and would need to be manually
   created via `odbcad32.exe` — and the DSN name was even inconsistent
   between forms (`Air` vs `AirReservation`).
2. A few leftover/orphaned files (`frmSearchBranch.frm`, `frmInquiry.frm`,
   `frmReservation0.frm`, `AirBus1.frm`, unused report designers, temp/log
   files, stray FoxPro system files, a source-control binding file) weren't
   even part of the real, compilable project — they've been removed to avoid
   confusion. (One of them had a hardcoded path to a student's `G:\` drive —
   harmless now since it's not part of the actual app.)

**The fix:** every connection in every form and Data Environment designer now
uses a direct, DSN-free connection string:

```
Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\AirlineApp\AirReservation.mdb;Persist Security Info=False
```

This means **zero ODBC configuration** — just put the project at
`C:\AirlineApp\` (or repoint it, see below) and run it.

## Setup (first run)

1. Install VB6 Professional/Enterprise (run its setup in Windows XP SP3
   compatibility mode, as Administrator, if installing on Windows 8+).
2. Clone or copy this repo to exactly: `C:\AirlineApp\`
   (so you end up with `C:\AirlineApp\AirReservation.vbp` and
   `C:\AirlineApp\AirReservation.mdb` side by side)
   - **If you want a different folder/drive**, run
     `tools\Set-DatabasePath.ps1 -NewPath "D:\wherever\AirReservation.mdb"`
     from PowerShell first — it rewrites the connection string across every
     form and designer file for you.
3. Open VB6 → **File → Open Project** → `AirReservation.vbp`
4. If any control shows as a red X / missing in the toolbox, go to
   **Project → Components** (Ctrl+T) and re-browse to the relevant `.ocx` in
   `C:\Windows\SysWOW64\` (register with `regsvr32` first if needed — see
   the components list below).
5. Press **F5**. The app opens on `frmPassword` (the login screen).

## Login

```
Username: kishan
Password: kishan@2026
```

(Change these in `frmPassword.frm`'s `cmdOk_Click` if you want something
else — it's a plain hardcoded string check, no database table involved.)

## Bonus fix: login logic bug

The original code checked
`If txtUser.Text <> "..." And txtPassword.Text <> "..." Then reject`.
With `And`, a login is only rejected when **both** fields are wrong — so the
correct username with *any* password (or vice versa) would get you in. This
copy uses `Or` instead, so both fields must actually match.

## Required OCX/DLL components (ship with a full VB6 Pro/Enterprise install)

If VB6 setup didn't register these automatically, register them manually
from an elevated Command Prompt:

```cmd
cd C:\Windows\SysWOW64
regsvr32 MSCOMCTL.OCX
regsvr32 MSCOMCT2.OCX
regsvr32 MSDATGRD.OCX
regsvr32 MSADODC.OCX
regsvr32 MSDATLST.OCX
regsvr32 MSSTDFMT.DLL
regsvr32 MSDBRPTR.DLL
regsvr32 MSDERUN.DLL
```

## Building a standalone .exe

**File → Make AirReservation.exe** in the VB6 IDE.

To hand the compiled app to someone else, use the **Package & Deployment
Wizard** (Start Menu → Microsoft Visual Basic 6.0 Tools) — point it at the
compiled `.exe` and it bundles the VB6 runtime and OCX dependencies into a
proper installer. You'll still need to either ship the `.mdb` at
`C:\AirlineApp\` on the target machine, or re-run `Set-DatabasePath.ps1`
against a different install location before compiling.

## About "deploying this publicly"

This is a Windows desktop app, not a web app — there's no way to "host it on
the internet" the way you would a website. The Package & Deployment Wizard
above is the correct notion of "deployment" for this kind of project: an
installer you distribute to run locally on other Windows PCs. Making this
reachable over the internet would mean rebuilding the front end as a web
app (e.g. ASP.NET or similar) against a server-hosted database — a separate
project, not a deployment step, if that's ever something you want to tackle.

## Project structure

```
AirReservation.vbp          Main project file — open this in VB6
AirReservation.mdb          Access database (must sit next to the project,
                             or wherever tools/Set-DatabasePath.ps1 points)
AirReservationMDI.frm       MDI parent window / main menu
frmPassword.frm             Login screen (startup form)
frmAirBus.frm               Aircraft management
frmBranchInfo.frm           Branch info
frmClassConfirm.frm         Seat class confirmation
frmControl.frm / frmControlInfo.frm   Admin/control panel
frmFare.frm                 Fare management
frmFlight.frm / frmFlightSchedule.frm Flight & schedule management
frmReservation.frm          Booking/reservation screen
frmSearch.frm               Search dialog
frmService.frm              Service management
modAirBus.bas               Shared module
env*.Dsr                    Data Environment designers
AllFlightFareReport.Dsr
CancelReservation.Dsr
FlightFareReport.Dsr        Printable reports (Data Report designer)
tools/Set-DatabasePath.ps1  Repoint the DB path if you move the project
```

## License

No license was included with the original student project. Add one
(MIT is a common choice for this kind of academic project) before treating
this as open source.
