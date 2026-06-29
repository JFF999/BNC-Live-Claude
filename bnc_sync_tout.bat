@echo off
REM ============================================================================
REM Wrapper : lance les deux synchros vers le Google Sheet, dans l'ordre.
REM   1) bnc_sync.bat          : Excel -> Sheet (A-H / A-C ; efface I-Q / D-I
REM                              pour les lignes dont le symbole a change)
REM   2) bnc_sync_affaires.bat : Surperformance -> Sheet (Pre Aff / MAJ Aff)
REM ============================================================================
set DOSSIER=C:\Users\jfilt\OneDrive\Documents\JFF\Claude\BNC-Live-Claude

echo === 1/2 : Synchro aller (Excel -^> Sheet) ===
call "%DOSSIER%\bnc_sync.bat"

echo === 2/2 : Mise a jour Les Affaires (Surperformance -^> Sheet) ===
call "%DOSSIER%\bnc_sync_affaires.bat"

echo === Termine ===
