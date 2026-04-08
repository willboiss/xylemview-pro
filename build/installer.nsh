; XylemView Pro — Custom NSIS hooks
; Posts "has left the chat." to shared chat.json on uninstall

!macro customUnInstall
  ; Post "left the chat" message via PowerShell (fire-and-forget, silent fail)
  nsExec::ExecToLog 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "\
    $$chatPaths = @(\"E:\\XylemView\\XylemView Pro\\chat.json\", \"\\\\01ckfp02-1\\Apps\\XylemView\\XylemView Pro\\chat.json\"); \
    $$chatFile = $$null; \
    foreach ($$p in $$chatPaths) { if (Test-Path $$p) { $$chatFile = $$p; break } }; \
    if ($$chatFile) { \
      try { \
        $$chat = Get-Content $$chatFile -Raw | ConvertFrom-Json; \
        $$msg = [PSCustomObject]@{ user = $$env:USERNAME; name = $$env:USERNAME; text = \"has left the chat.\"; ts = [long]((Get-Date).ToUniversalTime() - [datetime]\"1970-01-01\").TotalMilliseconds; system = $$true }; \
        $$chat += $$msg; \
        $$chat | Select-Object -Last 200 | ConvertTo-Json -Depth 3 | Set-Content $$chatFile -Encoding UTF8; \
      } catch {} \
    }"'
!macroend
