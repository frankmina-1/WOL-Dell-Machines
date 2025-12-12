
    This script configures Wake-on-LAN (WOL) on supported Dell systems.

    It performs the following actions:
      • Installs the DellBIOSProvider PowerShell module if missing
      • Applies BIOS settings:
            - Disables Deep Sleep Control
            - Enables Wake-on-LAN (LANOnly)
      • Configures all physical network adapters for Wake-on-LAN:
            - Enables Advanced Properties for Magic Packet wake
            - Enables NIC power management WOL settings
      • Logs all activity using Start-Transcript and appends a final result
      • Uses a global try/catch/finally to ensure the transcript is closed cleanly

    Requirements:
      • Dell business-class system with supported BIOS
      • PowerShell run as Administrator
      • VC++ 2015–2019 x64 Redistributable installed (required by DellBIOSProvider)
