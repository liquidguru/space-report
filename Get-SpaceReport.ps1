<#
.SYNOPSIS
    Explains what is eating your disk space, and whether each item is safe to delete.

.DESCRIPTION
    "WinDirStat, but with answers." Walks a drive or folder, finds the biggest
    directories and files, and annotates each one with:

        WHAT IT IS  - plain English description
        VERDICT     - SAFE / TOOL / REVIEW / KEEP / NEVER
        HOW         - the correct way to reclaim it, if there is one

    Anything the built-in knowledge base does not recognise can optionally be
    sent to local Ollama via delegate.py for a best-effort description, so it
    costs no Claude tokens.

    By default it only reads and reports. -Clean opens a tick-box picker of the
    files it found so you can delete a selection; see the DELETING section.

.PARAMETER Path
    Drive or folder to analyse. Default: C:\

.PARAMETER Top
    How many individual files to list. 0 lists every file over -MinSizeMB.
    Default: 40

.PARAMETER MinSizeMB
    Ignore files smaller than this. Default: 250

.PARAMETER TopFolders
    How many folders to list underneath the files. 0 hides them entirely.
    Default: 8

.PARAMETER MinFolderMB
    Ignore folders smaller than this. Default: 2000

.PARAMETER Ollama
    Ask local Ollama to describe unrecognised paths, via a delegate.py placed
    beside this script. Skipped silently if that file is not present.

.PARAMETER Html
    Also write a colour-coded HTML report and open it.

.PARAMETER Json
    Write the raw findings to a JSON file at this path.

.PARAMETER Clean
    After the report, open a tick-box list of the files found. Whatever you
    tick goes to the RECYCLE BIN (recoverable). Nothing is deleted without
    you both ticking it and confirming afterwards.

.PARAMETER Select
    With -Clean, select by path wildcard instead of opening the picker, e.g.
    -Select '*\pip\cache\*'. The same rails still apply. Confirmation is still
    required unless -Force.

.PARAMETER Permanent
    With -Clean, delete outright instead of recycling. Not recoverable.
    Needed if you want the space back immediately, since recycled files still
    occupy the drive until the bin is emptied.

.PARAMETER IncludeRisky
    With -Clean, also offer files classed KEEP. Files classed NEVER are never
    offered, with or without this switch.

.PARAMETER Force
    With -Clean, skip the typed confirmation. Use sparingly.

.NOTES
    DELETING
    Only files whose action is "delete file" can be ticked. Three categories
    are held back on purpose:
      BLOCKED   - verdict NEVER. Deleting breaks Windows or destroys data.
      use app   - belongs to a game/Store app/SDK. Deleting one file out of an
                  install corrupts it; the launcher just re-downloads. Uninstall
                  through the app instead.
      use cmd   - has a proper command (hiberfil.sys, .ost, WSL disks). The file
                  is locked or must be shrunk rather than removed.

.EXAMPLE
    .\Get-SpaceReport.ps1
    .\Get-SpaceReport.ps1 -Path F:\ -MinSizeMB 1000 -Html
    .\Get-SpaceReport.ps1 -MinSizeMB 5000 -Top 0 -TopFolders 0   # every file over 5GB
    .\Get-SpaceReport.ps1 -Ollama
#>

[CmdletBinding()]
param(
    [string]$Path        = 'C:\',
    [int]   $Top         = 40,
    [int]   $MinSizeMB   = 250,
    [int]   $TopFolders  = 8,
    [int]   $MinFolderMB = 2000,
    [switch]$Ollama,
    [switch]$Progress,
    [switch]$Html,
    [string]$Json,
    [switch]$Clean,
    [string[]]$Paths,
    [string]$Select,
    [switch]$Permanent,
    [switch]$IncludeRisky,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Knowledge base. First match wins, so specific patterns come before general.
#
# Verdicts:
#   SAFE   - disposable; rubbish, or it regenerates itself
#   TOOL   - reclaimable, but use the command/tool in 'H', not Explorer
#   REVIEW - your own data or your installed software; only you can decide
#   KEEP   - something installed needs it; deleting breaks or degrades it
#   NEVER  - deleting breaks Windows or loses data you cannot get back
# ---------------------------------------------------------------------------
$KB = @(
  # --- identity beats location. These sit at the very top so that a file is
  #     described by WHAT IT IS even when it happens to live somewhere that a
  #     broader rule below would otherwise claim (e.g. a .pst inside Temp). ---
  @{P='\.pst$'; N='Outlook data file (archive)'; V='NEVER';
    W='An Outlook personal folders file. Unlike a .ost this is frequently the ONLY copy of the mail in it - archived mail is usually not on any server.';
    H='Back it up. Compact it via File > Account Settings > Data Files > Settings > Compact Now.'}

  @{P='\.lrcat$'; N='Lightroom catalogue'; V='NEVER';
    W='Your edits, ratings, keywords and collections. Not the photos themselves, but losing it loses all your work on them.';
    H='Back it up.'}

  @{P='\\Saved Games\\|\\My Games\\'; N='Game saves'; V='NEVER';
    W='Your save files. Small, and gone forever if deleted without cloud sync.';
    H='Leave them. If anything, back them up.'}

  # --- structural containers. Anchored to match the folder EXACTLY, so they
  #     never swallow the specific rules for things underneath them. --------
  @{P='^[A-Za-z]:\\Windows\\?$'; N='Windows itself (container)'; V='NEVER';
    W='The operating system. Its size is dominated by WinSxS, which is mostly hard links and is double-counted by most scanners.';
    H='Nothing here is hand-deletable. Re-run scoped to break it down: Get-SpaceReport.ps1 -Path C:\Windows'}

  @{P='^[A-Za-z]:\\Program Files( \(x86\))?\\?$'; N='Installed programs (container)'; V='KEEP';
    W='A container for everything installed machine-wide, not a single thing. The total only tells you how much software you have.';
    H='Re-run scoped to see which application is actually large, then uninstall that one via Settings > Apps.'}

  @{P='^[A-Za-z]:\\Users\\?$'; N='All user profiles (container)'; V='KEEP';
    W='A container holding every user profile on the machine, including ones left behind by accounts you no longer use.';
    H='Re-run scoped. Old profiles can be removed via System Properties > Advanced > User Profiles.'}

  @{P='^[A-Za-z]:\\Users\\[^\\]+\\?$'; N='Your user profile (container)'; V='KEEP';
    W='Your whole profile - documents, AppData, caches, the lot. A container, not a finding.';
    H='Re-run scoped to break it down.'}

  @{P='^[A-Za-z]:\\Users\\[^\\]+\\AppData(\\Local|\\LocalLow)?\\?$'; N='AppData (container)'; V='KEEP';
    W='A container for per-application data. Holds both disposable caches and irreplaceable settings side by side, so the total is meaningless on its own.';
    H='Re-run scoped: Get-SpaceReport.ps1 -Path "$env:LOCALAPPDATA"'}

  @{P='^[A-Za-z]:\\ProgramData\\?$'; N='Shared app data (container)'; V='KEEP';
    W='A container for machine-wide application data.';
    H='Re-run scoped to see which application owns the bulk of it.'}

  @{P='\\Steam\\?$|\\Steam\\steamapps\\?$'; N='Steam library (container)'; V='REVIEW';
    W='Your Steam install and game library. Almost all of the size is installed games under steamapps\common.';
    H='Steam > Library, right-click a game > Manage > Uninstall. Everything is re-downloadable from your account.'}

  # --- the pseudo-files that scare everyone in WinDirStat ------------------
  @{P='\\pagefile\.sys$'; N='Virtual memory (page file)'; V='NEVER';
    W='Windows'' swap file. Windows offloads RAM here under pressure. Sized automatically, typically 1x-1.5x your RAM. It is held open by the kernel, which is why it looks undeletable.';
    H='Do not delete. To resize: System Properties > Advanced > Performance Settings > Advanced > Virtual memory. With 32GB RAM you can safely cap it around 8-16GB.'}

  @{P='\\swapfile\.sys$'; N='Store-app swap file'; V='NEVER';
    W='A second, small swap file used only by packaged/Store apps. Normally 16-256MB.';
    H='Do not delete. It only disappears if you disable the page file entirely.'}

  @{P='\\hiberfil\.sys$'; N='Hibernation image'; V='TOOL';
    W='A reserved block sized from your RAM, used to write memory to disk when hibernating. It also backs Fast Startup, so it exists even if you never hibernate.';
    H='powercfg /h off   (reclaims all of it; you lose hibernate AND fast startup). Or shrink it: powercfg /h /size 40'}

  @{P='\\DumpStack\.log'; N='Crash-dump stack stub'; V='NEVER';
    W='A placeholder Windows keeps open so it can write a stack trace during a bugcheck. Tiny and permanently locked.';
    H='Leave it alone.'}

  @{P='\\System Volume Information'; N='System Restore + shadow copies'; V='NEVER';
    W='This is the "critical volume information" folder. It holds System Restore points, Volume Shadow Copy snapshots (what Previous Versions restores from), disk quota and dedup databases, and indexing state. It is ACL-locked to SYSTEM, which is why scanners show it as inaccessible or report a nonsense size.';
    H='Never touch it by hand. Real size: vssadmin list shadowstorage (elevated). To cap or purge: System Protection > Configure > Max Usage / Delete, or Disk Cleanup > More Options > System Restore.'}

  @{P='\\\$Recycle\.Bin|\\\$RECYCLE\.BIN'; N='Recycle Bin for this drive'; V='TOOL';
    W='Files you deleted that are still occupying space. Every drive has its own hidden bin, including the ones you rarely look at.';
    H='Empty it from Explorer, or Settings > System > Storage > Temporary files.'}

  @{P='\\\$WinREAgent'; N='Feature-update rollback data'; V='TOOL';
    W='Working files a Windows feature update left behind so it could roll back.';
    H='Safe once the update has proven good. Disk Cleanup > Clean up system files, or Settings > Storage > Temporary files.'}

  @{P='\\\$SysReset'; N='"Reset this PC" logs'; V='SAFE';
    W='Leftover logs from a Reset This PC run, successful or abandoned.';
    H='Delete the folder.'}

  @{P='\\\$GetCurrent|\\\$Windows\.~(BT|WS)'; N='Windows Setup staging'; V='TOOL';
    W='Downloaded Windows installation or upgrade payload. Frequently several GB.';
    H='Disk Cleanup > Clean up system files > Temporary Windows installation files.'}

  @{P='\\Windows\.old'; N='Your previous Windows install'; V='TOOL';
    W='The whole of your old Windows, kept so an upgrade can be rolled back. Auto-deletes about 10 days after the upgrade. Usually 10-30GB.';
    H='Never delete by hand - the permissions will fight you and you can leave it half-removed. Settings > System > Storage > Temporary files > Previous Windows installation.'}

  # --- Windows internals ---------------------------------------------------
  @{P='\\Windows\\WinSxS'; N='Component store (side-by-side)'; V='NEVER';
    W='Every version of every Windows component, kept for updates, repair, and turning features on and off. WinDirStat OVER-REPORTS this badly: most of its files are hard links that also appear in System32, so the same bytes get counted twice. The true unique size is usually a third of what you see.';
    H='Never delete anything inside it. Real size: DISM /Online /Cleanup-Image /AnalyzeComponentStore. Shrink: DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase'}

  @{P='\\Windows\\Installer'; N='Cached MSI/MSP installers'; V='NEVER';
    W='The cached installer and patch files for every MSI-based program on the machine. Windows needs them to repair, modify, update or uninstall those programs. Commonly 10-25GB, and it is the folder people most often wreck their machine by "cleaning".';
    H='Deleting these silently breaks uninstall and repair for the affected apps. Only prune with a tool that verifies orphans (PatchCleaner or similar), never by hand.'}

  @{P='\\Windows\\SoftwareDistribution\\Download'; N='Windows Update download cache'; V='TOOL';
    W='Update payloads already downloaded, and usually already installed.';
    H='Disk Cleanup > Windows Update Cleanup. Or: net stop wuauserv, empty the folder, net start wuauserv.'}

  @{P='\\Windows\\SoftwareDistribution'; N='Windows Update working data'; V='KEEP';
    W='Update history database and state. Deleting it wipes your update history and forces a full rescan.';
    H='Only clear the Download subfolder, not this level.'}

  @{P='\\Windows\\Temp'; N='System temp folder'; V='SAFE';
    W='Scratch files from installers and services. Most of it is stale.';
    H='Delete what is not locked; in-use files will refuse, which is fine.'}

  @{P='\\Windows\\Prefetch'; N='App launch prefetch data'; V='SAFE';
    W='Small hints Windows uses to speed up application launches.';
    H='Deletable but pointless - tens of MB, rebuilds itself, and costs you launch speed meanwhile.'}

  @{P='\\Windows\\LiveKernelReports'; N='Kernel telemetry dumps'; V='SAFE';
    W='Partial dumps taken when a driver hiccuped without bluescreening. Can reach several GB on a machine with a flaky GPU driver.';
    H='Delete the .dmp files inside. Worth a glance first if you have been seeing display resets.'}

  @{P='\\Windows\\Minidump'; N='Bluescreen minidumps'; V='SAFE';
    W='Small crash dumps, one per bluescreen.';
    H='Delete, unless you are currently diagnosing crashes.'}

  @{P='\\MEMORY\.DMP$'; N='Full kernel crash dump'; V='SAFE';
    W='A complete memory image written by the last bluescreen. Can be as large as your RAM.';
    H='Delete, or Disk Cleanup > System error memory dump files.'}

  @{P='\\Windows\\Logs|\\Windows\\System32\\LogFiles'; N='Windows service logs'; V='SAFE';
    W='CBS, DISM, setup and service logs. Grows quietly for years.';
    H='Delete the old ones; leave anything from today if you are troubleshooting.'}

  @{P='\\Windows\\System32\\config'; N='Registry hives'; V='NEVER';
    W='The live Windows registry.';
    H='Do not touch, under any circumstances.'}

  @{P='\\Windows\\assembly|\\Windows\\Microsoft\.NET'; N='.NET runtimes / GAC'; V='NEVER';
    W='Installed .NET runtimes and the global assembly cache.';
    H='Remove a whole runtime via Apps & Features only if you are sure nothing uses it.'}

  @{P='\\Windows\\CSC'; N='Offline Files cache'; V='NEVER';
    W='Client-side cache of network shares marked "always available offline". May hold the ONLY copy of edits not yet synced back to the server.';
    H='Manage via Sync Center. Never by deleting files.'}

  @{P='\\Windows\\servicing'; N='Servicing stack'; V='NEVER';
    W='The machinery that installs Windows updates.';
    H='Leave alone.'}

  @{P='\\Windows\\SystemApps|\\Windows\\ImmersiveControlPanel'; N='Built-in Windows apps'; V='NEVER';
    W='Shell, Start menu, Settings and friends.';
    H='Leave alone.'}

  @{P='\\Recovery\\WindowsRE|\\Recovery\\ReAgent'; N='Windows Recovery Environment'; V='NEVER';
    W='The bootable repair image behind "Advanced startup".';
    H='Leave alone - removing it disables recovery and reset.'}

  @{P='\\ProgramData\\Microsoft\\Windows\\WER'; N='Error reporting queue'; V='SAFE';
    W='Queued application crash reports waiting to be sent to Microsoft.';
    H='Delete, or Disk Cleanup > Windows error reports.'}

  @{P='\\ProgramData\\Microsoft\\Search\\Data'; N='Windows Search index'; V='TOOL';
    W='Windows.edb - the full-text index of your files. Normally 1-10GB, and occasionally runs away to 50GB+.';
    H='Control Panel > Indexing Options > Advanced > Rebuild. Narrow the indexed locations first or it just regrows.'}

  @{P='Windows Defender\\(Scans|Definition Updates)'; N='Defender scan history / definitions'; V='TOOL';
    W='Antivirus scan history and accumulated signature versions.';
    H='"%ProgramFiles%\Windows Defender\MpCmdRun.exe" -RemoveDefinitions -DynamicSignatures'}

  @{P='DeliveryOptimization'; N='Update peer-sharing cache'; V='SAFE';
    W='Update chunks cached so other PCs on your LAN can pull them from you.';
    H='Disk Cleanup > Delivery Optimization Files.'}

  @{P='\\AppData\\Local\\Temp'; N='Your temp folder'; V='SAFE';
    W='Per-user scratch space. Installers, extractors and roughly half of all software dump here and never clean up.';
    H='Delete everything not currently locked.'}

  @{P='\\AppData\\Local\\CrashDumps'; N='App crash dumps'; V='SAFE';
    W='Per-application crash dumps, sometimes hundreds of MB each.';
    H='Delete.'}

  @{P='thumbcache|\\Explorer\\iconcache'; N='Thumbnail / icon cache'; V='SAFE';
    W='Cached thumbnails and icons for Explorer.';
    H='Disk Cleanup > Thumbnails. Rebuilds on demand.'}

  @{P='\\INetCache|\\WebCache'; N='WinINet / legacy browser cache'; V='SAFE';
    W='Cached web content used by Explorer, legacy Edge components, and any app built on WinINet.';
    H='Delete (close Explorer first if WebCache is locked).'}

  @{P='\\D3DSCache|DXCache|\\GLCache|ShaderCache|NV_Cache|OptixCache|ComputeCache'; N='GPU shader cache'; V='SAFE';
    W='Compiled shaders cached so games and GPU apps start faster. Can grow to many GB on a gaming machine.';
    H='Delete freely. The first launch of each game is slightly slower while it rebuilds.'}

  # --- vendor / driver -----------------------------------------------------
  @{P='NVIDIA Corporation\\Downloader|^[A-Za-z]:\\NVIDIA\\'; N='NVIDIA driver downloads'; V='SAFE';
    W='Extracted driver installer packages left behind by GeForce Experience - one per driver version, around 800MB each. Nothing needs them after installation.';
    H='Delete the whole folder.'}

  @{P='\\AMD\\.*Packages|\\ATI\\.*Packages'; N='AMD driver extraction'; V='SAFE';
    W='Extracted AMD driver package.';
    H='Delete once the driver is installed.'}

  @{P='\\Intel\\Logs|\\Intel\\Package Cache'; N='Intel driver leftovers'; V='SAFE';
    W='Intel installer logs and cached packages.';
    H='Delete.'}

  @{P='\\ProgramData\\Package Cache'; N='Bootstrapper installer cache'; V='KEEP';
    W='Visual Studio, .NET and VC++ redistributable installers, kept so those products can repair, update and uninstall themselves. Often 5-15GB if Visual Studio is installed.';
    H='Do not hand-delete. Uninstalling the product it belongs to clears its share properly.'}

  # --- virtualisation / containers -----------------------------------------
  @{P='ext4\.vhdx$'; N='WSL2 Linux disk image'; V='TOOL';
    W='The virtual disk for a WSL distro (or Docker Desktop''s backend). It GROWS but never shrinks by itself - deleting files inside Linux does not hand the space back to Windows.';
    H='wsl --shutdown, then either: Optimize-VHD -Path <file> -Mode Full (needs Hyper-V), or diskpart > select vdisk file="<file>" > compact vdisk. Deleting the file destroys the distro.'}

  @{P='DockerDesktop.*\.vhdx$|docker_data\.vhdx$'; N='Docker Desktop data disk'; V='TOOL';
    W='Every Docker image, container layer and volume lives inside this one file. It grows forever.';
    H='docker system prune -a --volumes (this does drop unused images and volumes - read it twice), then compact the vhdx as above.'}

  @{P='\.(vhd|vhdx|avhdx|vmdk|qcow2|vdi)$'; N='Virtual machine disk'; V='REVIEW';
    W='A virtual hard disk. Deleting it destroys that VM''s contents.';
    H='Delete the VM from its manager, or compact the disk rather than removing it.'}

  @{P='\\Hyper-V\\|\\Virtual Machines\\|\\VirtualBox VMs|\\Documents\\Virtual Machines'; N='Virtual machines'; V='REVIEW';
    W='Stored VMs and their snapshots. Snapshots and checkpoints are often bigger than the VM itself.';
    H='Merge or delete checkpoints from the VM manager first - that usually reclaims most of it without losing the VM.'}

  # --- dev toolchain caches ------------------------------------------------
  @{P='\\node_modules'; N='npm dependencies'; V='SAFE';
    W='Downloaded JavaScript packages for one project. Nothing in here is yours.';
    H='Delete; npm install (or pnpm/yarn) rebuilds it exactly from the lockfile.'}

  @{P='\\\.gradle|\\gradle\\wrapper|\\gradle\\caches'; N='Gradle cache'; V='SAFE';
    W='Downloaded Gradle distributions, dependency jars and build caches. On a machine that builds Android this is routinely 10-30GB.';
    H='Delete. The next build re-downloads (slow once, then fine). Gentler: gradle --stop, then delete only .gradle\caches\build-cache-*'}

  @{P='\\\.m2\\repository|\\\.ivy2'; N='Maven / Ivy repository'; V='SAFE';
    W='Downloaded Java dependency jars.';
    H='Delete; rebuilds on the next build.'}

  @{P='\\\.nuget\\packages|NuGetFallback'; N='NuGet package cache'; V='SAFE';
    W='Downloaded .NET packages.';
    H='dotnet nuget locals all --clear'}

  @{P='\\pip\\Cache|\\pip\\wheels'; N='pip download cache'; V='SAFE';
    W='Cached Python wheels.';
    H='pip cache purge'}

  @{P='npm-cache|\\Yarn\\Cache|\\\.pnpm-store'; N='JS package manager cache'; V='SAFE';
    W='Global download cache for npm/yarn/pnpm. Note: the pnpm store hard-links into your projects, so clearing it means reinstalling every pnpm project.';
    H='npm cache clean --force   /   yarn cache clean   /   pnpm store prune'}

  @{P='\\\.venv|\\venv\\|\\virtualenvs|\\envs\\'; N='Python virtual environment'; V='SAFE';
    W='An isolated set of Python packages for one project.';
    H='Delete; recreate with python -m venv .venv and pip install -r requirements.txt.'}

  @{P='__pycache__|\.pyc$'; N='Python bytecode cache'; V='SAFE';
    W='Compiled bytecode.';
    H='Delete; regenerates silently.'}

  @{P='\\Android\\Sdk\\system-images|\\\.android\\avd'; N='Android emulator images'; V='TOOL';
    W='Emulator system images and virtual device disks. 5-15GB each is normal, and old API levels linger forever.';
    H='Android Studio > SDK Manager (untick unused API levels) and Device Manager (delete unused AVDs). Do not hand-delete.'}

  @{P='\\Android\\Sdk'; N='Android SDK'; V='KEEP';
    W='Build tools, platforms and NDK for Android development.';
    H='Prune old platforms, build-tools and NDK versions via SDK Manager rather than deleting.'}

  @{P='\\build\\intermediates|\\build\\outputs|\\\.cxx\\|\\obj\\(Debug|Release)|\\bin\\(Debug|Release)'; N='Build output'; V='SAFE';
    W='Compiler output for a project. Fully regenerable from source.';
    H='gradlew clean / dotnet clean / cargo clean, or just delete the folder.'}

  @{P='\\target\\(debug|release)'; N='Rust / Maven build output'; V='SAFE';
    W='Compiled artefacts. Rust target directories are notoriously multi-GB.';
    H='cargo clean'}

  @{P='\\\.git\\objects|\\\.git\\lfs'; N='Git repository history'; V='KEEP';
    W='The actual commit history of a repo - not junk. LFS caches can be pruned, though.';
    H='git gc --aggressive --prune=now   /   git lfs prune. Never delete .git by hand unless you are certain the remote has everything.'}

  @{P='JetBrains\\.*(caches|index|log)|\\\.AndroidStudio|Google\\AndroidStudio'; N='IDE caches and indexes'; V='SAFE';
    W='JetBrains / Android Studio project indexes and logs.';
    H='File > Invalidate Caches, or just delete; the IDE re-indexes on next open.'}

  @{P='\\Code\\(Cache|CachedData|CachedExtension|Service Worker|logs)'; N='VS Code cache'; V='SAFE';
    W='Editor caches, old bundled versions and logs.';
    H='Delete; VS Code rebuilds them.'}

  @{P='\\\.ollama\\models|\\Ollama\\models'; N='Local LLM model weights'; V='REVIEW';
    W='Downloaded Ollama models. A 12B model is roughly 7-9GB. On a box that runs local models this is usually the single biggest lever.';
    H='ollama list, then ollama rm <model> for ones you no longer use. Check nothing else on the machine depends on a model before removing it.'}

  @{P='chocolatey\\lib-bkp|\\scoop\\cache|chocolatey\\.*cache'; N='Package manager cache'; V='SAFE';
    W='Downloaded installers for Chocolatey / Scoop packages.';
    H='choco cache remove   /   scoop cache rm *'}

  @{P='\\WindowsApps'; N='Installed Store apps'; V='KEEP';
    W='Store / packaged application binaries. ACL-locked, which is why scanners often show it oddly or not at all.';
    H='Uninstall the app normally; never delete from here.'}

  @{P='vm_bundles|\\Programs\\.*\\vm\\|WindowsSandbox'; N='App-managed VM image'; V='REVIEW';
    W='A virtual machine disk an installed application manages for its own sandbox. Deleting it out from under the app usually breaks that feature until it re-downloads several GB.';
    H='Turn the feature off inside the app, or uninstall the app. Do not delete the file directly.'}

  @{P='\\Packages\\[^\\]+\\LocalCache\\(Roaming|Local)'; N='Redirected app data (MSIX)'; V='KEEP';
    W='Despite the name "LocalCache", a packaged (MSIX/Store) app''s real AppData lives under here - Windows redirects %APPDATA% into it. This is settings and working data, not cache.';
    H='Do not delete. Reclaim from inside the app, or Settings > Apps > Advanced options > Reset (which wipes its data).'}

  @{P='\\Packages\\[^\\]+\\(LocalCache|TempState|AC\\Temp)'; N='Store app cache'; V='REVIEW';
    W='Cache and temp for one packaged app. Usually disposable, but some apps redirect real data in here - check what is inside before assuming.';
    H='Prefer the app''s own cache-clearing option, or Settings > Apps > Advanced options > Repair.'}

  @{P='\\Packages\\[^\\]+\\LocalState'; N='Store app data'; V='KEEP';
    W='The primary storage for a packaged app - settings and often real documents.';
    H='Uninstall or Reset the specific app instead.'}

  @{P='\\AppData\\Local\\Packages'; N='Store app data'; V='KEEP';
    W='Per-app storage for Store apps - includes real settings and sometimes real documents.';
    H='Uninstall or Reset the specific app instead.'}

  @{P='\\uv\\cache|archive-v0|\\AppData\\Local\\uv'; N='uv package cache'; V='SAFE';
    W='Python packages cached by uv. Fully rebuildable, and it grows fast because it stores whole unpacked wheels.';
    H='uv cache clean'}

  @{P='\\Programs\\Ollama|\\lib\\ollama\\|llama-cuda.*vendor|\\lib\\.*cuda.*\.dll$'; N='LLM runtime binaries'; V='KEEP';
    W='The CUDA and runtime libraries an installed AI tool needs in order to run at all. Big, but they are program files, not data.';
    H='Uninstall the application if you no longer want it.'}

  @{P='\.(gguf|safetensors)$|OnDeviceModel|\\models\\.*\.bin$|\\LM Studio\\models|lm-studio\\models'; N='AI model weights'; V='REVIEW';
    W='Downloaded machine-learning model weights. Individually multi-GB and they accumulate silently - usually the largest reclaimable category on a machine that runs local models.';
    H='Remove through whatever downloaded it (ollama rm, LM Studio''s model manager, etc.) so its index stays consistent. Re-downloadable, but large.'}

  @{P='-updater\\installer\.exe$|\\Update(r)?\\.*\.(exe|msi)$'; N='Downloaded app updater'; V='SAFE';
    W='An installer an application downloaded to update itself, then never cleaned up.';
    H='Delete. The app re-downloads it next time it updates.'}

  @{P='Raspberry Pi\\Imager\\cache|\\Imager\\.*download.*cache'; N='Pi Imager download cache'; V='SAFE';
    W='The last OS image Raspberry Pi Imager downloaded, kept so re-flashing is quicker.';
    H='Delete; it re-downloads on demand.'}

  @{P='\.ost$'; N='Outlook offline mailbox cache'; V='TOOL';
    W='A local copy of a mailbox that also exists on the server. Not your only copy - but rebuilding it means re-downloading the entire mailbox, which can take hours.';
    H='Do not delete while Outlook is running. Shrink properly: Account Settings > change "Mail to keep offline" to 3-12 months, or File > Account Settings > Data Files > Settings > Compact Now.'}

  @{P='\.pst$'; N='Outlook data file (archive)'; V='NEVER';
    W='An Outlook personal folders file. Unlike a .ost this is frequently the ONLY copy of the mail in it - archived mail is usually not on any server.';
    H='Back it up. Compact it via File > Account Settings > Data Files > Settings > Compact Now.'}

  # --- games ---------------------------------------------------------------
  @{P='\\steamapps\\(downloading|temp|shadercache)'; N='Steam download / shader scratch'; V='SAFE';
    W='Partially downloaded updates and compiled shaders.';
    H='Delete, or Steam > Settings > Downloads > Clear Download Cache.'}

  @{P='\\steamapps\\workshop'; N='Steam Workshop content'; V='REVIEW';
    W='Mods and community content. Re-downloadable.';
    H='Unsubscribe in the Workshop, or delete the appid folder.'}

  @{P='\\steamapps\\common'; N='Installed Steam game'; V='REVIEW';
    W='A game install. Reinstallable from your library any time - the only cost is the download.';
    H='Uninstall from Steam. Do not delete the folder directly; it leaves Steam confused about what is installed.'}

  @{P='\\Epic Games\\|GOG Galaxy\\Games|\\Battle\.net|\\Riot Games|\\Ubisoft|\\EA Games|Origin Games|Rockstar Games'; N='Installed game (other launcher)'; V='REVIEW';
    W='A game install from a non-Steam launcher. Reinstallable from your account.';
    H='Uninstall from that launcher, not from Explorer.'}

  @{P='\\EasyAntiCheat|\\BattlEye'; N='Anti-cheat runtime'; V='KEEP';
    W='Required by whichever game installed it.';
    H='It goes when the game goes.'}

  @{P='\\Saved Games|\\My Games'; N='Game saves'; V='NEVER';
    W='Your save files. Small, and gone forever if deleted without cloud sync.';
    H='Leave them. If anything, back them up.'}

  # --- creative / media ----------------------------------------------------
  @{P='Media Cache|Media Cache Files|Adobe\\Common\\Media'; N='Adobe media cache'; V='SAFE';
    W='Premiere / After Effects peak files, conformed audio and indexes. On a video machine this is one of the largest reclaimable items going - tens of GB is routine.';
    H='Premiere > Preferences > Media Cache > Delete Unused. Set a size cap while you are in there.'}

  @{P='Adobe\\.*(Cache|Temp)|Camera Raw\\Cache'; N='Adobe cache'; V='SAFE';
    W='Camera Raw / Bridge / Lightroom preview and scratch caches.';
    H='Purge from the application''s own preferences.'}

  @{P='CacheClip|\\ProxyMedia|\.gallery|BlackmagicRAW.*[Cc]ache|Resolve.*[Cc]ache'; N='DaVinci Resolve cache / proxies'; V='SAFE';
    W='Render cache, optimised media and proxies. All of it regenerates from your source footage.';
    H='Resolve > Playback > Delete Render Cache > All. Proxies: Media Pool > Delete Optimized Media.'}

  @{P='Affinity.*([Cc]ache|[Tt]emp|autosave)'; N='Affinity scratch / autosave'; V='REVIEW';
    W='Affinity Photo/Designer/Publisher scratch and autosave data. The autosave portion may hold unsaved work from a crash.';
    H='Open Affinity once and let it recover anything pending, then clear from its preferences.'}

  @{P='Previews\.lrdata'; N='Lightroom previews'; V='SAFE';
    W='Rendered preview pyramids for your catalogue. Frequently larger than the catalogue itself.';
    H='Delete the .lrdata folder with Lightroom closed; it rebuilds on demand. Do NOT delete the .lrcat next to it.'}

  @{P='\.lrcat$'; N='Lightroom catalogue'; V='NEVER';
    W='Your edits, ratings, keywords and collections. Not the photos themselves, but losing it loses all your work on them.';
    H='Back it up.'}

  @{P='Plex Media Server\\(Metadata|Media)'; N='Plex artwork and metadata'; V='REVIEW';
    W='Posters, fanart, thumbnails and analysis data. Rebuildable, but a full rebuild on a large library takes many hours.';
    H='Settings > Library > Empty Trash, and turn off video preview thumbnails - that is usually the bulk of it.'}

  @{P='transcoding-temp|\\Transcode\\|Cache\\Transcode'; N='Media server transcode scratch'; V='SAFE';
    W='Temporary transcoded segments from Plex / Emby / Jellyfin. Supposed to self-clean, frequently does not.';
    H='Delete while nothing is streaming.'}

  @{P='Emby-Server\\.*cache|\\Emby\\.*cache|Jellyfin\\.*cache'; N='Media server cache'; V='SAFE';
    W='Image and metadata cache.';
    H='Delete, or use the server''s own maintenance task.'}

  # --- cloud sync ----------------------------------------------------------
  @{P='\\OneDrive'; N='OneDrive synced files'; V='REVIEW';
    W='Your cloud files, mirrored locally. DELETING THEM HERE DELETES THEM IN THE CLOUD. This is the single most common way people lose data while "cleaning up".';
    H='Right-click > Free up space (keeps the file online-only), or enable Files On-Demand / Storage Sense. Never delete via Explorer to save space.'}

  @{P='\\Dropbox|\\Google Drive|GoogleDriveFS|\\pCloud|\\MEGAsync'; N='Cloud-synced folder'; V='REVIEW';
    W='Synced cloud storage. Deleting locally propagates the deletion to the cloud.';
    H='Use the client''s selective sync or online-only option instead of deleting.'}

  @{P='MobileSync\\Backup|iTunes\\.*Backup'; N='iPhone / iPad backup'; V='REVIEW';
    W='Full device backups. 30-150GB is common once several devices or several generations accumulate.';
    H='iTunes / Apple Devices > Preferences > Devices > delete the old ones. Keep the most recent for each device you still own.'}

  # --- browsers and chat ---------------------------------------------------
  @{P='(Chrome|Edge|Firefox|Brave|Vivaldi|Opera).*\\(Cache|Code Cache|GPUCache|Service Worker|CacheStorage)'; N='Browser cache'; V='SAFE';
    W='Cached web content and compiled scripts.';
    H='Clear from the browser, or delete the folder with it closed.'}

  @{P='Firefox\\Profiles|Chrome\\User Data\\Default|Edge\\User Data\\Default'; N='Browser profile'; V='KEEP';
    W='Your bookmarks, history, saved passwords, cookies and extensions.';
    H='Clear the cache from inside the browser instead of touching this.'}

  @{P='(Discord|Slack|Teams|Signal|WhatsApp|Telegram).*(Cache|GPUCache|Code Cache|logs)'; N='Chat app cache'; V='SAFE';
    W='Cached images, attachments and logs from a chat client.';
    H='Delete with the app closed. Telegram and WhatsApp can also clear media caches in-app.'}

  @{P='Spotify\\(Data|Storage)|Spotify\\.*[Cc]ache'; N='Spotify offline cache'; V='SAFE';
    W='Cached and downloaded tracks.';
    H='Spotify > Settings > Storage > Clear cache, or lower the offline cache limit.'}

  # --- user data -----------------------------------------------------------
  @{P='\\Users\\[^\\]+\\Downloads'; N='Your Downloads folder'; V='REVIEW';
    W='Whatever you have downloaded and forgotten about. Usually the easiest big win on any machine, but only you know what still matters.';
    H='Sort by size; installers and archives you have already used are the obvious candidates.'}

  @{P='\\Users\\[^\\]+\\(Videos|Movies)'; N='Your video files'; V='REVIEW';
    W='Your own footage and media. Irreplaceable unless it is backed up.';
    H='Archive to another drive before deleting anything.'}

  @{P='\\Users\\[^\\]+\\(Pictures|Music|Documents|Desktop)'; N='Your personal files'; V='REVIEW';
    W='Your own documents and media.';
    H='Only you can judge. Confirm your backup covers it first.'}

  @{P='\\AppData\\Local\\Programs'; N='Per-user installed applications'; V='REVIEW';
    W='Applications installed for your account only rather than machine-wide - VS Code, Ollama, Python and most Electron apps land here. These are program files, not data.';
    H='Uninstall the ones you do not use via Settings > Apps. Never delete the folders directly.'}

  @{P='\\AppData\\Local\\Microsoft'; N='Microsoft app data (mixed bag)'; V='KEEP';
    W='A bucket, not one thing. Outlook''s .ost, Edge''s profile, OneDrive, Teams and the Windows shell caches all live under here, so the total is meaningless on its own.';
    H='Re-run against this folder to break it down: Get-SpaceReport.ps1 -Path "$env:LOCALAPPDATA\Microsoft"'}

  @{P='\\AppData\\Local\\Google'; N='Chrome / Google app data'; V='KEEP';
    W='Chrome profiles, caches and on-device AI models. Profiles hold your bookmarks and passwords; the caches inside are disposable.';
    H='Clear caches from inside Chrome. Re-run against this folder to see which subfolder is actually large.'}

  @{P='\\AppData\\Roaming'; N='Per-app settings and data'; V='KEEP';
    W='Application settings, profiles, and often real user data.';
    H='Uninstall apps you do not use rather than deleting folders here.'}

  @{P='\\ProgramData\\'; N='Shared application data'; V='KEEP';
    W='Machine-wide data for installed programs.';
    H='Uninstall the owning program instead.'}

  @{P='\\Program Files( \(x86\))?\\'; N='Installed program'; V='REVIEW';
    W='An installed application.';
    H='Uninstall via Settings > Apps if you do not use it. Never delete the folder - it leaves the app half-registered.'}

  # --- file-type fallbacks -------------------------------------------------
  @{P='\.(iso|img|cue|nrg)$'; N='Disc image'; V='REVIEW';
    W='An ISO or disc image, usually a Windows or Linux installer kept "just in case".';
    H='Delete if you can re-download it.'}

  @{P='\.(zip|7z|rar|tar|gz|bz2|xz|zst)$'; N='Archive'; V='REVIEW';
    W='A compressed archive. Very often a duplicate of something already extracted right next to it.';
    H='Check whether the extracted copy exists, then delete one of the two.'}

  @{P='\.(mp4|mkv|mov|avi|mxf|braw|r3d|insv|m4v|ts|wmv|mts|m2ts)$'; N='Video file'; V='REVIEW';
    W='A video file - on this machine, quite possibly footage.';
    H='Archive to a spare drive before deleting. Check it is not the only copy.'}

  @{P='\.(psd|psb|afphoto|afdesign|afpub|tif|tiff|cr2|cr3|nef|arw|dng|raf|orf)$'; N='Image project / RAW'; V='REVIEW';
    W='A layered image project or a camera RAW file.';
    H='Archive rather than delete.'}

  @{P='\.(wav|aiff|flac|ape)$'; N='Uncompressed audio'; V='REVIEW';
    W='Lossless audio.';
    H='Your call; archive first.'}

  @{P='\.(dmp|mdmp|hdmp)$'; N='Crash dump'; V='SAFE';
    W='A memory dump from a crashed process.';
    H='Delete unless you are actively debugging it.'}

  @{P='\.(log|etl|evtx)$'; N='Log file'; V='SAFE';
    W='A log or trace file.';
    H='Delete the old ones.'}

  @{P='\.(bak|old|tmp|temp|partial|crdownload|part)$'; N='Backup / partial file'; V='REVIEW';
    W='A backup copy or an interrupted download.';
    H='Usually deletable - just confirm the .bak is not your only good copy.'}

  @{P='\.(msi|exe|appx|msix)$'; N='Installer'; V='REVIEW';
    W='A setup file, probably already used.';
    H='Delete if it is re-downloadable.'}

  @{P='\.(apk|aab)$'; N='Android build artefact'; V='REVIEW';
    W='A built Android package. Rebuildable from source.';
    H='Keep the released ones, bin the rest.'}

  @{P='\.(pak|pck|vpk|wad|bsa|ba2|assets|arc|sav)$'; N='Game / app data archive'; V='KEEP';
    W='A bulk data file belonging to an installed game or application.';
    H='Uninstall the owning app instead.'}

  @{P='\.(mdf|ldf|sqlite|edb|accdb)$'; N='Database file'; V='KEEP';
    W='A live database belonging to an application.';
    H='Use the owning app to compact or clear it.'}
)

# ---------------------------------------------------------------------------
# Fast filesystem walker. A PowerShell loop is far too slow for a whole C:\,
# so the walk itself is compiled. Reparse points are skipped deliberately:
# following junctions double-counts and can loop forever.
# ---------------------------------------------------------------------------
if (-not ('SpaceWalker' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;

public class SpaceWalker {
    public Dictionary<string,long> DirSizes = new Dictionary<string,long>(StringComparer.OrdinalIgnoreCase);
    public List<object[]> BigFiles = new List<object[]>();
    public List<string> Denied = new List<string>();
    public long MinFileBytes = 262144000;
    public long FileCount = 0;

    // Live progress for the GUI. Off by default so console runs stay clean.
    // Bytes seen so far is a good denominator for a whole-drive scan, because
    // the caller already knows how much of the drive is used.
    public bool Progress = false;
    public long SeenBytes = 0;
    private System.Diagnostics.Stopwatch _clock = System.Diagnostics.Stopwatch.StartNew();
    private long _lastMs = -1000;

    private void Tick(string where) {
        if (!Progress) return;
        long ms = _clock.ElapsedMilliseconds;
        if (ms - _lastMs < 200) return;      // at most 5 updates a second
        _lastMs = ms;
        Console.Out.WriteLine("##P\t" + SeenBytes + "\t" + FileCount + "\t" + where);
        Console.Out.Flush();
    }

    public long Walk(string path) {
        long total = 0;
        IEnumerator<FileSystemInfo> e;
        try {
            DirectoryInfo di = new DirectoryInfo(path);
            e = ((IEnumerable<FileSystemInfo>)di.EnumerateFileSystemInfos()).GetEnumerator();
        } catch (Exception) {
            if (Denied.Count < 500) Denied.Add(path);
            DirSizes[path] = -1;
            return 0;
        }

        List<string> subs = new List<string>();
        while (true) {
            FileSystemInfo fsi;
            try {
                if (!e.MoveNext()) break;
                fsi = e.Current;
            } catch (Exception) {
                if (Denied.Count < 500) Denied.Add(path);
                break;
            }
            try {
                if ((fsi.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                if ((fsi.Attributes & FileAttributes.Directory) != 0) {
                    subs.Add(fsi.FullName);
                } else {
                    long len = ((FileInfo)fsi).Length;
                    total += len;
                    FileCount++;
                    SeenBytes += len;
                    if (len >= MinFileBytes) BigFiles.Add(new object[]{ fsi.FullName, len });
                }
            } catch (Exception) { }
        }
        Tick(path);
        foreach (string s in subs) { total += Walk(s); }
        DirSizes[path] = total;
        return total;
    }
}
'@
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Format-Size {
    param([long]$Bytes)
    if ($Bytes -lt 0) { return '  (locked)' }
    $u = @('B','KB','MB','GB','TB'); $i = 0; $v = [double]$Bytes
    while ($v -ge 1024 -and $i -lt 4) { $v = $v / 1024; $i++ }
    '{0,7:N1} {1,-2}' -f $v, $u[$i]
}

function Format-Wrap {
    param([string]$Text, [int]$Width = 64)
    $out = New-Object System.Collections.Generic.List[string]
    $line = ''
    foreach ($w in ($Text -split '\s+')) {
        if ($w -eq '') { continue }
        if ($line -eq '')                              { $line = $w }
        elseif (($line.Length + 1 + $w.Length) -le $Width) { $line = "$line $w" }
        else { $out.Add($line); $line = $w }
    }
    if ($line -ne '') { $out.Add($line) }
    ,$out
}

# How an item may legitimately be removed. Matched against the classification
# name, so it lives in one place rather than being repeated across the KB.
#   never - not removable at all
#   app   - only via the owning application; deleting one file corrupts an install
#   cmd   - has a dedicated command; the file is locked or must be shrunk
#   file  - a plain file delete is the right move
$AppManaged = 'Steam|game|launcher|Anti-cheat|Store app|MSIX|WindowsApps|container|' +
              'Installed program|Android SDK|OneDrive|Cloud-synced|iPhone|iPad|' +
              'VM image|Plex|Browser profile|Redirected app data|Per-user installed'
$CmdManaged = 'Hibernation|Recycle Bin|previous Windows|rollback|Setup staging|' +
              'Update download|Search index|Defender|peer-sharing|WSL2|Docker|' +
              'Outlook|emulator images|Component store|Cached MSI|bootstrapper|' +
              'Virtual machine disk|Windows Recovery'

# Defence in depth. These are judged on the PATH, not the classification, so
# they hold even if a broad location rule matched first and called the file
# something harmless. Learned the hard way: a .pst sitting under AppData\Temp
# was classified "your temp folder, SAFE" and would have been deleted.
$NeverPlainDelete = '\\pagefile\.sys$|\\swapfile\.sys$|\\hiberfil\.sys$|\\DumpStack|' +
                    '\.(pst|lrcat|vhd|vhdx|avhdx|vmdk|qcow2|vdi)$|' +
                    '\\Saved Games\\|\\My Games\\|\\System Volume Information\\|' +
                    '\\Windows\\(WinSxS|Installer|servicing|System32\\config)\\'
$CmdByPath = '\.ost$'

function Get-DeleteMode {
    param([string]$Path, [string]$Name, [string]$Verdict)
    if ($Verdict -eq 'NEVER')           { return 'never' }
    if ($Path -match $NeverPlainDelete) { return 'never' }
    if ($Path -match $CmdByPath)        { return 'cmd' }
    if ($Name -match $CmdManaged)       { return 'cmd' }
    if ($Name -match $AppManaged)       { return 'app' }
    if ($Verdict -eq 'KEEP')            { return 'app' }   # something installed needs it
    'file'
}

function Resolve-Entry {
    param([string]$FullPath, [bool]$IsDir)
    foreach ($r in $KB) {
        if ($FullPath -match $r.P) {
            return [pscustomobject]@{
                Name = $r.N; Verdict = $r.V; What = $r.W; How = $r.H; Known = $true
                Mode = (Get-DeleteMode $FullPath $r.N $r.V)
            }
        }
    }
    $label = 'Unrecognised file'
    if ($IsDir) { $label = 'Unrecognised folder' }
    [pscustomobject]@{
        Name    = $label
        Verdict = 'UNKNOWN'
        What    = 'Not in the knowledge base, so no judgement is offered.'
        How     = 'Find out what owns it before deleting. Re-run with -Ollama for a best-effort local description.'
        Known   = $false
        Mode    = (Get-DeleteMode $FullPath $label 'UNKNOWN')
    }
}

$VerdictColour = @{
    'SAFE' = 'Green'; 'TOOL' = 'Cyan'; 'REVIEW' = 'Yellow'
    'KEEP' = 'DarkYellow'; 'NEVER' = 'Red'; 'UNKNOWN' = 'Magenta'
}
$VerdictLegend = [ordered]@{
    'SAFE'    = 'Delete freely - rubbish, or it regenerates itself'
    'TOOL'    = 'Reclaimable, but use the command shown - not Explorer'
    'REVIEW'  = 'Your data or your software - only you can decide'
    'KEEP'    = 'Something installed needs this; deleting degrades or breaks it'
    'NEVER'   = 'Deleting breaks Windows or loses data permanently'
    'UNKNOWN' = 'Not recognised - investigate before touching'
}

# ---------------------------------------------------------------------------
# Scan
#
# Skipped entirely when -Paths is supplied: the caller already knows which
# files it means, so there is nothing to discover and no report to print.
# The closing brace for this block is just above the Invoke-CleanUp function.
# ---------------------------------------------------------------------------
if (-not $Paths) {

$root = (Resolve-Path -LiteralPath $Path).Path
if (-not $root.EndsWith('\')) { $root = $root + '\' }
$minBytes       = [long]$MinSizeMB   * 1MB   # threshold for individual files
$minFolderBytes = [long]$MinFolderMB * 1MB   # threshold for folders

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

Write-Host ''
Write-Host '  DISK SPACE REPORT' -ForegroundColor White
Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
Write-Host "  Scanning $root ... " -NoNewline -ForegroundColor DarkGray

$sw = [Diagnostics.Stopwatch]::StartNew()
$walker = New-Object SpaceWalker
$walker.MinFileBytes = $minBytes
# Emits "##P <bytes> <files> <folder>" lines on stdout for SpaceReportApp to
# turn into a progress bar. Silent unless asked for.
$walker.Progress = [bool]$Progress
# Pass the trailing backslash: "C:" means "current directory on drive C:",
# whereas "C:\" is the root. Dropping it silently scans the wrong place.
$totalBytes = $walker.Walk($root)
$sw.Stop()

Write-Host ('{0:N0} files in {1:N0}s' -f $walker.FileCount, $sw.Elapsed.TotalSeconds) -ForegroundColor DarkGray
if (-not $isAdmin) {
    Write-Host '  Not elevated - System Volume Information and a few protected folders' -ForegroundColor DarkGray
    Write-Host '  will read as (locked). Re-run as admin to size them.' -ForegroundColor DarkGray
}

# --- drive figures ----------------------------------------------------------
$driveLetter = $root.Substring(0,2)
$vol = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
Write-Host ''
if ($vol) {
    $used = $vol.Size - $vol.FreeSpace
    $pct  = 0
    if ($vol.Size -gt 0) { $pct = [math]::Round($used / $vol.Size * 100) }
    $bar  = ('#' * [math]::Round($pct / 2.5)).PadRight(40, '.')
    $col  = 'Green'
    if     ($pct -ge 90) { $col = 'Red' }
    elseif ($pct -ge 75) { $col = 'Yellow' }
    Write-Host "  $driveLetter " -NoNewline -ForegroundColor White
    Write-Host $bar -NoNewline -ForegroundColor $col
    Write-Host ('  {0}% used   {1} free of {2}' -f $pct,
        (Format-Size $vol.FreeSpace).Trim(), (Format-Size $vol.Size).Trim())

    $unseen = $used - $totalBytes
    if ($unseen -gt 2GB -and $root -eq ($driveLetter + '\')) {
        Write-Host ('      {0} is in folders this scan could not read (protected system' -f (Format-Size $unseen).Trim()) -ForegroundColor DarkGray
        Write-Host  '      folders, or hard-linked bytes counted once here but twice by WinDirStat).' -ForegroundColor DarkGray
    }
    Write-Host ''
}

# ---------------------------------------------------------------------------
# Choose what to report
# ---------------------------------------------------------------------------
$rootDepth = ($root.TrimEnd('\') -split '\\').Count

$candidates = New-Object System.Collections.Generic.List[object]
foreach ($kv in $walker.DirSizes.GetEnumerator()) {
    $p = $kv.Key
    $sz = $kv.Value
    if ($sz -ge 0 -and $sz -lt $minFolderBytes) { continue }
    $depth = ($p.TrimEnd('\') -split '\\').Count - $rootDepth
    if ($depth -le 0) { continue }   # the scan root itself is not a finding
    $info  = Resolve-Entry -FullPath $p -IsDir $true
    # Report it if it is shallow enough to be a headline, or if we specifically
    # recognise it - naming the scary folders wherever they sit is the point.
    if ($depth -le 2 -or $info.Known) {
        $candidates.Add([pscustomobject]@{
            Path = $p; Bytes = $sz; Type = 'Folder'
            Name = $info.Name; Verdict = $info.Verdict; What = $info.What
            How = $info.How; Known = $info.Known; Mode = $info.Mode
        })
    }
}

# Suppress a child when an ancestor with the same classification is already
# listed, so 400 node_modules rows do not bury everything else.
# Sort by size, then shortest path first, so that when a folder and its only
# large child are the same size the parent wins and the child is suppressed.
$ordered = $candidates |
    Sort-Object -Property @{Expression='Bytes';Descending=$true},
                          @{Expression={$_.Path.Length};Descending=$false}
$kept = New-Object System.Collections.Generic.List[object]
foreach ($c in $ordered) {
    $covered = $false
    foreach ($anc in $kept) {
        if ($anc.Name -eq $c.Name -and
            $c.Path.StartsWith($anc.Path + '\', [StringComparison]::OrdinalIgnoreCase)) {
            $covered = $true; break
        }
    }
    if (-not $covered) { $kept.Add($c) }
}
$folders = @()
if ($TopFolders -gt 0) {
    $folders = @($kept | Sort-Object Bytes -Descending | Select-Object -First $TopFolders)
}

$fileList = New-Object System.Collections.Generic.List[object]
foreach ($f in $walker.BigFiles) {
    $info = Resolve-Entry -FullPath $f[0] -IsDir $false
    $fileList.Add([pscustomobject]@{
        Path = $f[0]; Bytes = [long]$f[1]; Type = 'File'
        Name = $info.Name; Verdict = $info.Verdict; What = $info.What
        How = $info.How; Known = $info.Known; Mode = $info.Mode
    })
}
# Individual files are the headline, so nothing is suppressed here - if two
# copies of the same 1GB DLL exist, seeing both listed is the point.
$allFiles = @($fileList | Sort-Object Bytes -Descending)
if ($Top -le 0) { $files = $allFiles }                              # 0 = list them all
else            { $files = @($allFiles | Select-Object -First $Top) }

$all = $files + $folders

# ---------------------------------------------------------------------------
# System cleanup areas - the things Disk Cleanup lists under "system files".
#
# These matter because they are invisible to a big-files report: Delivery
# Optimization or the update cache can be several GB spread over thousands of
# small files, none of which is individually large enough to appear.
#
# Sizes are read from the walk that already happened rather than re-measured,
# so this costs nothing. Anything outside the scanned root, or that the walk
# could not read, reports as unknown rather than zero.
# ---------------------------------------------------------------------------
$sysDrive = $env:SystemDrive
$SystemAreas = @(
  @{N='Windows Update cache'; P="$env:SystemRoot\SoftwareDistribution\Download"; V='TOOL';
    W='Update payloads already downloaded, and usually already installed.';
    H='Disk Cleanup > Windows Update Cleanup. Or: net stop wuauserv, empty the folder, net start wuauserv.'}
  @{N='Delivery Optimization files'; P="$env:SystemRoot\SoftwareDistribution\DeliveryOptimization"; V='SAFE';
    W='Update chunks cached so other PCs on your LAN can pull them from you rather than from Microsoft.';
    H='Disk Cleanup > Delivery Optimization Files. Or Settings > Windows Update > Advanced > Delivery Optimization.'}
  @{N='Delivery Optimization cache'; P="$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization"; V='SAFE';
    W='The other location Delivery Optimization uses. Same purpose, same disposability.';
    H='Disk Cleanup > Delivery Optimization Files.'}
  @{N='Previous Windows installation'; P="$sysDrive\Windows.old"; V='TOOL';
    W='Your entire previous Windows, kept so an upgrade can be rolled back. Auto-deletes about 10 days after the upgrade.';
    H='Settings > System > Storage > Temporary files > Previous Windows installation. Never delete by hand.'}
  @{N='Feature-update rollback data'; P="$sysDrive\`$WinREAgent"; V='TOOL';
    W='Working files a feature update left behind so it could roll back.';
    H='Disk Cleanup > Clean up system files. Safe once the update has proven good.'}
  @{N='Windows Setup staging'; P="$sysDrive\`$GetCurrent"; V='SAFE';
    W='Leftovers from a Windows setup or upgrade download.';
    H='Disk Cleanup > Temporary Windows installation files.'}
  @{N='Windows upgrade payload'; P="$sysDrive\`$Windows.~BT"; V='TOOL';
    W='Downloaded Windows installation payload, often several GB.';
    H='Disk Cleanup > Clean up system files > Temporary Windows installation files.'}
  @{N='Error reporting queue'; P="$env:ProgramData\Microsoft\Windows\WER"; V='SAFE';
    W='Queued application crash reports waiting to be sent to Microsoft.';
    H='Disk Cleanup > Windows error reports.'}
  @{N='System temp folder'; P="$env:SystemRoot\Temp"; V='SAFE';
    W='Scratch files from installers and services. Most of it is stale.';
    H='Delete what is not locked; in-use files will refuse, which is fine.'}
  @{N='Your temp folder'; P=$env:TEMP; V='SAFE';
    W='Per-user scratch space. Installers and half your software dump here and never clean up.';
    H='Disk Cleanup > Temporary files, or empty it directly.'}
  @{N='Kernel telemetry dumps'; P="$env:SystemRoot\LiveKernelReports"; V='SAFE';
    W='Partial dumps taken when a driver hiccuped without bluescreening. Several GB on a machine with a flaky GPU driver.';
    H='Delete the .dmp files inside.'}
  @{N='Bluescreen minidumps'; P="$env:SystemRoot\Minidump"; V='SAFE';
    W='Small crash dumps, one per bluescreen.';
    H='Delete, unless you are currently diagnosing crashes.'}
  @{N='Windows Search index'; P="$env:ProgramData\Microsoft\Search\Data"; V='TOOL';
    W='Windows.edb, the full-text index of your files. Normally 1-10GB and occasionally runs away.';
    H='Control Panel > Indexing Options > Advanced > Rebuild. Narrow the indexed locations first.'}
  @{N='Defender scan history'; P="$env:ProgramData\Microsoft\Windows Defender\Scans"; V='TOOL';
    W='Antivirus scan history and accumulated signature versions.';
    H='"%ProgramFiles%\Windows Defender\MpCmdRun.exe" -RemoveDefinitions -DynamicSignatures'}
  @{N='App launch prefetch'; P="$env:SystemRoot\Prefetch"; V='SAFE';
    W='Small hints Windows uses to speed up application launches.';
    H='Deletable but rarely worth it - tens of MB, and it costs you launch speed while it rebuilds.'}
)

$sysResults = New-Object System.Collections.Generic.List[object]
foreach ($a in $SystemAreas) {
    if (-not $a.P) { continue }
    $known = $walker.DirSizes.ContainsKey($a.P.TrimEnd('\'))
    $bytes = -1
    if ($known) {
        $bytes = $walker.DirSizes[$a.P.TrimEnd('\')]
    } else {
        # Test-Path throws rather than returning false on an ACL-protected path,
        # and "denied" means it is there - we just cannot measure it.
        $exists = $false
        try   { $exists = Test-Path -LiteralPath $a.P -ErrorAction Stop }
        catch { $exists = $true }
        if (-not $exists) { continue }        # genuinely not on this machine
    }
    $sysResults.Add([pscustomobject]@{
        Name = $a.N; Path = $a.P; Bytes = $bytes; Verdict = $a.V
        What = $a.W; How = $a.H
        # Always 'cmd': these are folders with a proper reclaim route. The tool
        # will not delete them file by file.
        Mode = 'cmd'
    })
}
$sysResults = @($sysResults | Sort-Object { if ($_.Bytes -lt 0) { -1 } else { $_.Bytes } } -Descending)

# ---------------------------------------------------------------------------
# Optional: describe unknowns using local Ollama (no Claude tokens)
# ---------------------------------------------------------------------------
if ($Ollama) {
    $unknown = @($all | Where-Object { -not $_.Known } | Select-Object -First 30)
    if ($unknown.Count -gt 0) {
        Write-Host "  Asking local Ollama about $($unknown.Count) unrecognised paths..." -ForegroundColor DarkGray
        $tmp = Join-Path $env:TEMP "spacereport_unknown_$PID.txt"
        ($unknown | ForEach-Object { '{0}  [{1}]' -f $_.Path, (Format-Size $_.Bytes).Trim() }) |
            Set-Content -LiteralPath $tmp -Encoding UTF8
        $q = 'Each line is a path on a Windows 11 developer PC followed by its size. ' +
             'For each line output exactly one row in this format: ' +
             'PATH|VERDICT|SHORT_NAME|one sentence on what it is and whether deleting it is safe. ' +
             'VERDICT must be one of SAFE, TOOL, REVIEW, KEEP, NEVER. ' +
             'No preamble, no markdown, one row per input line.'
        $delegate = Join-Path $PSScriptRoot 'delegate.py'
        $resp = @()
        # A native command's stderr under $ErrorActionPreference='Stop' becomes a
        # terminating error, which silently killed this step. Relax it here only.
        $prevEAP = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            if (-not (Test-Path -LiteralPath $delegate)) {
                Write-Host "  delegate.py not found at $delegate - skipping." -ForegroundColor DarkYellow
            } else {
                $resp = @(& python $delegate summarize $tmp $q)
            }
        } catch {
            Write-Host "  Ollama step failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
        $ErrorActionPreference = $prevEAP
        $global:LASTEXITCODE = 0

        $applied = 0
        foreach ($line in $resp) {
            $parts = "$line" -split '\|', 4
            if ($parts.Count -lt 4) { continue }
            $p = $parts[0].Trim().Trim('"', "'", '`')
            $target = $all | Where-Object { $_.Path -eq $p } | Select-Object -First 1
            if ($target) {
                $v = $parts[1].Trim().ToUpper()
                # Never let a model guess talk you into deleting something.
                # Clamp its most permissive verdicts up to REVIEW.
                if ($v -eq 'SAFE' -or $v -eq 'TOOL') { $v = 'REVIEW' }
                if ($VerdictColour.ContainsKey($v)) { $target.Verdict = $v }
                $target.Name = $parts[2].Trim()
                $target.What = $parts[3].Trim()
                $target.How  = 'This description came from local Ollama, not the curated list. Verify before deleting anything.'
                $applied++
            }
        }
        if ($applied -eq 0) {
            Write-Host '  Ollama returned nothing usable - entries left as UNKNOWN.' -ForegroundColor DarkYellow
        } else {
            Write-Host "  Described $applied of $($unknown.Count)." -ForegroundColor DarkGray
        }
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------
function Write-Section {
    param([string]$Title, $Items)
    if (-not $Items -or $Items.Count -eq 0) { return }
    Write-Host ''
    Write-Host "  $Title" -ForegroundColor White
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
    foreach ($i in $Items) {
        $short = $i.Path
        if ($short.Length -gt 60) { $short = '...' + $short.Substring($short.Length - 57) }
        Write-Host ('  {0} ' -f (Format-Size $i.Bytes)) -NoNewline
        Write-Host ('{0,-7}' -f $i.Verdict) -NoNewline -ForegroundColor $VerdictColour[$i.Verdict]
        Write-Host " $short"
        Write-Host ('             {0}' -f $i.Name) -ForegroundColor Gray
        foreach ($w in (Format-Wrap $i.What)) {
            Write-Host ('             {0}' -f $w) -ForegroundColor DarkGray
        }
        if ($i.How) {
            $first = $true
            foreach ($w in (Format-Wrap $i.How 61)) {
                if ($first) { Write-Host ('             -> {0}' -f $w) -ForegroundColor DarkCyan; $first = $false }
                else        { Write-Host ('                {0}' -f $w) -ForegroundColor DarkCyan }
            }
        }
        Write-Host ''
    }
}

Write-Section ('BIGGEST INDIVIDUAL FILES  (over {0} MB)' -f $MinSizeMB) $files

# --- rollup across every big file found, not just the ones listed above -----
if ($allFiles.Count -gt 0) {
    Write-Host ''
    Write-Host ('  BIG FILES BY TYPE  ({0:N0} files over {1} MB, {2} in total)' -f
        $allFiles.Count, $MinSizeMB,
        (Format-Size (($allFiles | Measure-Object Bytes -Sum).Sum)).Trim()) -ForegroundColor White
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
    $groups = $allFiles | Group-Object Name |
        Select-Object Name, Count, @{N='Bytes';E={($_.Group | Measure-Object Bytes -Sum).Sum}},
                      @{N='Verdict';E={$_.Group[0].Verdict}} |
        Sort-Object Bytes -Descending
    foreach ($g in $groups) {
        Write-Host ('  {0} ' -f (Format-Size $g.Bytes)) -NoNewline
        Write-Host ('{0,-7}' -f $g.Verdict) -NoNewline -ForegroundColor $VerdictColour[$g.Verdict]
        Write-Host (' {0,4} x  {1}' -f $g.Count, $g.Name)
    }
    Write-Host ''
}

if ($sysResults.Count -gt 0) {
    $sysKnown = @($sysResults | Where-Object { $_.Bytes -gt 0 })
    $sysTotal = 0L
    if ($sysKnown.Count -gt 0) { $sysTotal = ($sysKnown | Measure-Object Bytes -Sum).Sum }
    Write-Host ''
    Write-Host ('  SYSTEM CLEANUP AREAS  ({0} measurable)' -f (Format-Size $sysTotal).Trim()) -ForegroundColor White
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
    Write-Host '  Disk Cleanup territory - thousands of small files, so none of these show' -ForegroundColor DarkGray
    Write-Host '  up in the list above however low you set the threshold.' -ForegroundColor DarkGray
    Write-Host ''
    foreach ($s in $sysResults) {
        Write-Host ('  {0} ' -f (Format-Size $s.Bytes)) -NoNewline
        Write-Host ('{0,-7}' -f $s.Verdict) -NoNewline -ForegroundColor $VerdictColour[$s.Verdict]
        Write-Host (' {0}' -f $s.Name)
        if ($s.Bytes -lt 0) {
            Write-Host '              could not be read - re-run elevated to size it' -ForegroundColor DarkGray
        }
        foreach ($w in (Format-Wrap $s.How 61)) {
            Write-Host ('              -> {0}' -f $w) -ForegroundColor DarkCyan
        }
    }
    Write-Host ''
}

if ($TopFolders -gt 0) {
    Write-Section ('BIGGEST FOLDERS  (over {0} MB, for context)' -f $MinFolderMB) $folders
}

# Totals are files only - every big file found, not just the ones printed
# above. Folders are excluded deliberately: they are context, and adding a
# folder to its own contents double-counts every byte.
$topLevel = $allFiles

Write-Host ''
Write-Host '  WHAT YOU COULD ACTUALLY RECLAIM' -ForegroundColor White
Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray
Write-Host ('  (every file over {0} MB; the folders above are context only)' -f $MinSizeMB) -ForegroundColor DarkGray
foreach ($v in $VerdictLegend.Keys) {
    $grp = @($topLevel | Where-Object { $_.Verdict -eq $v })
    if ($grp.Count -eq 0) { continue }
    $sum = ($grp | Measure-Object Bytes -Sum).Sum
    Write-Host ('  {0} {1,-7} {2,3} item(s)   ' -f (Format-Size $sum), $v, $grp.Count) -NoNewline
    Write-Host $VerdictLegend[$v] -ForegroundColor $VerdictColour[$v]
}
$lowRisk = @($topLevel | Where-Object { $_.Verdict -eq 'SAFE' -or $_.Verdict -eq 'TOOL' })
$reclaim = 0
if ($lowRisk.Count -gt 0) { $reclaim = ($lowRisk | Measure-Object Bytes -Sum).Sum }
Write-Host ''
Write-Host ('  Low-risk total (SAFE + TOOL): {0}' -f (Format-Size $reclaim).Trim()) -ForegroundColor Green
Write-Host ''
if ($Clean) {
    Write-Host '  Nothing has been deleted yet - the picker comes next.' -ForegroundColor DarkGray
} else {
    Write-Host '  Nothing was deleted. Re-run with -Clean to pick files to remove.' -ForegroundColor DarkGray
}
Write-Host ''

# ---------------------------------------------------------------------------
# Optional outputs
# ---------------------------------------------------------------------------
if ($Json) {
    # Consumed by SpaceReportApp. Emits every big file (not just the printed
    # top N) and carries Mode, so the app can show padlocks without having to
    # re-implement any of the classification.
    $dSize = 0L; $dFree = 0L
    if ($vol) { $dSize = [long]$vol.Size; $dFree = [long]$vol.FreeSpace }
    $payload = [ordered]@{
        Root      = $root
        Drive     = $driveLetter
        ScannedAt = (Get-Date).ToString('s')
        FileCount = [long]$walker.FileCount
        MinSizeMB = $MinSizeMB
        DriveSize = $dSize
        DriveFree = $dFree
        Files     = @($allFiles  | Select-Object Path,Bytes,Name,Verdict,What,How,Mode)
        Folders   = @($folders   | Select-Object Path,Bytes,Name,Verdict,What,How,Mode)
        System    = @($sysResults | Select-Object Path,Bytes,Name,Verdict,What,How,Mode)
    }
    $payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Json -Encoding UTF8
    Write-Host "  JSON written to $Json" -ForegroundColor DarkGray
}

if ($Html -and $all.Count -eq 0) {
    Write-Host '  Nothing met the size threshold, so no HTML report was written.' -ForegroundColor DarkYellow
}
elseif ($Html) {
    Add-Type -AssemblyName System.Web
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $out   = Join-Path $env:TEMP ('space-report-{0}-{1}.html' -f $driveLetter.TrimEnd(':'), $stamp)
    $max = ($all | Measure-Object Bytes -Maximum).Maximum
    $actionLabelHtml = @{ 'file' = 'DELETE FILE'; 'app' = 'use app'; 'cmd' = 'use command'; 'never' = 'BLOCKED' }
    function Get-Rows {
        param($Items, [switch]$Tickable)
        foreach ($i in ($Items | Sort-Object Bytes -Descending)) {
            $pctw = 1
            if ($max -gt 0) { $pctw = [math]::Max(1, [math]::Round($i.Bytes / $max * 100)) }
            $ePath = [System.Web.HttpUtility]::HtmlEncode($i.Path)
            $eName = [System.Web.HttpUtility]::HtmlEncode($i.Name)
            $eWhat = [System.Web.HttpUtility]::HtmlEncode($i.What)
            $eHow  = [System.Web.HttpUtility]::HtmlEncode($i.How)
            $tick  = '<td class="ck"></td>'
            if ($Tickable) {
                if ($i.Mode -eq 'file') {
                    # data-p carries the raw path; the page never deletes, it
                    # only assembles a command for you to run.
                    $attr = [System.Web.HttpUtility]::HtmlAttributeEncode($i.Path)
                    $tick = "<td class='ck'><input type='checkbox' class='pick' data-p=""$attr"" data-b='$($i.Bytes)'></td>"
                } else {
                    $tick = "<td class='ck'><span class='lock' title='$($actionLabelHtml[$i.Mode]) - cannot be deleted as a plain file'>&#128274;</span></td>"
                }
            }
            "<tr class='v-$($i.Verdict)'>$tick<td class='sz'>$((Format-Size $i.Bytes).Trim())</td>" +
            "<td class='vd'><span>$($i.Verdict)</span></td>" +
            "<td class='pc'><div class='bar' style='width:$pctw%'></div>" +
            "<code>$ePath</code><div class='nm'>$eName</div>" +
            "<div class='wh'>$eWhat</div><div class='how'>$eHow</div></td></tr>"
        }
    }

    # Roll every big file up by classification, so the aggregate is visible
    # even when no single file is remarkable.
    $typeRows = foreach ($g in ($allFiles | Group-Object Name |
            Select-Object Name, Count,
                @{N='Bytes';E={($_.Group | Measure-Object Bytes -Sum).Sum}},
                @{N='Verdict';E={$_.Group[0].Verdict}} |
            Sort-Object Bytes -Descending)) {
        "<tr class='v-$($g.Verdict)'><td class='ck'></td><td class='sz'>$((Format-Size $g.Bytes).Trim())</td>" +
        "<td class='vd'><span>$($g.Verdict)</span></td>" +
        "<td class='pc'><code>$($g.Count) file(s)</code>" +
        "<div class='nm'>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</div></td></tr>"
    }
    $legend = foreach ($v in $VerdictLegend.Keys) {
        "<li class='v-$v'><span>$v</span> $($VerdictLegend[$v])</li>"
    }
    $head = @"
<!doctype html><html><head><meta charset="utf-8">
<title>Disk space report - $driveLetter</title><style>
:root{--bg:#12141a;--fg:#e6e8ee;--dim:#9aa1b0;--line:#262a35}
body{margin:0;padding:2rem;background:var(--bg);color:var(--fg);font:14px/1.5 "Segoe UI",system-ui,sans-serif}
h1{font-size:1.4rem;margin:0 0 .2rem}h2{font-size:.95rem;color:var(--dim);font-weight:400;margin:0 0 1.4rem}
h3{font-size:1.05rem;margin:2.2rem 0 .6rem;padding-bottom:.4rem;border-bottom:1px solid var(--line);max-width:1100px}
table{border-collapse:collapse;width:100%;max-width:1100px}
td{border-bottom:1px solid var(--line);padding:.7rem .6rem;vertical-align:top}
.sz{white-space:nowrap;font-variant-numeric:tabular-nums;font-weight:600;width:6rem}
.vd{width:5.5rem}
.vd span{display:inline-block;padding:.1rem .45rem;border-radius:3px;font-size:11px;font-weight:700;letter-spacing:.04em}
.pc{position:relative}
.bar{position:absolute;left:0;bottom:-.7rem;height:3px;opacity:.9;border-radius:2px}
code{font:12px ui-monospace,Consolas,monospace;position:relative}
.nm{font-weight:600;margin-top:.25rem;position:relative}
.wh{color:var(--dim);margin-top:.15rem;position:relative;max-width:72ch}
.how{color:#6fd3e0;margin-top:.3rem;position:relative;max-width:72ch;font-size:13px}
ul{list-style:none;padding:0;max-width:1100px;margin:0 0 2rem}
li{padding:.2rem 0;color:var(--dim)}
li span{display:inline-block;width:5.2rem;font-weight:700;font-size:11px;padding:.1rem .45rem;border-radius:3px;text-align:center}
.v-SAFE span{background:#1f7a3d;color:#fff}.v-TOOL span{background:#1a6b86;color:#fff}
.v-REVIEW span{background:#8a6a12;color:#fff}.v-KEEP span{background:#4a4a52;color:#fff}
.v-NEVER span{background:#8c2020;color:#fff}.v-UNKNOWN span{background:#6a2c86;color:#fff}
.v-SAFE .bar{background:#2ecc71}.v-TOOL .bar{background:#3ab7d8}.v-REVIEW .bar{background:#e0a92e}
.v-KEEP .bar{background:#8a8a95}.v-NEVER .bar{background:#e05252}.v-UNKNOWN .bar{background:#b06ad8}
.foot{color:var(--dim);margin-top:2rem;font-size:13px}
.hint{color:var(--dim);max-width:80ch;margin:.2rem 0 1rem;font-size:13px}
.ck{width:2.2rem;text-align:center}
.ck input{width:16px;height:16px;cursor:pointer;accent-color:#2ecc71}
.lock{opacity:.45;cursor:help;font-size:13px}
#bar{position:sticky;bottom:0;margin-top:2rem;background:#0d0f14;border-top:2px solid #2ecc71;
     padding:1rem;max-width:1100px;box-shadow:0 -8px 24px rgba(0,0,0,.5)}
#bar.empty{border-top-color:var(--line)}
#sum{font-weight:600;margin-bottom:.5rem}
#cmd{width:100%;min-height:7rem;background:#080a0e;color:#cfe;border:1px solid var(--line);
     border-radius:4px;padding:.6rem;font:12px ui-monospace,Consolas,monospace;resize:vertical}
#bar button{background:#1f7a3d;color:#fff;border:0;border-radius:4px;padding:.5rem 1rem;
     font-weight:600;cursor:pointer;margin-right:.5rem}
#bar button.sec{background:#333842}
#note{color:var(--dim);font-size:12px;margin-top:.5rem}
</style></head><body>
<h1>Disk space report &mdash; $driveLetter</h1>
<h2>$([System.Web.HttpUtility]::HtmlEncode($root)) &middot; $(Get-Date -Format 'ddd d MMM yyyy HH:mm') &middot; $('{0:N0}' -f $walker.FileCount) files scanned</h2>
<ul>$($legend -join '')</ul>
"@
    # NB: do not name this $html - PowerShell variables are case-insensitive,
    # so it would overwrite the -Html switch parameter and fail to bind.
    $page = $head +
        "<h3>Biggest individual files (over $MinSizeMB MB)</h3>" +
        "<p class='hint'>Tick the files you want gone, then copy the command at the bottom and run it in PowerShell. " +
        "A padlock means the file cannot be deleted on its own &mdash; hover it for the reason. " +
        "This page never deletes anything itself; it only writes the command.</p><table>" +
        ((Get-Rows $files -Tickable) -join '') + '</table>' +
        "<h3>Big files by type &mdash; $('{0:N0}' -f $allFiles.Count) files, $((Format-Size (($allFiles | Measure-Object Bytes -Sum).Sum)).Trim()) in total</h3><table>" +
        ($typeRows -join '') + '</table>'
    if ($folders.Count -gt 0) {
        $page += "<h3>Biggest folders (over $MinFolderMB MB, for context)</h3><table>" +
                 ((Get-Rows $folders) -join '') + '</table>'
    }
    $scriptSelf = $PSCommandPath
    if (-not $scriptSelf -and $PSScriptRoot) {
        $scriptSelf = Join-Path $PSScriptRoot 'Get-SpaceReport.ps1'
    }
    if (-not $scriptSelf) { $scriptSelf = '.\Get-SpaceReport.ps1' }
    # JSON-encode: a Windows path dropped raw into a JS string literal has its
    # backslashes eaten as escapes ("\tools" becomes a TAB + "ools").
    # ConvertTo-Json emits the surrounding quotes too, hence no quotes below.
    $selfJs = $scriptSelf | ConvertTo-Json -Compress

    $page += @"
<div id="bar" class="empty">
  <div id="sum">No files ticked</div>
  <textarea id="cmd" spellcheck="false" readonly>Tick one or more files above to build a command.</textarea>
  <div style="margin-top:.6rem">
    <button onclick="copyCmd()">Copy command</button>
    <button class="sec" onclick="clearAll()">Clear selection</button>
    <label style="color:var(--dim);font-size:12px;margin-left:.5rem">
      <input type="checkbox" id="perm" onchange="rebuild()"> delete permanently (skip Recycle Bin)
    </label>
  </div>
  <div id="note">Files go to the Recycle Bin unless you tick permanent. The script re-checks
    every file against its own safety rules before removing anything, and asks you to type
    DELETE to confirm.</div>
</div>
<script>
var SELF = $selfJs;
function fmt(b){var u=['B','KB','MB','GB','TB'],i=0;while(b>=1024&&i<4){b/=1024;i++}return b.toFixed(1)+' '+u[i];}
function picked(){return Array.prototype.slice.call(document.querySelectorAll('.pick:checked'));}
function rebuild(){
  var sel=picked(), bar=document.getElementById('bar'), cmd=document.getElementById('cmd');
  var total=sel.reduce(function(a,c){return a+parseInt(c.dataset.b,10)},0);
  if(!sel.length){
    bar.className='empty';
    document.getElementById('sum').textContent='No files ticked';
    cmd.value='Tick one or more files above to build a command.';
    return;
  }
  bar.className='';
  document.getElementById('sum').textContent=sel.length+' file(s) ticked  -  '+fmt(total);
  var quoted=sel.map(function(c){return "  '"+c.dataset.p.replace(/'/g,"''")+"'"}).join(',\n');
  var perm=document.getElementById('perm').checked?' -Permanent':'';
  cmd.value="& '"+SELF+"'"+perm+" -Paths @(\n"+quoted+"\n)";
}
function copyCmd(){
  var t=document.getElementById('cmd');
  t.removeAttribute('readonly'); t.select(); t.setSelectionRange(0,999999);
  var done=function(){t.setAttribute('readonly','');};
  if(navigator.clipboard){navigator.clipboard.writeText(t.value).then(done,function(){document.execCommand('copy');done();});}
  else{document.execCommand('copy');done();}
}
function clearAll(){picked().forEach(function(c){c.checked=false});rebuild();}
document.addEventListener('change',function(e){if(e.target.classList.contains('pick'))rebuild()});
rebuild();
</script>
<p class='foot'>Generated by Get-SpaceReport.ps1. Nothing has been deleted &mdash; this page only builds a command.</p></body></html>
"@
    Set-Content -LiteralPath $out -Value $page -Encoding UTF8
    Write-Host "  HTML report: $out" -ForegroundColor DarkGray
    Start-Process $out
}

}   # <- end of the "if (-not $Paths)" scan-and-report block

# ---------------------------------------------------------------------------
# -Clean : tick-box picker, then Recycle Bin (or -Permanent).
#
# Rails, in order of application:
#   1. Folders are never offered. Files only.
#   2. Verdict NEVER is never offered, with or without -IncludeRisky.
#   3. Only Mode 'file' can actually be deleted. 'app' and 'cmd' are shown so
#      you can see why they are held back, but are stripped from the selection.
#   4. Recycle Bin by default, so a mistake is recoverable.
#   5. A typed confirmation, unless -Force.
# ---------------------------------------------------------------------------
# Emitted for SpaceReportApp when -Json is used with a clean-up run, so the
# app can report exactly what happened - including what the rails refused.
function Out-CleanJson {
    param([string]$Outcome, $Removed, $Failed, $Refused, $Skipped, [long]$Freed, [bool]$Recycled)
    if (-not $Json) { return }
    $payload = [ordered]@{
        Outcome  = $Outcome
        Recycled = $Recycled
        Freed    = $Freed
        Removed  = @($Removed | Select-Object Path,Bytes,Name,Verdict)
        Failed   = @($Failed  | Select-Object Path,Reason)
        Refused  = @($Refused | Select-Object Path,Name,Verdict)
        Skipped  = @($Skipped | Select-Object Path,Name,Verdict,Mode)
    }
    $payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Json -Encoding UTF8
}

function Invoke-CleanUp {
    param($Items, [switch]$AlreadyChosen)
    $refusedAll = @()
    $skippedAll = @()

    Write-Host ''
    Write-Host '  CLEAN UP' -ForegroundColor White
    Write-Host ('  ' + ('-' * 76)) -ForegroundColor DarkGray

    $offer = @($Items | Where-Object {
        $_.Verdict -ne 'NEVER' -and ($IncludeRisky -or $_.Verdict -ne 'KEEP')
    })
    # When the caller named paths explicitly, say out loud what was refused
    # before we got as far as the picker - otherwise a hand-edited command
    # looks like it worked while quietly doing less than asked.
    if ($AlreadyChosen) {
        $dropped = @($Items | Where-Object {
            $_.Verdict -eq 'NEVER' -or (-not $IncludeRisky -and $_.Verdict -eq 'KEEP')
        })
        foreach ($d in $dropped) {
            Write-Host ('  REFUSED  {0}' -f $d.Path) -ForegroundColor Red
            Write-Host ('           {0} ({1})' -f $d.Name, $d.Verdict) -ForegroundColor Red
        }
        $refusedAll = $dropped
        if ($dropped.Count -gt 0) {
            $nDropped = [int]$dropped.Count
            # NB: @($Items) here trips a PowerShell binder bug when $Items is a
            # generic List ("Argument types do not match"). .Count directly is fine.
            $nSupplied = [int]$Items.Count
            Write-Host ("  {0} of {1} supplied path(s) refused outright." -f $nDropped, $nSupplied) -ForegroundColor Red
            Write-Host ''
        }
    }

    if ($offer.Count -eq 0) {
        Write-Host '  Nothing eligible to offer.' -ForegroundColor DarkYellow
        Out-CleanJson 'nothing-eligible' @() @() $refusedAll @() 0 (-not $Permanent)
        return
    }

    $actionLabel = @{ 'file' = 'DELETE FILE'; 'app' = 'use app'; 'cmd' = 'use command'; 'never' = 'BLOCKED' }
    $grid = $offer | ForEach-Object {
        [pscustomobject]@{
            Size    = (Format-Size $_.Bytes).Trim()
            Action  = $actionLabel[$_.Mode]
            Verdict = $_.Verdict
            What    = $_.Name
            Path    = $_.Path
            Guidance= $_.How
            Bytes   = $_.Bytes
            Mode    = $_.Mode
        }
    }

    if ($AlreadyChosen) {
        # Paths came from -Paths (e.g. ticked on the HTML report). They are
        # still put through every rail below - the caller chooses WHICH files
        # to consider, never WHETHER a file may be deleted.
        $picked = @($grid)
        Write-Host ("  {0} path(s) supplied via -Paths." -f $picked.Count) -ForegroundColor DarkGray
    } elseif ($Select) {
        # Non-interactive selection by path wildcard - skips the picker.
        $picked = @($grid | Where-Object { $_.Path -like $Select })
        Write-Host ("  -Select '{0}' matched {1} file(s)." -f $Select, $picked.Count) -ForegroundColor DarkGray
    } else {
        $title  = 'Tick files to delete  -  only rows marked DELETE FILE will be acted on'
        $picked = @($grid | Out-GridView -Title $title -PassThru)
    }
    if ($picked.Count -eq 0) {
        Write-Host '  Nothing selected. No changes made.' -ForegroundColor DarkGray
        Out-CleanJson 'nothing-selected' @() @() $refusedAll @() 0 (-not $Permanent)
        return
    }

    $held = @($picked | Where-Object { $_.Mode -ne 'file' })
    $togo = @($picked | Where-Object { $_.Mode -eq 'file' })

    foreach ($h in $held) {
        Write-Host ('  SKIPPED  {0}' -f $h.Path) -ForegroundColor DarkYellow
        Write-Host ('           {0} - {1}' -f $h.What, $actionLabel[$h.Mode]) -ForegroundColor DarkYellow
        # If the safety list caught it by path, say so - the classification
        # guidance may be about the folder, not about this file type.
        if ($h.Path -match $NeverPlainDelete) {
            Write-Host '           -> This file type is on the never-delete list regardless of' -ForegroundColor DarkCyan
            Write-Host '              where it sits. It holds data that cannot be regenerated,' -ForegroundColor DarkCyan
            Write-Host '              or must be shrunk/compacted rather than removed.' -ForegroundColor DarkCyan
        } elseif ($h.Path -match $CmdByPath) {
            Write-Host '           -> Shrink this through its own application rather than' -ForegroundColor DarkCyan
            Write-Host '              deleting the file.' -ForegroundColor DarkCyan
        } else {
            foreach ($w in (Format-Wrap $h.Guidance 61)) {
                Write-Host ('           -> {0}' -f $w) -ForegroundColor DarkCyan
            }
        }
    }
    if ($held.Count -gt 0) { Write-Host '' }

    # Grid rows call the classification "What"; normalise to Name for the JSON.
    $skippedAll = @($held | ForEach-Object {
        [pscustomobject]@{ Path = $_.Path; Name = $_.What; Verdict = $_.Verdict; Mode = $_.Mode }
    })
    if ($togo.Count -eq 0) {
        Write-Host '  Nothing in the selection can be deleted as a plain file. No changes made.' -ForegroundColor DarkYellow
        Out-CleanJson 'nothing-deletable' @() @() $refusedAll $skippedAll 0 (-not $Permanent)
        return
    }

    $sum  = ($togo | Measure-Object Bytes -Sum).Sum
    $verb = if ($Permanent) { 'PERMANENTLY DELETE' } else { 'send to the Recycle Bin' }
    Write-Host ("  About to $verb {0} file(s), {1}:" -f $togo.Count, (Format-Size $sum).Trim()) -ForegroundColor White
    foreach ($t in ($togo | Sort-Object Bytes -Descending)) {
        Write-Host ('    {0}  {1}' -f (Format-Size $t.Bytes), $t.Path)
    }
    Write-Host ''

    if (-not $Force) {
        if ($Permanent) {
            Write-Host '  This CANNOT be undone.' -ForegroundColor Red
        } else {
            Write-Host '  Recoverable from the Recycle Bin afterwards.' -ForegroundColor DarkGray
        }
        $answer = Read-Host '  Type DELETE to proceed, anything else to cancel'
        if ($answer -cne 'DELETE') {
            Write-Host '  Cancelled. No changes made.' -ForegroundColor DarkGray
            Out-CleanJson 'cancelled' @() @() $refusedAll $skippedAll 0 (-not $Permanent)
            return
        }
    }

    if (-not $Permanent) { Add-Type -AssemblyName Microsoft.VisualBasic }
    $freed = 0L; $ok = 0; $failed = 0
    $removedList = New-Object System.Collections.Generic.List[object]
    $failedList  = New-Object System.Collections.Generic.List[object]
    foreach ($t in $togo) {
        try {
            if (-not (Test-Path -LiteralPath $t.Path -PathType Leaf)) {
                throw 'no longer exists'
            }
            if ($Permanent) {
                Remove-Item -LiteralPath $t.Path -Force -ErrorAction Stop
            } else {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $t.Path,
                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)
            }
            $freed += $t.Bytes; $ok++
            $removedList.Add([pscustomobject]@{
                Path = $t.Path; Bytes = $t.Bytes; Name = $t.What; Verdict = $t.Verdict })
            Write-Host ('  removed  {0}' -f $t.Path) -ForegroundColor Green
        } catch {
            $failed++
            $failedList.Add([pscustomobject]@{ Path = $t.Path; Reason = $_.Exception.Message })
            Write-Host ('  FAILED   {0}' -f $t.Path) -ForegroundColor Red
            Write-Host ('           {0}' -f $_.Exception.Message) -ForegroundColor Red
        }
    }
    Out-CleanJson 'done' $removedList $failedList $refusedAll $skippedAll $freed (-not $Permanent)

    Write-Host ''
    Write-Host ('  {0} removed, {1} failed, {2}.' -f $ok, $failed, (Format-Size $freed).Trim()) -ForegroundColor White
    if (-not $Permanent -and $ok -gt 0) {
        Write-Host '  NOTE: recycled files still occupy the drive. The space is not actually' -ForegroundColor DarkYellow
        Write-Host '  free until you empty the Recycle Bin (check the contents there first).' -ForegroundColor DarkYellow
    }
    Write-Host ''
}

# Build items from an explicit path list (the HTML report's tick boxes emit
# one of these commands). Classification and every rail are applied here just
# as they are for a scanned file - the list only says WHICH files to consider.
function Get-ItemsForPaths {
    param([string[]]$List)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($p in $List) {
        $p = $p.Trim()
        if ($p -eq '') { continue }
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
            Write-Host "  not found (skipped): $p" -ForegroundColor DarkYellow
            continue
        }
        $len = 0L
        try { $len = (Get-Item -LiteralPath $p -Force).Length } catch { }
        $info = Resolve-Entry -FullPath $p -IsDir $false
        $out.Add([pscustomobject]@{
            Path = $p; Bytes = $len; Type = 'File'
            Name = $info.Name; Verdict = $info.Verdict; What = $info.What
            How = $info.How; Known = $info.Known; Mode = $info.Mode
        })
    }
    # Return a plain array, not the List - see the binder note in Invoke-CleanUp.
    ,$out.ToArray()
}

if ($Paths)      { Invoke-CleanUp (Get-ItemsForPaths $Paths) -AlreadyChosen }
elseif ($Clean)  { Invoke-CleanUp $allFiles }
