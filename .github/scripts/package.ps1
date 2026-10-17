param([Parameter(Mandatory=$true)][ValidateSet('x64','x86')][string]$Platform,[Parameter(Mandatory=$true)][string]$RefName)
$ErrorActionPreference='Stop'
[xml]$project=Get-Content SleepGuard/SleepGuard.csproj
$v=[string]$project.Project.PropertyGroup[0].Version
if($RefName.StartsWith('v') -and $RefName -cne "v$v"){throw 'Tag and product version differ'}
$bin="SleepGuard/bin/$Platform/Release/net8.0-windows/win-$Platform"
foreach($n in @('SleepGuard.exe','SleepGuard.dll','SleepGuard.deps.json','SleepGuard.runtimeconfig.json','Hardcodet.NotifyIcon.Wpf.dll')){if(!(Test-Path "$bin/$n")){throw "Missing dependency $n"}}
if([Diagnostics.FileVersionInfo]::GetVersionInfo("$bin/SleepGuard.dll").FileVersion -ne "$v.0"){throw 'Version mismatch'}
if((Get-FileHash README.md).Hash -ne (Get-FileHash README.txt).Hash){throw 'README mismatch'}
$out="package/SleepGuard_v${v}_$Platform"
if(Test-Path $out){throw 'Staging exists'}
New-Item -ItemType Directory -Force "$out/licenses"|Out-Null
Get-ChildItem $bin -Recurse -File|Where-Object {$_.Extension -notin @('.pdb','.log','.tmp','.bak')}|ForEach-Object{$rel=[IO.Path]::GetRelativePath((Resolve-Path $bin),$_.FullName);if($rel -match '^runtimes[\\/]([^\\/]+)' -and $Matches[1] -notin @("win-$Platform",'win')){return};$dest=Join-Path $out $rel;New-Item -ItemType Directory -Force (Split-Path $dest)|Out-Null;Copy-Item -LiteralPath $_.FullName -Destination $dest}
foreach($n in @('README.md','README.txt','LICENSE','THIRD_PARTY_NOTICES.txt')){Copy-Item $n $out}
Copy-Item licenses/* "$out/licenses"
$assets=Get-Content SleepGuard/obj/project.assets.json -Raw|ConvertFrom-Json
foreach($lib in $assets.libraries.PSObject.Properties){
$folder=$null
foreach($base in $assets.packageFolders.PSObject.Properties.Name){$candidate=Join-Path $base $lib.Value.path;if(Test-Path $candidate){$folder=$candidate;break}}
if(!$folder){throw "Package not found $($lib.Name)"}
$notices=@(Get-ChildItem $folder -File|Where-Object {$_.Name -match '^(LICENSE|THIRD.PARTY.NOTICES)(\..*)?$'})
if(!$notices.Count){throw "License missing $($lib.Name)"}
foreach($notice in $notices){Copy-Item -LiteralPath $notice.FullName -Destination (Join-Path "$out/licenses" ($lib.Name.Replace('/','_')+'_'+$notice.Name))}
}
@"
SleepGuard v$v ($Platform)

ビルド済み配布です。Visual Studioでのビルドは不要です。
Windows 10/11と同じアーキテクチャの.NET 8 Desktop Runtimeが必要です。
ZIPを展開したフォルダのファイル一式を保持してSleepGuard.exeを起動してください。
更新時はタスクトレイからSleepGuardを終了し、ファイル一式を差し替えてください。
ユーザー設定領域に保存された監視対象などの設定は削除しないでください。
詳しい操作はREADMEを参照してください。
"@|Set-Content "$out/INSTALL.txt" -Encoding utf8
Compress-Archive "$out/*" "package/SleepGuard_v${v}_$Platform.zip"

