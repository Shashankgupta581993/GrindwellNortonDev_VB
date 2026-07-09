# Build Signing

Local Debug builds do not require `GNOpcenterKey.snk.pfx`. The project defaults signing off, so `Debug|x64` can build without any local PFX/SNK file.

Client Release builds can still be signed by creating a local, ignored `LocalSigning.props` file next to `GrindwellNortonDev_VB.vbproj`. Point it at a secure local copy of the client PFX:

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <GNOpcenterSignAssembly>true</GNOpcenterSignAssembly>
    <GNOpcenterKeyFile>C:\Secure\Path\GNOpcenterKey.snk.pfx</GNOpcenterKeyFile>
  </PropertyGroup>
</Project>
```

Do not commit PFX/SNK files, certificate material, or passwords. Before pulling the cleanup commit, back up the existing client PFX to a secure location outside the Git working tree.

If the existing PFX is already tracked, remove it from Git only after backing it up:

```powershell
git rm --cached GNOpcenterKey.snk.pfx
```
