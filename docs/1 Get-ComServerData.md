# Phase 1 - Inventarize

Either specify a COM database file, a COM class object or a specific CLSID. If none of the three are specified, a COM database of the system (registry) is parsed.

## Usage
```
NAME
    Get-ComServerData

SYNTAX
    Get-ComServerData [[-ComClass] <COMRegistryEntry[]>] [[-ComDatabaseFile] <String>] [-OutPath] <String>
    [[-ClassContext] <String>] [[-CLSID] <String>] [[-CLSIDList] <String>] [<CommonParameters>]


OPTIONS
    -Output                 Path to export the data to
    -ComClass               Collect data from a ComClass object (OleViewDotNet.Database.COMRegistryEntry[])
    -ComDataBaseFile        Collect data from a COM database file
    -CLSID                  Collect data from a specific COM CLSID    
    -ClassContext           Collect data for a specific range of COM classes (remote, interactive, services)
    -CLSIDList              Collect data for each CLSID listed in parsed file
```

General example:
```powershell
Get-ComServerData -OutPath .\output\ -ClassContext <remote, interactive, services>
```

Load from COM database file (Get with Set-ComDatabase .\com.db):
```powershell
Get-ComServerData -ComDatabaseFile .\com.db -OutPath .\output\

$classes = Get-ComClass -ServiceName "SecurityHealthService"
$classes | Get-ComServerData -OutPath .\output\

Get-ComServerData -OutPath .\output\
```

Because the amount of COM classes, it is wise to specify a context on what to focus the fuzzer on. For example, filter classes that are initiated from a Windows service:
```powershell
Get-ComServerData -ComDatabaseFile .\com.db -OutPath .\output\ -ClassContext Services
```

Or that are initiated from a interactive user:
```powershell
Get-ComServerData -ComDatabaseFile .\com.db -OutPath .\output\ -ClassContext Interactive
```

Or a specific CLSID:
```powershell
Get-ComServerData -OutPath .\output\ -CLSID 13709620-C279-11CE-A49E-44455354000
```

Or a file containing a list of CLSID's:
```powershell
Get-ComServerData -OutPath .\output\ -CLSIDList .\clsids.txt
```