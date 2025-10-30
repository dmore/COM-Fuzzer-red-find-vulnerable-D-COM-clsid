# COM-Fuzzer

After the succes of the [MS-RPC Fuzzer](https://github.com/warpnet/MS-RPC-Fuzzer), I was wondering if the same approach could be applied to COM/DCOM. This involes fuzzing the COM classes and their interface definitions. Gain insights into COM/DCOM implementations that may be vulnerable using an automated approach and make it easy to visualize the data. By following this approach, a security researcher will hopefully identify interesting COM/DCOM classes in such a time, that would take a manual approach significantly more. 

> [!NOTE]
> The owner of this repository is not responsible for any damage of the usage made using these tools. These are for legal purposes only. Use at your own risks.

## Requirements
* [OleViewDotNet](https://github.com/tyranid/oleviewdotnet) PowerShell module
* PowerShell <7 (PS 7 is not supported)

## Install
Clone the repository and import the COM-Fuzzer module:
```powershell
Import-Module .\COM-Fuzzer.psm1
```
If the required PowerShell module `OleViewDotNet` is not installed, you will be asked to install it.

## Quick example
1. Get COM server data for CLSID `13709620-C279-11CE-A49E-444553540000`
```powershell
Get-ComServerData -OutPath .\output\ -CLSID 13709620-C279-11CE-A49E-444553540000
```
2. Execute calculator
```powershell
'.\output\ComServerData.json' | Invoke-ComFuzzer -Procedure ShellExecute -StringInput "calc.exe" -OutPath .\output\
```

For more examples see [Fuzzing examples](/docs/2%20Invoke-ComFuzzer.md).

## Global overview design
```mermaid
graph TD
    User([User])

    %% Input and output styling
    classDef input fill:#d4fcd4,stroke:#2b8a3e,stroke-width:2px,color:#000;
    classDef output fill:#fff3cd,stroke:#ffbf00,stroke-width:2px,color:#000;

    %% Phase 1: Gather COM Data
    User --> A1[Get-ComServerData]
    A1 --> A2[Target or context specified]
    A2 --> A3[ComServerData.json]
    A3 --> B1[Invoke-ComFuzzer]

    %% Phase 2: Fuzzing
    B1 --> B2[log.txt Call History]
    B1 --> B3[allowed.json]
    B1 --> B4[denied.json]

    %% All fuzzer outputs used in Phase 3
    B3 --> C1[Import-DataToNeo4j]
    B4 --> C1

    %% Phase 3: Analysis
    C1 --> C2[Neo4j Database]
    C2 --> C3[Graph Visualization & Querying]

    %% Apply styling
    class A3 input;
    class B3,B4,B2 output;

    %% Labels for clarity
    subgraph Phase1 [Phase 1: Initialize COM]
        A1
        A2
        A3
    end

    subgraph Phase2 [Phase 2: Fuzzing]
        B1
        B2
        B3
        B4
    end

    subgraph Phase3 [Phase 3: Analysis]
        C1
        C2
        C3
    end
```

## To-do
- Write cypher queries templates for Neo4j
- Implement time out for invoking procedures that take long

## Known bugs
- Find root cause to some PowerShell crashes and fix them

## Acknowledgement
This tool is heavily built upon [OleViewDotNet](https://github.com/tyranid/oleviewdotnet) by [James Forshaw](https://x.com/tiraniddo) with [Google Project Zero](https://googleprojectzero.blogspot.com/). This tool uses the OleViewDotNet module to do most tasks.