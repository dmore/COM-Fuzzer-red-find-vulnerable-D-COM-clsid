# Import-DataToNeo4j
`Invoke-ComFuzzer` exports its fuzzing results to two separated json files: Allowed.json and Denied.json based on the result of the COM call. These json can be imported to Neo4j using this function. It will do all the convertions and mapping automatically, only need to specify a Neo4j host and credentials.

## Usage
```
NAME
    Import-DataToNeo4j

SYNTAX
    Import-DataToNeo4j [[-jsonFile] <string>] [[-Neo4jHost] <string>] [[-Neo4jUsername] <string>] [<CommonParameters>]

OPTIONS
    -jsonFile               Path to the json file to import (can also be piped)
    -Neo4jHost              IPv4 + Port of the Neo4j host (e.g 192.168.178.89:7474)
    -Neo4jUsername          Username for the Neo4j database
```

## Examples
Export Allowed COM calls to Neo4j (with pipe):
```powershell
'.\output\Allowed.json' | Import-DatatoNeo4j -Neo4jHost 192.168.178.89:7474 -Neo4jUsername neo4j
Enter Neo4j Password: ***********
[+] Successfully connected to Neo4j
[+] Importing data to Neo4j...
```

Export Allowed COM calls to Neo4j (without pipe):
```powershell
Import-DatatoNeo4j -jsonFile '.\output\Allowed.json' -Neo4jHost 192.168.178.89:7474 -Neo4jUsername neo4j
Enter Neo4j Password: ***********
[+] Successfully connected to Neo4j
[+] Importing data to Neo4j...
```

Export Access Denied COM calls to Neo4j:
```powershell
'.\output\Denied.json' | Import-DatatoNeo4j -Neo4jHost 192.168.178.89:7474 -Neo4jUsername neo4j
Enter Neo4j Password: ***********
[+] Successfully connected to Neo4j
[+] Importing data to Neo4j...
```