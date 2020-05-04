# DBF to XLSX Table

index.js loads the DBF files and converts them into CSV storing them in `temp`
Then it spawns a child process `java -jar csvtoxlsx-jar-with-dependencies.jar temp/<filename>.csv <output>/<filename>.xlsx` 
Then it loops through all the files one by one doing that, can't do multiple at a time because there are some really large files that overload the memory

## Building the Java portion

`npm run buildJar`
