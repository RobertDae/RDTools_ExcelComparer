# RDTools_ExcelComparer

- Input: two xlsx files or csv files
- Output: xlsx-files, cli prompts the number of diffs found 

## Description
Tool which compares xlsx files in open office based on EPPlus and C# (csv files are also possible and will be converter to xlsx)
The tool  uses EPPlus 4.x which is opensource and the .netCore so that a compilation for linux is also possible.

## HowTo for CLI execution

```sh
$ ExcelComparer <fullpathname file1> <fullpathname file2> <fullpathname file3>
$ 
$ e.g.in Windows: ExcelComparer  C:\Inputfolder1\File1.xlsx C:\Inputfolder2\File2.xlsx C:\CompareResults\DiffResults.xlsx
```
