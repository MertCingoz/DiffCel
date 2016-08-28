1-Copy excelCompare folder to c:\ 
	(Java based excel comparasion tool
	downloaded from https://github.com/na-ka-na/ExcelCompare
	Used release ExcelCompare 0.6.0) 
	
2-Add fallowing lines to .git\config

[diff "excel"]
  command = "C:/excelCompare/exceldiff.cmd"


3-Add fallowing lines to .gitattributes

# Excel Compare
*.xlsx 	 diff=excel
*.xlsm	 diff=excel
*.xlsb	 diff=excel
*.xltx 	 diff=excel
*.xltm   diff=excel
*.xls	 diff=excel
*.xlt 	 diff=excel
*.xml    diff=excel
*.xlam	 diff=excel
*.xlw	 diff=excel

3-Include java path
	(Now git diff returns excel differences)

4-Create Temp folder on the root directory of the git repository and include fallowing line to .gitignore

[Tt]emp/

5-Include git path
	(To be able to use git command with DiffCel interface)

6-Run regedit.reg
	(To be able to open embedded excel in DiffCel.
	Otherwise it generates new excel process instance)

7-Open DiffCel release and select repository root directory

Usefull commands for git:

-git log --pretty=format:"%h|%an|%s|%ci" "path\to\file.extension"
	gets commits (commitId, Author Name, Commit Description, Commit Date) for specific file   

-git diff <commitID> <commitId> "path\to\file.extension"
	gets differences between two commit for specific file

-git cat-file -p <commitId> "path\to\file.extension" > "path\to\store\file.extension"
	gets original file before that commit (commitId) and saves it.