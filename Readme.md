# DiffCel
Visual excel differences for git-windows

## Installation
### Prerequisites
- [Git for Windows](https://git-scm.com/download/win)
- [Java Runtime Environment](https://java.com/en/download/)

### Usage
- **Copy excelCompare folder to c:\\**
	> Java based excel comparasion tool downloaded from: [na-ka-na/ExcelCompare](https://github.com/na-ka-na/ExcelCompare)  
	
	> Used release ExcelCompare 0.6.0 

- **Add following lines to .git\config**
```
[diff "excel"]
  command = "C:/excelCompare/exceldiff.cmd"
```
- **Add following lines to .gitattributes**
```
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
```
- **Include java path**
	> Now **git diff** returns excel differences

![alt tag](https://raw.githubusercontent.com/MertCingoz/DiffCel/master/Ss/git-diff.PNG)
- **Create Temp folder on the root directory of the git repository and include fallowing line to .gitignore**
```
[Tt]emp/
```
- **Include git path**
	> To be able to use git command with DiffCel interface

- **Run regedit.reg**
	> To be able to open **embedded** excel in DiffCel.

	> Otherwise it generates **new excel process** 

- **Open DiffCel release and select root directory of the git repository**

![alt tag](https://raw.githubusercontent.com/MertCingoz/DiffCel/master/Ss/screen.PNG)

## Usefull commands for git:
```
-git log --pretty=format:"%h|%an|%s|%ci" "path\to\file.extension"
```
gets commits (commitId, Author Name, Commit Description, Commit Date) for specific file   
```
-git diff <commitId> <commitId> "path\to\file.extension"
```
gets differences between two commit for specific file
```
-git cat-file -p <commitId> "path\to\file.extension" > "path\to\store\file.extension"
```
gets original file before that commit (commitId) and saves it.
