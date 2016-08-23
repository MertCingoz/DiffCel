1-copy excelCompare to c:\

2-append in .git\config
[diff "excel"]
  command = "C:/excelCompare/exceldiff.cmd"

3-Include path variable for java & git

3-Run regedit.reg

4-Used commands for git:
  git log --pretty=format:"%h|%an|%s" "Excel Files\\Kitap1.xls" >commits.txt
  git diff e22abd5 "Excel Files\\Kitap1.xls" >diff.txt
  git cat-file -p e22abd5 "Excel Files\\Kitap1.xls" > Temp.extension