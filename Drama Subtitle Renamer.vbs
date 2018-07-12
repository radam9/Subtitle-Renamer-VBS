dim di, si, xi, x, fd, fs

'Request user to input directory of drama, subtitles and episode count
di = InputBox("Enter the directory location of the drama    Example: C:\Dramas\Goblin", "Drama Location") & "\"
si = InputBox("Enter the directory location of the subtitles Example: C:\Dramas\Goblin\Subs", "Subtitles Location") & "\"
xi = InputBox("Enter the number of episodes of the drama", "Episode Count")

'setting up objects
Set fdd = CreateObject("Scripting.FileSystemObject")
Set fss = CreateObject("Scripting.FileSystemObject")
Set d = fdd.GetFolder(di)
Set s = fss.GetFolder(si)
x = xi

dim arrD(),arrS() 'creating array with episode count size
ReDim arrD(x),arrS(x) 'modifying array dimension
dim i,j 'declaring counter
i = 0 'resetting the counter
j = 0 'resetting the counter

'inputing drama file names into an array
for each fD in d.files
		filenameD = di & fD.name
		arrD(i) = fdd.GetBaseName(filenameD)
		i = i + 1
next		

'renaming the subtitle files according to the drama file names array
for each fS in s.files
		filenameS = si & fS.name
		arrS(j) = fss.GetBaseName(filenameS)
		newfilenameS = replace(filenameS, arrS(j), arrD(j))
		fss.MoveFile filenameS, newfilenameS
		j = j + 1
next

msgbox "Done!"
