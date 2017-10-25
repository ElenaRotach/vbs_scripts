Set objArgs = WScript.Arguments 
For I = 0 to objArgs.Count - 1 
LogPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
Set LogStream = objFSO.OpenTextFile(LogPath & "\" & objArgs(I)  & ".txt", 8, True)
   LogStream.WriteLine "test"
Next 
