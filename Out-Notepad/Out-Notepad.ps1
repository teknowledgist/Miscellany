<#
.Synopsis
  Display output in Notepad
.DESCRIPTION
  This function create a temporary file for output, displays the output in notepad, then deletes the temp file
.EXAMPLE
  Get-ADUser $env:USERNAME | Out-Notepad
#>
Function Out-Notepad {
   #this function is designed to take pipelined input
   #example: get-process | out-notepad

   Begin {
      #create the FileSystem COM object
      $tempfile = [IO.Path]::GetTempFileName() |
         Rename-Item -NewName { $_ -replace 'tmp$', 'txt' } -PassThru

      #initialize a placeholder array
      $data=@()
   } #end Begin scriptblock
   
   Process {
      #add pipelined data to $data
      $data+=$_
   } #end Process scriptblock
   
   End {
      #write data to the temp file
      $data | Out-File $tempfile
      
      #Use VB.Interaction Shell/Run with wait
      Add-Type -assemblyname Microsoft.visualBasic
      $command = "NOTEPAD.EXE $tempfile"
      $newProc = [Microsoft.VisualBasic.Interaction]::Shell($command,1)
      
      #sleep for 3 seconds to give Notepad a chance to open the file
      sleep 3
      
      # after notepad is opened, temp file is deleted
      Remove-Item $tempFile -ErrorAction SilentlyContinue
   } #end End scriptblock

} #end Function

Set-Alias on Out-Notepad
