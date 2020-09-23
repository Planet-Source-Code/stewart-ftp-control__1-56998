FTP OCX (Version 1.0.0)

This is the first release of FTPOCX, a freeware and open source ftp activex control.

Currently it supports the following:
  - Connect: You can connect to an ftp server. It handles all the major work for you.
	     Syntax: FTP1.Connect ConnectionName, URL, Port, User, Password
             
	     - ConnectionName is a string and can be anything. It's not to important.
	     - URL is a string and is the ftp url of the server your connecting to.
	     - Port is a string and is the port your connecting to. Usually 21.
	     - User is a string and is the username to connect to the ftp server.
	     - Password is a string and is the password to connect with the username.
	     
  - GetDirectoryListing:  This will allow you to easily list the files in a directory. The
                          files will be stored in a collection which you can get the data of.
                          The following information is stored in the collection.
                   
                            Boolean's:
                              ReadOnly   - If the file is ReadOnly or not.
                              Hiddn      - If the file is Hidden or not.
                              System     - If the file is a System file or not.
                              Directory  - If the file is a Directory or not.
                              Archive    - If the file is an Archive or not.
                              Normal     - If the file is Normal or not.
                              Temporary  - If the file is Temporary or not.
                              Compressed - If the file is Compressed or not.
                              Offline    - If the file is offline or not.
               
                            Date's:
                              CreationTime   - The date the file was created.
                              LastAccessTime - The date the file was last accessed.
                              LastWriteTime  - The date the file was last written to.
                   
                            Long's:
                              FileSize - The filesize of the file.
                 
                            String's:
                              FileName - The filename of the file.
                  
                            Syntax: FTP1.DoList Filter
               
                            Most often it would be used like this:
                            FTP1.GetDirectoryListing "*.*"
               
                            To Access the collection you would do this:
                            Dim Item As New ftpOCX.cDirItem
                            For Each Item in FTP1.Directory
                              MsgBox FTP1.FileName
                            Next
                            
  - DeleteSelection: This will allow you to easily delete a file. There's no need for
                     you to worry about determining if it's a folder or a file. The 
                     control will automaticly determine if it's a folder or file and
                     handle it approprietly. As usual you can not delete folders with
                     files in them.
                     
                     Syntax: FTP1.DeleteSelection File
                     
                       - File is a string. Just is what file you'd like deleted.

  - RenameSelection: This allows you to easily rename a selected file.
 
                     Syntax: FTP1.RenameSelection OldFileName, NewFileName
                     
                       - OldFileName is a string. Just set it to what file your renaming.
                       - NewFileName is a string. This is what the ocx will rename the old file to.

  - SetDirectory: This allows you to easily set the current ftp directory.
  		  Syntax: FTP1.SetDirectory Directory
  		  
  		   - Directory is a string. Just set it to what directory you want to open.

  		                                                 