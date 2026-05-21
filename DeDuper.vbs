' =======
' DeDuper
' =======
' Version 1.0.4.10 - May 16th 2025
' Copyright © Steve MacGuire 2011-2025
' http://samsoft.org.uk/iTunes/DeDuper.vbs
' Please visit http://samsoft.org.uk/iTunes/scripts.asp for updates

' =======
' Licence
' =======
' This program is free software: you can redistribute it and/or modify it under the terms
' of the GNU General Public License as published by the Free Software Foundation, either
' version 3 of the License, or (at your option) any later version.

' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License for more details.

' Please visit http://www.gnu.org/licenses/gpl-3.0-standalone.html to view the GNU GPLv3 licence.

' ===========
' Description
' ===========
' Aims to dedupe tracks where more than one logical entry exists for the same physical file
' or more than one physical file has the same metadata

' Written partly in response to this thread in Apple Support Communities
' https://discussions.apple.com/message/16038843#16038843
' Thanks to tcminyard for pointing out a potential problem with "/" vs. "\" in file paths.

' Related scripts: DeDuper, DeDupeKeepSmall, Duplicates, DuplicateSong&Artist, ExactDuplicates

' =========
' ChangeLog 
' =========
' Version 1.0.0.1 - Initial version
' Version 1.0.0.2 - Fixed race condition
' Version 1.0.0.3 - Extend to deal with files that have duplicate metadata & the same file size
'                 - Deleted files go to the recycle bin
' Version 1.0.0.4 - Fixed issue arising when there were more than two physical copies of any given file
' Version 1.0.0.5 - Improved distinction between logical and physical duplicates
'                 - Removes any folders emptied as a result of removing duplicates
' Version 1.0.0.6 - Update detection routines to cope with Mac/Linux style file paths that have been imported to the database
'		              - Rewrite recycle routine for clarity
' Version 1.0.1.1 - Updated to common code base with progress bar
' Version 1.0.1.2 - Updated recycle routine for Vista/Windows 7 compatibility
' Version 1.0.1.3 - Added feature to ensure playlist membership is preserved
' Version 1.0.1.4 - Extended to allow optional removal of alternate dupes, keeping largest file
' Version 1.0.1.5 - Fixed potential bug when running under UAC
' Version 1.0.1.6 - Option to keep smallest file, prevent processing the same file twice in a playlist, clean repeats
' Version 1.0.1.7 - Minor update to file delete code, bug hunting...
' Version 1.0.1.8 - Dutch support, can now easily extend support for other languages by updating the Recycle function
' Version 1.0.1.9 - Option to prefer certain file types over others when there are dupes with different size files
' Version 1.0.2.1 - Add automatic deletion of "missing" duplicates
' Version 1.0.2.2 - Tweaks to UAC/LUA detection for scanning phase
' Version 1.0.2.3 - Option to attempt to trash deleted files or delete directly
' Version 1.0.2.4 - Option to treat disc 0 as equivalent to disc 1, improve Recycle function for non-English Windows
' Version 1.0.2.5 - Detect pseudo-physical duplicates, prefer inside media folder
' Version 1.0.2.6 - Minor update to common code, add trap for possible error in Merge routine
' Version 1.0.2.7 - Added error handler for file deletion routine
' Version 1.0.2.8 - Fixed error in track confirmation code
' Version 1.0.2.9 - More error handling
' Version 1.0.3.1 - Fix missing Ext function
' Version 1.0.3.2 - Option to archive to <Media Folder>\Archive instead of deleting to recycle bin
' Version 1.0.3.3 - Attempt to fix sporadic missing object error
' Version 1.0.3.4 - Tweaks to dialog boxes when archiving, added some troubleshooting code that can be enabled with a simple edit
' Version 1.0.3.5 - Swap links to alternate dupes to preserve date added, uncheck dupes that will be removed to allow for manual review
' Version 1.0.3.6 - Add ability to match and dedupe streams, may not work correctly if favouring smallest file but not actually tested
' Version 1.0.4.1 - Add ability to keep newest or oldest file by date added or modified, improved start screen & report to show selected options
' Version 1.0.4.2 - Tweak to recycling code in case unable to locate folder object of item to recycle
' Version 1.0.4.3 - New ability to dedupe same song name/artist dupes, unchecking dupe songs and removing from regular playlists
'                 - Updates to alt. version rules and fix for selection reset after deleting from a playlist
' Version 1.0.4.4 - Optimise same song dupe playlist removal code             
' Version 1.0.4.5 - Modify Recycle function for network paths
' Version 1.0.4.6 - Test/Fix edge case in GetMediaPath
' Version 1.0.4.7 - Remove repeat spaces and trim before comparing values
' Version 1.0.4.8 - Attempt to trap the error mentioned in this thread: https://discussions.apple.com/thread/255969997
' Version 1.0.4.9 - And now this one: https://discussions.apple.com/thread/256013982
' Version 1.0.4.10 - And another: https://discussions.apple.com/thread/256060858



' ==========
' To-do List
' ==========
' Add things to do
' Eliminate risk of incorrect deletions when multiple paths are available to the same file.

' =============================
' Declare constants & variables
' =============================
' Variables for common code
' Modified 2014-04-09
Option Explicit	        ' Declare all variables before use
Dim Intro,Outro,Check   ' Manage confirmation dialogs
Dim PB,Prog,Debug       ' Control the progress bar
Dim Clock,T1,T2,Timing  ' The secret of great comedy
Dim Named,Source        ' Control use on named playlist
Dim Playlist,List       ' Name for any generated playlist, and the object itself
Dim iTunes              ' Handle to iTunes application
Dim Tracks              ' A collection of track objects
Dim Count               ' The number of tracks
Dim D,M,P,S,U,V         ' Counters
Dim nl,tab              ' New line/tab strings
Dim IDs                 ' A dictionary object used to ensure each object is processed once
Dim Rev                 ' Control processing order, usually reversed
Dim Quit                ' Used to abort script
Dim Title,Summary       ' Text for dialog boxes
Dim Tracing             ' Display/suppress tracing messages

' Values for common code
' Modified 2014-04-15
Const Kimo=False        ' True if script expects "Keep iTunes Media folder organised" to be disabled
Const Min=2             ' Minimum number of tracks this script should work with
Const Max=0             ' Maximum number of tracks this script should work with, 0 for no limit
Const Warn=500          ' Warning level, require confirmation for processing above this level
Intro=True              ' Set false to skip initial prompts, avoid if non-reversible actions
Outro=True              ' Produce summary report
Check=True              ' Track-by-track confirmation, can be set during Intro
Prog=True               ' Display progress bar, may be disabled by UAC/LUA settings
Debug=True              ' Include any debug messages in progress bar
Timing=True             ' Display running time in summary report
Source=""               ' Named playlist to process, use "Library" for entire library
Rev=False               ' Control processing order, usually reversed
Debug=True              ' Include any debug messages in progress bar
Tracing=True            ' Display tracing message boxes

' Additional variables for this particular script
' Modified 2024-09-23
Dim Paths               ' A dictionary object for comparing paths
Dim DeathRow            ' A dictionary object for tracks to remove
Dim PlayDupes           ' A dictionary object to organize playlist duplicate removal
Dim SameDupes,DupeLists ' Dictionary objects to organize same song duplicate removals
Dim FSO                 ' Handle to FileSystemObject
Dim Reg                 ' Handle to Registry object
Dim SH                  ' Handle to Shell application
Dim B(),C()             ' Arrays of counters
Dim CheckTypes          ' Check duplicate types to delete
Dim SameSongs           ' Include pass for same song duplicates
Dim IgnoreBraces        ' Ignore brackets when looking for same song dupes ()/[]
Dim MaxChars            ' Truncate long titles when looking for same song dupes
Dim KeepLarge           ' Which different sized dupe to keep? True=Largest, otherwise=Smallest, may be overridden by PrefAge or PrefKind
Dim PrefAge             ' Preferred file by age: 0=Disabled,1=Oldest Date Added,2=Oldest Date Modified,3=Newest Date Added,4=Newest Date Modified
Dim PrefChecked         ' When two possibilites, only one checked, keep checked, allows manual preference
Dim PrefKind,PrefKinds  ' Manage alternate dupes by file kind, overrides PrefAge
Dim PrefRating          ' Prefer highest rated track
Dim ProgScan            ' Control progress bar in scan phase, may be disabled by UAC/LUA settings
Dim UseTrash            ' Attempt to send local deleted files to trash
Dim Disc0as1            ' Treat disc 0 as disc 1, matches with otherwise identical properties
Dim Library,Root        ' Paths to library and media folders
Dim Thumbs              ' Generate "promoted" thumbnails for artists with single albums
Dim Archive             ' Flag to archive at <Media Folder>\Archive instead of deleting
Dim Uncheck             ' Uncheck potential discards, note might cause issues if working with a smart playlist
Dim Testing             ' Use to prevent tracks being removed and deleted when testing, use Uncheck to view potential discards
Dim SafeList            ' Independent list of selected items to workaround new issue with volatile lists

' Initialise variables for this particular script
' Modified 2024-09-23
CheckTypes=True         ' Confirm which types of dupes are to be removed
SameSongs=False         ' Include pass for same song duplicates
IgnoreBraces=False      ' Ignore brackets when looking for same song dupes ()/[]
MaxChars=0              ' Truncate long titles when looking for same song dupes
KeepLarge=True          ' Which different sized dupe to keep? True=Largest, otherwise=Smallest, may be overridden by PrefAge or PrefKind
PrefAge=0               ' Preferred file by age: 0=Disabled,1=Oldest Date Added,2=Oldest Date Modified,3=Newest Date Added,4=Newest Date Modified
PrefChecked=False       ' When two possibilites for same songs, only one checked, keep checked, allows manual preference
PrefKind=True           ' Set to prefer one format over another, overrides PrefAge
PrefKinds=".mp3.m4a"    ' One or more preferred file extensions, most preferred first, e.g. ".mp3.m4a"
PrefRating=False        ' Prefer highest rated track for same song dupes
ProgScan=Prog           ' Control progress bar in scan phase, may be disabled by UAC/LUA settings
UseTrash=True           ' Attempt to send local deleted files to trash 
Disc0as1=True           ' Treat disc 0 as disc 1, matches with otherwise identical properties
Root=""                 ' Path to media folder
Thumbs=False            ' Generate "promoted" thumbnails for artists with single albums
Archive=False           ' Flag to archive at <Media Folder>\Archive instead of deleting
Uncheck=True            ' Uncheck potential discards, note might cause issues if working with a smart playlist
Testing=False           ' Use to prevent tracks being removed and deleted when testing, use Uncheck to view potential discards

Title="DeDuper"         ' Alt. Title="DeDupe Keep Smallest/Largest/Mp3"
If Testing Then Title=Title & " - TEST MODE - No tracks removed or deleted"
Summary="Scan selected tracks or the current playlist for duplicates and optionally remove them. "
If Uncheck Then Summary=Summary & "Tracks that would be removed will be unchecked. " 
Summary=Summary & "Tracks in the cloud are ignored." & vbCrLf & vbCrLf _
  & "Plays & skips will be merged, most recent played/skipped dates will be " _
  & "used, and the track that is preserved will be added to any playlists " _
  & "that the removed tracks were in." & vbCrLf & vbCrLf 
If PrefKind Then
  Summary=Summary & "For alternate dupes these extensions are preferred, best first: " & vbCrLf & PrefKinds
ElseIf PrefAge Then
  Select Case PrefAge
    Case 1
      Summary=Summary & "For alternate dupes oldest date added is preferred."   
    Case 2
      Summary=Summary & "For alternate dupes oldest date modified is preferred."   
    Case 3
      Summary=Summary & "For alternate dupes newest date added is preferred."   
    Case 4
      Summary=Summary & "For alternate dupes newest date modified is preferred."   
  End Select
ElseIf KeepLarge Then
      Summary=Summary & "For alternate dupes largest file is preferred."   
Else
      Summary=Summary & "For alternate dupes smallest file is preferred."   
End If
Summary=Summary & vbCrLf & vbCrLf _  
  & "It should help to reduce the overall execution time if you use the " _
  & "iTunes feature to ''Show "
If Not(SameSongs) Then Summary=Summary & "Same Album "
Summary=Summary & "Duplicates'' and select the displayed tracks before running the script."


' ============
' Main program
' ============

GetTracks               ' Set things up
SetTracks               ' Copy tracks into a list of persistent objects
DedupeTracks            ' Main process 
Results                 ' Summary

' ===================
' End of main program
' ===================


' ===============================
' Declare subroutines & functions
' ===============================


' Note: The bulk of the code in this script is concerned with making sure that only suitable tracks are processed by
'       the following module and supporting numerous options for track selection, confirmation, progress and results.


' Loop through track selection processing suitable items
' Modified 2024-09-23
Sub DedupeTracks
  Dim A,Album,Artist,DN,F,File,I,ID,J,K,Key,Kind,L,List,N,Name,NewFolder,NewPath,O,OldFolder,Pass,PL,PlayDupes,PN,PT,Q,QS,R,S,SL,T,TN,Verb,W
  Set FSO=CreateObject("Scripting.FileSystemObject")
  Set SH=CreateObject("Shell.Application")
  Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	' Use . for local computer, otherwise could be computer name or IP address
  Set DeathRow=CreateObject("Scripting.Dictionary")
  Set Paths=CreateObject("Scripting.Dictionary")
  Set PlayDupes=CreateObject("Scripting.Dictionary")
  Set SameDupes=CreateObject("Scripting.Dictionary")
  Set DupeLists=CreateObject("Scripting.Dictionary")
  Set SL=iTunes.BrowserWindow.SelectedPlaylist
  PN=SL.Name
  ReDim B(6)
  ReDim C(6)
  
  Prog=ProgScan And (UAC=False) ' Control progress bar in scan phase
  If Prog Then                  ' Create ProgessBar
    Set PB=New ProgBar
    PB.Title=Title
    PB.Show
  End If
  A=0
  B(6)=0:C(6)=0
  Clock=0
  StartTimer
  If Testing Then
    If Archive Then Verb="Simulate archiving" Else Verb="Simulate cleaning"
  Else
    If Archive Then Verb="Archive" Else Verb="Clean"
  End If

  For Pass=0 To 5
    B(Pass)=0:C(Pass)=0
    If Pass=5 And Not(SameSongs) Then Exit For
    If Pass>1 and Pass<4 Then Paths.RemoveAll   ' Reset list of items to match  
    Select Case Pass
    Case 0
      Kind="playlist"
    Case 1
      Kind="logical"
  ' Case 1.6 Idea for later, for when one file appears on multiple paths...
  '   Kind="pseudo-physical" - Implemented in a different way as part of pass 2
    Case 2
      Kind="physical"
    Case 3
      Kind="alternate"
    Case 4
      Kind="missing"
    Case 5
      Kind="same song"
    End Select
    If Prog Then PB.Reset : PB.SetInfo "Scanning for " & Kind & " duplicates"
    For I=1 To Count                    ' Work forwards, so any playlist dupes are removed from later in the selection
      If Pass=4 Then P=P+1              ' Increment processed tracks
      If Prog Then
        PB.SetStatus "Pass " & Pass+1 & ", processing " & GroupDig(I) & " of " & GroupDig(Count)
        PB.Progress I,Count
      End If
      ' Set T=Tracks(I)
      Set T=ObjectFromID(SafeList.Item(I))      ' Workaround for volatile object collection
      ID=PersistentID(T)
      If (T.Kind=1 Or T.Kind=3) And PlayDupes.Exists(I)=False And DeathRow.Exists(ID)=False Then   ' Ignore previously identified dupes 
        With T
          If T.Kind=3 Then L=LCase(.URL) Else L=LCase(Replace(.Location & "","/","\"))
          DN=.DiscNumber : If DN=0 And Disc0as1 Then DN=1
          TN=.TrackNumber
          Album=Despace(.Album)
          Artist=Despace(.Artist)
          Name=Despace(.Name)
          ' Generate a key for this item
          Select Case Pass
          Case 0                        ' Playlist dupes, same PersistentID
            If IDs.Exists(ID) Then
              'Trace T,"Pass " & Pass & ", item #" & I & " is a playlist dupe of #" & IDs.Item(ID)
              PlayDupes.Add I,0         ' Note index to prevent reprocessing in future passes
              C(0)=C(0)+1               ' Increment count of playlist dupes
            Else
              IDs.Add ID,I              ' Note ID to recognize playlist dupes
            End If
          Case 1                        ' Logical dupes, same file path
            ' Key=L                     ' No need to reassign key
          Case 2                        ' Physical dupes, same details and size 
            If L<>"" Then Key=LCase(Artist & "\" & Album & "\" & DN & "." & TN & " " & Name & "=" & .Size)
          Case 3                        ' Alternate dupes, same details but different size 
            If L<>"" Then Key=LCase(Artist & "\" & Album & "\" & DN & "." & TN & " " & Name)
          Case 4                        ' Missing dupes, same details but file missing 
            If L="" Then Key=LCase(Artist & "\" & Album & "\" & DN & "." & TN & " " & Name) Else Key=""
          Case 5                        ' Same song dupes, same artist and name, ignore album 
            If L<>"" Then
              If IgnoreBraces Or (MaxChars>0) Then
                Key=LCase(Artist & "\" & Truncate(Name))
              Else
                Key=LCase(Artist & "\" & Name)
              End If
            End If
          End Select
          ' Check for match with existing list of keys, merge if found
          If Pass>0 And Key<>"" Then
            If Paths.Exists(Key) Then   ' Duplicate found, choose which to keep
              'Trace T,"Pass " & Pass & ", item #" & I & " is a dupe of #" & Paths.Item(Key) & " with key " & nl & Key
              Merge Key,Pass,I,Paths.Item(Key)
              C(Pass)=C(Pass)+1
            Else
              If Pass=4 Then            ' Missing file with no match
                'Trace T,"Pass " & Pass & ", missing original #" & I & " with key " & nl & Key
                M=M+1
              Else                      ' Add original to the list
                'Trace T,"Pass " & Pass & ", adding original #" & I & " with key " & nl & Key
                Paths.Add Key,I
              End If
            End If
          End If
        End With
      End If
      If Quit Then StopTimer : Exit Sub ' Abort loop on user request
    Next
  Next   
  
  If Check Then
    If Prog Then 
      PB.Close                          ' Hide progress bar during track-by-track confirmation
      Prog=False
    End If
  End If

  C(2)=C(2)-C(6)                        ' Adjust physical/pseudo-physical count
  For I=0 To 6 : A=A+C(I) : Next        ' All dupes
  D=A                                   ' Note dupes found for report
  F=True  

  If A>0 And (Check Or CheckTypes) Then
    If A=1 Then
      If C(0)=1 Then Q="playlist"
      If C(1)=1 Then Q="logical"
      If C(2)=1 Then Q="physical"
      If C(3)=1 Then Q="alternate"
      If C(4)=1 Then Q="missing"
      If C(5)=1 Then Q="same song"
      If C(6)=1 Then Q="pseudo-physical"
      Q="One " & Q & " duplicate was found." & nl & nl & "Would you like to " & LCase(Verb) & " it?"
    Else
      Q="There were " & GroupDig(A) & " duplicates found from " & GroupDig(Count) & " items:" & nl
      If C(0)>0 Then Q=Q & nl & Wrap(GroupDig(C(0)) & tab & "playlist" & tab & tab & "(multiple items in playlist with same database ID)",50," ",3)
      If C(1)>0 Then Q=Q & nl & Wrap(GroupDig(C(1)) & tab & "logical " & tab & tab & "(multiple entries in database for same file)",50," ",3)
      If C(6)>0 Then Q=Q & nl & Wrap(GroupDig(C(6)) & tab & "pseudo-physical" & tab & "(multiple links to same file via different paths)",50," ",3)
      If C(2)>0 Then Q=Q & nl & Wrap(GroupDig(C(2)) & tab & "physical"  & tab & tab &"(multiple copies with identical properties & size)",50," ",3)
      If C(3)>0 Then Q=Q & nl & Wrap(GroupDig(C(3)) & tab & "alternate" & tab & tab & "(multiple sized files with identical properties)",50," ",3)
      If C(4)>0 Then Q=Q & nl & Wrap(GroupDig(C(4)) & tab & "missing" & tab & tab & "(lost files with properties matching found files)",50," ",3)
      If C(5)>0 Then
				QS=GroupDig(C(5)) & tab & "same song" & tab & "(songs with same"
        If IgnoreBraces Or (MaxChars>0) Then QS=QS & "/similar"
        QS=QS & " name and artist)"
        Q=Q & nl & Wrap(QS,50," ",3)
      End If
      Q=Q & nl & "leaving:" & nl 
      If M>0 Then Q=Q & nl & GroupDig(M) & tab & "unmatched" & tab & tab & "(lost files with no match in found files)"
      Q=Q & nl & GroupDig((Count-A-M)) & tab & "original" & tab & tab & "(unique files after deduping)"
      Q=Q & nl & nl 
      If Check Then
        Q=Q & "Would you like to " & LCase(Verb) & " the duplicates with track-by-track confirmation?"
      Else
        Q=Q & "Would you like to " & LCase(Verb) & " the duplicates?"
      End If
      Q=Q & nl & nl 
      If C(3)>0 Then Q=Q & "Yes" Else Q=Q & "OK"
      Q=Q & tab & ": " & Verb & " all duplicate tracks"
      If Not Check Then Q=Q & " automatically"
      If C(3)>0 Then Q=Q & nl & "No" & tab & ": " & Verb & " all but alternate dupes (safer)"
      Q=Q & nl & "Cancel" & tab & ": Abort script"
    End If  
    StopTimer           ' Don't time user inputs
    If C(3)>0 Then R=MsgBox(Q,vbYesNoCancel,Title) Else R=MsgBox(Q,vbOKCancel,Title)
    StartTimer          ' Don't time user inputs
    If R=vbYes Then
      O=True            ' Delete alternate dupes
    ElseIf R=vbNo Then
      O=False
      If A=1 Then F=False
    ElseIf R=vbOK Then
      ' F=True          ' Already true
    Else                ' vbCancel
      F=False
    End If  
  End If  

  
  ' Delete all dupes

  If F=True And A>0 Then        ' Delete any dupes if found, unless cancelling
    
    If C(0)>0 Then              ' Remove any playlist duplicates
      If Prog Then PB.Reset : PB.SetInfo "Removing playlist duplicates"
      N=0
      A=A-C(0)                  ' All except playlist dupes
      For I=Count To 1 Step -1  ' Work backwards in case edit removes item from selection
        ' Set T=Tracks(I)       ' Can cause error if playlist resets
        Set T=ObjectFromID(SafeList.Item(I))      ' Workaround for volatile object collection
        ID=PersistentID(T)
        If Prog Then
          N=N+1    
          PB.SetStatus "Pass " & Pass+1 & ", processing " & GroupDig(N) & " of " & GroupDig(Count)
          PB.Progress N-1,Count
        End If
        If PlayDupes.Exists(I) Then
          F=True
          If Check Then
            Q="Delete playlist duplicate item #" & I & "?"
            Q=Q & nl & nl & Info(T)
            StopTimer           ' Don't time user inputs
            R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title)
            If R=vbCancel Then Quit=True : Exit Sub
            StartTimer          ' Don't time user inputs
            If R=vbNo Then F=False
          End If
          Set SL=iTunes.BrowserWindow.SelectedPlaylist        ' Reconnect to list which may reset after item deleted 
          If SL.Name<>PN Then F=False : MsgBox "Selected playlist has changed, playlist dupes won't be removed.",0,Title
          If F Then
            For K=SL.Tracks.Count To 1 Step -1
              If PersistentID(SL.Tracks.Item(K))=ID Then
                Set PT=SL.Tracks.Item(K)
                ' Trace PT, "Removing item " & K & " from playlist " & SL.Name
                PT.Delete       ' Remove repeated entry from current playlist
                Exit For        ' We're only removing one item in this pass of the outer loop
              End If
            Next
            ' T.Delete          ' Remove repeated entry from current playlist
            B(0)=B(0)+1         ' Increment update count
          End If
        End If
        If Quit Then Exit Sub	  ' Abort loop on user request
      Next
      Pass=Pass+1
    End If

    If A>0 Then                 ' Remove all other duplicates
      If Prog Then PB.Reset : PB.SetInfo "Removing non-playlist duplicates"
      N=0
      For Each ID In DeathRow.Keys        ' Process global objects to avoid errors if duplicates are removed from a playlist
        If Prog Then
          N=N+1    
          PB.SetStatus "Pass " & Pass+1 & ", processing " & N & " of " & A
          PB.Progress N-1,A
        End If
        Set T=ObjectFromID(ID)
        F=True
        If Not T is Nothing Then    ' Check we have an object
          W=DeathRow.Item(ID)
          If T.Kind=3 Then L=T.URL Else L=Replace(T.Location & "","/","\")
          If Check And (W<>3 OR O=True) Then
            Select Case W
            Case 1
              Q="Delete logical duplicate connected to file: " & nl & nl & L & "?"
            Case 2
              Q="Delete physical duplicate file: " & nl & nl & L & "?"
            Case 3
              Q="Delete alternate duplicate file: " & nl & nl & L & "?"
            Case 4
              Q="Delete missing duplicate file: " & nl & nl & Info(T) & "?"
            Case 5
              Q="Remove same song duplicate from regular playlists: " & nl & nl & Info(T) & "?"
            Case 6
              Q="Delete pseudo-physical duplicate file on path: " & nl & nl & L & "?"
            End Select            
            StopTimer               ' Don't time user inputs
            R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title)
            If R=vbCancel Then Quit=True : Exit Sub
            StartTimer              ' Don't time user inputs
            If R=vbNo Then
              F=False
            End If            
          ElseIf W=3 And O=False Then
            F=False
          End If
          If F Then
            If W=5 Then             ' Same song duplicates, remove from regular playlists
              SameDupes.Add ID,5    ' Queue up this item for playlist cleaning in the next phase
              Set PL=T.Playlists
              For J=1 To PL.Count
                If PL.Item(J).Kind=2 And PL.Item(J).Smart=False Then
                  PT=PersistentID(PL.Item(J))
                  If Not DupeLists.Exists(PT) Then DupeLists.Add PT,7
                End If
              Next
            ElseIf (W=2 or W=3) And T.Kind=1 Then   ' Recycle/archive physical/alternate dupes
              If Archive Then                   ' Move deleted files to <Media Folder>\Archive
                S=0
                If Root="" Then GetRoot("archive")
                OldFolder=FSO.GetParentFolderName(L)
                NewFolder=Root & "Archive\" & Relative(L)
                Set File=FSO.GetFile(L)
                MakePath(NewFolder)
                NewPath=NewFolder & "\" & FSO.GetFileName(L)
                Do While FSO.FileExists(NewPath)        ' Add a suffix as required to prevent overwrites
                  S=S+1
                  NewPath=NewFolder & "\" & FSO.GetBaseName(L) & " " & S & "." & FSO.GetExtensionName(L)
                Loop
                ' Trace T,"Archive from " & L & nl & nl & "to " & NewPath
                If Not Testing Then             ' Testing allows script to uncheck files that would be removed, without removing them
                  File.Move(NewPath)
                  MoveArt OldFolder,NewFolder   ' Move artwork if the last media file has been moved
                End If
              Else
                If Not Testing Then             ' Testing allows script to uncheck files that would be removed, without removing them
                  ' Trace T,"Recycle from " & L 
                  Recycle L
                End If
              End If
            End If
            If W<>5 Then                        ' Don't remove same songs dupes
              If Not Testing Then               ' Testing allows script to uncheck files that would be removed, without removing them
                'Trace T, "Removing type " & W & " dupe from library"
                T.Delete                        ' Delete this item from the library
                ' DeathRow.Remove ID            ' Superfluous? Worse... Might upset "For Each" loop
              End If
            End IF
            B(W)=B(W)+1                 ' Increment update count
          End If
        End If
        
        If Quit Then Exit Sub	          ' Abort loop on user request
      Next
    End If
  End If
  
  ' Clean up same song dupes from their playlists
  Set PL=iTunes.Sources.Item(1).Playlists
  For J=1 To PL.Count
    If DupeLists.Exists(PersistentID(PL.Item(J))) Then
      For K=PL.Item(J).Tracks.Count To 1 Step -1
        If SameDupes.Exists(PersistentID(PL.Item(J).Tracks.Item(K))) Then
          Set PT=PL.Item(J).Tracks.Item(K)
          ' Trace PT, "Removing item " & K & " from playlist " & PL.Item(J).Name
          PT.Delete
        End IF
      Next
    End If
  Next
  
  StopTimer
  If Prog Then
    PB.SetStatus "Finished!"
    PB.Progress Count,Count
    WScript.Sleep 1000
    PB.Close
  End If

End Sub


' Remove any repeated spaces and trim
' Modified 2024-09-23
Function Despace(S)
  S=RegRep(S,"\s+"," ")       ' Remove any repeat spaces
  Despace=Trim(S)
End Function


' Attempt to determine root of media path by inspecting location of media files
' Modified 2022-10-13
Function GetMediaPath
  Dim A,C,I,L,P,S,T,Tracks
  Set Tracks=iTunes.LibraryPlaylist.Tracks
  C=Tracks.Count
  If C>100 Then C=100		' Give up if can't find one valid location in the first 100 attempts
  I=1
  P=""
  Do
    Set T=Tracks.Item(I)
    If T.Kind=1 Then		' Only process "File" tracks
      With T
        L=.Location
        If L<>"" Then
          A=.AlbumArtist
          If A="" Then A=.Artist
          A=ValidiTunes(A,"")
          If .Compilation Then A="Compilations"
          If .Podcast Then
            A=ValidiTunes(.Album,"")
          ElseIf .VideoKind=1 Then
            A=ValidiTunes(.Name,"")
          ElseIf .VideoKind=3 Then
            A=ValidiTunes(.Show,"")
          End If
          If Instr(L,A) Then
            On Error Resume Next        ' Trap possible error
            P=Left(L,Instr(L,A)-2)
            If Err.Number<>0 Then
              Trace T,"Error parsing file properties..." & nl & nl & "Location=" & L & nl & nl & "Artist=" & A & "." & nl & nl &  Err.Description
              Err.Clear
            End If
            On Error Goto 0             ' End of error trap  
            P=Left(L,Instr(L,A)-2)
            S=Mid(P,InStrRev(P,"\"))
            If Instr("\Audiobooks\Books\iPod Games\iTunes U\Mobile Applications\Movies\Music\Podcasts\Ringtones\TV Shows",S) Then P=Left(P,Len(P)-Len(S))
          Else
            ' MsgBox "Artist:" & tab & .Artist & nl & "Name:" & tab & .Name & nl & "Location:" & tab & .Location
          End If
        Else
          ' MsgBox "Artist:" & tab & .Artist & nl & "Name:" & tab & .Name & nl & "Location:" & tab & .Location
        End If
      End With
    End If
    I=I+1
  Loop Until P<>"" OR I>C
  ' MsgBox "Media path is " & P & nl & "Found in " & I-1 & " step" & Plural(I-1,"s","")
  GetMediaPath=P
End Function


' Get iTunes Media folder
' Modified 2025-05-16
Sub GetRoot(Reason)
  Dim F,M,R
  If Root<>"" Then If FSO.FolderExists(Root)=False Then Root=""
  If Root="" Then
    Root=GetMediaPath
    If Root="" Then
      On Error Resume Next      ' Ignore potential error with this call
      Root=iTunes.LibraryXMLPath
      Root=Left(Root,InStrRev(Root,"\")-1)
      On Error Goto 0           ' Restore normal error handler
    End If
    F=False
    If Root="" Then Root=Library
    Do
      If FSO.FolderExists(Root & "\iTunes Media") Then Root=Root & "\iTunes Media"
      If FSO.FolderExists(Root & "\iTunes Music") Then Root=Root & "\iTunes Music"
      M="Please confirm/edit the location of your media folder. "
      If Reason="archive" Then
        M=M & "Archived files will be moved to" & nl & "<Media Folder>\Archive."
      Else
        M=M & "DeDuper has found pseudo-physical dupes and aims to favour those inside the media folder."
      End If
      Root=InputBox(M,Title,Root)
      If Right(Root,1)="\" Then Root=Left(Root,Len(Root)-1)
      If Root="" Then WScript.Quit
      If Not FSO.FolderExists(Root) Then
        R=MsgBox("The folder " & Root & " does not exist." & nl & "Shall I create it?",vbYesNoCancel,Title)
        If R=vbCancel Then WScript.Quit
        'If R=vbYes Then MakePath(Root)
        If R=vbYes Then F=True
      End If
    Loop Until F Or FSO.FolderExists(Root)
  End If  
  If Right(Root,1)<>"\" Then Root=Root & "\"
  If Right(Library,1)<>"\" Then Library=Library & "\"
  'Org=(Layout="1")
End Sub


' Custom info message for progress bar or trace reports
' Modified 2022-02-05
Function Info(T)
  On Error Resume Next
  Dim A,B
  With T
    A=.AlbumArtist & "" : If A="" Then A=.Artist & "" : If A="" Then A="Unknown Artist"
    B=.Album & "" : If B="" Then B="Unknown Album"
    Info=A & "\" & B & "\"
    If .DiscNumber>1 or .DiscCount>1 Then Info=Info & .DiscNumber & "-"
    If .TrackNumber<10 Then Info=Info & "0"
    Info=Info & .TrackNumber & " " & .Name
    If Err.Number>0 Then
      MsgBox "Problem with item " & .Name,0,Title
    End If
  End With
End Function


' Check to see if a track is in any regular playlists
' Modified 2022-02-01
Function InLists(T)
  Dim J,P
  InLists=False
  Set P=T.Playlists
  For J=1 To P.Count
    If P.Item(J).Kind=2 And P.Item(J).Smart=False Then
      InLists=True
      Exit For
    End If
  Next
End Function


' Create a folder path if it doesn't already exist
' Modified 2011-09-17
Function MakePath(Path)
  ' Default result
  MakePath=False
  ' Fail if drive is not valid
  If Not FSO.DriveExists(FSO.GetDriveName(Path)) Then Exit Function
  ' Succeed if folder exists
  If FSO.FolderExists(Path) Then
    MakePath=True
    Exit Function
  End If
  ' Call self to ensure parent path exists
  If Not MakePath(FSO.GetParentFolderName(Path)) Then Exit function
  ' Create folder
  On Error Resume Next
  FSO.CreateFolder Path
  MakePath=FSO.FolderExists(Path)    
End Function


' Merge metadata for two tracks, choosing which to preserve and prepare to delete the other
' Modified 2025-03-18
Sub Merge(A,W,X,Y)
  Dim E,F,I,J,K,L,P,Playlists,Pseudo,Swap,PPE
  PPE=False
  Pseudo=False
  Swap=False
  Set Playlists=CreateObject("Scripting.Dictionary")
  Select Case W
    Case 2                              ' Physical: Check for pseudo-physicals first
      If Tracks(X).Kind=3 Then E=LCase(Tracks(X).URL) Else E=LCase(FSO.GetFileName(Tracks(X).Location))
      If Tracks(Y).Kind=3 Then F=LCase(Tracks(Y).URL) Else F=LCase(FSO.GetFileName(Tracks(Y).Location))
      If E=F And Tracks(X).Kind=1 And Tracks(Y).Kind=1 Then       ' Possible problem, check more closely
        On Error Resume Next            ' Trap possible error
        Set E=FSO.GetFile(Tracks(X).Location)   ' Command that might error
        If Err.Number<>0 Then
          Trace Tracks(X),Err.Description
          Err.Clear
          PPE=True
        End If
        On Error Goto 0                 ' End of error trap  
        On Error Resume Next            ' Trap possible error
        Set F=FSO.GetFile(Tracks(Y).Location)   ' Command that might error
        If Err.Number<>0 Then
          Trace Tracks(Y),Err.Description
          Err.Clear
          PPE=True
        End If
        On Error Goto 0                 ' End of error trap  
        If PPE=False Then               ' If error occurred within traps above then likely not pseudo-physical dupe as, presumably, at least one track now missing 
          I=E.Attributes                ' Get attributes
          J=F.Attributes
          If I=J Then                   ' ignore if different
            E.Attributes=I XOR 32       ' See if flipping one archive bit affects the other
            If F.Attributes<>J Then     ' Pseudo-physical duplicate found
              Pseudo=True
            End If
            E.Attributes=I              ' Reset archive flag
          End If
        End If
      End If
      If Pseudo Then                    ' Pseudo-physical duplicate found, prefer the one in the media folder
        If Root="" Then GetRoot("pseudo-physical")
        ' E=LCase(Tracks(X).Location)   ' Not needed for test
        F=LCase(Tracks(Y).Location)
        If Instr(F,LCase(Root)) Then Swap=True
      Else                              ' Normal physical dupe, prefer the older one
        If Tracks(Y).DateAdded<Tracks(X).DateAdded Then Swap=True
      End If
    Case 3                              ' Alternate dupes: Keep preferred format/age or largest/smallest file
      If PrefKind Then                  ' Prefer file type
        If Tracks(X).Kind=3 Then E=Ext(LCase(Tracks(X).URL)) Else E=Ext(LCase(Tracks(X).Location))
        If Tracks(Y).Kind=3 Then F=Ext(LCase(Tracks(Y).URL)) Else F=Ext(LCase(Tracks(Y).Location))
        If E="" Then MsgBox "Error working with files that have no extension" & nl & Tracks(X).Location,vbCritical,Title : WScript.Quit
        If F="" Then MsgBox "Error working with files that have no extension" & nl & Tracks(Y).Location,vbCritical,Title : WScript.Quit
        If E<>F Then
          If (Instr(LCase(PrefKinds),E) And Instr(LCase(PrefKinds),F)) Then
            If Instr(LCase(PrefKinds),E)>Instr(LCase(PrefKinds),F) Then Swap=True
          Else
            If Instr(LCase(PrefKinds),F) Then Swap=True
          End If  
        End If  
        ' Trace Null,"Checking file kind: " & PrefKinds & nl & E & " vs. " & F & nl & "Swap=" & Swap
      ElseIf PrefAge>0 Then             ' Prefer by age
        Select Case PrefAge
          Case 1                        ' Oldest Date Added
            If Tracks(Y).DateAdded<Tracks(X).DateAdded Then Swap=True
          Case 2                        ' Oldest Date Modified
            If Tracks(Y).ModificationDate<Tracks(X).ModificationDate Then Swap=True
          Case 3                        ' Newest Date Added
            If Tracks(Y).DateAdded>Tracks(X).DateAdded Then Swap=True
          Case 4                        ' Newest Date Modified
            If Tracks(Y).ModificationDate>Tracks(X).ModificationDate Then Swap=True
        End Select
      Else                              ' Prefer by size
        If KeepLarge Then               ' Keep largest file
          If Tracks(Y).Size>Tracks(X).Size Then Swap=True
        Else                            ' Keep smallest file
          If Tracks(Y).Size<Tracks(X).Size Then Swap=True
        End If
      End If
    Case 4              ' Missing: Discard missing file
      'Trace Tracks(X),"Pass " & 4 & ", missing dupe #" & X & " merging with #" & Y
      If Tracks(X).Kind=3 Then E=LCase(Tracks(X).URL & "") Else E=LCase(Tracks(X).Location & "")
      If Tracks(Y).Kind=3 Then F=LCase(Tracks(Y).URL & "") Else F=LCase(Tracks(Y).Location & "")
      If E="" And F<>"" Then Swap=True
      If (E="" And F="") Or (E<>"" And F<>"") Then MsgBox "Error merging missing files",vbCritical,Title : WScript.Quit
    Case 5                              ' Same song dupes
      If PrefChecked Then
        If ((Tracks(X).Enabled And Not(Tracks(Y).Enabled)) Or (Tracks(Y).Enabled And Not(Tracks(X).Enabled))) Then
          If Tracks(Y).Enabled Then Swap=True     ' When two possibilites, only one checked, keep checked
          ' Trace Null,"Keeping checked file" & nl & Tracks(X).Enabled & nl & Tracks(Y).Enabled & nl & "Swap=" & Swap
        End If
      ElseIf PrefRating Then
        If Tracks(Y).RatingKind=0 And Tracks(Y).Rating>Tracks(X).Rating Then
          Swap=True ' Keep track with best rating
        End If
      Else
        If Tracks(Y).TrackNumber<Tracks(X).TrackNumber Then 
          Swap=True ' Keep track with lowest track number
        End If
      End If  
      ' Trace Tracks(X),"Pass " & 5 & ", same song dupe #" & X & " and #" & Y & nl & "Swap=" & Swap     
    Case Else           ' (0,1) Keep oldest file by date added, applies to physical & logical dupes
      If Tracks(Y).DateAdded<Tracks(X).DateAdded Then Swap=True
  End Select

  If W=3 And PrefAge=0 Then             ' Swap file pointers if needed to preserve date added, best not used if choosing by age?
    ' Trace Null,"X=" & Tracks(X).Location & nl & Tracks(X).DateAdded & nl & "Y=" & Tracks(Y).Location & nl & Tracks(Y).DateAdded & nl & "Swap=" & Swap
    If (Swap=False And Tracks(X).DateAdded>Tracks(Y).DateAdded) Or (Swap=True And Tracks(X).DateAdded<Tracks(Y).DateAdded) Then
      E=Tracks(X).Location & ""
      F=Tracks(Y).Location & ""
      Tracks(X).Location=F
      Tracks(Y).Location=E
      Swap=Not Swap
    End If
  End If

  ' Track I is to be kept, track J discarded
  If Swap Then
    I=Y
    J=X
  Else
    I=X
    J=Y
  End If
  ' Find all playlists track I is a member of
  On Error Resume Next        ' Trap possible error
  Set P=Tracks(I).Playlists
  If Err.Number<>0 Then
    Trace Tracks(I),Err.Description
    Err.Clear
  End If
  On Error Goto 0             ' End of error trap  
  For L=1 To P.Count
    If P.Item(L).Kind=2 And P.Item(L).Smart=False Then
      K=PersistentID(P.Item(L))
      ' If Playlists.Exists(K)=False Then Playlists.Add K,1
      Playlists.Add K,1
    End If
  Next
  ' Add track I to any playlists that J is in
  Set P=Tracks(J).Playlists
  For L=1 To P.Count
    If P.Item(L).Kind=2 And P.Item(L).Smart=False Then
      K=PersistentID(P.Item(L))
      If Playlists.Exists(K)=False Then P.Item(L).AddTrack(Tracks(I))
    End If
  Next
  ' Consolidate played & skipped info
  If Tracks(J).PlayedCount>0 Then
    If Tracks(J).PlayedDate>Tracks(I).PlayedDate Then Tracks(I).PlayedDate=Tracks(J).PlayedDate
    Tracks(I).PlayedCount=Tracks(I).PlayedCount+Tracks(J).PlayedCount
    On Error Resume Next                        ' Trap possible error
    If W<>4 Then Tracks(J).PlayedCount=0        ' In case we don't end up deleting this item, can't alter missing tracks
    If Err.Number<>0 Then
      Trace Tracks(J),Err.Description
    End If
    On Error Goto 0                             ' End of error trap
  End If
  If Tracks(I).Kind=1 And Tracks(J).Kind=1 Then
    If Tracks(J).SkippedCount>0 Then
      If Tracks(J).SkippedDate>Tracks(I).SkippedDate Then Tracks(I).SkippedDate=Tracks(J).SkippedDate
      Tracks(I).SkippedCount=Tracks(I).SkippedCount+Tracks(J).SkippedCount
      If W<>4 Then Tracks(J).SkippedCount=0     ' In case we don't end up deleting this item, can't alter missing tracks
    End If
  End If
  ' Take the highest manual rating, move lower rating to discard if manual
  If Tracks(J).RatingKind=0 And Tracks(J).Rating>Tracks(I).Rating Then
    K=0 : If Tracks(J).RatingKind=0 Then K=Tracks(I).Rating
    Tracks(I).Rating=Tracks(J).Rating
    Tracks(J).Rating=K
  End If
  Paths.Item(A)=I                               ' Make sure key now matches item we are keeping
  If Pseudo Then
    DeathRow.Add PersistentID(Tracks(J)),6      ' Discard the other item as a logical dupe
    C(6)=C(6)+1                                 ' Count pseudo-physicals here
  Else
    DeathRow.Add PersistentID(Tracks(J)),W      ' Discard the other item, with duplicate type
  End If
  ' Make checked status mark tracks to keep or remove, could cause unwanted effects in smart playlists
  If Uncheck Then
    If Tracks(I).Enabled=False Then Tracks(I).Enabled=True
    If Tracks(J).Enabled=True Then Tracks(J).Enabled=False
  End If
End Sub


' Test for media files or subfolders, if none found move remaining files to new path, then delete folder
' Modified 2014-05-04
Sub MoveArt(ByVal OldPath,ByVal NewPath)
  Dim Files,E,F,M,NF,NP,OF,OP
  If FSO.FolderExists(OldPath)=False Then Exit Sub      ' Nothing to do... 
  Set OF=FSO.GetFolder(OldPath)
  If FSO.FolderExists(NewPath)=False Then
    MsgBox "iTunes has changed the path of the last file that was copied from" & nl & OldPath & " to " & nl & NewPath & "." _
      & nl & nl& "Please disable the ''Keep iTunes Media folder organised'' option" & nl & "or choose another target folder.",vbInformation,title
      Quit=True
    Exit Sub
  End If
  Set NF=FSO.GetFolder(NewPath)
  ' Allow for special case when moving files from Artist folder to Arist\Album folder
  If OF.Subfolders.Count=0 OR (OldPath=FSO.GetParentFolderName(NewPath) And NF.SubFolders.Count=0) Then
    Set Files=OF.Files
    If Files.Count>0 Then
      ' There are some files, any media ones?
      M=False
      For Each F in Files
        E=LCase(Right(F.Name,4))
        If Instr(".mp3.mp4.m4a.m4b.m4p.m4v.mov.mpg.mpeg.wav.aif.aiff.mid.ipa.ipg.ite.itlp.m4r.epub.pdf",E) Then M=True : Exit For
      Next
      ' If no media files shift everything else
      If M=False Then
        For Each F in Files
	        ' If target folder already has a Folder.jpg image it is likely to be "fresher" so delete the one from the source folder
	        If LCase(F.Name)="folder.jpg" And FSO.FileExists(NewPath & "\Folder.jpg") Then
	          F.Delete
	        ElseIf LCase(F.Name)="thumbs.db" And FSO.FileExists(NewPath & "\Thumbs.db") Then
	          F.Delete
	        ElseIf LCase(Left(F.Name,8))="albumart" And FSO.FileExists(NewPath & "\" & F.Name) Then
	          F.Delete
	        ElseIf FSO.FileExists(NewPath & "\" & F.Name)=False Then
            F.Move(NewPath & "\")
          Else
            If M=False Then
              M=True
              SH.Explore OldPath
              MsgBox "Unable to move all remaining non-media files from folder" & nl & OldPath & nl & nl & "Please check and tidy if required.",vbInformation,title
            End If
          End If
        Next
      End If
    End If
    If Files.Count=0 Then
      ' The folder is now/was empty of art so remove it unless it contains subfolders (the special case above)
      If OF.Subfolders.Count=0 Then
        RmDir OldPath
      ElseIf OF.Subfolders.Count>1 Then
        SH.Explore NewPath
        MsgBox "There may be artwork for more than one album in the folder" & nl & NewPath & nl & nl & "Please check and tidy if required.",vbInformation,title
      End If
      OP=FSO.GetParentFolderName(OldPath)
      NP=FSO.GetParentFolderName(NewPath)
      ' See if parent folders no longer contain media or subfolders, move art if needed, then delete
      MoveArt OP,NP
      ' Promote or remove promoted art if required
      If NF.ParentFolder.SubFolders.Count=1 Then
        ' Only one album subfolder for artist, promote thumbnail if possible
        If FSO.FileExists(NP & "\Folder.jpg")=False Then
          If FSO.FileExists(NewPath & "\Folder.jpg") And Thumbs Then
            FSO.CopyFile NewPath & "\Folder.jpg",NP & "\"
          End If
        End If
      Else
        ' More than one album subfolder for artist, remove thumbnail in artist folder if it exists and is not read-only
        If FSO.FileExists(NP & "\Folder.jpg") Then
          Set F=FSO.GetFile(NP & "\Folder.jpg")
          If (F.Attributes AND 1)=0 Then F.Delete
	      End If
      End If
    End If
  End If
End Sub


' File location relative to media folder, or three parent folders if not in the media folder, e.g. <Media Kind>\<Artist>\<Album>
' Modified 2015-02-23
Function Relative(L)
  Dim P,S
  S=Replace(L & "","/","\")     ' Fix bad paths
  P=InStrRev(S,"\")       
  If P>0 Then S=Left(S,P-1)     ' Path to parent folder
  If Instr(LCase(S),LCase(Root))=1 Then
    Relative=Mid(S,Len(Root)+1) ' Folder offset from media folder
  Else
    P=Len(S)
    P=InStrRev(S,"\",P-1)       ' Second parent
    P=InStrRev(S,"\",P-1)       ' Third parent
    P=InStrRev(S,"\",P-1)       ' Fourth parent
    Relative=Mid(S,P+1)         ' Up to three parent folders, e.g. Music\Artist\Album, less if path not so deeply nested, and not inside media folder
  End If    
End Function


' Recycled from http://gallery.technet.microsoft.com/scriptcenter/191eb207-3a7e-4dbc-884d-5f4498440574
' Modified to recursively remove any emptied folders. Rewritten to simplify and use global objects/declarations
' Needs FSO,Reg,SH objects. If UseTrash is false delete directly without attempting to recycle.

' Send file or folder to recycle bin, return status
' Modified 2022-06-15
Function Recycle(FilePath)
  Const HKEY_CURRENT_USER=&H80000001 
  Const KeyPath="Software\Microsoft\Windows\CurrentVersion\Explorer" 
  Const KeyName="ShellState" 
  Dim File,FileName,Folder,FolderName,I,Parent,State,Value,Verb,Q,R
  Recycle=False
  If Not(FSO.FileExists(FilePath) Or FSO.FolderExists(FilePath)) Then Exit Function     ' Can't delete something that isn't there
  If UseTrash And Left(FilePath,2)<>"\\" Then										' Can't recycle network files, will cause unwanted confirmation prompt
    ' Make sure recycle bin properties are set to NOT display request for delete confirmation 
    Reg.GetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Get current shell state 
    State=Value(4)	 							' Preserve current option
    Value(4)=39									  ' Set new option 
    Reg.SetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Update shell state
   
    ' Use the Shell to send the file to the recycle bin 
    FileName=FSO.GetFileName(FilePath)
    FolderName=FSO.GetParentFolderName(FilePath)
    Set Folder=SH.NameSpace(FolderName)
    If Folder is Nothing Then
      Set File=Nothing
    Else
      Set File=Folder.ParseName(FileName)
    End If
    If Not File Is Nothing Then  
      'File.InvokeVerb("&Delete")	' Delete file, sending to recycle bin - fails for Vista/Windows 7
      I=File.Verbs.Count          ' Use DoIt instead of InvokeVerb - http://forums.wincustomize.com/322016
      Do
        I=I-1
        Verb=Replace(LCase(File.Verbs.Item(I).Name),"&","")
        ' Add lower case localised words for delete here, separated by |
        If Instr("|delete|löschen|verwijderen|","|" & Verb & "|") Then
          On Error Resume Next          ' Trap potential error
          File.Verbs.Item(I).DoIt()     ' Possible error generating line
          ' Err.Raise 1,Title,"Test"    ' Test error handler
          If Err.Number<>0 Then         ' Handle any error
            Q="An error occurred recycling a file. (Recycle #1)" & nl
            Q=Q & nl & "Script:" & tab & Wrap(WScript.ScriptFullName,50,"\",1)
            Q=Q & nl & "Error:" & tab & Err.Description
            Q=Q & nl & "Number:" & tab & "&" & Right("0000000" & Hex(Err.Number),8)
            Q=Q & nl & "Source:" & tab & Err.Source
            Q=Q & nl & "Path:" & tab & Wrap(FilePath,50,"\",1)
            Q=Q & nl & nl & "Press Cancel to abort this run."
            R=MsgBox(Q,vbCritical+vbOKCancel,Title)
            If R=vbCancel Then Quit=True
            Err.Clear
          End If                        'End of conditional block
          On Error Goto 0               'Reset default error handler
          Exit Do
        End If
      Loop Until I=0
      If I=0 Then             ' Assume non-English settings and predict which verb is delete
        I=File.Verbs.Count-3  ' Because delete should be the third item from the end of the list, ignoring separators
        Verb=Replace(LCase(File.Verbs.Item(I).Name),"&","")
        Trace Null,"Verb for delete not recognized" & nl & nl & "I=" & I & ", Count=" & File.Verbs.Count & ", Verb=" & Verb
        Tracing=False
        File.Verbs.Item(I).DoIt()
      End If
			Value(4)=State					' Restore the user's property settings for the Recycle Bin 
			Reg.SetBinaryValue HKEY_CURRENT_USER,KeyPath,KeyName,Value			' Update shell state
    End If
  Else    ' Delete via FSO instead of Shell
    FolderName=FSO.GetParentFolderName(FilePath)
    On Error Resume Next                ' iTunes operation with error handling
    FSO.DeleteFile FilePath,True        ' Possible error generating code
    ' Err.Raise 1,Title,"Test"          ' Test error handler
    If Err.Number<>0 Then               ' Handle any error
      Q="An error occurred deleting a file. (Recycle #2)" & nl
      Q=Q & nl & "Script:" & tab & Wrap(WScript.ScriptFullName,50,"\",1)
      Q=Q & nl & "Error:" & tab & Err.Description
      Q=Q & nl & "Number:" & tab & "&" & Right("0000000" & Hex(Err.Number),8)
      Q=Q & nl & "Source:" & tab & Err.Source
      Q=Q & nl & "Path:" & tab & Wrap(FilePath,50,"\",1)
      Q=Q & nl & nl & "Press Cancel to abort this run."
      R=MsgBox(Q,vbCritical+vbOKCancel,Title)
      If R=vbCancel Then Quit=True
      Err.Clear
    End If                              'End of conditional block
    On Error Goto 0                     'Reset default error handler
  End If
  If FSO.FileExists(FilePath) Then
    MsgBox "There was a problem deleting the file: (Recycle #3)" & nl & FilePath,vbCritical,Title
  Else
    Recycle=True
    ' Delete folder using FileSystem if now empty, repeat for parent folders
    Set Folder=FSO.GetFolder(FolderName)
    While Folder.Files.Count=0 And Folder.SubFolders.Count=0
      Set Parent=Folder.ParentFolder
      Folder.Delete
      Set Folder=Parent
    Wend
  End If
End Function


' Use RegEx for string replacement, input, pattern, replacement text, allow multiple matches, case sensitve
' Created 2024-09-23
Function RegRep(I,P,R)
  Dim Reg
  Set Reg=New RegExp
  Reg.Pattern=P
  Reg.Global=True
  RegRep=Reg.Replace(I,R)
End Function


' Output results
' Modified 2022-02-01
Sub Results
  If Not Outro Then Exit Sub
  Dim A,I,L,T
  A=0
  For I=0 To 6 : A=A+B(I) : Next        ' All dupes
  If Quit Then T="Script aborted!" & nl & nl Else T=""
  If D=0 Then T=T & "No" Else T=T & GroupDig(D) 
  T=T & " duplicate" & Plural(D,"s were"," was") & " found from" & nl
  T=T & GroupDig(P) & " items processed"
  If A>0 Then
    T=T & " of which " & nl & GroupDig(A)
    If Testing Then T=T & " would have been" Else T=T & Plural(A," were"," was")
    If Archive Then T=T & " archived" Else T=T & " cleaned"
    T=T & " as follows:"
    L=""
    If B(0)>0 Then L=PrettyList(L,GroupDig(B(0)) & " playlist duplicate" & Plural(B(0),"s",""))
    If B(1)>0 Then L=PrettyList(L,GroupDig(B(1)) & " logical duplicate" & Plural(B(1),"s",""))
    If B(6)>0 Then L=PrettyList(L,GroupDig(B(6)) & " pseudo-physical duplicate" & Plural(B(6),"s",""))
    If B(2)>0 Then L=PrettyList(L,GroupDig(B(2)) & " physical duplicate" & Plural(B(2),"s",""))
    If B(3)>0 Then
      If PrefKind Then
        L=PrettyList(L,GroupDig(B(3)) & " alternate duplicate" & Plural(B(3),"s",""))
      ElseIf PrefAge Then
        If PrefAge<3 Then
          L=PrettyList(L,GroupDig(B(3)) & " newer duplicate" & Plural(B(3),"s",""))
        Else
          L=PrettyList(L,GroupDig(B(3)) & " older duplicate" & Plural(B(3),"s",""))
        End If
      ElseIf KeepLarge Then 
        L=PrettyList(L,GroupDig(B(3)) & " smaller duplicate" & Plural(B(3),"s",""))
      Else
        L=PrettyList(L,GroupDig(B(3)) & " larger duplicate" & Plural(B(3),"s",""))
      End If
    End If
    If B(4)>0 Then L=PrettyList(L,GroupDig(B(4)) & " missing duplicate" & Plural(B(4),"s",""))
    If B(5)>0 Then L=PrettyList(L,GroupDig(B(5)) & " same song duplicate" & Plural(B(5),"s",""))
    T=T & nl & nl & L
  End If
  L=""
  S=D-A                         ' Correct skip total for aborted script
  IF S>0 Then L=PrettyList(L,GroupDig(S) & Plural(S," items were"," item was") & " skipped")
  If M>0 Then L=PrettyList(L,GroupDig(M) & Plural(M," items were"," item was") & " missing")
  IF L<>"" Then T=T & nl & nl & L
  If Timing Then T=T & nl & nl & "Processing time: " & FormatTime(Clock)
  MsgBox T,vbInformation,Title
End Sub


' Remove folder even if marked as Read only
' Modified 2011-09-17
Sub RmDir(F)
  On Error Resume Next
  FSO.DeleteFolder F, True
  If FSO.FolderExists(F) Then MsgBox "There was a problem deleing the folder" & nl & F & nl & nl & "Please delete by hand, probably after rebooting.",0,Title
End Sub


' Copy tracks into a list of persistent objects
' Modified 2022-02-05
Sub SetTracks
  Dim I,T
  Set SafeList=CreateObject("Scripting.Dictionary")
  For I=1 To Count
    Set T=Tracks.Item(I)
    SafeList.Add I,PersistentID(T)
  Next
End Sub


' Custom trace messages for troubleshooting, T is the current track if needed, Null otherwise 
' Modified 2022-02-05
Sub Trace(T,M)
  If Tracing Then
    Dim R,Q
    If IsNull(T) Then
      Q=M & nl & nl
    Else
      Q="Checking: " & Info(T) & nl & nl & M & nl & nl
    End If    
    Q=Q & "Yes" & tab & ": Continue tracing" & nl
    Q=Q & "No" & tab & ": Skip further tracing" & nl
    Q=Q & "Cancel" & tab & ": Abort script"
    R=MsgBox(Q,vbYesNoCancel,Title)
    If R=vbCancel Then Quit=True : Results : WScript.Quit
    If R=vbNo Then Tracing=False
  End If
End Sub


' Truncate brackets from song title for comparison and/or trim to MaxChars
' Modified 2024-09-23
Function Truncate(S)
  Dim P,T
  T=S
  If IgnoreBraces Then
    T=RegRep(T,"\(.*?\)","")    ' Remove any () brackets
    T=RegRep(T,"\[.*?\]","")    ' Remove any [] brackets
    T=RegRep(T,"\s+"," ")       ' Remove any repeat spaces
    T=Trim(T)                   ' And trim
    If Len(T)<5 Then T=S        ' Arbitrary min. characters to avoid empty/silly results, change if needed, or use MaxChars
  End If
  If MaxChars>0 And Len(T)>MaxChars Then T=Trim(Left(T,MaxChars))
  Truncate=T
End Function


' Replace invalid filename characters: \ / : * ? " < > | and also ;
' Replace leading space or period, strip trailing spaces, trailing periods allowed
' Limit to 40 characters inclusive of extenion. No tailing period for folder name 
' Modified 2011-09-17
Function ValidiTunes(N,E)
  N=Left(N,40-Len(E))
  N=Replace(N,"\","_")
  N=Replace(N,"/","_")
  N=Replace(N,":","_")
  N=Replace(N,"*","_")
  N=Replace(N,"?","_")
  N=Replace(N,"""","_")
  N=Replace(N,"<","_")
  N=Replace(N,">","_")
  N=Replace(N,"|","_")
  N=Replace(N,";","_")
  Do While Right(N,1)=" "
    N=Left(N,Len(N)-1)
  Loop
  If Left(N,1)=" " Or Left(N,1)="." Then N="_" & Mid(N,2)
  If E="" And Right(N,1)="." Then N=Left(N,Len(N)-1) & "_"
  ValidiTunes=N
End Function
  

' ============================================
' Reusable Library Routines for iTunes Scripts
' ============================================
' Modified 2015-01-24


' Get extension from file path
' Modified 2025-03-19
Function Ext(P)
  Ext=""
	On Error Resume Next            			' Trap possible error
	Ext=LCase(Mid(P,InStrRev(P,".")))     ' Command that might error
	If Err.Number<>0 Then
		Trace Null,"Error with path " & P & "."
		Err.Clear
	End If
	On Error Goto 0                 			' End of error trap
End Function


' Format time interval from x.xxx seconds to hh:mm:ss
' Modified 2011-11-07
Function FormatTime(T)
  If T<0 Then T=T+86400         ' Watch for timer running over midnight
  If T<2 Then
    FormatTime=FormatNumber(T,3) & " seconds"
  ElseIf T<10 Then
    FormatTime=FormatNumber(T,2) & " seconds"
  ElseIf T<60 Then
    FormatTime=Int(T) & " seconds"
  Else
    Dim H,M,S
    S=T Mod 60
    M=(T\60) Mod 60             ' \ = Div operator for integer division
    'S=Right("0" & (T Mod 60),2)
    'M=Right("0" & ((T\60) Mod 60),2)  ' \ = Div operator for integer division
    H=T\3600
    If H>0 Then
      FormatTime=H & Plural(H," hours "," hour ") & M & Plural(M," mins"," min")
      'FormatTime=H & ":" & M & ":" & S
    Else
      FormatTime=M & Plural(M," mins "," min ") & S & Plural(S," secs"," sec")
      'FormatTime=M & " :" & S
      'If Left(FormatTime,1)="0" Then FormatTime=Mid(FormatTime,2)
    End If
  End If
End Function


' Initialise track selections, quit script if track selection is out of bounds or user aborts
' Modified 2014-05-05
Sub GetTracks
  Dim Q,R
  ' Initialise global variables
  nl=vbCrLf : tab=Chr(9) : Quit=False
  D=0 : M=0 : P=0 : S=0 : U=0 : V=0
  ' Initialise global objects
  Set IDs=CreateObject("Scripting.Dictionary")
  Set iTunes=CreateObject("iTunes.Application")
  Set Tracks=iTunes.SelectedTracks      ' Get current selection
  If iTunes.BrowserWindow.SelectedPlaylist.Source.Kind<>1 And Source="" Then Source="Library" : Named=True      ' Ensure section is from the library source
  'If iTunes.BrowserWindow.SelectedPlaylist.Name="Ringtones" And Source="" Then Source="Library" : Named=True    ' and not ringtones (which cannot be processed as tracks???)
  If iTunes.BrowserWindow.SelectedPlaylist.Name="Radio" And Source="" Then Source="Library" : Named=True        ' or radio stations (which cannot be processed as tracks)
  If iTunes.BrowserWindow.SelectedPlaylist.Name=Playlist And Source="" Then Source="Library" : Named=True       ' or a playlist that will be regenerated by this script
  If Named Or Tracks Is Nothing Then    ' or use a named playlist
    If Source<>"" Then Named=True
    If Source="Library" Then            ' Get library playlist...
      Set Tracks=iTunes.LibraryPlaylist.Tracks
    Else                                ' or named playlist
      On Error Resume Next              ' Attempt to fall back to current selection for non-existent source
      Set Tracks=iTunes.LibrarySource.Playlists.ItemByName(Source).Tracks
      On Error Goto 0
      If Tracks is Nothing Then         ' Fall back
        Named=False
        Source=iTunes.BrowserWindow.SelectedPlaylist.Name
        Set Tracks=iTunes.SelectedTracks
        If Tracks is Nothing Then
          Set Tracks=iTunes.BrowserWindow.SelectedPlaylist.Tracks
        End If
      End If
    End If
  End If  
  If Named And Tracks.Count=0 Then      ' Quit if no tracks in named source
    If Intro Then MsgBox "The playlist " & Source & " is empty, there is nothing to do.",vbExclamation,Title
    WScript.Quit
  End If
  If Tracks.Count=0 Then Set Tracks=iTunes.LibraryPlaylist.Tracks
  If Tracks.Count=0 Then                ' Can't select ringtones as tracks?
    MsgBox "This script cannot process " & iTunes.BrowserWindow.SelectedPlaylist.Name & ".",vbExclamation,Title
    WScript.Quit
  End If
  ' Check there is a suitable number of suitable tracks to work with
  Count=Tracks.Count
  If Count<Min Or (Count>Max And Max>0) Then
    If Max=0 Then
      MsgBox "Please select " & Min & " or more tracks in iTunes before calling this script!",0,Title
      WScript.Quit
    Else
      MsgBox "Please select between " & Min & " and " & Max & " tracks in iTunes before calling this script!",0,Title
      WScript.Quit
    End If
  End If
  ' Check if the user wants to proceed and how
  Q=Summary
  If Q<>"" Then Q=Q & nl & nl
  If Warn>0 And Count>Warn Then
    Intro=True
    Q=Q & "WARNING!" & nl & "Are you sure you want to process " & GroupDig(Count) & " tracks"
    If Named Then Q=Q & nl
  Else
    Q=Q & "Process " & GroupDig(Count) & " track" & Plural(Count,"s","")
  End If
  If Named Then Q=Q & " from the " & Source & " playlist"
  Q=Q & "?"
  If Intro Or (Prog And UAC) Then
    If Check Then
      Q=Q & nl & nl 
      Q=Q & "Yes" & tab & ": Process track" & Plural(Count,"s","") & " automatically" & nl
      Q=Q & "No" & tab & ": Preview & confirm each action" & nl
      Q=Q & "Cancel" & tab & ": Abort script"
    End If
    If Kimo Then Q=Q & nl & nl & "NB: Disable ''Keep iTunes Media folder organised'' preference before use."
    If Prog And UAC Then
      Q=Q & nl & nl & "NB: Use the EnableLUA script to allow the progress bar to function "
      Q=Q & "or change the declaration ""Prog=True"" to ""Prog=False"" to hide this message. "
      Prog=False
    End If
    If Check Then
      R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title)
    Else
      R=MsgBox(Q,vbOKCancel+vbQuestion,Title)
    End If
    If R=vbCancel Then WScript.Quit
    If R=vbYes or R=vbOK Then
      Check=False
    Else
      Check=True
    End If
  End If 
  If Check Then Prog=False      ' Suppress progress bar if prompting for user input
End Sub


' Group digits and separate with commas
' Modified 2014-04-29
Function GroupDig(N)
  GroupDig=FormatNumber(N,0,-1,0,-1)
End Function


' Return the persistent object representing the track from its ID as a string
' Modified 2014-09-26 - CLng works better than Eval 
Function ObjectFromID(ID)
  Set ObjectFromID=iTunes.LibraryPlaylist.Tracks.ItemByPersistentID(CLng("&H" & Left(ID,8)),CLng("&H" & Right(ID,8)))
End Function


' Create a string representing the 64 bit persistent ID of an iTunes object
' Modified 2012-08-24
Function PersistentID(T)
  PersistentID=Right("0000000" & Hex(iTunes.ITObjectPersistentIDHigh(T)),8) & "-" & Right("0000000" & Hex(iTunes.ITObjectPersistentIDLow(T)),8)
End Function


' Return the persistent object representing the track
' Keeps hold of an object that might vanish from a smart playlist as it is updated
' Modified 2015-01-24
Function PersistentObject(T)
  Dim E,L
  Set PersistentObject=T
  On Error Resume Next  ' Trap possible error
  If Instr(T.KindAsString,"audio stream") Then
    L=T.URL 
  ElseIf T.Kind=5 Then
    L="iCloud/Shared"
  Else
    L=T.Location
  End If
  If Err.Number<>0 Then
    Trace T,"Error reading location property from object."
  ElseIf L<>"" Then
    E=Ext(L)
    If Instr(".ipa.ipg.m4r",E)=0 Then   ' Method below fails for apps, games & ringtones
      Set PersistentObject=iTunes.LibraryPlaylist.Tracks.ItemByPersistentID(iTunes.ITObjectPersistentIDHigh(T),iTunes.ITObjectPersistentIDLow(T))
    End If  
  End If  
End Function


' Return relevant string depending on whether value is plural or singular
' Modified 2011-10-04
Function Plural(V,P,S)
  If V=1 Then Plural=S Else Plural=P
End Function


' Format a list of values for output
' Modified 2012-08-25
Function PrettyList(L,N)
  If L="" Then
    PrettyList=N & "."
  Else
    PrettyList=Replace(Left(L,Len(L)-1)," and" & nl,"," & nl) & " and" & nl & N & "."
  End If
End Function


' Loop through track selection processing suitable items
' Modified 2015-01-06
Sub ProcessTracks
  Dim C,I,N,Q,R,T
  Dim First,Last,Steps
  If IsEmpty(Rev) Then Rev=True
  If Rev Then
    First=Count : Last=1 : Steps=-1
  Else
    First=1 : Last=Count : Steps=1
  End If
  N=0
  If Prog Then                  ' Create ProgessBar
    Set PB=New ProgBar
    PB.SetTitle Title
    PB.Show
  End If
  Clock=0 : StartTimer
  For I=First To Last Step Steps        ' Usually work backwards in case edit removes item from selection
    N=N+1                 
    If Prog Then
      PB.SetStatus Status(N)
      PB.Progress N-1,Count
    End If
    Set T=Tracks.Item(I)
    If T.Kind=1 Then            ' Ignore tracks which can't change
      Set T=PersistentObject(T) ' Attach to object in library playlist
      If Prog Then PB.SetInfo Info(T)
      If Updateable(T) Then     ' Ignore tracks which won't change
        If Check Then           ' Track by track confirmation
          Q=Prompt(T)
          StopTimer             ' Don't time user inputs 
          R=MsgBox(Q,vbYesNoCancel+vbQuestion,Title & " - " & GroupDig(N) & " of " & GroupDig(Count))
          StartTimer
          Select Case R
          Case vbYes
            C=True
          Case vbNo
            C=False
            S=S+1               ' Increment skipped tracks
          Case Else
            Quit=True
            Exit For
          End Select          
        Else
          C=True
        End If
        If C Then               ' We have a valid track, now do something with it
          Action T
        End If
      End If
    End If 
    P=P+1                       ' Increment processed tracks
    ' WScript.Sleep 500         ' Slow down progress bar when testing
    If Quit Then Exit For       ' Abort loop on user request
  Next
  StopTimer
  If Prog And Not Quit Then
    PB.Progress Count,Count
    WScript.Sleep 250
  End If
  If Prog Then PB.Close
End Sub


' Output report
' Modified 2014-04-29
Sub Report
  If Not Outro Then Exit Sub
  Dim L,T
  L=""
  If Quit Then T="Script aborted!" & nl & nl Else T=""
  T=T & GroupDig(P) & " track" & Plural(P,"s","")
  If P<Count Then T=T & " of " & GroupDig(Count)
  T=T & Plural(P," were"," was") & " processed of which " & nl
  If D>0 Then L=PrettyList(L,GroupDig(D) & Plural(D," were duplicates"," was a duplicate") & " in the list")
  If V>0 Then L=PrettyList(L,GroupDig(V) & " did not need updating")
  If U>0 Or V=0 Then L=PrettyList(L,GroupDig(U) & Plural(U," were"," was") & " updated")
  If S>0 Then L=PrettyList(L,GroupDig(S) & Plural(S," were"," was") & " skipped")
  If M>0 Then L=PrettyList(L,GroupDig(M) & Plural(M," were"," was") & " missing")
  T=T & L
  If Timing Then 
    T=T & nl & nl
    If Check Then T=T & "Processing" Else T=T & "Running"
    T=T & " time: " & FormatTime(Clock)
  End If
  MsgBox T,vbInformation,Title
End Sub


' Return iTunes like sort name
' Modified 2011-01-27
Function SortName(N)
  Dim L
  N=LTrim(N)
  L=LCase(N)
  SortName=N
  If Left(L,2)="a " Then SortName=Mid(N,3)
  If Left(L,3)="an " Then SortName=Mid(N,4)
  If Left(L,3)="""a " Then SortName=Mid(N,4)
  If Left(L,4)="the " Then SortName=Mid(N,5)
  If Left(L,4)="""an " Then SortName=Mid(N,5)
  If Left(L,5)="""the " Then SortName=Mid(N,6)
End Function


' Start timing event
' Modified 2011-10-08
Sub StartEvent
  T2=Timer
End Sub


' Start timing session
' Modified 2011-10-08
Sub StartTimer
  T1=Timer
End Sub


' Stop timing event and display elapsed time in debug section of Progress Bar
' Modified 2011-11-07
Sub StopEvent
  If Prog Then
    T2=Timer-T2
    If T2<0 Then T2=T2+86400            ' Watch for timer running over midnight
    If Debug Then PB.SetDebug "<br>Last iTunes call took " & FormatTime(T2) 
  End If  
End Sub


' Stop timing session and add elapased time to running clock
' Modified 2011-10-08
Sub StopTimer
  Clock=Clock+Timer-T1
  If Clock<0 Then Clock=Clock+86400     ' Watch for timer running over midnight
End Sub


' Detect if User Access Control is enabled, UAC (or rather LUA) prevents use of progress bar
' Modified 2011-10-18
Function UAC
  Const HKEY_LOCAL_MACHINE=&H80000002
  Const KeyPath="Software\Microsoft\Windows\CurrentVersion\Policies\System"
  Const KeyName="EnableLUA"
  Dim Reg,Value
  Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	  ' Use . for local computer, otherwise could be computer name or IP address
  Reg.GetDWORDValue HKEY_LOCAL_MACHINE,KeyPath,KeyName,Value	  ' Get current property
  If IsNull(Value) Then UAC=False Else UAC=(Value<>0)
End Function


' Wrap & tab long strings, break string S on first separator C found at or before character W adding T tabs to each new line
' Modified 2014-05-29
Function Wrap(S,W,C,T)
  Dim P,Q
  P=InstrRev(S," ",W)
  Q=InstrRev(S,"\",W)
  If Q>P Then P=Q
  If P Then
    Wrap=Left(S,P) & nl & String(T,tab) & Wrap(Mid(S,P+1),W,C,T)
  Else
    Wrap=S
  End If
End Function


' ==================
' Progress Bar Class
' ==================

' Progress/activity bar for vbScript implemented via IE automation
' Can optionally rebuild itself if closed or abort the calling script
' Modified 2014-05-04
Class ProgBar
  Public Cells,Height,Width,Respawn,Title,Version
  Private Active,Blank,Dbg,Filled(),FSO,IE,Info,NextOn,NextOff,Status,SHeight,SWidth,Temp

' User has closed progress bar, abort or respwan?
' Modified 2011-10-09
  Public Sub Cancel()
    If Respawn And Active Then
      Active=False
      If Respawn=1 Then
        Show                    ' Ignore user's attempt to close and respawn
      Else
        Dim R
        StopTimer               ' Don't time user inputs 
        R=MsgBox("Abort Script?",vbExclamation+vbYesNoCancel,Title)
        StartTimer
        If R=vbYes Then
          On Error Resume Next
          CleanUp
          Respawn=False
          Quit=True             ' Global flag allows main program to complete current task before exiting
        Else
          Show                  ' Recreate box if closed
        End If  
      End If        
    End If
  End Sub

' Delete temporary html file  
' Modified 2011-10-04
  Private Sub CleanUp()
    FSO.DeleteFile Temp         ' Delete temporary file
  End Sub
  
' Close progress bar and tidy up
' Modified 2011-10-04
  Public Sub Close()
    On Error Resume Next        ' Ignore errors caused by closed object
    If Active Then
      Active=False              ' Ignores second call as IE object is destroyed
      IE.Quit                   ' Remove the progess bar
      CleanUp
    End If    
 End Sub
 
' Initialize object properties
' Modified 2012-09-05
  Private Sub Class_Initialize()
    Dim I,Items,strComputer,WMI
    ' Get width & height of screen for centering ProgressBar
    strComputer="."
    Set WMI=GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set Items=WMI.ExecQuery("Select * from Win32_OperatingSystem",,48)
    'Get the OS version number (first two)
    For Each I in Items
      Version=Left(I.Version,3)
    Next
    Set Items=WMI.ExecQuery ("Select * From Win32_DisplayConfiguration")
    For Each I in Items
      SHeight=I.PelsHeight
      SWidth=I.PelsWidth
    Next
    If Debug Then
      Height=160                ' Height of containing div
    Else
      Height=120                ' Reduce height if no debug area
    End If
    Width=300                   ' Width of containing div
    Respawn=True                ' ProgressBar will attempt to resurect if closed
    Blank=String(50,160)        ' Blanks out "Internet Explorer" from title
    Cells=25                    ' No. of units in ProgressBar, resize window if using more cells
    ReDim Filled(Cells)         ' Array holds current state of each cell
    For I=0 To Cells-1
      Filled(I)=False
    Next
    NextOn=0                    ' Next cell to be filled if busy cycling
    NextOff=Cells-5             ' Next cell to be cleared if busy cycling
    Dbg="&nbsp;"                ' Initital value for debug text
    Info="&nbsp;"               ' Initital value for info text
    Status="&nbsp;"             ' Initital value for status text
    Title="Progress Bar"        ' Initital value for title text
    Set FSO=CreateObject("Scripting.FileSystemObject")          ' File System Object
    Temp=FSO.GetSpecialFolder(2) & "\ProgBar.htm"               ' Path to Temp file
  End Sub

' Tidy up if progress bar object is destroyed
' Modified 2011-10-04
  Private Sub Class_Terminate()
    Close
  End Sub
 
' Display the bar filled in proportion X of Y
' Modified 2011-10-18
  Public Sub Progress(X,Y)
    Dim F,I,L,S,Z
    If X<0 Or X>Y Or Y<=0 Then
      MsgBox "Invalid call to ProgessBar.Progress, variables out of range!",vbExclamation,Title
      Exit Sub
    End If
    Z=Int(X/Y*(Cells))
    If Z=NextOn Then Exit Sub
    If Z=NextOn+1 Then
      Step False
    Else
      If Z>NextOn Then
        F=0 : L=Cells-1 : S=1
      Else
        F=Cells-1 : L=0 : S=-1
      End If
      For I=F To L Step S
        If I>=Z Then
          SetCell I,False
        Else
          SetCell I,True
        End If
      Next
      NextOn=Z
    End If
  End Sub

' Clear progress bar ready for reuse  
' Modified 2011-10-16
  Public Sub Reset
    Dim C
    For C=Cells-1 To 0 Step -1
      IE.Document.All.Item("P",C).classname="empty"
      Filled(C)=False
    Next
    NextOn=0
    NextOff=Cells-5   
  End Sub
  
' Directly set or clear a cell
' Modified 2011-10-16
  Public Sub SetCell(C,F)
    On Error Resume Next        ' Ignore errors caused by closed object
    If F And Not Filled(C) Then
      Filled(C)=True
      IE.Document.All.Item("P",C).classname="filled"
    ElseIf Not F And Filled(C) Then
      Filled(C)=False
      IE.Document.All.Item("P",C).classname="empty"
    End If
  End Sub 
 
' Set text in the Dbg area
' Modified 2011-10-04
  Public Sub SetDebug(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Dbg=T
    IE.Document.GetElementById("Debug").InnerHTML=T
  End Sub

' Set text in the info area
' Modified 2011-10-04
  Public Sub SetInfo(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Info=T
    IE.Document.GetElementById("Info").InnerHTML=T
  End Sub

' Set text in the status area
' Modified 2011-10-04
  Public Sub SetStatus(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Status=T
    IE.Document.GetElementById("Status").InnerHTML=T
  End Sub

' Set title text
' Modified 2011-10-04
  Public Sub SetTitle(T)
    On Error Resume Next        ' Ignore errors caused by closed object
    Title=T
    IE.Document.Title=T & Blank
  End Sub
  
' Create and display the progress bar  
' Modified 2014-05-04
  Public Sub Show()
    Const HKEY_CURRENT_USER=&H80000001
    Const KeyPath="Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_LOCALMACHINE_LOCKDOWN"
    Const KeyName="iexplore.exe"
    Dim File,I,Reg,State,Value
    Set Reg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 	' Use . for local computer, otherwise could be computer name or IP address
    'On Error Resume Next        ' Ignore possible errors
    ' Make sure IE is set to allow local content, at least while we get the Progress Bar displayed
    Reg.GetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Get current property
    State=Value	 							  ' Preserve current option
    Value=0		    							' Set new option 
    Reg.SetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	' Update property
    'If Version<>"5.1" Then Prog=False : Exit Sub      ' Need to test for Vista/Windows 7 with UAC
    Set IE=WScript.CreateObject("InternetExplorer.Application","Event_")
    Set File=FSO.CreateTextFile(Temp, True)
    With File
      .WriteLine "<!doctype html>"
      '.WriteLine "<!-- saved from url=(0014)about:internet -->"
      .WriteLine "<!-- saved from url=(0016)http://localhost -->"      ' New "Mark of the web"
      .WriteLine "<html><head><title>" & Title & Blank & "</title>"
      .WriteLine "<style type='text/css'>"
      .WriteLine ".border {border: 5px solid #DBD7C7;}"
      .WriteLine ".debug {font-family: Tahoma; font-size: 8.5pt;}"
      .WriteLine ".empty {border: 2px solid #FFFFFF; background-color: #FFFFFF;}"
      .WriteLine ".filled {border: 2px solid #FFFFFF; background-color: #00FF00;}"
      .WriteLine ".info {font-family: Tahoma; font-size: 8.5pt;}"
      .WriteLine ".status {font-family: Tahoma; font-size: 10pt;}"
      .WriteLine "</style>"
      .WriteLine "</head>"
      .WriteLine "<body scroll='no' style='background-color: #EBE7D7'>"
      .WriteLine "<div style='display:block; height:" & Height & "px; width:" & Width & "px; overflow:hidden;'>"
      .WriteLine "<table border-width='0' cellpadding='2' width='" & Width & "px'><tr>"
      .WriteLine "<td id='Status' class='status'>" & Status & "</td></tr></table>"
      .WriteLine "<table class='border' cellpadding='0' cellspacing='0' width='" & Width & "px'><tr>"
      ' Write out cells
      For I=0 To Cells-1
	      If Filled(I) Then
          .WriteLine "<td id='p' class='filled'>&nbsp;</td>"
        Else
          .WriteLine "<td id='p' class='empty'>&nbsp;</td>"
        End If
      Next
	    .WriteLine "</tr></table>"
      .WriteLine "<table border-width='0' cellpadding='2' width='" & Width & "px'><tr><td>"
      .WriteLine "<span id='Info' class='info'>" & Info & "</span><br>"
      .WriteLine "<span id='Debug' class='debug'>" & Dbg & "</span></td></tr></table>"
      .WriteLine "</div></body></html>"
    End With
    ' Create IE automation object with generated HTML
    With IE
      .width=Width+35           ' Increase if using more cells
      .height=Height+60         ' Increase to allow more info/debug text
      If Version>"5.1" Then     ' Allow for bigger border in Vista/Widows 7
        .width=.width+10
        .height=.height+10
      End If        
      .left=(SWidth-.width)/2
      .top=(SHeight-.height)/2
      .navigate "file://" & Temp
      '.navigate "http://samsoft.org.uk/progbar.htm"
      .addressbar=False
      .resizable=False
      .toolbar=False
      On Error Resume Next      
      .menubar=False            ' Causes error in Windows 8 ? 
      .statusbar=False          ' Causes error in Windows 7 or IE 9
      On Error Goto 0
      .visible=True             ' Causes error if UAC is active
    End With
    Active=True
    ' Restore the user's property settings for the registry key
    Value=State		    					' Restore option
    Reg.SetDWORDValue HKEY_CURRENT_USER,KeyPath,KeyName,Value	  ' Update property 
    Exit Sub
  End Sub
 
' Increment progress bar, optionally clearing a previous cell if working as an activity bar
' Modified 2011-10-05
  Public Sub Step(Clear)
    SetCell NextOn,True : NextOn=(NextOn+1) Mod Cells
    If Clear Then SetCell NextOff,False : NextOff=(NextOff+1) Mod Cells
  End Sub

' Self-timed shutdown
' Modified 2011-10-05 
  Public Sub TimeOut(S)
    Dim I
    Respawn=False                ' Allow uninteruppted exit during countdown
    For I=S To 2 Step -1
      SetDebug "<br>Closing in " & I & " seconds" & String(I,".")
      WScript.sleep 1000
    Next
      SetDebug "<br>Closing in 1 second."
      WScript.sleep 1000
    Close
  End Sub 
    
End Class


' Fires if progress bar window is closed, can't seem to wrap up the handler in the class
' Modified 2011-10-04
Sub Event_OnQuit()
  PB.Cancel
End Sub


' ==============
' End of listing
' ==============