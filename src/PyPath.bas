Attribute VB_Name = "PyPath"
'@Folder "PyPathProject.src"
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" ( _
      ByVal lpszShortPath As String, _
      ByVal lpszLongPath As String, _
      ByVal cchBuffer As Long) As Long
  Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
#Else
  Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" ( _
      ByVal lpszShortPath As String, _
      ByVal lpszLongPath As String, _
      ByVal cchBuffer As Long) As Long
  Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
#End If

Private Const INVALID_FILE_ATTRIBUTES As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Public Const CUR_DIR As String = "."
Public Const PAR_DIR As String = ".."
Public Const EX_SEP As String = "."
Public Const SEP As String = "\"
Public Const ALT_SEP As String = "/"
Public Const PATH_SEP As String = ";"
Public Const DEV_NULL As String = "nul"

'@Description "Return a normalized absolutized version of the pathname path."
Public Function AbsPath(ByVal Path As String) As String
Attribute AbsPath.VB_Description = "Return a normalized absolutized version of the pathname path."
    Dim CWD As String
    CWD = FileSystem.CurDir()

    If Not PyPath.IsAbs(Path) Then
        Path = PyPath.Join(CWD, Path)
    End If

    AbsPath = PyPath.NormPath(Path)
End Function

'@Description "Return the base name of pathname path. This is the second element of the pair returned by passing path to the function split()."
Public Function Basename(ByVal Path As String) As String
Attribute Basename.VB_Description = "Return the base name of pathname path. This is the second element of the pair returned by passing path to the function split()."
    Basename = PyPath.Split(Path)(1)
End Function

'@Description "Return the longest common sub-path of each pathname in the iterable paths."
Public Function CommonPath(ByRef Paths As Variant) As String
Attribute CommonPath.VB_Description = "Return the longest common sub-path of each pathname in the iterable paths."
    Dim DriveSplits() As Variant
    ReDim DriveSplits(UBound(Paths))
    Dim i As Long
    Dim p As Variant
    For Each p In Paths
        DriveSplits(i) = PyPath.SplitRoot(PyPath.NormCase(p))
        i = i + 1
    Next

    Dim SplitPaths() As Variant
    ReDim SplitPaths(UBound(DriveSplits))
    i = 0
    For Each p In DriveSplits
        SplitPaths(i) = Strings.Split(p(2), SEP)
        i = i + 1
    Next

    ' Check that all drive letters or UNC paths match. The check is made only
    ' now otherwise type errors for mixing strings and bytes would not be
    ' caught.
    Dim DrivesSet As Object
    Set DrivesSet = CreateObject("Scripting.Dictionary")
    For Each p In DriveSplits
        DrivesSet(p(0)) = True
    Next

    If UBound(DrivesSet.Keys()) <> 0 Then
        Information.Err().Raise _
            Number:=5, _
            Source:="CommonPath", _
            Description:="Paths don't have the same drive"
    End If

    Dim Splited() As String
    Splited = PyPath.SplitRoot(Strings.Replace(Paths(0), ALT_SEP, SEP))
    Dim Drive As String
    Drive = Splited(0)
    Dim Root As String
    Root = Splited(1)
    Dim Path As String
    Path = Splited(2)

    Dim RootsSet As Object
    Set RootsSet = CreateObject("Scripting.Dictionary")
    Dim DriveSplit As Variant
    For Each DriveSplit In DriveSplits
        RootsSet(DriveSplit(1)) = True
    Next

    If UBound(RootsSet.Keys()) <> 0 Then
        If Strings.Len(Drive) > 0 Then
            Information.Err().Raise _
                Number:=5, _
                Source:="CommonPath", _
                Description:="Can't mix absolute and relative paths"
        Else
            Information.Err().Raise _
                Number:=5, _
                Source:="CommonPath", _
                Description:="Can't mix rooted and not-rooted paths"
        End If
    End If

    Dim Common() As String
    ReDim Common(UBound(Strings.Split(Path, SEP)))

    i = 0
    Dim c As Variant
    For Each c In Strings.Split(Path, SEP)
        If c <> CUR_DIR Then
            Common(i) = c
            i = i + 1
        End If
    Next
    ReDim Preserve Common(i - 1)

    Dim Buffer() As Variant
    ReDim Buffer(UBound(SplitPaths))
    i = 0
    Dim s As Variant
    For Each s In SplitPaths
        Dim NestedBuffer() As String
        ReDim NestedBuffer(UBound(s))
        Dim j As Long
        j = 0
        For Each c In s
            If c <> CUR_DIR Then
                NestedBuffer(j) = c
                j = j + 1
            End If
        Next
        ReDim Preserve NestedBuffer(j - 1)
        Buffer(i) = NestedBuffer
        i = i + 1
    Next
    ReDim Preserve Buffer(i - 1)

    Dim s1 As Variant
    s1 = Buffer(0)
    
    For i = 1 To UBound(Buffer)
        If CompareLists(Buffer(i), s1) < 0 Then
            s1 = Buffer(i)
        End If
    Next
    Dim s2 As Variant
    s2 = Buffer(0)
    
    For i = 1 To UBound(Buffer)
        If CompareLists(Buffer(i), s2) > 0 Then
            s2 = Buffer(i)
        End If
    Next

    For i = 0 To UBound(s1)
        c = s1(i)
        If c <> s2(i) Then
            ReDim Preserve Common(i - 1)
            Exit For
        Else
            ReDim Preserve Common(UBound(s1))
        End If
    Next

    CommonPath = Drive & Root & Strings.Join(Common, SEP)
End Function

'@Description "Return the longest path prefix (taken character-by-character) that is a prefix of all paths in list. If list is empty, return the empty string ('')."
Public Function CommonPrefix(ByRef List As Variant)
Attribute CommonPrefix.VB_Description = "Return the longest path prefix (taken character-by-character) that is a prefix of all paths in list. If list is empty, return the empty string ('')."
    If LBound(List) = UBound(List) Then
        CommonPrefix = List(LBound(List))
        Exit Function
    End If

    Dim Common As String
    Common = List(LBound(List))
    Dim i As Long
    For i = LBound(List) + 1 To UBound(List)
        Dim Path As String
        Path = List(i)
        Dim j As Long
        For j = 1 To Strings.Len(Path)
            Dim Char As String
            Char = Strings.Mid(Path, j, 1)
            If i <> LBound(List) And _
            j = 1 And _
            Strings.Left(Common, Strings.Len(Char)) <> Char Then
                Exit Function
            End If
            If Strings.Left(Common, Strings.Len(Char)) = Char Then
                Do While Strings.Left(Common, Strings.Len(Char)) = Char
                    If j = Strings.Len(Path) Then Exit Do
                    j = j + 1
                    Char = Char & Strings.Mid(Path, j, 1)
                Loop
                If Strings.Left(Common, Strings.Len(Char) - 1) = Strings.Left(Char, Strings.Len(Char) - 1) Then
                    If Strings.Left(Common, Strings.Len(Char)) = Char Then
                        Common = Char
                    Else
                        Common = Strings.Left(Char, Strings.Len(Char) - 1)
                    End If
                End If
            End If
        Next
    Next

    CommonPrefix = Common
End Function

'@Description "Return the directory name of pathname path. This is the first element of the pair returned by passing path to the function split()."
Public Function Dirname(ByVal Path As String) As String
Attribute Dirname.VB_Description = "Return the directory name of pathname path. This is the first element of the pair returned by passing path to the function split()."
    Dirname = PyPath.Split(Path)(0)
End Function

'@Description "Return True if path refers to an existing path."
Public Function Exists(ByVal Path As String) As Boolean
Attribute Exists.VB_Description = "Return True if path refers to an existing path."
    Path = PyPath.AbsPath(Path)
    Exists = GetFileAttributes(Path) <> INVALID_FILE_ATTRIBUTES
End Function

'@Description "Return the argument with an initial component of ~ or ~user replaced by that user's home directory. USERPROFILE will be used if set, otherwise a combination of HOMEPATH and HOMEDRIVE will be used. An initial ~user is handled by checking that the last directory component of the current user's home directory matches USERNAME, and replacing it if so. If the expansion fails or if the path does not begin with a tilde, the path is returned unchanged."
Public Function ExpandUser(ByVal Path As String) As String
Attribute ExpandUser.VB_Description = "Return the argument with an initial component of ~ or ~user replaced by that user's home directory. USERPROFILE will be used if set, otherwise a combination of HOMEPATH and HOMEDRIVE will be used. An initial ~user is handled by checking that the last directory component of the current user's home directory matches USERNAME, and replacing it if so. If the expansion fails or if the path does not begin with a tilde, the path is returned unchanged."
    Const TILDE As String = "~"
    Const SEPS As String = SEP & ALT_SEP

    If Strings.Left(Path, 1) <> TILDE Then
        ExpandUser = Path
        Exit Function
    End If

    Dim i As Long
    i = 1
    Dim n As Long
    n = Strings.Len(Path)

    Do While i < n And Strings.InStr(1, SEPS, Strings.Mid(Path, i + 1, 1)) = 0
        i = i + 1
    Loop

    Dim UserHome As String
    UserHome = Interaction.Environ("USERPROFILE")

    If Strings.Len(UserHome) = 0 Then
        Dim Drive As String
        Drive = Interaction.Environ("HOMEDRIVE")
        UserHome = PyPath.Join(Drive, Interaction.Environ("HOMEPATH"))

        If Strings.Len(UserHome) = 0 Then
            ExpandUser = Path
            Exit Function
        End If
    End If

    If i <> 1 Then '~user
        Dim TargetUser As String
        TargetUser = Strings.Mid(Path, 2, i - 1)
        Dim CurrentUser As String
        CurrentUser = Interaction.Environ("USERNAME")

        If TargetUser <> CurrentUser Then
            ' Try to guess user home directory.  By default all user
            ' profile directories are located in the same place and are
            ' named by corresponding usernames.  If userhome isn't a
            ' normal profile directory, this guess is likely wrong,
            ' so we bail out.
            If CurrentUser <> PyPath.Basename(UserHome) Then
                ExpandUser = Path
                Exit Function
            End If
            UserHome = PyPath.Join(PyPath.Dirname(UserHome), TargetUser)
        End If
    End If

    ExpandUser = UserHome & Strings.Mid(Path, i + 1)
End Function

'@Description "Return the argument with environment variables expanded. Substrings of the form %name% are replaced by the value of environment variable name. Malformed variable names and references to non-existing variables are left unchanged."
Public Function ExpandVars(ByVal Path As String) As String
Attribute ExpandVars.VB_Description = "Return the argument with environment variables expanded. Substrings of the form %name% are replaced by the value of environment variable name. Malformed variable names and references to non-existing variables are left unchanged."
    Dim i As Long
    i = 1
    Do While True
        Dim Var As String
        Var = ""

        Dim StartVar As Long
        StartVar = i
        If Strings.Mid(Path, i, 1) = "%" Then
            i = i + 1
            Dim Char As String
            Char = Strings.Mid(Path, i, 1)
            Do While Char <> " "
                Char = Strings.Mid(Path, i, 1)
                If Char = "%" Then
                    Var = Strings.Mid(Path, StartVar + 1, i - StartVar - 1)
                    Dim Value As String
                    Value = Interaction.Environ(Var)
                    Path = Strings.Replace(Path, "%" & Var & "%", Value, Compare:=VbCompareMethod.vbTextCompare)
                    Exit Do
                End If
                i = i + 1
                If i > Strings.Len(Path) Then Exit Do
            Loop
        End If
        i = i + 1
        If i > Strings.Len(Path) Then Exit Do
    Loop

    ExpandVars = Path
End Function

'@Description "Return the time of last access of path."
Public Function GetATime(ByVal Path As String) As Double
Attribute GetATime.VB_Description = "Return the time of last access of path."
    GetATime = GetFSO.GetFile(Path).DateLastAccessed
End Function

'@Description "Return the creation time for path."
Public Function GetCTime(ByVal Path As String) As Double
Attribute GetCTime.VB_Description = "Return the creation time for path."
    GetCTime = GetFSO.GetFile(Path).DateCreated
End Function

'@Description "Return the time of last modification of path."
Public Function GetMTime(ByVal Path As String) As Double
Attribute GetMTime.VB_Description = "Return the time of last modification of path."
    GetMTime = GetFSO.GetFile(Path).DateLastModified
End Function

'@Description "Return the size, in bytes, of path."
Public Function GetSize(ByVal Path As String) As Long
Attribute GetSize.VB_Description = "Return the size, in bytes, of path."
    GetSize = GetFSO.GetFile(Path).Size
End Function

'@Description "Return True if path is an absolute pathname. That it begins with two (back)slashes, or a drive letter, colon, and (back)slash together."
Public Function IsAbs(ByVal Path As String) As Boolean
Attribute IsAbs.VB_Description = "Return True if path is an absolute pathname. That it begins with two (back)slashes, or a drive letter, colon, and (back)slash together."
    Const COLON_SEP As String = ":\"
    Const DOUBLE_SEP As String = "\\"

    Dim PathDrive As String
    PathDrive = Strings.Left(Path, 3)
    Dim CorrectPathDrive As String
    CorrectPathDrive = Strings.Replace(PathDrive, ALT_SEP, SEP)
    Path = Strings.Replace(Path, PathDrive, CorrectPathDrive, Count:=1)

    IsAbs = Strings.Mid(Path, 2, 2) = COLON_SEP Or _
            Strings.Left(Path, 2) = DOUBLE_SEP
End Function

'@Description "Return True if path is an existing directory."
Public Function IsDir(ByVal Path As String) As Boolean
Attribute IsDir.VB_Description = "Return True if path is an existing directory."
    Path = PyPath.AbsPath(Path)

    Dim attr As Long
    attr = GetFileAttributes(Path)

    IsDir = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function

'@Description "Return True if path is an existing regular file."
Public Function IsFile(ByVal Path As String) As Boolean
Attribute IsFile.VB_Description = "Return True if path is an existing regular file."
    Path = PyPath.AbsPath(Path)

    Dim attr As Long
    attr = GetFileAttributes(Path)

    IsFile = (attr <> INVALID_FILE_ATTRIBUTES) And ((attr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

'@Description "Join one or more path segments intelligently. The return value is the concatenation of path and all members of Paths(), with exactly one directory separator following each non-empty part, except the last. That is, the result will only end in a separator if the last part is either empty or ends in a separator. If a segment is an absolute path (which on Windows requires both a drive and a root), then all previous segments are ignored and joining continues from the absolute path segment."
Public Function Join(ByVal Path As String, ParamArray Paths() As Variant) As String
Attribute Join.VB_Description = "Join one or more path segments intelligently. The return value is the concatenation of path and all members of Paths(), with exactly one directory separator following each non-empty part, except the last. That is, the result will only end in a separator if the last part is either empty or ends in a separator. If a segment is an absolute path (which on Windows requires both a drive and a root), then all previous segments are ignored and joining continues from the absolute path segment."
    Const COLON_SEP As String = ":\"
    Const SEPS As String = SEP & ALT_SEP
    Const COLON_SEPS As String = COLON_SEP & ALT_SEP

    Dim Result() As String
    Result = PyPath.SplitRoot(Path)
    Dim ResultDrive As String
    ResultDrive = Result(0)
    Dim ResultRoot As String
    ResultRoot = Result(1)
    Dim ResultPath As String
    ResultPath = Result(2)

    Dim p As Variant
    For Each p In Paths
        Dim p_Result() As String
        p_Result = PyPath.SplitRoot(p)
        Dim p_Drive As String
        p_Drive = p_Result(0)
        Dim p_Root As String
        p_Root = p_Result(1)
        Dim p_Path As String
        p_Path = p_Result(2)

        If Strings.Len(p_Root) > 0 Then
            ' Second path is absolute
            If Strings.Len(p_Drive) > 0 Or Strings.Len(ResultDrive) = 0 Then
                ResultDrive = p_Drive
            End If
            ResultRoot = p_Root
            ResultPath = p_Path
            GoTo Continue
        ElseIf Strings.Len(p_Drive) > 0 And p_Drive <> ResultDrive Then
            If Strings.LCase(p_Drive) <> Strings.LCase(ResultDrive) Then
                ' Different drives => ignore the first path entirely
                ResultDrive = p_Drive
                ResultRoot = p_Root
                ResultPath = p_Path
                GoTo Continue
            End If

            ' Same drive in different case
            ResultDrive = p_Drive
        End If

        ' Second path is relative to the first
        If Strings.Len(ResultPath) > 0 And Strings.InStr(1, SEPS, Strings.Right(ResultPath, 1)) = 0 Then
            ResultPath = ResultPath + SEP
        End If
        ResultPath = ResultPath + p_Path
Continue:
    Next
    ' add separator between UNC and non-absolute path
    If (Strings.Len(ResultPath) > 0 And Strings.Len(ResultRoot) = 0 And _
        Strings.Len(ResultDrive) > 0 And Strings.InStr(1, COLON_SEPS, Strings.Right(ResultDrive, 1)) = 0) Then
        Join = ResultDrive + SEP + ResultPath
        Exit Function
    End If

    Join = ResultDrive + ResultRoot + ResultPath
End Function

'@Description "Normalize the case of a pathname. Convert all characters in the pathname to lowercase, and also convert forward slashes to backward slashes."
Public Function NormCase(ByVal Path As String) As String
Attribute NormCase.VB_Description = "Normalize the case of a pathname. Convert all characters in the pathname to lowercase, and also convert forward slashes to backward slashes."
    NormCase = Strings.LCase(Strings.Replace(Path, ALT_SEP, SEP))
End Function

'@Description "Normalize a path, e.g. A//B, A/./B and A/foo/../B all become A\B."
Public Function NormPath(ByVal Path As String) As String
Attribute NormPath.VB_Description = "Normalize a path, e.g. A//B, A/./B and A/foo/../B all become A\\B."
    Path = Strings.Replace(Path, ALT_SEP, SEP)
    Dim Splited() As String
    Splited = PyPath.SplitRoot(Path)
    Dim Drive As String
    Drive = Splited(0)
    Dim Root As String
    Root = Splited(1)
    Path = Splited(2)
    Dim PREFIX As String
    PREFIX = Drive & Root

    Dim Comps() As String
    Comps = Strings.Split(Path, SEP)

    Dim Result() As String
    ReDim Result(UBound(Comps))
    Dim j As Long
    j = 0
    Dim i As Long
    i = 0
    While i <= UBound(Comps)
        If Strings.Len(Comps(i)) = 0 Or Comps(i) = CUR_DIR Then
            Comps(i) = Empty
            i = i + 1
        ElseIf Comps(i) = PAR_DIR Then
            If i > 0 Then
                If Comps(i - 1) <> PAR_DIR Then
                    Comps(i - 1) = Empty
                    Comps(i) = Empty
                    Result(j - 1) = Empty
                    j = j - 1
                End If
                i = i + 1
            ElseIf i = 0 And Strings.Len(Root) > 0 Then
                Comps(i) = Empty
                i = i + 1
            Else
                Result(j) = Comps(i)
                j = j + 1
                i = i + 1
            End If
        Else
            Result(j) = Comps(i)
            j = j + 1
            i = i + 1
        End If
    Wend

    ReDim Preserve Result(j - 1)
    ' If the path is now empty, substitute '.'
    If Strings.Len(PREFIX) = 0 And Strings.Len(Strings.Join(Result)) = 0 Then
        ReDim Preserve Result(0)
        Result(0) = CUR_DIR
    End If

    NormPath = PREFIX & Strings.Join(Result, SEP)
End Function

'@Description "Return the canonical path of the specified filename. This function will also resolve MS-DOS (also called 8.3) style names such as C:\PROGRA~1 to C:\Program Files."
Public Function RealPath(ByVal Path As String) As String
Attribute RealPath.VB_Description = "Return the canonical path of the specified filename. This function will also resolve MS-DOS (also called 8.3) style names such as C:\\PROGRA~1 to C:\\Program Files."
    Path = PyPath.NormPath(Path)
    Const PREFIX As String = "\\?\"
    Const UNC_PREFIX = "\\?\UNC\"
    Const NEW_UNC_PREFIX = "\\"

    Dim CWD As String
    CWD = FileSystem.CurDir()

    If PyPath.NormCase(Path) = DEV_NULL Then
        RealPath = "\\.\NULL"
        Exit Function
    End If

    Dim HadPrefix As Boolean
    HadPrefix = Strings.Left(Path, Strings.Len(PREFIX)) = PREFIX

    If Not HadPrefix And Not PyPath.IsAbs(Path) Then
        Path = PyPath.Join(CWD, Path)
    End If
    Path = PyPath.NormPath(Path)

    Dim SPath As String
    If Not HadPrefix And Strings.Left(Path, Strings.Len(PREFIX)) = PREFIX Then
        ' For UNC paths, the prefix will actually be \\?\UNC\
        ' Handle that case as well.
        If Strings.Left(Path, Strings.Len(UNC_PREFIX)) = UNC_PREFIX Then
            SPath = NEW_UNC_PREFIX & Strings.Mid(Path, Strings.Len(UNC_PREFIX) + 1)
        Else
            SPath = Strings.Mid(Path, Strings.Len(PREFIX) + 1)
        End If
        SPath = Path
    End If

    Dim LongPath As String
    Dim Size As Long

    LongPath = Strings.String(260, 0)
    Size = GetLongPathName(Path, LongPath, 260)

    If Size > 0 Then
        RealPath = Strings.Left(LongPath, Size)
    Else
        RealPath = Path
    End If
End Function

'@Description "Return a relative filepath to path either from the current directory or from an optional start directory."
Public Function RelPath(ByVal Path As String, Optional ByVal Start As String = CUR_DIR) As String
Attribute RelPath.VB_Description = "Return a relative filepath to path either from the current directory or from an optional start directory."
    If Strings.Len(Path) = 0 Then
        Information.Err().Raise _
            Number:=5, _
            Source:="RelPath", _
            Description:="Path no specified"
    End If

    If Strings.Len(Start) = 0 Then
        Start = CUR_DIR
    End If

    Dim StartAbs As String
    StartAbs = PyPath.AbsPath(Start)
    Dim PathAbs As String
    PathAbs = AbsPath(Path)

    Dim Splited() As String
    Splited = PyPath.SplitRoot(StartAbs)

    Dim StartDrive As String
    StartDrive = Splited(0)
    Dim StartRest As String
    StartRest = Splited(2)

    Splited = PyPath.SplitRoot(PathAbs)

    Dim PathDrive As String
    PathDrive = Splited(0)
    Dim PathRest As String
    PathRest = Splited(2)

    If PyPath.NormCase(StartDrive) <> PyPath.NormCase(PathDrive) Then
        Information.Err().Raise _
            Number:=5, _
            Source:="RelPath", _
            Description:="path is on mount" & PathDrive & ", start on mount " & StartDrive
    End If

    Dim StartList() As String
    StartList = Strings.Split(StartRest, SEP)
    Dim PathList() As String
    PathList = Strings.Split(PathRest, SEP)

    ' Work out how much of the filepath is shared by start and path.
    Dim i As Long
    For i = 0 To Application.Min(UBound(StartList), UBound(PathList))
        If PyPath.NormCase(StartList(i)) <> PyPath.NormCase(PathList(i)) Then
            Exit For
        End If
    Next

    Dim RelList() As String
    ReDim RelList(i + UBound(PathList))
    Dim j As Long
    For j = 0 To UBound(StartList) - i
        RelList(j) = PAR_DIR
    Next

    For j = j To UBound(PathList) - i + 1
        RelList(j) = PathList(j + i - 1)
    Next
    ReDim Preserve RelList(j - 1)

    If Strings.Len(Strings.Join(RelList)) = 0 Then
        RelPath = CUR_DIR
    Else
        RelPath = Strings.Join(RelList, SEP)
    End If
End Function

'@Description "Split the pathname path into a pair, (head, tail) where tail is the last pathname component and head is everything leading up to that."
Public Function Split(ByVal Path As String) As String()
Attribute Split.VB_Description = "Split the pathname path into a pair, (head, tail) where tail is the last pathname component and head is everything leading up to that."
    Const SEPS As String = SEP & ALT_SEP

    Dim Result() As String
    Result = PyPath.SplitRoot(Path)
    Path = Result(2)

    Dim i As Long
    i = Strings.Len(Path)

    Do While i > 0 And Strings.InStr(1, SEPS, Strings.Mid(Path, i, 1)) = 0
        i = i - 1
    Loop
    Dim Head As String
    Dim Tail As String
    Head = Strings.Left(Path, i)
    Tail = Strings.Mid(Path, i + 1)

    If InStr(1, SEPS, Strings.Right(Head, 1)) > 0 Then
        Head = Strings.Left(Head, Strings.Len(Head) - 1)
    End If
    Result(0) = Result(0) & Result(1) & Head
    Result(1) = Tail
    ReDim Preserve Result(1)

    Split = Result
End Function

'@Description "Split the pathname path into a pair (drive, tail) where drive is either a mount point or the empty string. "
Public Function SplitRoot(ByVal Path As String) As String()
Attribute SplitRoot.VB_Description = "Split the pathname path into a pair (drive, tail) where drive is either a mount point or the empty string. "
    Dim Result() As String
    ReDim Result(2)

    Const UNC_PREFIX As String = "\\?\UNC\"
    Const COLON As String = ":"
    Dim NormPath As String
    NormPath = Strings.Replace(Path, ALT_SEP, SEP)

    If Strings.Left(NormPath, 1) = SEP Then
        If Strings.Mid(NormPath, 2, 1) = SEP Then
            ' UNC drives, e.g. \\server\share or \\?\UNC\server\share
            ' Device drives, e.g. \\.\device or \\?\device
            Dim Start As Long
            Start = Interaction.IIf(Strings.UCase(Strings.Left(NormPath, 8)) = UNC_PREFIX, 8, 2)
            Dim Index As Long
            Index = Strings.InStr(Start, NormPath, SEP)
            If Index = 0 Then
                Result(0) = Path
                SplitRoot = Result
                Exit Function
            End If
    
            Dim Index2 As Long
            Index2 = Strings.InStr(Index + 1, NormPath, SEP)
            If Index2 = 0 Then
                Result(0) = Path
                SplitRoot = Result
                Exit Function
            End If

            Result(0) = Strings.Left(Path, Index2)
            Result(1) = Strings.Mid(Path, Index2, 1)
            Result(2) = Strings.Mid(Path, Index2 + 1)
            SplitRoot = Result
            Exit Function
        Else
            Result(1) = Strings.Left(Path, 1)
            Result(2) = Strings.Mid(Path, 2)
            SplitRoot = Result
            Exit Function
        End If
    ElseIf Strings.Mid(NormPath, 2, 1) = COLON Then
        If Strings.Mid(NormPath, 3, 1) = SEP Then
            ' Absolute drive-letter path, e.g. X:\Windows
            Result(0) = Strings.Left(Path, 2)
            Result(1) = Strings.Mid(Path, 3, 1)
            Result(2) = Strings.Mid(Path, 4)
            SplitRoot = Result
            Exit Function
        Else
            ' Relative path with drive, e.g. X:Windows
            Result(0) = Strings.Left(Path, 2)
            Result(2) = Strings.Mid(Path, 3)
            SplitRoot = Result
            Exit Function
        End If
    Else
        ' Relative path, e.g. Windows
        Result(2) = Path
        SplitRoot = Result
        Exit Function
    End If
End Function

'@Description "Split the pathname path into a 3-item tuple (drive, root, tail) where drive is a device name or mount point, root is a string of separators after the drive, and tail is everything after the root. Any of these items may be the empty string. In all cases, drive + root + tail will be the same as path."
Public Function SplitDrive(ByVal Path As String) As String()
Attribute SplitDrive.VB_Description = "Split the pathname path into a 3-item tuple (drive, root, tail) where drive is a device name or mount point, root is a string of separators after the drive, and tail is everything after the root. Any of these items may be the empty string. In all cases, drive + root + tail will be the same as path."
    Dim Result() As String

    Result = PyPath.SplitRoot(Path)
    Result(1) = Result(1) & Result(2)
    ReDim Preserve Result(1)

    SplitDrive = Result
End Function

'@Description "Split the pathname path into a pair (root, ext) such that root + ext = path, and the extension, ext, is empty or begins with a period and contains at most one period. If the path contains no extension, ext will be ''."
Public Function SplitExt(ByVal Path As String) As String()
Attribute SplitExt.VB_Description = "Split the pathname path into a pair (root, ext) such that root + ext = path, and the extension, ext, is empty or begins with a period and contains at most one period. If the path contains no extension, ext will be ''."
    Dim Result() As String
    ReDim Result(1)

    Dim LastSep As Long
    LastSep = Strings.InStrRev(Path, CUR_DIR)

    If LastSep = 0 Then
        Result(0) = Path
    Else
        Result(0) = Strings.Left(Path, LastSep - 1)
        Result(1) = Strings.Mid(Path, LastSep)
    End If

    SplitExt = Result
End Function

Private Function CompareLists(ByRef List1 As Variant, ByRef List2 As Variant) As Long
    Dim i As Integer
    Dim MinLength As Long
    MinLength = Min(UBound(List1), UBound(List2))

    For i = 0 To MinLength
        If List1(i) < List2(i) Then
            CompareLists = -1
            Exit Function
        ElseIf List1(i) > List2(i) Then
            CompareLists = 1
            Exit Function
        End If
    Next i

    If UBound(List1) < UBound(List2) Then
        CompareLists = -1
    ElseIf UBound(List1) > UBound(List2) Then
        CompareLists = 1
    Else
        CompareLists = 0
    End If
End Function

Private Function Min(ParamArray Args() As Variant) As Variant
    Dim MinValue As Variant

    MinValue = Args(LBound(Args))
    Dim i As Long
    For i = LBound(Args) + 1 To UBound(Args)
        If Args(i) < MinValue Then MinValue = Args(i)
    Next

    Min = MinValue
End Function

Private Function GetFSO() As Object
    Static FSO As Object
    If FSO Is Nothing Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
    End If

    Set GetFSO = FSO
End Function

Private Function GetFileAttributes(ByVal Path As String) As Long
    Dim PathW As String
    If Strings.Len(Path) >= 248 Then
        PathW = "\\?\" & Path
    Else
        PathW = Path
    End If

    GetFileAttributes = GetFileAttributesW(StrPtr(PathW))
End Function
