Attribute VB_Name = "Project"
''
' This module is created to easily export and import VBA code
' into and from `./src` directory and run basic git commands.
'
' This is needed as Git can't read Excel files directly, but can read the
' source files that are exported.
'
' The export should be ran either on every save or on certain events.
' Note, can't use other modules with importing as they will be deleted and
' won't be able to be accessed.
'
' @author Robert Todar <robert@roberttodar.com>
' @status Development - Exporting seems to be working fine right now,
'                       but still having issues with importing. Looks Like
'                       I can import everything but Sheets, ThisWorkbook, and .frx files.
'                       Maybe create a solution to store just the text of code for sheets
'                       and workbook vs the code export.
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponets
' @ref {Microsoft Visual Basic For Applications Extensibility 5.3} VBComponet
' @ref {Microsoft Scripting Runtime} Scripting.FileSystemObject
''
Option Explicit

' Root Directory of this Project.
Public Property Get Dirname() As String
    Dirname = ThisWorkbook.path
End Property

' Directory where all source code will be stored. `./src`
Public Property Get SourceDirectory() As String
    SourceDirectory = joinPaths(Dirname, "src")
End Property

' This Projects VB thisProjectsVBComponents.
' @NOTE: Should this be a single project, or should I use this
'        for any project/workbook? For now will leave as the
'        current
Private Property Get thisProjectsVBComponents() As VBComponents
    Set thisProjectsVBComponents = ThisWorkbook.VBProject.VBComponents
End Property

' Helper function to run scripts from the root directory.
Public Function Bash(script As String, Optional keepCommandWindowOpen As Boolean = False) As Double
    ' cmd.exe Opens the command prompt.
    ' /S      Modifies the treatment of string after /C or /K (see below)
    ' /C      Carries out the command specified by string and then terminates
    ' /K      Carries out the command specified by string but remains
    ' cd      Change directory to the root directory.
    Bash = Shell("cmd.exe /S /" & IIf(keepCommandWindowOpen, "K", "C") & " cd " & ThisWorkbook.path & " && " & script)
End Function

' Initiates a new Git Project in the current folder.
' Safe to run even if project is initialized.
Public Sub InitializeProject()
    Dim fso As New Scripting.FileSystemObject
    ' Create a default .gitignore file if it doesn't exist already
    ' @see https://git-scm.com/docs/gitignore
    Dim gitignorePath As String
    gitignorePath = joinPaths(Dirname, ".gitignore")
    If Not fso.FileExists(gitignorePath) Then
        With fso.OpenTextFile(gitignorePath, ForWriting, True)
            .WriteLine ("# Packages")
            .WriteLine ("node_modules")
            .WriteBlankLines 1
            .WriteLine ("# Excel's Backup copies")
            .Write ("~$*.xl*")
            .Close
        End With
    End If
    
    ' Initialie git (safe even if it already exists)
    ' @see https://git-scm.com/docs/git-init
    Bash script:="git init", keepCommandWindowOpen:=False
End Sub

' Get the file extension for a VBComponent. That is the component name and the proper extension.
Private Function getVBComponentFilename(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentFilename = component.Name & ".cls"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentFilename = component.Name & ".bas"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentFilename = component.Name & ".frm"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentFilename = component.Name & ".cls"
            
        Case Else
            ' @TODO: Need to think of possible throwing an error?
            ' Is it possible to get something else?? I don't think so
            ' Will need to double check this.
            Debug.Print "Unknown component"
    End Select
End Function

' Check to see if component exits in this current Project
Private Function componentExists(ByVal filename As String) As Boolean
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        If getVBComponentFilename(component) = filename Then
            componentExists = True
            Exit Function
        End If
    Next index
End Function

' Export all modules in this current workbook into a src dir
Public Sub ExportComponentsToSourceFolder()
    ' Make sure the source directory exists before adding to it.
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(SourceDirectory) Then
        fso.CreateFolder SourceDirectory
    Else
        Dim file As file
        For Each file In fso.GetFolder(SourceDirectory).Files
            file.Delete
        Next file
    End If
    
    ' Loop each component within this project and export to source directory.
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        ' Export component to the source directory using the components name and file extension.
        component.Export joinPaths(SourceDirectory, getVBComponentFilename(component))
    Next index
End Sub

' Import source code from the source Directory.
' This works by first deleting all current components,
' then importing all the components from the source directory.
'
' @status Testing && Development
' @warn This will cause files to overwrite that already exists.
' @warn This will also remove files not found in the source component.
Public Sub DangerouslyImportComponentsFromSourceFolder()
    If MsgBox("Are you sure you want to import from source folder? There is no going back!!!", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim fso As New Scripting.FileSystemObject
    
    ' Remove current components to make room for the imported ones.
    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        ' If the component already, it needs to be deleted in order to
        ' import the file, otherwise an error is thrown.
        If componentExists(file.Name) And file.Name <> "Project.bas" Then
            Dim component As VBComponent
            Set component = thisProjectsVBComponents.item(fso.GetBaseName(file.Name))
            
            ' Unable to remove document type components (Sheets, workbook)
            If component.Type <> vbext_ct_Document Then
                ' This removes the component but doesn't from memory until
                ' after all code execution has completed.
                thisProjectsVBComponents.Remove component
            End If
        End If
    Next file
    
    ' After all code is finished executing, the components removed above will
    ' finally be removed from memory.
    Application.OnTime Now, "saftleyImportAfterCleanup"
End Sub

Private Sub saftleyImportAfterCleanup()
    Dim fso As New Scripting.FileSystemObject

    Dim file As file
    For Each file In fso.GetFolder(SourceDirectory).Files
        If Not componentExists(file.Name) And fso.GetExtensionName(file.Name) <> "frx" Then
            ' Safe to import the source file as there are no conflicts of names.
            thisProjectsVBComponents.Import joinPaths(SourceDirectory, file.Name)
        End If
    Next file
End Sub

' Converts the VBComponent enum to a string representation of type of component.
Private Function getVBComponentTypeName(ByRef component As VBComponent) As String
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_ClassModule
            getVBComponentTypeName = "Class Module"
            
        Case vbext_ComponentType.vbext_ct_StdModule
            getVBComponentTypeName = "Module"
            
        Case vbext_ComponentType.vbext_ct_MSForm
            getVBComponentTypeName = "Form"
            
        Case vbext_ComponentType.vbext_ct_Document
            getVBComponentTypeName = "Document"
            
        Case Else
            ' All components should be accounted for, this is just in case ;)
            Debug.Print "Unknown type: " & component.Type
    End Select
End Function

' Prints out details about a specific VBComponent. Used for
' @status Production
Private Function getComponentDetails(ByRef component As VBComponent) As String
    getComponentDetails = component.Name & vbTab _
                          & getVBComponentTypeName(component) & vbTab _
                          & getVBComponentFilename(component)
End Function

' Prints out details about all VBComponents in the current project
' @status Production
Public Property Get ComponentsDetails() As String
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        ComponentsDetails = ComponentsDetails & getComponentDetails(component) & vbNewLine
    Next index
End Property

' Prints out details about all VBComponents in the current project
' @status Development
Private Sub printDiffFromSourceFolder()
    Dim index As Long
    For index = 1 To thisProjectsVBComponents.count
        Dim component As VBComponent
        Set component = thisProjectsVBComponents(index)
        
        Debug.Print getVBComponentFilename(component)
    Next index
End Sub

' Helper function to join paths...
Private Function joinPaths(ParamArray paths() As Variant) As String
    Dim fso As New Scripting.FileSystemObject
    Dim index As Long
    For index = LBound(paths) To UBound(paths)
        joinPaths = fso.BuildPath(joinPaths, Replace(paths(index), "/", "\"))
    Next
End Function

' Compares a VBA module's code in the current workbook against a .bas file
' Returns True if identical (by checksum), False otherwise
' This is to enable automatic project change detetection of a target safe file. Like a tests completed output.
` or to identify which internal module has changed on save to keep from writing all outputs every time.
' Compare a single module to its exported file (.bas, .cls, .frm)
Public Function CompareModuleToBasFile(ByVal moduleName As String, ByVal basFilePath As String) As Boolean
    Dim moduleContent As String
    Dim fileContent As String
    Dim hashModule As String
    Dim hashFile As String
    
    On Error GoTo ErrHandler
    
    moduleContent = GetModuleCode(ThisWorkbook, moduleName)
    fileContent = ReadTextFile(basFilePath)
    
    hashModule = GetMD5(moduleContent)
    hashFile = GetMD5(fileContent)
    
    CompareModuleToBasFile = (hashModule = hashFile)
    Exit Function
    
ErrHandler:
    Debug.Print "CompareModuleToBasFile error (" & moduleName & "): " & Err.Description
    CompareModuleToBasFile = False
End Function


' Compare all VBA components in the current workbook to a folder of exported files
' Returns True if all match, False if any differ
Public Function CompareAllModulesToFolder(ByVal folderPath As String) As Boolean
    Dim vbComp As Object
    Dim fso As Object
    Dim folderMatch As Boolean
    Dim filePath As String
    Dim ext As String
    Dim allMatch As Boolean
    Dim diffCount As Long
    
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    allMatch = True
    
    Debug.Print "=== VBA Project Checksum Report ==="
    Debug.Print "Folder: " & folderPath
    Debug.Print "-----------------------------------"
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1: ext = ".bas" ' Standard Module
            Case 2: ext = ".cls" ' Class Module
            Case 3: ext = ".frm" ' UserForm
            Case Else: ext = ".bas"
        End Select
        
        filePath = fso.BuildPath(folderPath, vbComp.Name & ext)
        
        If fso.FileExists(filePath) Then
            folderMatch = CompareModuleToBasFile(vbComp.Name, filePath)
            
            If folderMatch Then
                Debug.Print vbComp.Name & " — OK"
            Else
                Debug.Print vbComp.Name & " — DIFFERENT"
                allMatch = False
                diffCount = diffCount + 1
            End If
        Else
            Debug.Print vbComp.Name & " — Missing file (" & filePath & ")"
            allMatch = False
        End If
    Next vbComp
    
    Debug.Print "-----------------------------------"
    If allMatch Then
        Debug.Print " All modules match exported files."
    Else
        Debug.Print " " & diffCount & " module(s) differ or missing."
    End If
    
    CompareAllModulesToFolder = allMatch
    Exit Function
    
ErrHandler:
    Debug.Print "CompareAllModulesToFolder error: " & Err.Description
    CompareAllModulesToFolder = False
End Function


' Returns a dictionary (late-bound) with [ModuleName] = MD5 checksum
Public Function GetAllModuleChecksums() As Object
    Dim vbComp As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        dict(vbComp.Name) = GetMD5(GetModuleCode(ThisWorkbook, vbComp.Name))
    Next vbComp
    
    Set GetAllModuleChecksums = dict
End Function


'=== Internal Utilities ====================================================

' Get all lines of code from a VBA component
Private Function GetModuleCode(ByVal wb As Workbook, ByVal moduleName As String) As String
    Dim vbComp As Object
    Set vbComp = wb.VBProject.VBComponents(moduleName)
    GetModuleCode = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
End Function


' Read text from file (UTF-8 safe for ASCII content)
Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1)
    ReadTextFile = ts.ReadAll
    ts.Close
End Function


' Compute MD5 hash from text
Private Function GetMD5(ByVal text As String) As String
    Dim enc As Object, bytes() As Byte
    Dim hash() As Byte, i As Long
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    bytes = StrConv(text, vbFromUnicode)
    hash = enc.ComputeHash_2((bytes))
    For i = 0 To UBound(hash)
        GetMD5 = GetMD5 & LCase(Right("0" & Hex(hash(i)), 2))
    Next i
End Function
