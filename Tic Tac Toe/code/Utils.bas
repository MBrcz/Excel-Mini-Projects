Attribute VB_Name = "utils"
Option Explicit

Public Function GetPathToImagesFile() As String
' Returns path to the file, where the images are stored.

    Dim path As String
    Dim directoryPath As String
    
    path = ThisWorkbook.FullName
    directoryPath = Left(path, InStrRev(path, "\") - 1) & "\img\"
    
    GetPathToImagesFile = directoryPath

End Function

Public Function GetMiddlePointInScreen() As Variant
' Returns position of the middle pixel in the User's Screen.

    Dim screenWidth As Long
    Dim screenHeight As Long
    Dim middleX As Long
    Dim middleY As Long
    
    ' Get the screen dimensions
    screenWidth = Application.Width
    screenHeight = Application.Height
    
    ' Calculate the middle point
    middleX = screenWidth / 2
    middleY = screenHeight / 2
    
    ' Return the middle point as a variant array
    GetMiddlePointInScreen = Array(middleX, middleY)
End Function

Public Function RemoveElementFromArray(arr As Variant, el As Variant) As Variant
' Removes and element from an array and returns the array without element.
    
    Dim result() As Variant
    Dim element As Variant
    Dim count As Byte
    Dim el_num As Byte
    
    ReDim Preserve result(0 To UBound(arr) - 1)
    
    For count = LBound(arr) To UBound(arr)
        If arr(count) <> el Then
            result(el_num) = arr(count)
            el_num = el_num + 1
        End If
    Next count
    
    RemoveElementFromArray = result
End Function

Public Function GetRandomIndexFromArray(arr As Variant) As Integer
' As name implies, returns random element from an array.
    
    Dim arr_index As Integer
    
    arr_index = Application.WorksheetFunction.RandBetween(LBound(arr), UBound(arr))
    GetRandomIndexFromArray = arr_index

End Function

Public Function PrintArray(ByVal arr As Variant, Optional ByVal arr_depth As Long = 1, Optional ByVal separator As String = ", ")
' Debug function Prints the array in the one line.

    Dim i As Long, j As Long
    Dim subArray As Variant
    Dim output As String
    
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If arr_depth > 1 Then
                If IsArray(arr(i)) Then
                    subArray = arr(i)
                    output = output & PrintArray(subArray, arr_depth - 1, separator & " ") & separator
                Else
                    output = output & arr(i) & separator
                End If
            Else
                output = output & arr(i) & separator
            End If
        Next i
    Else
        output = "Not an array: " & arr
    End If
    
    PrintArray = Left(output, Len(output) - Len(separator))
    
End Function

Public Function GetAllFileNamesFromDir(dir_path As String) As Variant
' Gets all files from chosen directory.
    
    Dim fso As Object
    Dim folder As Object
    Dim files As Object
    Dim file As Object
    
    Dim num As Byte
    Dim files_names() As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(dir_path)
    
    Set files = folder.files
    
    ReDim files_names(1 To files.count)
    num = 1
    
    For Each file In files
        files_names(num) = file.Name
        num = num + 1
    Next file
    
    GetAllFileNamesFromDir = files_names
    
    Set fso = Nothing
    Set folder = Nothing
    Set files = Nothing

End Function

