Attribute VB_Name = "basOfficeFileDialog"
Option Compare Database
Option Explicit


Public Function DimeCarpeta(Optional strInitialFileName As String = "", Optional strTitle As String = "") As Variant
   ' Requires reference to Microsoft Office 11.0 Object Library.

   Dim fDialog As Office.FileDialog
   Dim varF As Variant

   ' Clear listbox contents.
   'Me.FileList.RowSource = ""

   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)

   With fDialog

      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
            
      ' Set the title of the dialog box.
      .Title = "Seleccione " & IIf(strTitle = "", "Carpeta", strTitle)

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      '.Filters.Add "Access Databases", "*.MDB"
      '.Filters.Add "Access Projects", "*.ADP"
      '.Filters.Add "All Files", "*.*"
      .InitialFileName = strInitialFileName
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then

         'Loop through each file selected and add it to our list box.
         For Each varF In .SelectedItems
            'Me.FileList.AddItem varF
            DimeCarpeta = varF
         Next
        
      Else
         'MsgBox "You clicked Cancel in the file dialog box."
         DimeCarpeta = Null
      End If
   End With
End Function


Public Function DimeFicheros(Optional strInitialFileName As String = "", Optional strTitle As String = "" _
        , Optional strFilterTexto As String = "Tipo de fichero", Optional strFilterExtension As String = "*.*") As Variant
   ' Requires reference to Microsoft Office 11.0 Object Library.
    'Devuelve un string con un fichero en cada línea
   Dim fDialog As Office.FileDialog
   Dim varF As Variant, strR As String

   ' Clear listbox contents.
   'Me.FileList.RowSource = ""

   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog

      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = True
            
      ' Set the title of the dialog box.
      .Title = "Seleccione " & IIf(strTitle = "", "Fichero/s", strTitle)

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      '.Filters.Add "Access Databases", "*.MDB"
      .Filters.Add strFilterTexto, strFilterExtension
      '.Filters.Add "Imágenes JPG", "*.JPG"
      '.Filters.Add "All Files", "*.*"
      .InitialFileName = strInitialFileName
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then

         'Loop through each file selected and add it to our list box.
         For Each varF In .SelectedItems
            'Me.FileList.AddItem varF
            strR = strR & varF & vbCrLf
         Next
         If Len(strR) > 2 Then strR = RecDerTop(strR, 2, 0)
         DimeFicheros = strR
        
      Else
         'MsgBox "You clicked Cancel in the file dialog box."
         DimeFicheros = Null
      End If
   End With
End Function



Public Function DimeFichero(Optional strInitialFileName As String = "", Optional strTitle As String = "" _
                , Optional strFilterTexto As String = "Tipo de fichero", Optional strFilterExtension As String = "*.*") As Variant
   ' Requires reference to Microsoft Office 11.0 Object Library.
    'Devuelve un string con un fichero en cada línea
   Dim fDialog As Office.FileDialog
   Dim i As FileSystemObject
   Dim varF As Variant, strR As String

   ' Clear listbox contents.
   'Me.FileList.RowSource = ""

   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog

      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
            
      ' Set the title of the dialog box.
      .Title = "Seleccione " & IIf(strTitle = "", "Fichero", strTitle)

      ' Clear out the current filters, and add our own.
      .Filters.Clear
      '.Filters.Add "Access Databases", "*.MDB"
      .Filters.Add strFilterTexto, strFilterExtension
      '.Filters.Add "All Files", "*.*"
      .InitialFileName = strInitialFileName
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then

         'Loop through each file selected and add it to our list box.
         For Each varF In .SelectedItems
            'Me.FileList.AddItem varF
            strR = strR & varF & vbCrLf
         Next
         If Len(strR) > 2 Then strR = RecDerTop(strR, 2, 0)
         DimeFichero = strR
        
      Else
         'MsgBox "You clicked Cancel in the file dialog box."
         DimeFichero = Null
      End If
   End With
End Function

Public Function DimeFileSaveAs(strInitialName As String, strTitle As String) As String
    'Declare a variable as a FileDialog object

    'Create a FileDialog object as a File Picker dialog box.
    Dim fd As FileDialog, strFile As String
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    'Declare a variable to contain the path
    'of each selected item. Even though the path is a String,
    'the variable must be a Variant because For Each...Next
    'routines only work with Variants and Objects.
    Dim vrtSelectedItem As Variant

    'Use a With...End With block to reference the FileDialog object.
    With fd
        'Set the initial path to the C:\ drive.
        .InitialFileName = strInitialName
        .Title = strTitle
        .AllowMultiSelect = False
        'Use the Show method to display the File Picker dialog box and return the user's action.
        'If the user presses the action button...
        If .Show = -1 Then

            'Step through each string in the FileDialogSelectedItems collection.

            For Each vrtSelectedItem In .SelectedItems

                'vrtSelectedItem is a String that contains the path of each selected item.
                'You can use any file I/O functions that you want to work with this path.
                'This example simply displays the path in a message box.
                'MsgBox "Selected item's path: " & vrtSelectedItem
                strFile = vrtSelectedItem

            Next vrtSelectedItem
        'If the user presses Cancel...
        Else
            Exit Function
        End If
    End With
    'Set the object variable to Nothing.
    Set fd = Nothing
    DimeFileSaveAs = strFile
End Function




