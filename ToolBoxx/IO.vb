Imports System.IO
Imports System.Xml.Serialization

Public Class IO

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="DirPath"></param>
    ''' <param name="IncludeSubFolders"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function GetFolderSize(ByVal DirPath As String, Optional IncludeSubFolders As Boolean = True) As Object

        Dim lngDirSize As Long
        Dim objFileInfo As FileInfo
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo

        Try

            'add length of each file
            For Each objFileInfo In objDir.GetFiles()
                lngDirSize += objFileInfo.Length
            Next

            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    lngDirSize += GetFolderSize(objSubFolder.FullName)
                Next
            End If

            Return lngDirSize
        Catch Ex As Exception
            Return Ex
        End Try

    End Function

    ''' <summary>
    ''' Used to deserialize an xml file into an object
    ''' </summary>
    ''' <param name="path">Full path to xml file</param>
    ''' <param name="type">Type (use GetType(object)) of object to deserialize</param>
    ''' <returns>Returns an object filled with the contents of the xml pointed to with 'path'.  Returns exception if error occurs</returns>
    ''' <remarks></remarks>
    Shared Function DeserializeXML(ByVal path As String, ByVal type As Type) As Object

        Dim reader As StreamReader = New StreamReader(path)
        Dim serializer As XmlSerializer = New XmlSerializer(type)

        Try
            Dim returnObj = serializer.Deserialize(reader)
            Return returnObj
        Catch e As Exception
            Return e
        Finally
            reader.Close()
        End Try

    End Function

    ''' <summary>
    ''' Used to serialize an object to an xml file
    ''' </summary>
    ''' <param name="path">Full path to save xml file</param>
    ''' <param name="obj">The object to serialize</param>
    ''' <param name="append">set append to TRUE to append object values to file instead of overwriting</param>
    ''' <returns>Returns TRUE if successful.  Returns exception if error occurs</returns>
    ''' <remarks></remarks>
    Shared Function SerializeXML(ByVal path As String, ByVal obj As Object, Optional ByVal append As Boolean = False) As Object
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(Object))
        Dim writer As StreamWriter = New StreamWriter(path, append)
        Try
            serializer.Serialize(writer, obj)
            Return True
        Catch e As Exception
            Return e
        Finally
            writer.Close()
        End Try
    End Function
End Class
