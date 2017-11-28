Imports System.Reflection

Public Class AssemblyVersionUtil

    Private Shared m_Description As String = ""

    Public Shared Function GetAssemblyDescription() As String
        'https://visualstudiomagazine.com/articles/2016/12/01/assembly-by-name-and-number.aspx
        If String.IsNullOrEmpty(m_Description) Then
            Dim a As Assembly = Assembly.GetExecutingAssembly()
            Dim attrs As Object() = a.GetCustomAttributes(GetType(AssemblyDescriptionAttribute), True)
            If attrs.Length > 0 Then
                Dim attr As AssemblyDescriptionAttribute = CType(attrs(0), AssemblyDescriptionAttribute)
                m_Description = attr.Description
            End If
        End If
        Return m_Description
    End Function

End Class
