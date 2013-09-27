Imports System.Security

Module Module1

    Sub Main()
        Console.WriteLine("Killing Open Item Process...")
        Dim cn As New OAConnection.Connection

        Dim Process() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcessesByName("OpenItems")
        On Error Resume Next

        For Each P In Process
            P.Kill()
        Next

        Console.WriteLine("Restarting Open Item Process...")
        System.Diagnostics.Process.Start(My.Application.Info.DirectoryPath & "\OpenItems.exe")
    End Sub

End Module
