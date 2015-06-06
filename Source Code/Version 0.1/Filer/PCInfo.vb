' NOT FINISHED
Friend Class PCInfo
#Region "CPU"
    Shared Function GetCPUName()
        Try
            Dim name As String
            name = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "ProcessorNameString", Nothing)
            Return name
        Catch ex As Exception : End Try
    End Function

    Shared Function GetCPUClockSpeed()
        Dim speed As String
        speed = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "~MHz", Nothing)
        Return speed
    End Function
#End Region
End Class
' NOT FINISHED