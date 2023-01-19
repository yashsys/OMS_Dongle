Imports System.ComponentModel
Imports System.Configuration.Install

<RunInstaller(True)> Public Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

    <System.Diagnostics.DebuggerNonUserCode()> Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    Private components As System.ComponentModel.IContainer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

        Me.ServiceProcessInstaller1 = New System.ServiceProcess.ServiceProcessInstaller
        Me.ServiceInstaller1 = New System.ServiceProcess.ServiceInstaller

        Me.ServiceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem

        Me.ServiceProcessInstaller1.Password = Nothing

        Me.ServiceProcessInstaller1.Username = Nothing
        Me.ServiceInstaller1.ServiceName = "OMS_Dongle"

        Me.ServiceInstaller1.StartType = System.ServiceProcess.ServiceStartMode.Automatic
        Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller1, Me.ServiceInstaller1})

    End Sub

    Friend WithEvents ServiceProcessInstaller1 As System.ServiceProcess.ServiceProcessInstaller
    Friend WithEvents ServiceInstaller1 As System.ServiceProcess.ServiceInstaller

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub
End Class

