Option Explicit On
Option Strict On

Imports System
Imports NLog
Imports Topshelf
Imports FluentScheduler

Module modMain
#Region "Global Variables"
    Public m_logger As NLog.Logger = LogManager.GetCurrentClassLogger()
    Public retCode As Integer = 0
    Public ActiveJobs As Integer
#End Region
#Region "Startup"
    Public Sub Main()
        m_logger.Info("Starting Service: " & DateTime.Now.ToUniversalTime.ToString)
        HostFactory.Run(
        Sub(x)
            x.Service(Of ServiceClass)(
                Sub(s)
                    s.ConstructUsing(Function(name) New ServiceClass())
                    s.WhenStarted(Sub(tc) tc.StartService())
                    s.WhenStopped(Sub(tc) tc.StopService())
                    's.WhenPaused(Sub(tc) tc.PauseService())
                    's.WhenContinued(Sub(tc) tc.ContinueService())
                End Sub)
            'x.RunAsLocalSystem()
            x.SetDescription(My.Settings.ServiceDescription)
            x.SetDisplayName(My.Settings.ServiceDisplayName)
            x.SetServiceName(My.Settings.ServiceName)
            x.DependsOnEventLog
        End Sub)
    End Sub
#End Region

#Region "Scheduler Event Handlers"
    Sub JobManager_TaskException(a As JobExceptionInfo)
        m_logger.Fatal(a.Exception)
        ActiveJobs -= 1
    End Sub

    Sub JobManager_JobStart(a As JobStartInfo)
        m_logger.Debug("Job Start: [" & a.Name & "] ")
        ActiveJobs += 1
    End Sub

    Sub JobManager_JobEnd(a As JobEndInfo)
        m_logger.Debug("Job End: [" & a.Name & "] Duration - " & a.Duration.TotalMilliseconds.ToString & " (ms) Next Run - " & a.NextRun.ToString)
        ActiveJobs -= 1
    End Sub

#End Region

#Region "Service Control"
    Public Class ServiceClass

        Public Sub StartService()
            m_logger.Info("Service starting")
            JobManager.Initialize(New ScheduleRegistryRegistrySample())
            AddHandler JobManager.JobException, AddressOf JobManager_TaskException
            AddHandler JobManager.JobStart, AddressOf JobManager_JobStart
            AddHandler JobManager.JobEnd, AddressOf JobManager_JobEnd

            m_logger.Info("Service started")
        End Sub

        Public Sub StopService()
            m_logger.Info("Service stopping")
            JobManager.StopAndBlock()
            JobManager.RemoveAllJobs()
            m_logger.Info("Service stopped")
        End Sub
        'Public Sub PauseService()
        '    m_logger.Info("Service pausing")
        '    JobManager.StopAndBlock()
        '    m_logger.Info("Service paused")
        'End Sub

        'Public Sub ContinueService()
        '    JobManager.Start()
        '    m_logger.Info("Service Restarted")
        'End Sub
    End Class
#End Region

#Region "Schedule Setup"
    Public Class ScheduleRegistryRegistrySample
        Inherits Registry
        Public Sub New()
            Schedule(Of StartUpProcess)().WithName("StartUp Job").NonReentrant.ToRunNow.DelayFor(My.Settings.SchedulePauseBeforeFirstRunSeconds).Seconds()
            Schedule(Of MainProcess)().WithName("Main Job").NonReentrant.ToRunEvery(My.Settings.ScheduleRunSyncMinutes).Minutes.Between(My.Settings.ScheduleMinTime, 0, My.Settings.ScheduleMaxTime, 0)
            Schedule(Of ProcessOne)().WithName("Step Job").AndThen(Of ProcessTwo).AndThen(Of ProcessThree).NonReentrant.ToRunEvery(My.Settings.ScheduleRunSyncMinutes + 1).Minutes.Between(My.Settings.ScheduleMinTime, 0, My.Settings.ScheduleMaxTime, 0)
        End Sub
    End Class

    Public Class MainProcess
        Implements IJob

        Private Sub IJob_Execute() Implements IJob.Execute
            MainProcessToStart()
        End Sub
    End Class

    Public Class StartUpProcess
        Implements IJob

        Private Sub IJob_Execute() Implements IJob.Execute
            MainProcessToStart()
        End Sub
    End Class

    Public Class ProcessOne
        Implements IJob

        Private Sub IJob_Execute() Implements IJob.Execute
            ProcessOneSub()
        End Sub
    End Class
    Public Class ProcessTwo
        Implements IJob

        Private Sub IJob_Execute() Implements IJob.Execute
            ProcessTwoSub()
        End Sub
    End Class
    Public Class ProcessThree
        Implements IJob

        Private Sub IJob_Execute() Implements IJob.Execute
            ProcessThreeSub()
        End Sub
    End Class
#End Region

#Region "Main Job"
    Sub MainProcessToStart()
        Try
            ' Main Process Start
            m_logger.Info("Main process starts")
        Catch ex As Exception
            m_logger.Fatal(ex)
        End Try
    End Sub
    Sub ProcessOneSub()
        Try
            m_logger.Info("Process One starts")
        Catch ex As Exception
            m_logger.Fatal(ex)
        End Try
    End Sub
    Sub ProcessTwoSub()
        Try
            m_logger.Info("Process Two starts")
        Catch ex As Exception
            m_logger.Fatal(ex)
        End Try
    End Sub
    Sub ProcessThreeSub()
        Try
            m_logger.Info("Process Three starts")
        Catch ex As Exception
            m_logger.Fatal(ex)
        End Try
    End Sub

#End Region
End Module
