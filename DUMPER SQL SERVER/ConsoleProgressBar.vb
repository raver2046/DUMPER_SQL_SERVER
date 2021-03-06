
Imports System
Imports System.Text
Imports Microsoft.VisualBasic.Strings


Public Class ConsoleProgressBar
    Private m_length As Integer
    Private m_left As Integer
    Private m_right As Integer
    Private m_progressBarRow As Integer
    Private m_messageBarRow As Integer
    Private m_percentPosition As Integer
    Private m_maximumValue As Long
    Private m_currentValue As Long
    Private m_currentBarLength As Integer

#Region "Properties"
    Public ReadOnly Property Length() As Integer
        Get
            Return m_length
        End Get
    End Property

    Public ReadOnly Property Left() As Integer
        Get
            Return m_left
        End Get
    End Property

    Public ReadOnly Property Right() As Integer
        Get
            Return m_right
        End Get
    End Property

    Public ReadOnly Property ProgressBarRow() As Integer
        Get
            Return m_progressBarRow
        End Get
    End Property

    Public ReadOnly Property MessageBarRow() As Integer
        Get
            Return m_messageBarRow
        End Get
    End Property

    Public ReadOnly Property PercentPosition() As Integer
        Get
            Return m_percentPosition
        End Get
    End Property

    Public ReadOnly Property MaximumValue() As Long
        Get
            Return m_maximumValue
        End Get
    End Property

    Public Property CurrentValue() As Long
        Get
            Return m_currentValue
        End Get
        Set(ByVal value As Long)
            m_currentValue = value
        End Set
    End Property

    Public Property CurrentBarLength() As Integer
        Get
            Return m_currentBarLength
        End Get
        Set(ByVal value As Integer)
            m_currentBarLength = value
        End Set
    End Property
#End Region

    '----------------------------------------------------------------------
    ' NAME    : New
    ' PURPOSE : Constructor for this class.
    '----------------------------------------------------------------------
    Public Sub New(ByVal MaximumValue As Long)
        m_length = Console.WindowWidth - 10
        m_left = 7
        m_right = m_left + m_length + 1
        m_progressBarRow = 5
        m_messageBarRow = m_progressBarRow + 1
        m_percentPosition = 4
        m_maximumValue = MaximumValue
        m_currentValue = 0
        Initialize()
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : Initialize
    ' PURPOSE : Call the corresponding initialization methods for the
    '           percent complete, the progress bar, and the message bar.
    '----------------------------------------------------------------------
    Private Sub Initialize()
        InitializePercentComplete()
        InitializeProgressBar()
        InitializeMessageBar()
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : InitializePercentComplete
    ' PURPOSE : Initialize the percent complete area by making Magenta the
    '           ForegroundColor and printing a '%'.
    '----------------------------------------------------------------------
    Private Sub InitializePercentComplete()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor

        Console.ForegroundColor = ConsoleColor.Magenta
        Console.SetCursorPosition(m_percentPosition, m_progressBarRow)
        Console.Write("%")
        Console.ForegroundColor = originalForegroundColor
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : InitializeProgressBar
    ' PURPOSE : Initialize the progress bar by printing a green '[' at the
    '           left edge and a green ']' at the right edge.  Then, print a
    '           string of spaces to fill in the background color.
    '----------------------------------------------------------------------
    Private Sub InitializeProgressBar()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor
        Dim originalBackgroundColor As ConsoleColor = Console.BackgroundColor

        Console.ForegroundColor = ConsoleColor.Black
        Console.BackgroundColor = ConsoleColor.Green
        Console.SetCursorPosition(m_left, m_progressBarRow)
        Console.Write("[")
        Console.SetCursorPosition(m_right, m_progressBarRow)
        Console.Write("]")

        Console.SetCursorPosition(m_left + 1, m_progressBarRow)
        Console.Write(New String(Space(1), m_length))

        Console.ForegroundColor = originalForegroundColor
        Console.BackgroundColor = originalBackgroundColor
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : InitializeMessageBar
    ' PURPOSE : Print a Magenta '0' at the left edge of the progress bar,
    '           and m_maximumValue at the right edge.
    '----------------------------------------------------------------------
    Private Sub InitializeMessageBar()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor

        Console.ForegroundColor = ConsoleColor.Magenta
        Console.SetCursorPosition(m_left, MessageBarRow)
        Console.Write(m_currentValue.ToString)
        Console.SetCursorPosition(m_right - 5, MessageBarRow)
        Console.Write("{0,6}", m_maximumValue.ToString)

        Console.ForegroundColor = originalForegroundColor
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : Update
    ' PURPOSE : Update m_currentValue and calculate the length of the
    '           progress bar.
    '----------------------------------------------------------------------
    Public Sub Update(ByVal CurrentValue As Long)
        m_currentValue = CurrentValue
        m_currentBarLength = CInt((m_currentValue / m_maximumValue) * m_length)
        Refresh()
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : Refresh
    ' PURPOSE : Call the corresponding methods to refresh the percent, the
    '           progress bar, and the message bar.
    '----------------------------------------------------------------------
    Private Sub Refresh()
        UpdatePercentComplete()
        UpdateProgressBar()
        UpdateMessageBar()
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : UpdatePercentComplete
    ' PURPOSE : Update the percent counter to show the current percent
    '           complete.
    '----------------------------------------------------------------------
    Private Sub UpdatePercentComplete()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor

        Console.ForegroundColor = ConsoleColor.Magenta
        Console.SetCursorPosition(0, m_progressBarRow)
        Console.Write("{0,3}", CInt((m_currentValue / m_maximumValue) * 100).ToString)
        Console.ForegroundColor = originalForegroundColor
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : UpdateProgressBar
    ' PURPOSE : Update the progress bar to show the percent complete.
    '----------------------------------------------------------------------
    Private Sub UpdateProgressBar()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor
        Dim originalBackgroundColor As ConsoleColor = Console.BackgroundColor

        Console.ForegroundColor = ConsoleColor.Black
        Console.BackgroundColor = ConsoleColor.Green
        Console.SetCursorPosition(m_left + 1, m_progressBarRow)
        Dim progress As New String("▒", m_currentBarLength)
        Console.Write(progress)

        Console.ForegroundColor = originalForegroundColor
        Console.BackgroundColor = originalBackgroundColor
    End Sub

    '----------------------------------------------------------------------
    ' NAME    : UpdateMessageBar
    ' PURPOSE : Update the file counter in the message bar.
    '----------------------------------------------------------------------
    Private Sub UpdateMessageBar()
        Dim originalForegroundColor As ConsoleColor = Console.ForegroundColor

        Console.ForegroundColor = ConsoleColor.Magenta
        Console.SetCursorPosition((m_right / 2) - 5, MessageBarRow)
        Console.Write("")
        Console.ForegroundColor = ConsoleColor.Green
        Console.Write(m_currentValue.ToString)
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.Write(" / " & m_maximumValue.ToString)
        Console.ForegroundColor = originalForegroundColor
    End Sub
End Class