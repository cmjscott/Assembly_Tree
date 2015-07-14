Public Class ListBoxItem

    Private mvarItemText As String
    Private mvarItemText2 As String
    Private mvarItemIntegerData As Integer
    Private mvarItemIntegerData2 As Integer
    Private mvarItemIntegerData3 As Integer
    Private mvarItemIntegerData4 As Integer
    Private mvarItemIntegerData5 As Integer
    Private mvarItemRealData As Single
    Private mvarItemBooleanData As Boolean


    Public Property ItemText() As String

        Get
            Return mvarItemText
        End Get

        Set(ByVal Value As String)
            mvarItemText = Value
        End Set

    End Property


    Public Property ItemText2() As String

        Get
            Return mvarItemText2
        End Get

        Set(ByVal Value As String)
            mvarItemText2 = Value
        End Set

    End Property


    Public Property ItemIntegerData() As Integer

        Get
            Return mvarItemIntegerData
        End Get

        Set(ByVal Value As Integer)
            mvarItemIntegerData = Value
        End Set

    End Property


    Public Property ItemIntegerData2() As Integer

        Get
            Return mvarItemIntegerData2
        End Get

        Set(ByVal Value As Integer)
            mvarItemIntegerData2 = Value
        End Set

    End Property


    Public Property ItemIntegerData3() As Integer

        Get
            Return mvarItemIntegerData3
        End Get

        Set(ByVal Value As Integer)
            mvarItemIntegerData3 = Value
        End Set

    End Property


    Public Property ItemIntegerData4() As Integer

        Get
            Return mvarItemIntegerData4
        End Get

        Set(ByVal Value As Integer)
            mvarItemIntegerData4 = Value
        End Set

    End Property


    Public Property ItemIntegerData5() As Integer

        Get
            Return mvarItemIntegerData5
        End Get

        Set(ByVal Value As Integer)
            mvarItemIntegerData5 = Value
        End Set

    End Property


    Public Property ItemRealData() As Single

        Get
            Return mvarItemRealData
        End Get

        Set(ByVal Value As Single)
            mvarItemRealData = Value
        End Set

    End Property


    Public Property ItemBooleanData() As Boolean

        Get
            Return mvarItemBooleanData
        End Get

        Set(ByVal Value As Boolean)
            mvarItemBooleanData = Value
        End Set

    End Property


    Public Sub New()

        mvarItemText = ""
        mvarItemIntegerData = 0
        mvarItemIntegerData2 = 0
        mvarItemRealData = 0.0
        mvarItemBooleanData = False

    End Sub

End Class
