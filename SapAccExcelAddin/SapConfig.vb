Imports System.Configuration
Public Class SapConnectionSection
    Inherits ConfigurationSection

    ' Declare the SapConnectionCollection collection property.
    <ConfigurationProperty("connections", IsDefaultCollection:=False), ConfigurationCollection(GetType(SapConnectionCollection), AddItemName:="add", ClearItemsName:="clear", RemoveItemName:="remove")>
    Public Property SapConnections() As SapConnectionCollection
        Get
            Dim sapConnectionCollection As SapConnectionCollection = CType(MyBase.Item("connections"), SapConnectionCollection)
            Return sapConnectionCollection
        End Get

        Set(ByVal value As SapConnectionCollection)
            Dim sapConnectionCollection As SapConnectionCollection = value
        End Set
    End Property

    Public Sub New()
        Dim sapConnection As New SapConnectionConfigElement()
        SapConnections.Add(sapConnection)
    End Sub

End Class

Public Class SapConnectionCollection
        Inherits System.Configuration.ConfigurationElementCollection

        Public Overrides ReadOnly Property CollectionType() As ConfigurationElementCollectionType
            Get
                Return ConfigurationElementCollectionType.AddRemoveClearMap
            End Get
        End Property

        Protected Overrides Function CreateNewElement() As ConfigurationElement
            Return New SapConnectionConfigElement()
        End Function

        Protected Overrides Function GetElementKey(ByVal element As ConfigurationElement) As Object
            Return (CType(element, SapConnectionConfigElement)).Name
        End Function

        Default Public Shadows Property Item(ByVal index As Integer) As SapConnectionConfigElement
            Get
                Return CType(BaseGet(index), SapConnectionConfigElement)
            End Get
            Set(ByVal value As SapConnectionConfigElement)
                If BaseGet(index) IsNot Nothing Then
                    BaseRemoveAt(index)
                End If
                BaseAdd(value)
            End Set
        End Property
        Default Public Shadows ReadOnly Property Item(ByVal Name As String) As SapConnectionConfigElement
            Get
                Return CType(BaseGet(Name), SapConnectionConfigElement)
            End Get
        End Property
        Public Function IndexOf(ByVal sapConnection As SapConnectionConfigElement) As Integer
            Return BaseIndexOf(sapConnection)
        End Function

        Public Sub Add(ByVal sapConnection As SapConnectionConfigElement)
            BaseAdd(sapConnection)
        End Sub

    End Class

    Public Class SapConnectionConfigElement
        Inherits ConfigurationElement
        Public Sub New(ByVal Name As String, ByVal AppServerHost As String, ByVal SystemNumber As String, ByVal SystemID As String,
                   Optional ByVal Client As String = "", Optional ByVal Language As String = "",
                   Optional ByVal SncMode As String = "0", Optional ByVal SncPartnerName As String = "")
            Me.Name = Name
            Me.AppServerHost = AppServerHost
            Me.SystemNumber = SystemNumber
            Me.SystemID = SystemID
            Me.Client = Client
            Me.Language = Language
            Me.SncMode = SncMode
            Me.SncPartnerName = SncPartnerName
        End Sub
        Public Sub New()
        End Sub

        <ConfigurationProperty("Name", DefaultValue:="", IsRequired:=True, IsKey:=True)>
        Public Property Name() As String
            Get
                Return CStr(Me("Name"))
            End Get
            Set(ByVal value As String)
                Me("Name") = value
            End Set
        End Property

        <ConfigurationProperty("AppServerHost", DefaultValue:="", IsRequired:=True, IsKey:=False)>
        Public Property AppServerHost() As String
            Get
                Return CStr(Me("AppServerHost"))
            End Get
            Set(ByVal value As String)
                Me("AppServerHost") = value
            End Set
        End Property

        <ConfigurationProperty("SystemNumber", DefaultValue:="", IsRequired:=True, IsKey:=False)>
        Public Property SystemNumber() As String
            Get
                Return CStr(Me("SystemNumber"))
            End Get
            Set(ByVal value As String)
                Me("SystemNumber") = value
            End Set
        End Property

        <ConfigurationProperty("SystemID", DefaultValue:="", IsRequired:=True, IsKey:=False)>
        Public Property SystemID() As String
            Get
                Return CStr(Me("SystemID"))
            End Get
            Set(ByVal value As String)
                Me("SystemID") = value
            End Set
        End Property

        <ConfigurationProperty("Client", DefaultValue:="", IsRequired:=False, IsKey:=False)>
        Public Property Client() As String
            Get
                Return CStr(Me("Client"))
            End Get
            Set(ByVal value As String)
                Me("Client") = value
            End Set
        End Property

        <ConfigurationProperty("Language", DefaultValue:="", IsRequired:=False, IsKey:=False)>
        Public Property Language() As String
            Get
                Return CStr(Me("Language"))
            End Get
            Set(ByVal value As String)
                Me("Language") = value
            End Set
        End Property

        <ConfigurationProperty("SncMode", DefaultValue:="0", IsRequired:=False, IsKey:=False)>
        Public Property SncMode() As String
            Get
                Return CStr(Me("SncMode"))
            End Get
            Set(ByVal value As String)
                Me("SncMode") = value
            End Set
        End Property

        <ConfigurationProperty("SncPartnerName", DefaultValue:="", IsRequired:=False, IsKey:=False)>
        Public Property SncPartnerName() As String
            Get
                Return CStr(Me("SncPartnerName"))
            End Get
            Set(ByVal value As String)
                Me("SncPartnerName") = value
            End Set
        End Property
    End Class
