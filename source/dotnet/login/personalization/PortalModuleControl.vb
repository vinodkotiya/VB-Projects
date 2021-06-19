Imports System
Imports System.Web
Imports System.Web.UI
Imports Personalization

Public Class PortalModuleControl : Inherits UserControl
  
    Public ReadOnly Property UserState As UserState    
        Get
            Dim myState As UserState = CType(Context.Items("UserState"), UserState)
            if (myState Is Nothing) Then
                Throw New Exception("No UserState Loaded!!!")
			Else
				Return myState
            End If
        End Get
    End Property

End Class