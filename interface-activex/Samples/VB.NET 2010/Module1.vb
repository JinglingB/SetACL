Module Module1

    Sub Main()

		ListPermissions("c:\windows")

    End Sub

	Public Function ListPermissions(ByVal sFileName As String) As Boolean

		' Declare
		Dim objSetACL As New SETACLLib.SetACL
		Dim nResult As Integer

		' Add message handler
		AddHandler objSetACL.MessageEvent, AddressOf SetAclMessage

		' Enumerate users and permissions
		nResult = objSetACL.SetObject(sFileName, SETACLLib.SE_OBJECT_TYPE.SE_FILE_OBJECT)
		If nResult <> 0 Then Return False
		nResult = objSetACL.SetAction(SETACLLib.ACTIONS.ACTN_LIST)
		If nResult <> 0 Then Return False
		nResult = objSetACL.SetListOptions(SETACLLib.LISTFORMATS.LIST_CSV, SETACLLib.SDINFO.ACL_DACL, True, SETACLLib.LISTNAMES.LIST_NAME_SID)
		If nResult <> 0 Then Return False
		nResult = objSetACL.Run
		If nResult <> 0 Then Return False

		' Return
		Return True
	End Function

	Public Sub SetAclMessage(ByVal sMessage As String)

		MsgBox(sMessage)

	End Sub


End Module
