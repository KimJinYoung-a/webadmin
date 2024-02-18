<%
Function GetGuid()
    Dim TypeLib, Guid
    Set TypeLib = Server.CreateObject("Scriptlet.Typelib") 
    Guid = Left(CStr(TypeLib.Guid), 38)
    Guid = Replace(Guid, "{", "")
    GetGuid = Replace(Guid, "}", "") 
    Set TypeLib = Nothing 
End Function 
%>