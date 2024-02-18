<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim sqlStr, buf

dim mode, orderserial
dim arrchk
dim chk, itemid, itemoption

mode = requestCheckVar(html2db(request("mode")),32)
orderserial = requestCheckVar(html2db(request("orderserial")),32)
arrchk = request("arrchk")
chk = requestCheckVar(html2db(request("chk")),32)
itemid = requestCheckVar(html2db(request("itemid")),32)
itemoption = requestCheckVar(html2db(request("itemoption")),32)

if (mode = "cancelselected") then

	arrchk = "-1" + arrchk

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr + " set matchState = 'D', etcFinUser = '" + CStr(session("ssBctId")) + "' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and outmallorderseq in (" + CStr(arrchk) + ") "
	''rw sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('삭제 되었습니다.'); opener.location.reload();window.close();</script>"
	dbget.Close() : response.end

elseif (mode = "modifymatchitem") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr + " set matchState = 'A', matchitemid = " + CStr(itemid) + ", matchitemoption = '" + CStr(itemoption) + "', etcFinUser = '" + CStr(session("ssBctId")) + "' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and outmallorderseq = " + CStr(chk) + " "
	''rw sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('변경 되었습니다.'); opener.location.reload();window.close();</script>"
	dbget.Close() : response.end

elseif (mode = "modifymatchitemno") then

	sqlStr = " update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr + " set changeitemid = " + CStr(itemid) + ", changeitemoption = '" + CStr(itemoption) + "', etcFinUser = '" + CStr(session("ssBctId")) + "' "
	sqlStr = sqlStr + " where orderserial = '" + CStr(orderserial) + "' and outmallorderseq = " + CStr(chk) + " "
	''rw sqlStr
	dbget.Execute sqlStr

	response.write "<script>alert('변경 되었습니다.'); opener.location.reload();window.close();</script>"
	dbget.Close() : response.end

end if

''on error Goto 0

''response.end
%>

<script>alert('저장되었습니다.');</script>
<script>location.replace('<%= refer %>');</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
