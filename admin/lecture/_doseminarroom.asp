<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim idx,tdate,ttime,usetime,usestart
dim groupname,username,userphone
dim usepeople,isusing,roomid,etc
dim mode

idx = request("idx")
roomid = request("roomid")
mode = request("mode")
tdate = request("tdate")
usetime	= request("usetime")
groupname = request("groupname")
username = request("username")
userphone	= request("userphone")
etc = request("etc")
usepeople	= request("usepeople")
isusing	= request("isusing")

usestart = Cstr(tdate)

dim sqlStr,Fusername

if isusing <> "N" then
	sqlStr = "select top 1 username from [db_shop].[dbo].tbl_seminar_room" + VBCrlf
	sqlStr = sqlStr + " where usestart='" + tdate + "'" + VBCrlf
	sqlStr = sqlStr + " and usetime > " + Cstr(usetime-4) + "" + VBCrlf
	sqlStr = sqlStr + " and usetime < " + Cstr(usetime+4) + "" + VBCrlf
	sqlStr = sqlStr + " and roomid='" + Cstr(roomid) + "'"
	sqlStr = sqlStr + " and isusing<>'N'"
	rsget.Open sqlStr,dbget,1
   if  not rsget.EOF  then
		  Fusername   =  rsget("username")
   end if
   rsget.close

	if Fusername <> "" then
		response.write "<script language='JavaScript'>alert('" + Fusername + "님의 예약과 겹칩니다.\n다시 확인하시고 선택해주세요...');history.back(-1);</script>"
		dbget.close()	:	response.End
	end if
end if

if (idx<>"") then
	sqlStr = "update [db_shop].[dbo].tbl_seminar_room" + VBCrlf
	sqlStr = sqlStr + " set roomid ='" + roomid + "'" + VBCrlf
	sqlStr = sqlStr + " ,usestart ='" + usestart + "'" + VBCrlf
	sqlStr = sqlStr + " ,usetime = " + usetime + "" + VBCrlf
	sqlStr = sqlStr + " ,groupname='" + groupname + "'" + VBCrlf
	sqlStr = sqlStr + " ,username='" + username + "'" + VBCrlf
	sqlStr = sqlStr + " ,userphone='" + userphone + "'" + VBCrlf
	sqlStr = sqlStr + " ,usepeople=" + usepeople + "" + VBCrlf
	sqlStr = sqlStr + " ,etc='" + html2db(etc) + "'" + VBCrlf
	sqlStr = sqlStr + " ,isusing='" + isusing + "'" + VBCrlf
	sqlStr = sqlStr + " where idx=" + CStr(idx)
	rsget.Open sqlStr,dbget,1
else
	sqlStr = "insert into [db_shop].[dbo].tbl_seminar_room" + VBCrlf
	sqlStr = sqlStr + " (roomid,usestart,usetime,groupname,username,userphone" + VBCrlf
	sqlStr = sqlStr + " ,usepeople,etc)" + VBCrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " '" + Cstr(roomid) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(usestart) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(usetime) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(groupname) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(username) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(userphone) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(usepeople) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + html2db(etc) + "'" + VBCrlf
	sqlStr = sqlStr + " )"
	rsget.Open sqlStr,dbget,1
end if

%>

<script language="javascript">
alert('수정되었습니다.');
window.opener.location.reload();
self.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->