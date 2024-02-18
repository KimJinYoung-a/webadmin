<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
dim idx,tdate,ttime,usetime,usestart
dim groupname,username,userphone
dim usepeople,isusing,roomid,etc,btime

idx = request("idx")
roomid = request("roomid")
tdate = request("tdate")
usetime	= request("usetime")
btime	= request("btime")
groupname = request("groupname")
username = request("username")
userphone	= request("userphone")
etc = request("etc")
usepeople	= request("usepeople")
isusing	= request("isusing")

usestart = Cstr(tdate)

dim sqlStr

if (idx<>"") then
	sqlStr = "update [db_shop].[dbo].tbl_seminar_room" + VBCrlf
	sqlStr = sqlStr + " set roomid ='" + roomid + "'" + VBCrlf
	sqlStr = sqlStr + " ,usestart ='" + usestart + "'" + VBCrlf
	sqlStr = sqlStr + " ,usetime = " + usetime + "" + VBCrlf
	sqlStr = sqlStr + " ,basictime = " + btime + "" + VBCrlf
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
	sqlStr = sqlStr + " (roomid,usestart,usetime,basictime,groupname,username,userphone" + VBCrlf
	sqlStr = sqlStr + " ,usepeople,etc)" + VBCrlf
	sqlStr = sqlStr + " values("
	sqlStr = sqlStr + " '" + Cstr(roomid) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(usestart) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(usetime) + "'" + VBCrlf
	sqlStr = sqlStr + " ,'" + Cstr(btime) + "'" + VBCrlf
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