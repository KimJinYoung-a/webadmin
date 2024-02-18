<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim idx, stitle , reservationdate , state , mode , lastupdate
Dim viewno , worktext
Dim partwdid , partmdid , lastadminid
Dim simg1 ,simg2 ,simg3 ,simg4 ,simg5
Dim surl1 ,surl2 ,surl3 ,surl4 ,surl5
Dim salt1 ,salt2 ,salt3 ,salt4 ,salt5
Dim listidx, subidx , itemid , sortnum
Dim mainTopBGColor , subtitle

lastadminid = session("ssBctId")

idx		= RequestCheckVar(request("idx"),10)
viewno	= RequestCheckVar(request("viewno"),10)
stitle	= RequestCheckVar(request("stitle"),50)
reservationdate = RequestCheckVar(request("reservationdate"),10)
state	 = RequestCheckVar(request("state"),2)
worktext = RequestCheckVar(request("worktext"),800)

simg1 = RequestCheckVar(request("simg1"),200)
simg2 = RequestCheckVar(request("simg2"),200)
simg3 = RequestCheckVar(request("simg3"),200)
simg4 = RequestCheckVar(request("simg4"),200)
simg5 = RequestCheckVar(request("simg5"),200)

surl1 = RequestCheckVar(request("surl1"),50)
surl2 = RequestCheckVar(request("surl2"),50)
surl3 = RequestCheckVar(request("surl3"),50)
surl4 = RequestCheckVar(request("surl4"),50)
surl5 = RequestCheckVar(request("surl5"),50)

salt1 = RequestCheckVar(request("salt1"),150)
salt2 = RequestCheckVar(request("salt2"),150)
salt3 = RequestCheckVar(request("salt3"),150)
salt4 = RequestCheckVar(request("salt4"),150)
salt5 = RequestCheckVar(request("salt5"),150)

partmdid = RequestCheckVar(request("partmdid"),32)
partwdid = RequestCheckVar(request("partwdid"),32)

mode = RequestCheckVar(request("mode"),10)

mainTopBGColor	= RequestCheckVar(request("mainTopBGColor"),10) 'colorcode

listidx		= Request("listidx")
subidx		= Request("subidx")
itemid		= Request("subItemid")
sortnum		= Request("sortnum")

subtitle = RequestCheckVar(request("subtitle"),100)

if idx = "" then
	idx = 0
end If

If listidx = "" then
	if idx = 0 then
		mode = "add"
	else
		mode = "edit"
	end if
End If 

dim sqlStr

Select Case mode

	Case "add"
		sqlStr = " insert into db_sitemaster.dbo.tbl_mobile_showbanner_list " + VbCrlf
		sqlStr = sqlStr + " (stitle , simg1 , simg2 , simg3 , simg4 , simg5 , reservationdate " + VbCrlf
		sqlStr = sqlStr + " ,worktext , partMDid , partWDid , state , surl1 , surl2 , surl3 , surl4 , surl5 " + VbCrlf
		sqlStr = sqlStr + " ,salt1 , salt2 , salt3 , salt4 , salt5 , colorcode , viewno , subtitle )" + VbCrlf
		sqlStr = sqlStr + " values(" + VbCrlf
		sqlStr = sqlStr + " '" + stitle + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + simg1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + simg2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + simg3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + simg4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + simg5 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + reservationdate + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + worktext + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partmdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + partwdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + state + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + surl1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + surl2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + surl3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + surl4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + surl5 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + salt1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + salt2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + salt3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + salt4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,'" + salt5 + "'" + VbCrlf    
		sqlStr = sqlStr + " ,'" + mainTopBGColor + "'" + VbCrlf
		sqlStr = sqlStr + " ," + viewno + "" + VbCrlf
		sqlStr = sqlStr + " ,'" + subtitle + "'" + VbCrlf
		sqlStr = sqlStr + " )"

		'response.write sqlStr
		dbget.Execute sqlStr

		sqlStr = "select IDENT_CURRENT('db_sitemaster.dbo.tbl_mobile_showbanner_list') as idx"
		rsget.Open sqlStr, dbget, 1
		If Not Rsget.Eof then
			idx = rsget("idx")
		end if
		rsget.close

	Case "edit"

		sqlStr = " update  db_sitemaster.dbo.tbl_mobile_showbanner_list " + VbCrlf
		sqlStr = sqlStr + " set " + VbCrlf
		sqlStr = sqlStr + " stitle='" + stitle + "'" + VbCrlf
		sqlStr = sqlStr + " ,simg1='" + simg1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,simg2='" + simg2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,simg3='" + simg3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,simg4='" + simg4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,simg5='" + simg5 + "'" + VbCrlf
		sqlStr = sqlStr + " ,reservationdate='" + reservationdate + "'" + VbCrlf
		sqlStr = sqlStr + " ,state='" + state + "'" + VbCrlf
		sqlStr = sqlStr + " ,worktext='" + worktext + "'" + VbCrlf
		sqlStr = sqlStr + " ,partmdid='" + partmdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,partwdid='" + partwdid + "'" + VbCrlf
		sqlStr = sqlStr + " ,lastadminid='" + lastadminid + "'" + VbCrlf
		sqlStr = sqlStr + " ,lastupdate=getdate()" + VbCrlf
		sqlStr = sqlStr + " ,surl1='" + surl1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,surl2='" + surl2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,surl3='" + surl3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,surl4='" + surl4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,surl5='" + surl5 + "'" + VbCrlf
		sqlStr = sqlStr + " ,salt1='" + salt1 + "'" + VbCrlf
		sqlStr = sqlStr + " ,salt2='" + salt2 + "'" + VbCrlf
		sqlStr = sqlStr + " ,salt3='" + salt3 + "'" + VbCrlf
		sqlStr = sqlStr + " ,salt4='" + salt4 + "'" + VbCrlf
		sqlStr = sqlStr + " ,salt5='" + salt5 + "'" + VbCrlf
		sqlStr = sqlStr + " ,colorcode='" + mainTopBGColor + "'" + VbCrlf
		sqlStr = sqlStr + " ,viewno=" + viewno + "" + VbCrlf
		sqlStr = sqlStr + " ,subtitle='" + subtitle + "'" + VbCrlf

		sqlStr = sqlStr + " where showidx=" + CStr(idx)
		dbget.Execute sqlStr

	Case "subadd"
		'subitem 신규 등록
		sqlStr = "Insert Into [db_sitemaster].[dbo].tbl_mobile_showbanner_subitem " &_
					" (showidx, itemid , sortnum) values " &_
					" ('" & listidx  & "'" &_
					" ,'" & itemid &"'" &_
					" ,'" & sortnum &"')"
'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

	Case "submodify"
		'subitem 수정
		sqlStr = "Update [db_sitemaster].[dbo].tbl_mobile_showbanner_subitem " &_
				" Set itemid='" & itemid & "'" &_
				" 	,sortnum='" & sortnum & "'" &_
				" 	,isusing='" & isusing & "'" &_
				" Where showitemidx=" & subidx
		dbget.Execute(sqlStr)
'		response.write sqlStr
'		response.end
		dbget.Execute(sqlStr)

End Select

dim referer
referer = request.ServerVariables("HTTP_REFERER")
If mode = "subadd"  Or mode = "submodify" then
	Response.write "<script>alert('저장했습니다.');window.opener.document.location.href = window.opener.document.URL;self.close();</script>"
Else
	response.write "<script>alert('저장되었습니다.');</script>"
	response.write "<script>location.href='" & manageUrl & "/admin/mobile/showbanner/popShowbannerEdit.asp?idx=" + Cstr(idx) + "&reload=on'</script>"
End If 

%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->	
