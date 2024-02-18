<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 당첨자
' History : 2009.04.17 최초생성자 모름
'			2016.06.30 한용민 수정
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim username, reqname,reqphone,reqhp,reqzip, reqaddr1,reqaddr2,reqetc, makerid, isupchebeasong, songjangdiv, songjangno
dim id, mode, issended, jungsan, jungsanValue, reqdeliverdate, prizetitle, gubuncd, userid, gubunname, i, idarr
dim useDefaultAddr, loRs, evtprize_code, evtprize_enddate
	id = request("id")
	username    = html2db(request("username"))
	reqname     = html2db(request("reqname"))
	reqphone    = request("reqphone1") & "-" & request("reqphone2") & "-" & request("reqphone3")
	reqhp       = request("reqhp1") & "-" & request("reqhp2") & "-" & request("reqhp3")
	reqzip      = request("zipcode")
	reqaddr1    = html2db(request("addr1"))
	reqaddr2    = html2db(request("addr2"))
	reqetc      = html2db(request("reqetc"))
	prizetitle  = html2db(request("prizetitle"))
	makerid     = request("makerid")
	isupchebeasong  = request("isupchebeasong")
	songjangdiv     = request("songjangdiv")
	songjangno      = request("songjangno")
	reqdeliverdate  = request("reqdeliverdate")
	mode            = request("mode")
	issended        = request("issended")
	jungsan            = request("jungsan")
	jungsanValue       = request("jungsanValue")
	idarr = request("idarr")
	userid  	= html2db(request("userid"))
	useDefaultAddr  	= html2db(request("useDefaultAddr"))
	evtprize_enddate  	= html2db(request("evtprize_enddate"))

If jungsan = "" Then
	jungsan = "N"
Else
	jungsan = "Y"
End If

dim sqlStr
if mode="del" then
	sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang" + VbCrlf
	sqlStr = sqlStr + " set deleteyn='Y'" + VbCrlf
	sqlStr = sqlStr + " where id=" + id

	dbget.Execute sqlStr

	sqlStr = "update  [db_event].[dbo].[tbl_event_prize]"+ VbCrlf
	sqlStr = sqlStr + " set evtprize_status=9"+ VbCrlf
	sqlStr = sqlStr + " ,lastupdate=getdate()"+ VbCrlf
	sqlStr = sqlStr + " where evtprize_code in (select top 1 evtprize_code from [db_sitemaster].[dbo].tbl_etc_songjang where id=" + id + ")"

	dbget.Execute sqlStr

elseif mode="delarr" then
	idarr = Mid(idarr,2,Len(idarr))
	idarr = replace(idarr,"|",",")

	sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang" + VbCrlf
	sqlStr = sqlStr & " set deleteyn='Y' where" + VbCrlf
	sqlStr = sqlStr & " id in (" + idarr + ")"

	dbget.Execute sqlStr

	sqlStr = "update  [db_event].[dbo].[tbl_event_prize]"+ VbCrlf
	sqlStr = sqlStr & " set evtprize_status=9"+ VbCrlf
	sqlStr = sqlStr & " ,lastupdate=getdate() where"+ VbCrlf
	sqlStr = sqlStr & " evtprize_code in ("+ VbCrlf
	sqlStr = sqlStr & " 	select top 1 evtprize_code"+ VbCrlf
	sqlStr = sqlStr & " 	from [db_sitemaster].[dbo].tbl_etc_songjang"+ VbCrlf
	sqlStr = sqlStr & " 	where id in (" + idarr + ")"+ VbCrlf
	sqlStr = sqlStr & " )"

	dbget.Execute sqlStr

elseif mode = "I" Then
	gubuncd = request("gubuncd")
	gubunname = request("gubunname")

	sqlStr = "insert into  [db_sitemaster].[dbo].tbl_etc_songjang"    + VbCrlf
	sqlStr = sqlStr + " (username, userid, reqname, reqphone, reqhp, reqzipcode, reqaddress1," + VbCrlf
	sqlStr = sqlStr + " reqaddress2, gubuncd, gubunname, prizetitle, reqetc, " + VbCrlf
	if (reqaddr2 <> "") then
		sqlStr = sqlStr + " inputdate, " + VbCrlf
	end if
	sqlStr = sqlStr + " reqdeliverdate, jungsanYN, jungsan, isupchebeasong, delivermakerid)" + VbCrlf
	sqlStr = sqlStr + " values(" + VbCrlf
	sqlStr = sqlStr + " '" + username + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + userid + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + reqname + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqphone + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqhp + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqzip + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqaddr1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,'" + reqaddr2 + "'" + VbCrlf
	sqlStr = sqlStr + " ,'"&gubuncd&"'" + VbCrlf
	sqlStr = sqlStr + " ,'" + gubunname + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + prizetitle + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + reqetc + "'" + VbCrlf
	if (reqaddr2 <> "") then
		sqlStr = sqlStr + " ,getdate()" + VbCrlf
	end if
	sqlStr = sqlStr + " ,'" + reqdeliverdate + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + jungsan + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + jungsanValue + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + isupchebeasong + "'" + VbCrlf
	sqlStr = sqlStr + " ,'" + makerid + "'" + VbCrlf
	sqlStr = sqlStr + " )"

	dbget.Execute sqlStr

	sqlStr = "SELECT @@IDENTITY AS NewID"
	Set loRs = dbget.Execute(sqlStr)
	id = loRs.Fields("NewID").value
	Set loRs = Nothing

	if (useDefaultAddr = "Y") and (userid <> "") then
		sqlStr = " update e "
		sqlStr = sqlStr + " set e.username = u.username, e.reqname = u.username, e.reqphone = u.userphone, e.reqhp = u.usercell, e.reqzipcode = u.zipcode, e.reqaddress1 = u.zipaddr, e.reqaddress2 = u.useraddr, e.inputdate = getdate() "
		sqlStr = sqlStr + " 		from "
		sqlStr = sqlStr + " 			[db_sitemaster].[dbo].tbl_etc_songjang e "
		sqlStr = sqlStr + " 			join [db_user].[dbo].[tbl_user_n] u on e.userid = u.userid "
		sqlStr = sqlStr + " 		where e.id = " & id
		dbget.Execute sqlStr
	elseif (useDefaultAddr = "N") and (userid <> "") then
		IF application("Svr_Info")="Dev" THEN
			sqlStr = " insert into [db_event].[dbo].[tbl_event_prize](evt_code, evt_ranking, evt_rankname, evt_winner, evt_regdate, adminid, evtprize_name, evtprize_type, evtprize_startdate, evtprize_enddate, evtprize_status, evtgroup_code, giftkind_code, lastupdate) "
			sqlStr = sqlStr + " values(90398, 0, '당첨', '" & userid & "', getdate(), '" & session("ssBctId") & "', '" & prizetitle & "', 3, '" & Left(Now(), 10) & " 00:00:00.000', '" & evtprize_enddate & " 00:00:00.000', 0, " & id & ", 0, getdate()) "
		else
			sqlStr = " insert into [db_event].[dbo].[tbl_event_prize](evt_code, evt_ranking, evt_rankname, evt_winner, evt_regdate, adminid, evtprize_name, evtprize_type, evtprize_startdate, evtprize_enddate, evtprize_status, evtgroup_code, giftkind_code, lastupdate) "
			sqlStr = sqlStr + " values(97605, 0, '당첨', '" & userid & "', getdate(), '" & session("ssBctId") & "', '" & prizetitle & "', 3, '" & Left(Now(), 10) & " 00:00:00.000', '" & evtprize_enddate & " 00:00:00.000', 0, " & id & ", 0, getdate()) "
		end if
		dbget.Execute sqlStr

		sqlStr = "SELECT @@IDENTITY AS NewID"
		Set loRs = dbget.Execute(sqlStr)
		evtprize_code = loRs.Fields("NewID").value
		Set loRs = Nothing

		sqlStr = " update "
		sqlStr = sqlStr + " [db_sitemaster].[dbo].[tbl_etc_songjang] "
		sqlStr = sqlStr + " set evtprize_code = " & evtprize_code
		sqlStr = sqlStr + " where id = " & id
		dbget.Execute sqlStr
	end if
else
    sqlStr = "update [db_sitemaster].[dbo].tbl_etc_songjang" + VbCrlf
    sqlStr = sqlStr + " set username='" + username + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqname='" + reqname + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqphone='" + reqphone + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqhp='" + reqhp + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqzipcode='" + reqzip + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqaddress1='" + reqaddr1 + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqaddress2='" + reqaddr2 + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqetc='" + reqetc + "'" + VbCrlf
    sqlStr = sqlStr + " ,prizetitle='" + prizetitle + "'" + VbCrlf
    sqlStr = sqlStr + " ,delivermakerid='" + makerid + "'" + VbCrlf
    sqlStr = sqlStr + " ,isupchebeasong='" + isupchebeasong + "'" + VbCrlf
    sqlStr = sqlStr + " ,songjangdiv='" + songjangdiv + "'" + VbCrlf
    sqlStr = sqlStr + " ,songjangno='" + songjangno + "'" + VbCrlf
    sqlStr = sqlStr + " ,reqdeliverdate='" + reqdeliverdate + "'" + VbCrlf
    sqlStr = sqlStr + " ,inputdate=IsNULL(inputdate,getdate())" + VbCrlf
    sqlStr = sqlStr + " ,jungsanYN='" + jungsan + "'" + VbCrlf
    sqlStr = sqlStr + " ,jungsan='" + jungsanValue + "'" + VbCrlf

    if (issended="Y") then
        sqlStr = sqlStr + " ,issended='Y'"
        sqlStr = sqlStr + " ,senddate=IsNULL(senddate,getdate())"
    elseif (issended="N") then
        sqlStr = sqlStr + " ,issended='N'"
    end if
    sqlStr = sqlStr + " where id=" + id

    dbget.Execute sqlStr

    ''상태변경
    sqlStr = "update  [db_event].[dbo].[tbl_event_prize]"+ VbCrlf
    sqlStr = sqlStr + " set evtprize_status=3"+ VbCrlf
    sqlStr = sqlStr + " ,lastupdate=getdate()"+ VbCrlf
    sqlStr = sqlStr + " where evtprize_status=0"+ VbCrlf
    sqlStr = sqlStr + " and evtprize_code in (select top 1 evtprize_code from [db_sitemaster].[dbo].tbl_etc_songjang where id=" + id + ")"

    dbget.Execute sqlStr
end if

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if mode="del" then
    response.write "<script>alert('삭제 되었습니다.'); opener.location.reload(); window.close();</script>"
elseif mode="delarr" then
    response.write "<script>alert('삭제 되었습니다.'); parent.location.reload();</script>"
elseif mode="I" then
    response.write "<script>alert('저장 되었습니다.'); opener.location.reload(); window.close();</script>"
else
    response.write "<script>alert('저장 되었습니다.'); location.replace('" + referer + "');</script>"
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
