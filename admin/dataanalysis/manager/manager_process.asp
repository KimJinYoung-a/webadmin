<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 데이터분석 방문수
' History : 2016.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->

<%
dim strSql, mode, menupos, i, lastuserid, kind, measure, measurename, goapiurl, dimensiongubun, pretypegubun, ordtypegubun
dim mainidx, channeltype, charttype, position, positionpretype, chartsortno, isusing, chartidx, comment, groupcd
dim groupsortno, searchgroupcd, option1, option2, shchannelgubun, shmakeridgubun, shdategubun, shdateunit, shdatetermgubun
dim groupcdnewreg, groupcdnamenewreg
	menupos = getNumeric(requestcheckvar(request("menupos"),10))
	mainidx = getNumeric(requestcheckvar(request("mainidx"),10))
	channeltype = requestcheckvar(request("channeltype"),32)
	charttype = requestcheckvar(request("charttype"),32)
	position = requestcheckvar(request("position"),32)
	positionpretype = requestcheckvar(request("positionpretype"),32)
	chartsortno = getNumeric(requestcheckvar(request("chartsortno"),10))
	isusing = requestcheckvar(request("isusing"),1)
	mode = requestcheckvar(request("mode"),32)
	kind = requestcheckvar(request("kind"),32)
	measure = requestcheckvar(request("measure"),32)
	measurename = requestcheckvar(request("measurename"),32)
	goapiurl = requestcheckvar(request("goapiurl"),32)
	dimensiongubun = requestcheckvar(request("dimensiongubun"),10)
	pretypegubun = requestcheckvar(request("pretypegubun"),10)
	ordtypegubun = requestcheckvar(request("ordtypegubun"),10)
	comment = request("comment")
	option1 = requestcheckvar(request("option1"),800)
	option2 = requestcheckvar(request("option2"),800)
	groupcd = requestcheckvar(request("groupcd"),32)
	groupsortno = getNumeric(requestcheckvar(request("groupsortno"),10))
	searchgroupcd = requestcheckvar(request("searchgroupcd"),32)
	shchannelgubun = requestcheckvar(request("shchannelgubun"),10)
	shmakeridgubun = requestcheckvar(request("shmakeridgubun"),10)
	shdategubun = requestcheckvar(request("shdategubun"),10)
	shdateunit = requestcheckvar(request("shdateunit"),10)
	shdatetermgubun = requestcheckvar(request("shdatetermgubun"),10)
	groupcdnewreg = requestcheckvar(request("groupcdnewreg"),32)
	groupcdnamenewreg = requestcheckvar(request("groupcdnamenewreg"),32)

lastuserid=session("ssBctId")

dim referer
	referer = request.ServerVariables("HTTP_REFERER")

if mode="inchartreg" then
	if isusing="" then isusing="Y"
	if chartsortno="" then chartsortno=100

	strSql = "insert into db_analyze.dbo.tbl_analysis_chart (" & vbcrlf
	strSql = strSql & " mainidx, channeltype, charttype, position, positionpretype, option1, option2, isusing, chartsortno, regdate) values (" & vbcrlf
	strSql = strSql & " " & trim(mainidx) & ", '"& html2db(trim(channeltype)) &"', '"& html2db(trim(charttype)) &"', '" & html2db(trim(position)) & "'" & vbcrlf
	strSql = strSql & " ,'" & html2db(trim(positionpretype)) & "','" & html2db(trim(option1)) & "','" & html2db(trim(option2)) & "','" & trim(isusing) & "'" & vbcrlf
	strSql = strSql & " ," & trim(chartsortno) & ", getdate()" & vbcrlf
	strSql = strSql & " )"

	'response.write strSql & "<br>"
	dbanalget.Execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	'response.write "	self.close();"
	response.write "	location.replace('/admin/dataanalysis/manager/chart_edit.asp?mainidx="& mainidx &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

elseif mode="chartlistedit" then

	for i=1 to request.form("chartidx").count
		chartidx = getNumeric(request.form("chartidx")(i))
		channeltype = request.form("channeltype_"&chartidx)
		charttype = request.form("charttype_"&chartidx)
		position = request.form("position_"&chartidx)
		positionpretype = request.form("positionpretype_"&chartidx)
		isusing = request.form("isusing_"&chartidx)
		chartsortno = getNumeric(request.form("chartsortno_"&chartidx))
		option1 = request.form("option1_"&chartidx)
		option2 = request.form("option2_"&chartidx)
		option1 = replace(option1,"""","'")
		option1 = replace(option1," ","")
		option1 = replace(option1,vbcrlf,"")
		option1 = replace(option1,"	","")
		option2 = replace(option2,"""","'")
		option2 = replace(option2," ","")
		option2 = replace(option2,vbcrlf,"")
		option2 = replace(option2,"	","")

		if isusing="" then isusing="Y"
		if chartsortno="" then chartsortno=100

		strSql = " Update db_analyze.dbo.tbl_analysis_chart" & vbcrlf
		strSql = strSql & " Set channeltype='" & html2db(trim(channeltype)) & "'" & vbcrlf
		strSql = strSql & " ,charttype='" & html2db(trim(charttype)) & "'" & vbcrlf
		strSql = strSql & " ,position='" & html2db(trim(position)) & "'" & vbcrlf
		strSql = strSql & " ,positionpretype='" & html2db(trim(positionpretype)) & "'" & vbcrlf
		strSql = strSql & " ,option1='" & html2db(trim(option1)) & "'" & vbcrlf
		strSql = strSql & " ,option2='" & html2db(trim(option2)) & "'" & vbcrlf
		strSql = strSql & " ,chartsortno=" & trim(chartsortno) & "" & vbcrlf
		strSql = strSql & " ,isusing='" & trim(isusing) & "' Where " & vbcrlf
		strSql = strSql & " mainidx='"& trim(mainidx) &"'" & vbcrlf
		strSql = strSql & " and chartidx="& trim(chartidx) &""

		'response.write strSql & "<br>"
		dbanalget.Execute strSql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	'response.write "	self.close();"
	response.write "	location.replace('/admin/dataanalysis/manager/chart_edit.asp?mainidx="& mainidx &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

elseif mode="maindatareg" then
	if groupcd="" or isnull(groupcd) then
		response.write "그룹이 지정되지 않았습니다."
		dbget.close()	:	response.end
	end if

	'/그룹 신규등록일경우 구분 테이블에 먼저 등록
	if groupcd="NEWREG" then
		strSql = "insert into db_analyze.dbo.tbl_analysis_gubun (" & vbcrlf
		strSql = strSql & " gubun, gubunkey, gubunname, sortno, regdate) values (" & vbcrlf
		strSql = strSql & " 'groupcd', '"& html2db(trim(groupcdnewreg)) &"', '"& html2db(trim(groupcdnamenewreg)) &"', 100, getdate()" & vbcrlf
		strSql = strSql & " )"

		'response.write strSql & "<br>"
		dbanalget.Execute strSql
		groupcd=trim(groupcdnewreg)
	end if

	if isusing="" then isusing="Y"
	if dimensiongubun="" or isnull(dimensiongubun) or dimensiongubun="0" then
		dimensiongubun="NULL"
	end if
	if pretypegubun="" or isnull(pretypegubun) or pretypegubun="0" then
		pretypegubun="NULL"
	end if
	if shchannelgubun="" or isnull(shchannelgubun) or shchannelgubun="0" then
		shchannelgubun="NULL"
	end if
	if shmakeridgubun="" or isnull(shmakeridgubun) or shmakeridgubun="0" then
		shmakeridgubun="NULL"
	end if
	if shdategubun="" or isnull(shdategubun) or shdategubun="0" then
		shdategubun="NULL"
	end if
	if shdateunit="" or isnull(shdateunit) or shdateunit="0" then
		shdateunit="NULL"
	end if
	if shdatetermgubun="" or isnull(shdatetermgubun) or shdatetermgubun="0" then
		shdatetermgubun="NULL"
	end if
	if ordtypegubun="" or isnull(ordtypegubun) or ordtypegubun="0" then
		ordtypegubun="NULL"
	end if

	strSql = "insert into db_analyze.dbo.tbl_analysis_main (" & vbcrlf
	strSql = strSql & " kind, measure, measurename, apiurl, dimensiongubun, pretypegubun, shchannelgubun, shmakeridgubun, shdategubun" & vbcrlf
	strSql = strSql & " , shdateunit,shdatetermgubun, ordtypegubun, comment, isusing, regdate) values (" & vbcrlf
	strSql = strSql & " '" & trim(kind) & "', '"& html2db(trim(measure)) &"', '"& html2db(trim(measurename)) &"', '" & html2db(trim(goapiurl)) & "'" & vbcrlf
	strSql = strSql & " ," & trim(dimensiongubun) & "," & trim(pretypegubun) & "," & trim(shchannelgubun) & "," & trim(shmakeridgubun) & "" & vbcrlf
	strSql = strSql & " ," & trim(shdategubun) & "," & trim(shdateunit) & "," & trim(shdatetermgubun) & ", "& trim(ordtypegubun) &", '"& html2db(trim(comment)) &"'" & vbcrlf
	strSql = strSql & " , '" & trim(isusing) & "', getdate()" & vbcrlf
	strSql = strSql & " )"

	'response.write strSql & "<br>"
	dbanalget.Execute strSql

	strSql = " select top 1" & vbcrlf
	strSql = strSql & " m.mainidx" & vbcrlf
	strSql = strSql & " from db_analyze.dbo.tbl_analysis_main m" & vbcrlf
	strSql = strSql & " order by m.mainidx desc"

	'response.write strSql & "<br>"
	rsAnalget.Open strSql,dbAnalget,1
	if not rsAnalget.EOF  then
		mainidx = rsAnalget("mainidx")
	end if
	rsAnalget.close

	if getNumeric(mainidx)="" then
		response.write "메인데이터 번호가 없습니다."
		dbget.close()	:	response.end
	end if

	strSql = "insert into db_analyze.dbo.tbl_analysis_group (" & vbcrlf
	strSql = strSql & " groupcd, mainidx, groupsortno, regdate) values (" & vbcrlf
	strSql = strSql & " '" & html2db(trim(groupcd)) & "', "& trim(mainidx) &", "& trim(groupsortno) &", getdate()" & vbcrlf
	strSql = strSql & " )"

	'response.write strSql & "<br>"
	dbanalget.Execute strSql

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	'response.write "	self.close();"
	response.write "	location.replace('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd="& groupcd &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end

elseif mode="maindatalistedit" then

	for i=1 to request.form("mainidx").count
		mainidx = getNumeric(request.form("mainidx")(i))
		kind = request.form("kind_"&mainidx)
		measure = request.form("measure_"&mainidx)
		measurename = request.form("measurename_"&mainidx)
		goapiurl = request.form("goapiurl_"&mainidx)
		dimensiongubun = getNumeric(request.form("dimensiongubun_"&mainidx))
		pretypegubun = getNumeric(request.form("pretypegubun_"&mainidx))
		shchannelgubun = getNumeric(request.form("shchannelgubun_"&mainidx))
		shmakeridgubun = getNumeric(request.form("shmakeridgubun_"&mainidx))
		shdategubun = getNumeric(request.form("shdategubun_"&mainidx))
		shdateunit = getNumeric(request.form("shdateunit_"&mainidx))
		shdatetermgubun = getNumeric(request.form("shdatetermgubun_"&mainidx))
		ordtypegubun = request.form("ordtypegubun_"&mainidx)
		comment = request.form("comment_"&mainidx)
		isusing = request.form("isusing_"&mainidx)
		groupsortno = getNumeric(request.form("groupsortno_"&mainidx))
		'groupcd = request.form("groupcd_"&mainidx)

		if isusing="" then isusing="Y"
		if groupsortno="" then groupsortno=100

		if dimensiongubun="" or isnull(dimensiongubun) or dimensiongubun="0" then
			dimensiongubun="NULL"
		end if
		if pretypegubun="" or isnull(pretypegubun) or pretypegubun="0" then
			pretypegubun="NULL"
		end if
		if shchannelgubun="" or isnull(shchannelgubun) or shchannelgubun="0" then
			shchannelgubun="NULL"
		end if
		if shmakeridgubun="" or isnull(shmakeridgubun) or shmakeridgubun="0" then
			shmakeridgubun="NULL"
		end if
		if shdategubun="" or isnull(shdategubun) or shdategubun="0" then
			shdategubun="NULL"
		end if
		if shdateunit="" or isnull(shdateunit) or shdateunit="0" then
			shdateunit="NULL"
		end if
		if shdatetermgubun="" or isnull(shdatetermgubun) or shdatetermgubun="0" then
			shdatetermgubun="NULL"
		end if
		if ordtypegubun="" or isnull(ordtypegubun) or ordtypegubun="0" then
			ordtypegubun="NULL"
		end if

		strSql = "Update db_analyze.dbo.tbl_analysis_main" & vbcrlf
		strSql = strSql & " Set kind='" & html2db(trim(kind)) & "'" & vbcrlf
		strSql = strSql & " ,measure='" & html2db(trim(measure)) & "'" & vbcrlf
		strSql = strSql & " ,measurename='" & html2db(trim(measurename)) & "'" & vbcrlf
		strSql = strSql & " ,apiurl='" & html2db(trim(goapiurl)) & "'" & vbcrlf
		strSql = strSql & " ,dimensiongubun=" & trim(dimensiongubun) & "" & vbcrlf
		strSql = strSql & " ,pretypegubun=" & trim(pretypegubun) & "" & vbcrlf
		strSql = strSql & " ,shchannelgubun=" & trim(shchannelgubun) & "" & vbcrlf
		strSql = strSql & " ,shmakeridgubun=" & trim(shmakeridgubun) & "" & vbcrlf
		strSql = strSql & " ,shdategubun=" & trim(shdategubun) & "" & vbcrlf
		strSql = strSql & " ,shdateunit=" & trim(shdateunit) & "" & vbcrlf
		strSql = strSql & " ,shdatetermgubun=" & trim(shdatetermgubun) & "" & vbcrlf
		strSql = strSql & " ,ordtypegubun=" & trim(ordtypegubun) & "" & vbcrlf
		strSql = strSql & " ,comment='" & html2db(trim(comment)) & "'" & vbcrlf
		strSql = strSql & " ,isusing='" & trim(isusing) & "' Where " & vbcrlf
		strSql = strSql & " mainidx='"& trim(mainidx) &"'"

		'response.write strSql & "<br>"
		dbanalget.Execute strSql

		'//사용안함으로 돌렸을경우 그룹테이블 삭제
		if trim(isusing)="N" then
			strSql = "if exists(" & vbcrlf
			strSql = strSql & " 	select top 1 groupcd" & vbcrlf
			strSql = strSql & " 	from db_analyze.dbo.tbl_analysis_group" & vbcrlf
			strSql = strSql & " 	where groupcd='" & html2db(trim(searchgroupcd)) & "'" & vbcrlf
			strSql = strSql & " 	and mainidx=" & trim(mainidx) & "" & vbcrlf
			strSql = strSql & " )" & vbcrlf
			strSql = strSql & " 	begin" & vbcrlf
			strSql = strSql & " 	delete from db_analyze.dbo.tbl_analysis_group where"& vbcrlf
			strSql = strSql & " 	groupcd='" & html2db(trim(searchgroupcd)) & "'" & vbcrlf
			strSql = strSql & " 	and mainidx=" & trim(mainidx) & "" & vbcrlf
			strSql = strSql & " 	end" & vbcrlf
		else
			strSql = "if exists(" & vbcrlf
			strSql = strSql & " 	select top 1 groupcd" & vbcrlf
			strSql = strSql & " 	from db_analyze.dbo.tbl_analysis_group" & vbcrlf
			strSql = strSql & " 	where groupcd='" & html2db(trim(searchgroupcd)) & "'" & vbcrlf
			strSql = strSql & " 	and mainidx=" & trim(mainidx) & "" & vbcrlf
			strSql = strSql & " )" & vbcrlf
			strSql = strSql & " 	begin" & vbcrlf
			strSql = strSql & " 	update db_analyze.dbo.tbl_analysis_group" & vbcrlf
			strSql = strSql & " 	set groupsortno="& trim(groupsortno) &" where" & vbcrlf
			strSql = strSql & " 	groupcd='" & html2db(trim(searchgroupcd)) & "'" & vbcrlf
			strSql = strSql & " 	and mainidx=" & trim(mainidx) & "" & vbcrlf
			strSql = strSql & " 	end" & vbcrlf
			strSql = strSql & " else" & vbcrlf
			strSql = strSql & " 	begin" & vbcrlf
			strSql = strSql & " 	insert into db_analyze.dbo.tbl_analysis_group (" & vbcrlf
			strSql = strSql & " 	groupcd, mainidx, groupsortno, regdate) values (" & vbcrlf
			strSql = strSql & " 	'" & html2db(trim(searchgroupcd)) & "', "& trim(mainidx) &", "& trim(groupsortno) &", getdate()" & vbcrlf
			strSql = strSql & " 	)"
			strSql = strSql & " 	end" & vbcrlf
		end if

		'response.write strSql & "<br>"
		dbanalget.Execute strSql
	next

	response.write "<script type='text/javascript'>"
	response.write "	alert('OK');"
	response.write "	opener.location.reload();"
	'response.write "	self.close();"
	response.write "	location.replace('/admin/dataanalysis/manager/maindata_edit.asp?searchgroupcd="& searchgroupcd &"&menupos="& menupos &"');"
	response.write "</script>"
	dbget.close()	:	response.end
else
	response.write "<script type='text/javascript'>"
	response.write "	alert('구분자가 없습니다.');"
	response.write "</script>"
	dbget.close()	:	response.end
end if
%>

<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->