<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 80
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

function FormatStr(n,orgData)
	dim tmp
	if (n-Len(CStr(orgData))) < 0 then
		FormatStr = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	FormatStr = tmp
end Function

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim orderserial, sitename
dim differencekey, workgroup, songjangdiv, baljutype
dim ems
dim epostmilitary
dim extSiteName

mode        = RequestCheckvar(request("mode"),16)
orderserial = RequestCheckvar(request("orderserial"),16)
sitename    = RequestCheckvar(request("sitename"),16)
workgroup   = RequestCheckvar(request("workgroup"),1)
baljutype   = RequestCheckvar(request("baljutype"),1)
extSiteName   = RequestCheckvar(request("extSiteName"),32)

'// not used
'' ems				= request("ems")
'' epostmilitary	= request("epostmilitary")
'' cn10x10			= request("cn10x10")
'' ecargo			= request("ecargo")


''출고 택배사 추가 int
songjangdiv = RequestCheckvar(request("songjangdiv"),10)

dummiserial = orderserial
orderserial = split(orderserial,"|")
sitename    = split(sitename,"|")

dim sqlStr,i
dim iid
dim dummiserial
dim obaljudate
dim errcode

if mode="arr" then
	''유효성체크.
	dummiserial = Mid(dummiserial,2,Len(dummiserial))
	dummiserial = replace(dummiserial,"|","','")
	sqlStr = " select top 1 orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
	sqlStr = sqlStr + " where orderserial in ('" + dummiserial + "')"

	rsACADEMYget.Open sqlStr,dbACADEMYget,1
	if Not rsACADEMYget.Eof then
		response.write "<script language='javascript'>"
		response.write "alert('" + rsACADEMYget("orderserial") + " : 이미 발주된 주문입니다. \n\n이미 발주한 주문건은 중복 발주 할 수 없습니다.');"
		response.write "location.replace('" + CStr(refer) + "');"
		response.write "</script>"
		dbACADEMYget.close()	:	response.End
	end if
	rsACADEMYget.Close

	sqlStr = " select (count(id) + 1) as differencekey"
	sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_baljumaster"
	sqlStr = sqlStr + " where convert(varchar(10),baljudate,21)=convert(varchar(10),getdate(),21)"

	rsACADEMYget.Open sqlStr,dbACADEMYget,1
		differencekey = rsACADEMYget("differencekey")
	rsACADEMYget.close


	On Error Resume Next
	dbACADEMYget.beginTrans


	If Err.Number = 0 Then
        errcode = "001"

		'#######발주 마스터###############
		sqlStr = "insert into [db_academy].[dbo].tbl_academy_baljumaster(baljudate,differencekey,workgroup, songjangdiv, baljutype, extSiteName)"
		sqlStr = sqlStr + " values(getdate()," + CStr(differencekey) + ",'" + workgroup + "','" + songjangdiv + "','" + baljutype + "', '" + CStr(extSiteName) + "')"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		sqlStr = "select top 1 id, convert(varchar(19),baljudate,21) as baljudate from [db_academy].[dbo].tbl_academy_baljumaster order by id desc"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	    iid = rsACADEMYget("id")
	    obaljudate = rsACADEMYget("baljudate")
		rsACADEMYget.Close

		'#######발주 디테일###############
		for i=0 to Ubound(orderserial)
			if orderserial(i)<>"" then
				sqlStr = "insert into [db_academy].[dbo].tbl_academy_baljudetail(baljuid,orderserial,sitename,userid)"
				sqlStr = sqlStr + " values(" + CStr(iid) + ","
				sqlStr = sqlStr + " '" + orderserial(i) + "',"
				sqlStr = sqlStr + " '" + sitename(i) + "',"
				sqlStr = sqlStr + " '')"
				rsACADEMYget.Open sqlStr,dbACADEMYget,1
			end if
		next

		''** [db_academy].[dbo].tbl_academy_baljudetail.baljusongjangno is NULL 인경우 업체배송으로 인식 (Logics 기존 시스템)
		''텐바이텐 배송의 경우 송장번호를 not null 값으로 입력..

		sqlStr = "update [db_academy].[dbo].tbl_academy_baljudetail" + VbCrlf
		sqlStr = sqlStr + " set baljusongjangno=''"
		sqlStr = sqlStr + " where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " and orderserial in "
		sqlStr = sqlStr + " (select distinct bd.orderserial "
		sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_baljudetail bd, "
		sqlStr = sqlStr + "     [db_academy].[dbo].tbl_academy_order_detail od "
		sqlStr = sqlStr + "     where bd.baljuid=" + CStr(iid)
		sqlStr = sqlStr + "     and bd.orderserial=od.orderserial "
		sqlStr = sqlStr + "     and od.isupchebeasong='N' "
		sqlStr = sqlStr + "     and od.itemid<>0 and "
		sqlStr = sqlStr + "     od.cancelyn<>'Y' "
		sqlStr = sqlStr + "  ) "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if


	If Err.Number = 0 Then
        errcode = "002"

		'// 플라워 주문건 CASE - 먼저 확인후 출고된경우;
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
		sqlStr = sqlStr + " where ipkumdiv>4"
		sqlStr = sqlStr + " and baljudate is NULL"
		sqlStr = sqlStr + " and orderserial in "
		sqlStr = sqlStr + " (select orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
		sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " )"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

		'#######주문 통보 진행 ############### (발주일 지정)
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " set ipkumdiv='5'"
		sqlStr = sqlStr + " ,baljudate='" + CStr(obaljudate) + "'"
		sqlStr = sqlStr + " where ipkumdiv=4"
		sqlStr = sqlStr + " and orderserial in "
		sqlStr = sqlStr + " (select orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
		sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " )"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1


		''' 업체배송 확인이 위 두 쿼리 사이에 끼이게 되면 발주일이 입력이 안되는 경우가 발생한다.
		''' 발주일 미입력된 모든 주문에 발주일 입력
		''' 10042631803
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " set baljudate='" + CStr(obaljudate) + "'"
		sqlStr = sqlStr + " where baljudate is NULL"
		sqlStr = sqlStr + " and orderserial in "
		sqlStr = sqlStr + " (select orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
		sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " )"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "003"

		'#######  택배사입력 (Master) ############### (송장은 출고시 입력으로 변경 , 택배사만 입력함.)
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_master"
		sqlStr = sqlStr + " set songjangdiv=" + CStr(songjangdiv)
		sqlStr = sqlStr + " where orderserial in "
		sqlStr = sqlStr + " (select distinct bd.orderserial "
		sqlStr = sqlStr + "     from [db_academy].[dbo].tbl_academy_baljudetail bd, "
		sqlStr = sqlStr + "     [db_academy].[dbo].tbl_academy_order_detail od "
		sqlStr = sqlStr + "     where bd.baljuid=" + CStr(iid)
		sqlStr = sqlStr + "     and bd.orderserial=od.orderserial "
		sqlStr = sqlStr + "     and od.isupchebeasong='N' "
		sqlStr = sqlStr + "     and od.itemid<>0 and "
		sqlStr = sqlStr + "     od.cancelyn<>'Y' "
		sqlStr = sqlStr + "  ) "
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "004"
		'###### Order Detail 텐바이텐 배송 발주일 저장 ############
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_detail"
		sqlStr = sqlStr + " set upcheconfirmdate= '" + CStr(obaljudate) + "'"
		sqlStr = sqlStr + " ,currstate='2'"
		sqlStr = sqlStr + " ,songjangdiv=" + CStr(songjangdiv)
		sqlStr = sqlStr + " where orderserial in "
		sqlStr = sqlStr + " (select orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
		sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and isupchebeasong='N'"
		sqlStr = sqlStr + " and itemid<>0"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "005"

		'###### Order Detail 업체 배송 업체주문통보 flag 변경 : NULL 인 경우만 ############
		sqlStr = "update [db_academy].[dbo].tbl_academy_order_detail"
		sqlStr = sqlStr + " set currstate='2'"
		sqlStr = sqlStr + " where orderserial in "
		sqlStr = sqlStr + " (select orderserial from [db_academy].[dbo].tbl_academy_baljudetail"
		sqlStr = sqlStr + " 	where baljuid=" + CStr(iid)
		sqlStr = sqlStr + " )"
		sqlStr = sqlStr + " and isupchebeasong='Y'"
		sqlStr = sqlStr + " and itemid<>0"
		sqlStr = sqlStr + " and IsNULL(currstate,0)=0"

		'' 업체배송의 경우 업체주문통보 or NULL 인경우 경우 확인 가능함.
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "006"

		''발주 아이템 목록
		sqlStr = " insert into [db_academy].[dbo].[tbl_academy_baljuitem]" + VbCrlf
		sqlStr = sqlStr + " (baljuid,itemid,itemoption,rackcode,makerid," + VbCrlf
		sqlStr = sqlStr + " itemname,itemoptionname,orgsellprice,baljuno," + VbCrlf
		sqlStr = sqlStr + " ipgono,smallimage,listimage,itemrackcode)" + VbCrlf
		sqlStr = sqlStr + " select " + CStr(iid) + ",d.itemid,d.itemoption,0,'',"
		sqlStr = sqlStr + " '', '',0, sum(d.itemno)," + VbCrlf
		sqlStr = sqlStr + " 0,'','',''" + VbCrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_academy_baljudetail lo," + VbCrlf
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_master m," + VbCrlf
		sqlStr = sqlStr + " [db_academy].[dbo].tbl_academy_order_detail d" + VbCrlf
		sqlStr = sqlStr + " where lo.baljuid=" + CStr(iid)  + VbCrlf
		sqlStr = sqlStr + " and lo.orderserial=d.orderserial" + VbCrlf
		sqlStr = sqlStr + " and m.orderserial=d.orderserial"
		sqlStr = sqlStr + " and m.cancelyn='N'"
		sqlStr = sqlStr + " and d.itemid<>0"
		sqlStr = sqlStr + " and d.isupchebeasong<>'Y'"
		sqlStr = sqlStr + " and d.cancelyn<>'Y'"
		sqlStr = sqlStr + " group by d.itemid,d.itemoption" + VbCrlf
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
		errcode = "008"

		''상품테이블관련
		sqlStr = " update [db_academy].[dbo].[tbl_academy_baljuitem]" + VbCrlf
		sqlStr = sqlStr + " set makerid = T.makerid" + VbCrlf
		sqlStr = sqlStr + " ,itemname =  T.itemname" + VbCrlf
		sqlStr = sqlStr + " ,orgsellprice = T.sellcash" + VbCrlf
		sqlStr = sqlStr + " ,smallimage = T.smallimage" + VbCrlf
		sqlStr = sqlStr + " ,listimage = T.listimage" + VbCrlf
		sqlStr = sqlStr + " ,itemrackcode = '' " + VbCrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].[tbl_diy_item] T" + VbCrlf
		sqlStr = sqlStr + " where [db_academy].[dbo].[tbl_academy_baljuitem].baljuid=" + CStr(iid) + VbCrlf
		sqlStr = sqlStr + " and [db_academy].[dbo].[tbl_academy_baljuitem].itemid=T.itemid" + VbCrlf
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
		errcode = "009"

		sqlStr = " update [db_academy].[dbo].[tbl_academy_baljuitem]" + VbCrlf
		sqlStr = sqlStr + " set itemrackcode='9999'" + VbCrlf
		sqlStr = sqlStr + " where baljuid=" + CStr(iid)  + VbCrlf
		sqlStr = sqlStr + " and (itemrackcode is null or itemrackcode='')"
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "011"

		''옵션테이블관련
		sqlStr = " update [db_academy].[dbo].[tbl_academy_baljuitem]" + VbCrlf
		sqlStr = sqlStr + " set itemoptionname = IsNULL(T.optionname,'')" + VbCrlf
		sqlStr = sqlStr + " from [db_academy].[dbo].tbl_diy_item_option T" + VbCrlf
		sqlStr = sqlStr + " where [db_academy].[dbo].[tbl_academy_baljuitem].baljuid=" + CStr(iid) + VbCrlf
		sqlStr = sqlStr + " and [db_academy].[dbo].[tbl_academy_baljuitem].itemid=T.itemid" + VbCrlf
		sqlStr = sqlStr + " and [db_academy].[dbo].[tbl_academy_baljuitem].itemoption=T.itemoption" + VbCrlf
		rsACADEMYget.Open sqlStr,dbACADEMYget,1

	end if

	If Err.Number = 0 Then
        errcode = "012"

		sqlStr = " update  [db_academy].[dbo].tbl_academy_baljumaster"
		sqlStr = sqlStr + " set songjanginputed='Y'"
		sqlStr = sqlStr + " where id=" + CStr(iid)

		rsACADEMYget.Open sqlStr,dbACADEMYget,1
	end if

	If Err.Number = 0 Then
        dbACADEMYget.CommitTrans
	Else
        dbACADEMYget.RollBackTrans
        response.write "<script>alert('데이타를 저장하는 도중에 에러가 발생하였습니다.\r\n관리자 문의 요망 (에러코드 : " + CStr(errcode) + ")');</script>"
        response.write "<script>history.back()</script>"
        dbACADEMYget.close()	:	response.End
	End If
	on error Goto 0

	''추가 작업. 트랜잭션에서 뺌.. 주문내역 잠김 시간이 오래걸림 : 사은품작성시간.
	''response.write "<font color=red>발주 생성시 오류나는 경우 서팀장께 문의요망!! - 사은품내역 </font><br>"
    '' 오래된 내역 삭제
    sqlStr = " delete from [db_academy].[dbo].[tbl_academy_baljuitem] " + VbCrlf
    sqlStr = sqlStr + " where baljuid<" + CStr(iid-100)

    dbACADEMYget.execute sqlStr

end if


%>

<script language="javascript">
alert('발주서가 생성 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
