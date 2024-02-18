<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 강좌 바코드 검색 ,페이지 부하와, 링크드 서버 사용안하기 위해 생성
' History : 2011.02.25 한용민 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<%
dim OrderSerial ,osearch , userid , sql , RowNum , itemid ,oroom , barcoderoomid
dim startdate,entryname ,lec_title ,SubTotalPrice ,barcodelecprice ,barcodematprice
	OrderSerial = RequestCheckvar(request("OrderSerial"),16)
	itemid = RequestCheckvar(request("itemid"),10)
	startdate = RequestCheckvar(request("startdate"),10)
	entryname = RequestCheckvar(request("entryname"),32)
	lec_title = request("lec_title")
	SubTotalPrice = RequestCheckvar(request("SubTotalPrice"),10)
	barcodelecprice = RequestCheckvar(request("barcodelecprice"),10)
	barcodematprice = RequestCheckvar(request("barcodematprice"),10)
	
	RowNum = 0
  	if lec_title <> "" then
		if checkNotValidHTML(lec_title) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
if OrderSerial = "" or itemid = "" then
	response.write "<script>alert('주문 번호가 없거나 강좌번호가 없습니다.\n관리자 문의 하세요');</script>"
end if

set osearch = new CLectureFingerOrder
	osearch.frectOrderSerial = OrderSerial
	osearch.flecturer_search()
	
	'//주문번호로 해당 회원의 아이디를 가져온다
	if osearch.ftotalcount > 0 then
		userid = osearch.FOneItem.fuserid
	else
		response.write "<script>alert('주문내역이 없습니다.관리자 문의 하세요');</script>"
		response.end
	end if

	sql = ";WITH TMP_LIST AS"
	sql = sql & " (" 
	sql = sql & " select ROW_NUMBER() OVER (ORDER BY o.lecstartdate) AS RowNum,"
	sql = sql & " d.orderserial, o.*"  
	sql = sql & " from db_academy.dbo.tbl_academy_order_master m"
	sql = sql & " Join [db_academy].[dbo].tbl_academy_order_detail d"
	sql = sql & " on m.orderserial=d.orderserial" 
	sql = sql & " and m.ipkumdiv>3"
	sql = sql & " and m.cancelyn='N'"
	sql = sql & " left Join [db_academy].dbo.tbl_lec_item_option o"
	sql = sql & " on d.itemid=o.lecIdx"
	sql = sql & " and d.itemoption=o.lecOption"
	sql = sql & " where m.userid='"&userid&"'"
	sql = sql & " )"
	
	sql = sql & " select RowNum from TMP_LIST"
	sql = sql & " where orderserial='"&OrderSerial&"'"
	
	'response.write sql &"<br>"
	rsACADEMYget.Open sql, dbACADEMYget, 1
	
	'//해당 회원의 핑거의 참여 횟수를 가져온다
	if Not rsACADEMYget.Eof then		
		RowNum = rsACADEMYget("RowNum")
	else
		RowNum = 1	
	end if
	
	rsACADEMYget.Close
	
	'//비회원일 경우 참여횟수 1 고정
	if userid = "" then RowNum = 1

set oroom = new CLectureFingerOrder
	oroom.frectitemid = itemid
	oroom.flecturer_room()
	
	'/해당 강좌의 강의실을 가져온다
	barcoderoomid = oroom.FOneItem.barcoderoomid
%>

	<script language="javascript">

	function DrawReceiptPrintobj_TEC(elementid,printname){
	    var objstring = "";
	    var e;
	    objstring = '<OBJECT name="' + elementid + '" ';
	    objstring = objstring + ' classid="clsid:E76C9051-A8C4-458E-9F60-3C14DB9EECF9" ';
	    objstring = objstring + ' codebase="http://billyman/Tec_dol.cab#version=1,5,0,0" ';
	    objstring = objstring + ' width=0 ';
	    objstring = objstring + ' height=0 ';
	    objstring = objstring + ' align=center ';
	    objstring = objstring + ' hspace=0 ';
	    objstring = objstring + ' vspace=0 ';
	    objstring = objstring + ' > ';
	    objstring = objstring + ' <PARAM Name="PrinterName" Value="' + printname + '"> ';
	    objstring = objstring + ' </OBJECT>';
	    
	    document.write(objstring);
	}

	DrawReceiptPrintobj_TEC("TEC_DO2","TEC B-SV4");

	var OrderSerial = '<%= OrderSerial %>';
	var startdate = '<%= startdate %>';
	var entryname = '<%= entryname %>';
	var lec_title = '<%= lec_title %>';
	var SubTotalPrice = '<%= SubTotalPrice %>';
	var barcodelecprice = '<%= barcodelecprice %>';
	var barcodematprice = '<%= barcodematprice %>';
	var RowNum = '<%= RowNum %>';
	var barcoderoomid = '<%= barcoderoomid %>';

	var X = 7.9; //1.5;
	var Y = 8; //1.5;
	var F = 1; //1.4;
	
	if (TEC_DO2.IsDriver == 1){
		//TEC_DO2.SetCutter(1, 0, 0, 0 );
		//TEC_DO2.SetDriverStock(0, 1, 5, 0, 1);  // Default 지정한것임.
	   TEC_DO2.SetPaper(575,1450);
	   
       TEC_DO2.OffsetX = 0;
       TEC_DO2.OffsetY = 0;
       
       TEC_DO2.PrinterOpen();
       //TEC_DO2.SetDriverStock(0, 1, 2.0, 0, 3);
       //left
              
       TEC_DO2.PrintText(13*X, 48*Y, "arial", 22*F, 0, 0, OrderSerial);		//주문번호
              
       TEC_DO2.PrintText(100*X, 15*Y, "arial", 22*F, 0, 0, entryname);		//이름
       
       TEC_DO2.PrintText(100*X, 21*Y, "arial", 22*F, 0, 0, startdate);		//수강일자
       
       TEC_DO2.PrintText(93*X, 29*Y, "HY견고딕", 16*F, 0, 0, lec_title);		//강좌명
       
       TEC_DO2.PrintText(108*X, 34*Y, "arial", 22*F, 0, 0, barcoderoomid);		//강의실
       
       TEC_DO2.PrintText(93*X, 44*Y, "arial", 22*F, 0, 0, SubTotalPrice);		//합계
       
       TEC_DO2.PrintText(93*X, 51*Y, "arial", 22*F, 0, 0, barcodelecprice);		//수강료
       
       TEC_DO2.PrintText(120*X, 51*Y, "arial", 22*F, 0, 0, barcodematprice);		//재료비
              
       TEC_DO2.PrintText(56*X, 30*Y, "arial", 22*F, 0, 0, RowNum);		//참여횟수
       
       //right 주문내역
             
       TEC_DO2.PrinterClose();
	}else {
		alert('TEC B-EV4-G 드라이버를 설치해 주세요')		
	}
	</script>
<%
	set osearch = nothing
	set oroom = nothing
%>	
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->