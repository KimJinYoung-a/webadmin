<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/photo_req/sheduleCls.asp"-->
<%
'################################# �޷½��� #################################
Dim ThisDate,ThisYear,ThisMonth,ThisDay,ThisToday
Dim PrevYear,PrevMonth,NextYear,NextMonth
Dim FirstDay,FirstWeekDay
Dim PrintDay,LastDay,LoopWeek,LoopDay,Stop_Flag
Dim 	PrevThisYearDate,NextThisYearDate,PrevThisMonthDate,NextThisMonthDate
Dim NowToday,NowThisDay
Dim r_photo
	r_photo = request("req_photo")
	ThisDate = Request("ThisDate")		' �Ѿ�� �� �ޱ� �� - �� - ��
	ThisYear = Request("ThisYear")		' �Ѿ�� ���� �⵵
	ThisMonth = Request("ThisMonth")		' �Ѿ�� ���� ��
	ThisDay = Request("ThisDay")		' �Ѿ�� ���� ��

	menupos = request("menupos")		'�޴�������ȣ

	IF ThisDate = "" THEN			' �Ѿ�� ���� ���ٸ� (ó�� �������� �������� ���)
		ThisToday = DATE()		' ����� ������ �� (������ �� - �� - ��)
		ThisYear = YEAR(ThisToday)		' ������ �⵵
		ThisMonth = MONTH(ThisToday)	' ������ ��
		ThisDay = DAY(ThisToday)		' ������ ��
	ELSE					' �Ѿ�� ���� �ִٸ� (�ش� �⵵�� ���� ã�����)
		ThisToday = CDATE(ThisDate)	' �Ѿ�� ���� ����
		ThisYear = YEAR(ThisToday)		' �Ѿ�� ���� �⵵
		ThisMonth = MONTH(ThisToday)	' �Ѿ�� ���� ��
		ThisDay = DAY(ThisToday)		' �Ѿ�� ���� ��
	END IF

	IF ThisMonth = 1 THEN			' �۳�, ����, ������, ������ ���ϱ� (1��, 12��, ������ �޵�)
		PrevYear = ThisYear -1
		PrevMonth = 12
		NextYear = ThisYear +1
		NextMonth = 2
	ELSEIF ThisMonth = 12 THEN
		PrevYear = ThisYear -1
		PrevMonth = 11
		NextYear = ThisYear +1
		NextMonth= 1
	ELSE
		PrevYear = ThisYear -1
		PrevMonth = ThisMonth -1
		NextYear = ThisYear +1
		NextMonth = ThisMonth +1
	END IF

	FirstDay = DateSerial(ThisYear,ThisMonth,1)	' �Ѱܹ��� ���� �� �ʱⰪ (��-��-1)
	FirstWeekDay = WeekDay(FirstDay,vbSunday)	' ù���� ������ ���� (�����ϱ��� : �� 2)

	PrintDay = 1				' ��� �ʱⰪ�� 1

	IF ThisMonth = 4 OR ThisMonth =6 OR ThisMonth = 9 OR ThisMonth = 11 THEN		' ���� ���� ���� �� ���
		LastDay = 30
	ELSEIF ThisMonth = 2 AND NOT (ThisYear MOD 4) = 0 THEN
		LastDay = 28
	ELSEIF ThisMonth = 2 AND (ThisYear MOD 4) = 0 THEN
		IF (ThisYear MOD 100) = 0 THEN
			IF (ThisYear MOD 400) = 0 THEN
				LastDay = 29
			ELSE
				LastDay = 30
			END IF
		ELSE
			LastDay = 29
		END IF
	ELSE
		LastDay = 31
	END IF

	ThisDate  = CDATE(ThisYear&"-"&ThisMonth&"-"&ThisDay)			' ������ �̵��� �Ѱ��� ����
	PrevThisYearDate = PrevYear&"-"&ThisMonth&"-"&ThisDay
	NextThisYearDate =  NextYear&"-"&ThisMonth&"-"&ThisDay
	PrevThisMonthDate = ThisYear&"-"&PrevMonth&"-01"
	NextThisMonthDate = ThisYear&"-"&NextMonth&"-01"

	NowToday = DATE()						' ���� ��¥(��) �� ���ϱ�
	NowThisDay = DAY(NowToday)

'################################# �޷³� #################################
Dim iz, osemi

Dim Thisyyyymm
Thisyyyymm = left(CDATE(ThisYear&"-"&ThisMonth&"-"&ThisDay),7)

Set osemi = new CSeminarRoomCalendar
osemi.FRectYYYYMM = Thisyyyymm
osemi.FReq_photo = r_photo
osemi.list
%>
<!-- �Ʒ��� ��ũ��Ʈ���� ���� �־�� �Ѵ�. -->
<div ID="viewDIV" STYLE="position:absolute; visibility:hide;"></div>
<!-- �� ������ ��ũ��Ʈ���� ���� �־�� �Ѵ�. -->
<script language="JavaScript">
<!--
var fcolor = "#ffffff";        // Main background color
var textcolor = "#000000";        // Text color
var border_size = "1";                // border size, 1-3
var border_color = "#000000";        // Border color
var width = "300";                // �˾� �ڽ��� ����, 100 - 300
var palign = 0;                // �˾� �ڽ��� ��ġ, 0:center/1:right/2:left

ns4 = (document.layers)? true:false
ie4 = (document.all)? true:false
ie5 = false;

// Microsoft Stupidity Check.
if (ie4) {
        if (navigator.userAgent.indexOf('MSIE 5')>0) {
                ie5 = true;
        } else {
                ie5 = false; }
} else {
        ie5 = false;
}

var x = 0;
var y = 0;
var offsetx = 10;
var offsety = 10;
var popup_on = 0;
var over;

if ( (ns4) || (ie4) ) {
    if (ns4) over = document.viewDIV;
    if (ie4) over = viewDIV.style;
    document.onmousemove = mouseMove;
    if (ns4) document.captureEvents(Event.MOUSEMOVE);
}
// Clears popups if appropriate
function viewoff() {
    if ( (ns4) || (ie4) ) {
            popup_on = 0;
            hideObject(over);
    }
}
function myColor(num) {
	if (document.all) {
		eval('document.all.cell'+num+'.style.background = "#E6E6F2"');
	}
}

function myColorOut(num) {
	if (document.all) {
		eval('document.all.cell'+num+'.style.background = "#FFFFFF"');
	}
}

function GoPhotoreq(thisday) {
    location.href = "request_cal_day.asp?getday=" + thisday;
}

// Simple popup
function viewon(text) {
    var txt = "<TABLE WIDTH="+width+" STYLE='filter:alpha(opacity=100); border:0 ' BORDER=0 CELLPADDING="+border_size+" CELLSPACING=0 BGCOLOR="+border_color+"><TR><TD><TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR="+fcolor+"><TR><TD><FONT FACE='Arial,Helvetica' COLOR="+textcolor+" SIZE='2'>"+text+"</FONT></TD></TR></TABLE></TD></TR></TABLE>"
    layerWrite(txt);
    disp();
}

// Common calls
function disp() {
    if ( (ns4) || (ie4) ) {
        if (popup_on == 0)         {
            if (palign == 0) { // Center
                    moveTo(over,x+offsetx-(width/2),y+offsety);
            }
            if (palign == 1) { // Right
                    moveTo(over,x+offsetx,y+offsety);
            }
            if (palign == 2) { // Left
                    moveTo(over,x-offsetx-width,y+offsety);
            }
            showObject(over);
            popup_on = 1;
        }
    }
// Here you can make the text goto the statusbar.
}
//-->

// Moves the layer
function mouseMove(e) {
    if (ns4) {x=e.pageX; y=e.pageY;}
    if (ie4) {x=event.x; y=event.y;}
    if (ie5) {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
    if (popup_on) {
        if (palign == 0) { // Center
                moveTo(over,x+offsetx-(width/2),y+offsety);
        }
        if (palign == 1) { // Right
                moveTo(over,x+offsetx,y+offsety);
        }
        if (palign == 2) { // Left
                moveTo(over,x-offsetx-width,y+offsety);
        }
    }
}

// Writes to a layer
function layerWrite(txt) {
    if (ns4) {
        var lyr = document.viewDIV.document
        lyr.write(txt)
        lyr.close()
    }
    else if (ie4) document.all["viewDIV"].innerHTML = txt
}

// Make an object visible
function showObject(obj) {
    if (ns4) obj.visibility = "show"
    else if (ie4) obj.visibility = "visible"
}

// Hides an object
function hideObject(obj) {
    if (ns4) obj.visibility = "hide"
    else if (ie4) obj.visibility = "hidden"
}

// Move a layer
function moveTo(obj,xL,yL) {
    obj.left = xL
    obj.top = yL
}
function fnSearch(){
	var frm = document.sFrm;
	frm.submit();
}
</script>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td>
		<form name="sFrm" method="get" action="?">
		<input type="hidden" name="ThisDate" value="<%=ThisDate%>">
		<input type="hidden" name="menupos" value="<%=menupos%>">
		<input type="hidden" name="r_photo" value="<%=r_photo%>">
		<table width="100%" border="0" cellspacing="0" cellpadding="0" class="font" height="25">
		<tr>
			<td align="right"><% call SelectUser("req_photo", ""&r_photo&"") %></td>
		</tr>
		<tr align="center" valign="middle">
			<td height="25">
				<A HREF="request_cal.asp?ThisDate=<%=PrevThisYearDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">��</A>&nbsp;&nbsp;<%=ThisYear%>&nbsp;&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=NextThisYearDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">��</A>&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=PrevThisMonthDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">��</A>&nbsp;&nbsp;<%=ThisMonth%>&nbsp;&nbsp;
				<A HREF="request_cal.asp?ThisDate=<%=NextThisMonthDate%>&menupos=<%=menupos%>&req_photo=<%=r_photo%>">��</A>
			</td>
		</tr>
		</table>
		</form>
		<br>
		<table border="1" cellspacing="0" cellpadding="0"  bordercolordark="White" bordercolorlight="black">
		<tr>
			<td width="120" height="80" align="center" class="verdana-mid"><font color="#FF0000">��</font></td>
			<td width="120" height="80" align="center" class="verdana-mid">��</td>
			<td width="120" height="80" align="center" class="verdana-mid">ȭ</td>
			<td width="120" height="80" align="center" class="verdana-mid">��</td>
			<td width="120" height="80" align="center" class="verdana-mid">��</td>
			<td width="120" height="80" align="center" class="verdana-mid">��</td>
			<td width="120" height="80" align="center" class="verdana-mid"><font color=#8FB5DA>��</font></td>
		</tr>
<%
	FOR LoopWeek = 1 TO 6						' �� ���� LOOP  �ִ� 6��
		Response.Write "<tr>"&vbCR			' vbCR = ch<13>  : ����
		FOR LoopDay = 1 TO 7					' ���� LOOP ���� : ������
			IF FirstWeekDay > 1 THEN					' ���� ������ ��Ÿ���°��� 1(�Ͽ���)���� Ŭ��� ������ ������ֱ� ���� �ڵ� ; ���̺� ĭ���� ���߱� ����
				Response.Write "<td height=80 align=left valign=top>&nbsp;"
				FirstWeekDay = FirstWeekDay - 1
			ELSE
				IF PrintDay > LastDay THEN				' ���� ���� �������� Ŭ��� (����ó��) ; ���̺� ĭ���� ���߱� ����
  					Response.Write "<td height=80 align=left valign=top>"
  				ELSE
  					IF PrintDay = NowThisDay THEN		' ���� ���� ������ ��� (����� ũ�⸦ �ٸ��� ����)
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#990000><b>"&PrintDay&"</b></font></td>"
						For iz = 0 to osemi.FResultCount - 1

							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
								</tr><tr><td class=verdana-mid onMouseOver = "viewon('���� : <%=osemi.FItemList(iz).FReqStatus%><br> ��ǰ�� : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> �μ� :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"��Ʈ","") %><br> ��û�� : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
					ELSEIF LoopDay = 1 THEN			' �Ͽ����� ���� �ٸ���
%>
  						<td height=80 align=left valign=top id=cell<%=Cstr(PrintDay)%> style='cursor:hand' onclick=GoPhotoreq(<%=Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay)))%>)><table class=verdana-small><tr><td><font color=#FF0000><%=PrintDay%></font></td></tr>
<%
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
							</tr><tr><td class=verdana-mid onMouseOver = "viewon('���� : <%=osemi.FItemList(iz).FReqStatus%><br> ��ǰ�� : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> �μ� :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"��Ʈ","") %><br> ��û�� : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
					ELSEIF LoopDay = 7 THEN			' ������� ���� �ٸ���
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#3366CC>"&PrintDay&"</font></td></tr>"
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
							</tr><tr><td class=verdana-mid onMouseOver = "viewon('���� : <%=osemi.FItemList(iz).FReqStatus%><br> ��ǰ�� : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> �μ� :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"��Ʈ","") %><br> ��û�� : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End If
						Next
						response.write "</table>"
	  				ELSE					' ������ �ϵ��� ���̺� �ȿ� �־���
  						Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " style='cursor:hand' onclick=GoPhotoreq('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td>"&PrintDay & "</td></tr>"
						For iz = 0 to osemi.FResultCount - 1
							If PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) Then
%>
								</tr><tr><td class=verdana-mid onMouseOver = "viewon('���� : <%=osemi.FItemList(iz).FReqStatus%><br> ��ǰ�� : <%= Replace(osemi.FItemList(iz).FPrdName, "'", "") %> <br> �μ� :<%= chkIIF(osemi.FItemList(iz).FReqDepartment <> "",osemi.FItemList(iz).FReqDepartment&"��Ʈ","") %><br> ��û�� : <%=osemi.FItemList(iz).FUsername%>'); return true;" onMouseOut = 'viewoff(); return true;'><font color="<%=osemi.FItemList(iz).FReqStatusColor%>"><%=DDotFormat(Replace(osemi.FItemList(iz).FPrdName, "'", ""),10)%>(<%=osemi.FItemList(iz).FPrd_type2%>)</font></td></tr>
<%
							End if
						Next
						response.write "</table>"
	  				END IF
  				END IF
	  			PrintDay = PrintDay + 1				' �ϼ��� 1�� ����
				IF PrintDay>LastDay THEN  				' �������� Stop_Flag = 1 �̶�� ����
					Stop_Flag=1
				END IF
  			END IF
  			Response.Write "</td>" &vbCR
  		NEXT							' �ϴ��� (7ĭ) ����
  		Response.Write "</tr>" &vbCR
		IF Stop_Flag=1 THEN					' 6�ֱ��� �Ȱ��� Stop_Flag = 1 �̸� �ִ��� ���� ���� / �̺κ� �����ϸ� �޷� �ڽ��� �����ϰ� ������
			EXIT FOR
		END IF
  	NEXT								' �ִ��� (�ִ� 6��) ����
%>
		</table>
	</td>
</tr>
</table>
<%set osemi = Nothing%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->