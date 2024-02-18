<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/academy/seminar_roomcls.asp"-->
<%
'################################# �޷½��� #################################

dim ThisDate,ThisYear,ThisMonth,ThisDay,ThisToday
dim PrevYear,PrevMonth,NextYear,NextMonth
dim FirstDay,FirstWeekDay
dim PrintDay,LastDay,LoopWeek,LoopDay,Stop_Flag
dim 	PrevThisYearDate,NextThisYearDate,PrevThisMonthDate,NextThisMonthDate
dim NowToday,NowThisDay

	ThisDate = Request("ThisDate")		' �Ѿ�� �� �ޱ� �� - �� - ��
	ThisYear = Request("ThisYear")		' �Ѿ�� ���� �⵵
	ThisMonth = Request("ThisMonth")		' �Ѿ�� ���� ��
	ThisDay = Request("ThisDay")		' �Ѿ�� ���� ��

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

	ThisDate  = dateserial(ThisYear,ThisMonth,ThisDay)			' ������ �̵��� �Ѱ��� ����
	PrevThisYearDate = dateserial(PrevYear,ThisMonth,ThisDay)
	NextThisYearDate = dateserial(NextYear,ThisMonth,ThisDay)
	PrevThisMonthDate = dateserial(ThisYear,PrevMonth,1)
	NextThisMonthDate = dateserial(ThisYear,NextMonth,1)

	NowToday = DATE()						' ���� ��¥(��) �� ���ϱ�
	NowThisDay = DAY(NowToday)

'################################# �޷³� #################################



dim iz, osemi

dim Thisyyyymm
Thisyyyymm = left(CDATE(ThisYear&"-"&ThisMonth&"-"&ThisDay),7)

set osemi = new CSeminarRoomCalendar
osemi.FRectYYYYMM = Thisyyyymm
osemi.list
%>
<script language="JavaScript">
<!--

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

function GoSeminar(thisday) {
    location.href = "seminar_room_daily.asp?getday=" + thisday;
}
//-->
</script>
<table border="0" cellpadding="0" cellspacing="0">
<tr>
	<td>
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="font" height="25">
  <tr align="center" valign="middle">
    <td height="25">
    <A HREF="college_calendar.asp?ThisDate=<%=PrevThisYearDate%>">��</A>&nbsp;&nbsp;<%=ThisYear%>&nbsp;&nbsp;
    <A HREF="college_calendar.asp?ThisDate=<%=NextThisYearDate%>">��</A>&nbsp;
    <A HREF="college_calendar.asp?ThisDate=<%=PrevThisMonthDate%>">��</A>&nbsp;&nbsp;<%=ThisMonth%>&nbsp;&nbsp;
    <A HREF="college_calendar.asp?ThisDate=<%=NextThisMonthDate%>">��</A></td>
  </tr>
</table>
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
  				Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " onMouseOver='myColor(" + Cstr(PrintDay) + ")' onMouseOut='myColorOut(" + Cstr(PrintDay) + ")' style='cursor:hand' onclick=GoSeminar('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#990000><b>"&PrintDay&"</b></font></td>"

					for iz = 0 to osemi.FResultCount - 1
						if PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) then
							Response.Write "</tr><tr><td class=verdana-mid>���� : " & Cstr(osemi.FItemList(iz).FCount)  & "��</td></tr>"
						end if
					next
				response.write "</table>"

  				ELSEIF LoopDay = 1 THEN			' �Ͽ����� ���� �ٸ���
  				Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " onMouseOver='myColor(" + Cstr(PrintDay) + ")' onMouseOut='myColorOut(" + Cstr(PrintDay) + ")' style='cursor:hand' onclick=GoSeminar('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#FF0000>"&PrintDay&"</font></td></tr>"

					for iz = 0 to osemi.FResultCount - 1
						if PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) then
							Response.Write "</tr><tr><td class=verdana-mid>���� : " & Cstr(osemi.FItemList(iz).FCount)  & "��</td></tr>"
						end if
					next
				response.write "</table>"

				ELSEIF LoopDay = 7 THEN			' ������� ���� �ٸ���
  				Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " onMouseOver='myColor(" + Cstr(PrintDay) + ")' onMouseOut='myColorOut(" + Cstr(PrintDay) + ")' style='cursor:hand' onclick=GoSeminar('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td><font color=#3366CC>"&PrintDay&"</font></td></tr>"

					for iz = 0 to osemi.FResultCount - 1
						if PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) then
							Response.Write "</tr><tr><td class=verdana-mid>���� : " & Cstr(osemi.FItemList(iz).FCount)  & "��</td></tr>"
						end if
					next
				response.write "</table>"

  				ELSE					' ������ �ϵ��� ���̺� �ȿ� �־���
  				Response.Write "<td height=80 align=left valign=top id=cell" + Cstr(PrintDay) + " onMouseOver='myColor(" + Cstr(PrintDay) + ")' onMouseOut='myColorOut(" + Cstr(PrintDay) + ")' style='cursor:hand' onclick=GoSeminar('" + Cstr(DateSerial(Cstr(ThisYear),Cstr(ThisMonth),Cstr(PrintDay))) + "')><table class=verdana-small><tr><td>"&PrintDay & "</td></tr>"

					for iz = 0 to osemi.FResultCount - 1
						if PrintDay = Clng(right(osemi.FItemList(iz).FDate,2)) then
							Response.Write "</tr><tr><td class=verdana-mid>���� : " & Cstr(osemi.FItemList(iz).FCount)  & "��</td></tr>"
						end if
					next
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

<%
set osemi = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->