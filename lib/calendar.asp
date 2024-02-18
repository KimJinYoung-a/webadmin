<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim caption, defaultval
caption = request("caption")
defaultval = request("defaultval")

dim yyyymmdd
dim yyyy,mm,dd

on Error resume next
yyyymmdd = CDate(defaultval)
on Error goto 0

If Not Err then
	if yyyymmdd<>"" then
		yyyy = Format00(4,(year(yyyymmdd)))
		mm = Format00(2,(month(yyyymmdd)))
		dd = Format00(2,(day(yyyymmdd)))
	end if
end if

'response.write CStr(yyyy) + "-" + CStr(mm) + "-" + CStr(dd)
%>

<HTML>
<HEAD>
	<TITLE><%= caption %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TITLE>
	<META http-equiv="Content-Type" content="text/html; charset=euc-kr">
      	<!-- This code can be reused as long as this copyright notice is not removed -->
      	<!-- Copyright 1999 InsideDHTML.com, LLC.  All rights reserved.
           See www.siteexperts.com for more information.
      	-->
	<LINK rel="stylesheet" type="text/css" href="/css/seoul_cp.css">
    <STYLE TYPE="text/css">
         	.today {color:#ff6600; font-weight:bold;font-size:12}
         	.days {font-weight:bold }
         	.tempday {font-size:12 }
    </STYLE>
    <SCRIPT LANGUAGE="JavaScript">
         	// Initialize arrays.
         	var months = new Array("1월", "2월", "3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월");
         	var daysInMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
         	var days = new Array("&nbsp;&nbsp;일", "&nbsp;&nbsp;월", "&nbsp;&nbsp;화", "&nbsp;&nbsp;수", "&nbsp;&nbsp;목", "&nbsp;&nbsp;금", "&nbsp;&nbsp;토" );

         	function getDays(month, year)
         	{
            	// Test for leap year when February is selected.
            	if (1 == month)
              	 return ((0 == year % 4) && (0 != (year % 100))) || (0 == year % 400) ? 29 : 28;
            	else
               	return daysInMonth[month];
        	 }

         	function getToday()
         	{
            	// Generate today's date.

            	<% if yyyy<>"" then %>
				this.now = new Date(<%= yyyy %>,<%= mm %>-1,<%= dd %>,0,0,0);
				<% else %>
				this.now = new Date();
				<% end if %>

            	this.year = this.now.getFullYear();
            	this.month = this.now.getMonth();
            	this.day = this.now.getDate();
 				this.hh = this.now.getHours();
				this.mm = this.now.getMinutes();

         	}

         	// Start with a calendar for today.
         	today = new getToday();

         	function newCalendar() {
            	document.all.ret.value = "";
            	today = new getToday();

            	var parseYear = parseInt(document.all.year[document.all.year.selectedIndex].text);
            	var newCal = new Date(parseYear, document.all.month.selectedIndex, 1);
            	var day = -1;
            	var startDay = newCal.getDay();
            	var daily = 0;

            	if ((today.year == newCal.getFullYear()) && (today.month == newCal.getMonth()))
                 	day = today.day;

            	// Cache the calendar table's tBody section, dayList.
            	var tableCal = document.all.calendar.tBodies.dayList;
            	var intDaysInMonth = getDays(newCal.getMonth(), newCal.getFullYear());

            	for (var intWeek = 0; intWeek < tableCal.rows.length; intWeek++){
                	for (var intDay = 0; intDay < tableCal.rows[intWeek].cells.length; intDay++) {
                  		var cell = tableCal.rows[intWeek].cells[intDay];

                  		// Start counting days.
                  		if ((intDay == startDay) && (0 == daily))
                       		daily = 1;

                  		// Highlight the current day.
                  		cell.className = (day == daily) ? "today" : "tempday";

                  		// Output the day number into the cell.
                  		if ((daily > 0) && (daily <= intDaysInMonth)) {
                     		var str = daily++;
                      		if (str <= 9 )
                         		cell.innerText = "0" + str;
                     	 	else
                         		cell.innerText = str;
                  		}else
                     		cell.innerText = "";
               		}//for문 끝
            	}//for문 끝
        	}

        	function getDate() {
        		// This code executes when the user clicks on a day in the calendar.
            	if ("TD" == event.srcElement.tagName){
               		if ("" != event.srcElement.innerText){
		  				//sDate =   document.all.month.value + "/" + event.srcElement.innerText  + "/" + document.all.year.value;
		  				sDate = document.all.year.value + "-" + document.all.month.value + "-" + event.srcElement.innerText;
                  		//sDate =   document.all.month.value + "." + event.srcElement.innerText ;

	          			document.all.ret.value = sDate;
 		  				window.close();
 					}//if문 끝
 	    		}//if문 끝
        	}
    </SCRIPT>
</HEAD>
<BODY bgcolor="#ffffff" ONLOAD="newCalendar()" OnUnload="window.returnValue = document.all.ret.value;">
   	<INPUT type="hidden" name="ret">
<!----------------------------------------------------------------------------->
<TABLE width="227" border="0" align="center" cellpadding="5" cellspacing="0">
<TR>
<TD>
	<TABLE width="217" border="0" cellspacing="1" cellpadding="0" bgcolor="CCCCCC">
	<TR>
    <TD bgcolor="ffffff">
		<TABLE width="217" border="0" cellspacing="0" cellpadding="0">
        <TR align="center" valign="bottom">
        <TD height="57" colspan="3">


<!----------------------------------------------------------------------------->
 	<TABLE width="203" height="53" border="0" cellspacing="0" cellpadding="0">
   		<TABLE ID="calendar" align="center" border="0">
			<THEAD>
            	<TR>
          			<TD COLSPAN=7 ALIGN=CENTER height=25 width="197"  class="cal2">
                  		<!-- Year combo box -->
                  		<SELECT ID="year" ONCHANGE="newCalendar()" style="width:90">
                     		<SCRIPT LANGUAGE="JavaScript">
                        	// Output years into the document.
                        	// Select current year.
                        		for (var intLoop = 2000; intLoop < 2021; intLoop++)
                           			document.write("<OPTION VALUE= " + intLoop + " " + (today.year == intLoop ? "Selected" : "") +
                                           ">" + intLoop);
                     		</SCRIPT>
                  		</SELECT>&nbsp;
                  		<!-- Month combo box -->
                  		<SELECT ID="month" ONCHANGE="newCalendar()" style="width:90">
                     		<SCRIPT LANGUAGE="JavaScript">
                        		// Output months into the document.
                        		// Select current month.
                        		for (var intLoop = 0; intLoop < months.length;intLoop++)
                        		{
                           			if (intLoop < 9)
                               			document.write("<OPTION VALUE= 0" + (intLoop + 1) + " " +
                                               (today.month == intLoop ? "Selected" : "") + ">" +
                                                months[intLoop]);
                          			else
                               			document.write("<OPTION VALUE= " + (intLoop + 1) + " " +
                                               (today.month == intLoop ?
                                                "Selected" : "") + ">" +
                                                months[intLoop]);
                        		}
                     		</SCRIPT>
                  		</SELECT>
               		</TD>
            	</TR>
				<TR><TD COLSPAN="7" ></TD></TR>
				<tr><td colspan="7"></td></tr>

            	<TR align="center" bgcolor="E6E6E6">

               		<!-- Generate column for each day. -->
               		<SCRIPT LANGUAGE="JavaScript">
                  		// Output days.
                  		for (var intLoop = 0; intLoop < days.length; intLoop++){
                        	if (intLoop==0){
                             	document.write("<TD width='17' class='calr'>" + days[intLoop] + "</TD>");
                        	}else{
		         				document.write("<TD width='17' class='cal2' >" + days[intLoop] + "</TD>");
                        	}
							//if(intLoop != days.length)
								//document.write("<td width='15'></td>");
                  		}
               		</SCRIPT>

            	</TR>
				<!--<tr><td background="/image/cal_bg.gif" height="1" colspan="7"></td></tr>-->
				</THEAD>
         		<TBODY ID="dayList" ALIGN=CENTER ONCLICK="getDate()" >
            	<!-- Generate grid for individual days. -->
            	<SCRIPT LANGUAGE="JavaScript">
               		for (var intWeeks = 0; intWeeks < 6; intWeeks++)
               		{
                  		document.write("<TR align='center'>");

                  		for (var intDays = 0; intDays < days.length; intDays++){
                        	document.write("<TD class='cal'></TD>");
                  		}
                  		document.write("</TR>");
               		}

            	</SCRIPT>
         		</TBODY>
     	</TABLE>
    </TABLE>

<!--------------------------------------------------------------------------->
		</TD>
		</TR>
    	</TABLE>
    </TD>
    </TR>
    </TABLE>
</TD>
</TR>
</TABLE>
<!--------------------------------------------------------------------------->
</BODY>
</HTML>
<Script Language="JavaScript">
	function Cancel() {
		document.all.ret.value = "";
		window.close();
	}
</script>
