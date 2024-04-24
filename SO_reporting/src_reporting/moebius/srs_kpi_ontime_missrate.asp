<%option explicit%>
<%server.scripttimeout=10*60 '10minutes%>
<!--#include file="modFunctions.inc" -->
<html>
<head>
<script>
function downloadCSV(tbl,filename)
{
	//loop through table rows & cols
	var obj = document.getElementById(tbl);
	var row;
	var col;
	var csv='';
	for (var i = 0; row = obj.rows[i]; i++)
	{
	   //iterate through rows
	   //rows would be accessed using the "row" variable assigned in the for loop
	   for (var j = 0; col = row.cells[j]; j++)
	   {
	     //iterate through columns
	     //columns would be accessed using the "col" variable assigned in the for loop
	     var str = col.innerText;
	     str = str.replace("#", encodeURI("#"));
	     csv+='"'+str+'";';
	   }
	   csv+='\n';
	}
	filename = filename;
	if (!csv.match(/^data:text\/csv/i))
	{
	csv = 'data:text/csv;charset=utf-8,' + encodeURI('\uFEFF'+csv); //force the file to be UTF8
	}
	var data = csv;
	var link = document.createElement('a');
	link.setAttribute('href', data);
	link.setAttribute('download', filename);
	document.body.appendChild(link); //needed for FF
	link.click();
}
</script>
</head>
<body style='font-family:verdana;font-size:8pt;'>
<%
Dim tmp
Dim tmp2
Dim url
dim id
dim prj
'dim arr
dim issuetype
dim event_
dim sum_
dim desc_
dim i
dim j

dim days
dim tot_days
dim vali
dim invali

dim arr
dim a
dim arr2
dim a2
dim arrP
dim aP

dim xml
dim objXML
set objXML = createobject("MSXML.DOMDocument")
Dim item_
Dim field_in_item_
dim nodelist

dim label
dim hist_
dim comp
dim bln
dim bln2

dim duedate
dim target_duedate
dim sprint
dim sprintenddate

dim yearmonth
yearmonth = request.querystring("yearmonth")
dim settodone

dim fso
dim f

arrP = split("STSPES,STSCNT,MOEBIUS,RES", ",")
'arr = split("STSPES", ",")
'arr = split("STSCNT,STSPES", ",")
arrP = split(request.querystring("project"), ",")
'response.write request.querystring("project") & "<br>"
'response.write ubound(arr) & "<br>"

dim dt
dim from_
dim to_

yearmonth = request.querystring("yearmonth")
if (yearmonth = "") then yearmonth = right("0000" & year(date),4) & right("00" & Month(date),2)

response.write "<b>On time missrate</b><br>"
response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table'>"
response.write "<tr>"
response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
response.write "<td>" & "issuetype" & "</td>"
response.write "<td>" & "labels" & "</td>"
response.write "<td>" & "summary" & "</td>"
response.write "<td>" & "state"   & "</td>"
response.write "<td>" & "created"   & "</td>"
response.write "<td>" & "updated"   & "</td>"
response.write "<td>" & "resolutiondate"   & "</td>"
response.write "<td>" & "state change overview"   & "</td>"
response.write "<td>" & "date set to done"   & "</td>"
response.write "<td>" & "item duedate"   & "</td>"
response.write "<td>" & "Late?" & "</td>"
response.write "</tr>"

dt = dateserial(mid(yearmonth,1,4), mid(yearmonth,5,2), 1)
'format into 2018/06/30
from_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)
dt = dateadd("m", 1, dt)
dt = dateadd("d", -1, dt)
to_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)

dim jql
jql = ""
jql = jql & " project = ""MOEBIUS"" "
'jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
'jql = jql & " AND ""Assigned Unit"" in (""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
jql = jql & " AND ("
jql = jql & " status changed to ""done"" during (""" & from_ & """, """ & to_ & """)"
'jql = jql & " OR status changed to ""closed"" during (""" & from_ & """"", """ & to_ & """)"
'jql = jql & " OR status changed to ""planned"" during (""" & from_ & """"", """ & to_ & """)"
jql = jql & " )"

'response.write jql
'response.end

xml = getJiraItemsMoebius(jql)
'response.write xml & "<hr noshade>"
	objXML.LoadXML xml

	xml = ""
	xml = xml & "<item_result>" & vbcrlf
	'now rebuild the XML using addiotnal filters
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	i = 1
	For Each item_ In nodelist
		bln = true
		For Each field_in_item_ In item_.ChildNodes
			'check if correct labels (exclude entries with label incident
			if field_in_item_.BaseName = "labels" then
			arr = split(field_in_item_.Text, "||")
			for a = lbound(arr) to ubound(arr)
			if arr(a) <> "" then
				if instr(lcase("||incident||"),lcase("||" & arr(a) & "||")) > 0 then
					bln = false
				end if
			end if
			next
			end if
		Next
		if bln = true then
		xml = xml & "<item>" & vbcrlf
		For Each field_in_item_ In item_.ChildNodes
			xml = xml & "<" & field_in_item_.BaseName & ">" & xmlsafe("" & field_in_item_.Text) & "</" & field_in_item_.BaseName & ">" & vbcrlf
		Next
		xml = xml & "</item>" & vbcrlf
		end if

		i = i + 1
	Next
	xml = xml & "</item_result>" & vbcrlf
	'response.write (replace(xml, "<" , "["))
'response.write xml & "<hr noshade>"
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0

	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		'determine enddate
		target_duedate = getFieldValue(item_,"duedate")
		settodone = getSettoDoneMoebius(getFieldValue(item_,"transitions_history"))

		if target_duedate = "" then
			bln = false
		end if

		if bln = true then
		if settodone > target_duedate then
			j = j + 1
			response.write "<tr style='background-color:red'>"
		else
			response.write "<tr>"
		end if
		response.write "<td>" & getFieldValue(item_,"key") & "</td>"
		response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
		response.write "<td>" & getFieldValue(item_,"labels") & "</td>"
		response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
		response.write "<td>" & getFieldValue(item_,"status") & "</td>"
		'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"created")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"updated")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"resolutiondate")) & "</td>"
		response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getSettoDoneMoebius(getFieldValue(item_,"transitions_history"))) & "</td>"
		'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
		'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"duedate")) & "</td>"

		'For Each field_in_item_ In item_.ChildNodes
		'response.write "<td>" & field_in_item_.text & "</td>"
		'next
		if settodone > target_duedate then
			response.write "<td>1</td>"
		else
			response.write "<td>0</td>"
		end if
		response.write "</tr>"
		i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"

'end if 'prj loop
'next 'prj loop

'response.write tmp
%>
</body>
</html>
<%
function getSettoDoneMoebius(allstates)
getSettoDoneMoebius = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Done" then
	getSettoDoneMoebius = arr2(0)
	end if
end if
next

end function

function donethismonthMoebius(allstates, ym)
donethismonthMoebius = false
'allstates = date||from||to@@...
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	'response.write ym & " - " & arr2(0) & "-" & arr2(1) & "-" & arr2(2) & "-" & "<br>"
	if arr2(2) = "Done" then
	if mid(arr2(0), 1, 6) = ym then
		donethismonthMoebius = true
	end if
	end if
end if
next
end function

function itemRejectedMoebius(allstates)
'in = 20180226||Open||Breakdown@@20180226||Breakdown||Breakdown Done@@20180302||Breakdown Done||Development@@20180302||Development||Development Done@@20180305||Development Done||Validation@@20180305||Validation||Done
dim arr
dim arr2
dim a
itemRejectedMoebius = false
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(1) = "Validation" and arr2(2) <> "Done" then
	itemRejectedMoebius = true
	end if
end if
next
end function

Function getJiraItemsMoebius(jql)
'response.write "http://nlsrvwp-pcd02.net-10-67-0-0.tt3.com/src_reporting/getJiraitemsMoebius.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", "http://nlsrvwp-pcd02.net-10-67-0-0.tt3.com/src_reporting/getJiraitemsMoebius.aspx?jql=" & jql & "&rnd=" & rnd & "", False
xmlhttp.send
getJiraItemsMoebius = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

%>