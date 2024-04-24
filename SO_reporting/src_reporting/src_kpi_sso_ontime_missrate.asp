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
	     str = str.replace("#", encodeURIComponent("#"));
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
response.write "<td>" & "segment" & "</td>"
response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
response.write "<td>" & "issuetype" & "</td>"
response.write "<td>" & "component" & "</td>"
response.write "<td>" & "summary" & "</td>"
response.write "<td>" & "team" & "</td>"
response.write "<td>" & "state"   & "</td>"
response.write "<td>" & "resolution"   & "</td>"
response.write "<td>" & "created"   & "</td>"
response.write "<td>" & "updated"   & "</td>"
response.write "<td>" & "resolutiondate"   & "</td>"
response.write "<td>" & "state change overview"   & "</td>"
response.write "<td>" & "date set to done"   & "</td>"
response.write "<td>" & "item duedate"   & "</td>"
response.write "<td>" & "item target duedate"   & "</td>"
response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
response.write "<td>" & "Late?" & "</td>"

response.write "<td>" & "set to Backlog"   & "</td>"
response.write "<td>" & "set to Breakdown"   & "</td>"
response.write "<td>" & "set to Ready for development"   & "</td>"
response.write "<td>" & "set to Development"   & "</td>"
response.write "<td>" & "set to Ready For Validation"   & "</td>"
response.write "<td>" & "set to Validation"   & "</td>"
response.write "<td>" & "set to Done"   & "</td>"
response.write "<td>" & "set to On Hold"   & "</td>"
response.write "</tr>"

dt = dateserial(mid(yearmonth,1,4), mid(yearmonth,5,2), 1)
'format into 2018/06/30
from_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)
dt = dateadd("m", 1, dt)
dt = dateadd("d", -1, dt)
to_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)

dim jql
jql = ""
jql = jql & " project = ""SSE"" "
jql = jql & " AND (status changed to ""done"" during (""" & from_ & """, """ & to_ & """)) "
'jql = jql & " AND not ((status = Done and resolution = Cancelled)) "
jql = jql & " and issuetype in (""User Story"",""Sub-task"") "

'response.write jql
'response.end

xml = getJiraItems(jql)
objXML.LoadXML xml

	xml = ""
	xml = xml & "<item_result>" & vbcrlf
	'now rebuild the XML using addiotnal filters
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	i = 1
	For Each item_ In nodelist
		bln = true
		For Each field_in_item_ In item_.ChildNodes

			if false then
			'check if correct issuetype
			if field_in_item_.BaseName = "issuetype" then
			if instr("||User Story||Sub-task||", "||" & field_in_item_.Text & "||") > 0 then
				bln = false
			end if
			end if
			end if

			if false then 'need to check with Rodrigo/Dagmara
			'check if correct components
			if field_in_item_.BaseName = "components" then
			arr = split(field_in_item_.Text, "||")
			for a = lbound(arr) to ubound(arr)
			if arr(a) <> "" then
				if instr(lcase("||Innovation platforms||STS Support||SO Regions||Events organization||Hosting Sharepoint sites||R&D||STS internal processes||STS Newsletter||TD platform||"),lcase("||" & arr(a) & "||")) > 0 then
					bln = false
				end if
			end if
			next
			end if
			end if

			if false then 'this condition is done by getJiraItems.aspx
			'check if done in this month
			if field_in_item_.BaseName = "transitions_history" then
			    bln2 = False
			    arr = Split(field_in_item_.Text, "@@")
			    For a = LBound(arr) To UBound(arr)
			    If arr(a) <> "" Then
				arr2 = Split(arr(a), "||")
				'For a2 = LBound(arr2) To UBound(arr2)
				'If arr2(a2) <> "" Then
				    If arr2(2) = "Done" Then
				    If Mid(arr2(0), 1, 6) = yearmonth Then
					bln2 = True
				    End If
				    End If
				'End If
				'Next
			    End If
			    Next
			    If bln2 = True Then ' it is set to done in this month
				'do nothing
			    Else
				'not a good item
				bln = False
			    End If
			end if
			end if 'set to done in this month

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
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0

	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		'determine enddate
		target_duedate = getFieldValue(item_,"target_duedate")
		if target_duedate = "" then
			target_duedate = getFieldValue(item_,"sprintcompletedate")
		end if
		settodone = getSettoDone(getFieldValue(item_,"transitions_history"))

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
		response.write "<td>" & "SSE" & "</td>"
		response.write "<td>" & getFieldValue(item_,"key") & "</td>"
		response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
		response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
		response.write "<td>" & getFieldValue(item_,"team") & "</td>"
		response.write "<td>" & getFieldValue(item_,"status") & "</td>"
		response.write "<td>" & getFieldValue(item_,"resolution") & "</td>"
		'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"created")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"updated")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"resolutiondate")) & "</td>"
		response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getSettoDone(getFieldValue(item_,"transitions_history"))) & "</td>"
		'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
		'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"duedate")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"target_duedate")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"sprintcompletedate")) & "</td>"

		'For Each field_in_item_ In item_.ChildNodes
		'response.write "<td>" & field_in_item_.text & "</td>"
		'next
		if settodone > target_duedate then
			response.write "<td>1</td>"
		else
			response.write "<td>0</td>"
		end if

		'function returns tds so no need to put td here
		response.write (displayTransitions(getFieldValue(item_,"transitions_history")))
		response.write "</tr>"
		i = i + 1
		end if
	next

'14JUN2019 - disable MOEBIUS (mail Mireille)
if false then

'RUN MOEBIUS HERE

jql = ""
jql = jql & " project = ""MOEBIUS"" "
'jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
'jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
jql = jql & " AND not ((status = Done and resolution = Cancelled) or (status = cancelled)) "
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
		response.write "<td>" & "SSO" & "</td>"
		response.write "<td>" & getFieldValue(item_,"key") & "</td>"
		response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
		response.write "<td>" & getFieldValue(item_,"labels") & "</td>"
		response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
		response.write "<td>" & getFieldValue(item_,"status") & "</td>"
		response.write "<td>" & getFieldValue(item_,"resolution") & "</td>"
		'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"created")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"updated")) & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"resolutiondate")) & "</td>"
		response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
		response.write "<td>" & toDD_MM_YYYY(getSettoDone(getFieldValue(item_,"transitions_history"))) & "</td>"
		'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
		'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
		response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"duedate")) & "</td>"
		response.write "<td>" & "" & "</td>" 'target duedate : N/A for MOEBIUS
		response.write "<td>" & "" & "</td>" 'duedate via sprint : N/A for MOEBIUS

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
end if 'moebius = false

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"

'end if 'prj loop
'next 'prj loop

'response.write tmp
%>
</body>
</html>
<%




function getSettoDone(allstates)
getSettoDone = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Done" then
	getSettoDone = arr2(0)
	end if
end if
next

end function

function donethismonth(allstates, ym)
donethismonth = false
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
		donethismonth = true
	end if
	end if
end if
next
end function

Function getJiraItems(jql)
'response.write "https://soreporting.azurewebsites.net/src_reporting/getJiraitems.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", "https://soreporting.azurewebsites.net/src_reporting/getJiraitemsSSE.aspx?jql=" & jql & "&rnd=" & rnd & "", False
xmlhttp.send
getJiraItems = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

Function getJiraItemsMoebius(jql)
'response.write "https://soreporting.azurewebsites.net/src_reporting/getJiraitemsMoebius.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", "https://soreporting.azurewebsites.net/src_reporting/getJiraitemsMoebius.aspx?jql=" & jql & "&rnd=" & rnd & "", False
xmlhttp.send
getJiraItemsMoebius = xmlhttp.responseText
Set xmlhttp = Nothing
End Function


function itemRejected(allstates)
'in = 20180226||Open||Breakdown@@20180226||Breakdown||Breakdown Done@@20180302||Breakdown Done||Development@@20180302||Development||Development Done@@20180305||Development Done||Validation@@20180305||Validation||Done
dim arr
dim arr2
dim a
itemRejected = false
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(1) = "Validation" and arr2(2) <> "Done" then
	itemRejected = true
	end if
end if
next
end function



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

Function displayTransitions(allstates)
Dim set_to_Backlog
Dim set_to_Breakdown
Dim set_to_Readyfordevelopment
Dim set_to_Development
Dim set_to_ReadyForValidation
Dim set_to_Validation
Dim set_to_Done
Dim set_to_OnHold

set_to_Backlog = ""
set_to_Breakdown = ""
set_to_Readyfordevelopment = ""
set_to_Development = ""
set_to_ReadyForValidation = ""
set_to_Validation = ""
set_to_Done = ""
set_to_OnHold = ""

'in = 20180226||Open||Breakdown@@20180226||Breakdown||Breakdown Done@@20180302||Breakdown Done||Development@@20180302||Development||Development Done@@20180305||Development Done||Validation@@20180305||Validation||Done
Dim arr
Dim arr2
Dim a
arr = Split(allstates, "@@")
For a = LBound(arr) To UBound(arr)
If arr(a) <> "" Then
    arr2 = Split(arr(a), "||")
    If arr2(2) = "Backlog" 					Then set_to_Backlog = 				set_to_Backlog & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Breakdown" 				Then set_to_Breakdown = 			set_to_Breakdown & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Ready for development" 	Then set_to_Readyfordevelopment = 	set_to_Readyfordevelopment & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Development" 				Then set_to_Development = 			set_to_Development & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Ready For Validation" 	Then set_to_ReadyForValidation = 	set_to_ReadyForValidation & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Validation" 				Then set_to_Validation = 			set_to_Validation & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "Done" 					Then set_to_Done = 					set_to_Done & toDD_MM_YYYY(arr2(0)) & "<br>"
    If arr2(2) = "On Hold" 					Then set_to_OnHold = 				set_to_OnHold & toDD_MM_YYYY(arr2(0)) & "<br>"
End If
Next

If set_to_Backlog <> "" Then set_to_Backlog = Mid(set_to_Backlog, 1, Len(set_to_Backlog) - 4)
If set_to_Breakdown <> "" Then set_to_Breakdown = Mid(set_to_Breakdown, 1, Len(set_to_Breakdown) - 4)
If set_to_Readyfordevelopment <> "" Then set_to_Readyfordevelopment = Mid(set_to_Readyfordevelopment, 1, Len(set_to_Readyfordevelopment) - 4)
If set_to_Development <> "" Then set_to_Development = Mid(set_to_Development, 1, Len(set_to_Development) - 4)
If set_to_ReadyForValidation <> "" Then set_to_ReadyForValidation = Mid(set_to_ReadyForValidation, 1, Len(set_to_ReadyForValidation) - 4)
If set_to_Validation <> "" Then set_to_Validation = Mid(set_to_Validation, 1, Len(set_to_Validation) - 4)
If set_to_Done <> "" Then set_to_Done = Mid(set_to_Done, 1, Len(set_to_Done) - 4)
If set_to_OnHold <> "" Then set_to_OnHold = Mid(set_to_OnHold, 1, Len(set_to_OnHold) - 4)

displayTransitions = ""
displayTransitions = displayTransitions & "<td>" & set_to_Backlog & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_Breakdown & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_Readyfordevelopment & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_Development & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_ReadyForValidation & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_Validation & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_Done & "</td>"
displayTransitions = displayTransitions & "<td>" & set_to_OnHold & "</td>"

End Function


%>