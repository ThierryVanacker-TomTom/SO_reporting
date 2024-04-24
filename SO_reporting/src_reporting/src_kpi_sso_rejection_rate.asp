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

dim dt
dim from_
dim to_

yearmonth = request.querystring("yearmonth")
if (yearmonth = "") then yearmonth = right("0000" & year(date),4) & right("00" & Month(date),2)

response.write "<b>Rejection rate</b><br>"
response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table'>"
response.write "<tr>"
response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
response.write "<td>" & "issuetype" & "</td>"
response.write "<td>" & "component" & "</td>"
response.write "<td>" & "summary" & "</td>"
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
response.write "<td>" & "rejected"   & "</td>"
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

			if false then 'done by JQL query
			'check if correct issuetype
			if field_in_item_.BaseName = "issuetype" then
			'Task Epic Bug
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
			end if 'set to done in this month
			end if 'FALSE - done by GetJiraItems

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
'	response.write (replace(xml, "<" , "["))
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		if bln = true then
			response.write "<tr>"
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
			response.write "<td>" & getFieldValue(item_,"components") & "</td>"
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
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"target_duedate")) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"sprintcompletedate")) & "</td>"

			if itemRejected(getFieldValue(item_,"transitions_history")) then
				j = j + 1
				response.write "<td>" & "yes" & "</td>"
			else
				response.write "<td>" & "no" & "</td>"
			end if

			'For Each field_in_item_ In item_.ChildNodes
			'response.write "<td>" & field_in_item_.text & "</td>"
			'next
			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"
	'response.write i & " items were set to done in " & yearmonth & " however " & j & " items have been rejected<br>"
	'response.write "<br>"

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
'response.write "https://soreporting.azurewebsites.net/src_reporting/getJiraitemsSSE.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", "https://soreporting.azurewebsites.net/src_reporting/getJiraitemsSSE.aspx?jql=" & jql & "&rnd=" & rnd & "", False
xmlhttp.send
getJiraItems = xmlhttp.responseText
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
	if lcase(arr2(1)) = lcase("Validation") and lcase(arr2(2)) <> lcase("Done") then
	'17DEC2019 - from Validation to On hold is not a rejection (cfr mail Ryno Botes)
	if lcase(arr2(1)) = lcase("Validation") and lcase(arr2(2)) = lcase("On hold") then
	else
	itemRejected = true
	end if
	end if
end if
next
end function
%>