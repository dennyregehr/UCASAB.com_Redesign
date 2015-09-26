<%
Option Explicit
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

Const adUseServer = 2
Const adUseClient = 3

' Uncomment the next line when you are debugging.
 On Error Resume Next

Dim oConn             ' Database connection object
Dim oRS               ' Recordset object
Dim sSQL              ' SQL query string
Dim sObjectName         ' NEW
Dim gsResultsPageName      ' NEW
Dim sWhere                ' WHERE clause for the SQL Query
Dim sSort             ' Sort string used in ORDER BY in SQL query
Dim iCounter          ' Looping counter for interating through the records on current page
Dim sMsg              ' Anyy general messages to be posted to user
Dim sErr              ' Error message to be returned to the user
Dim iTotalRecords     ' Total number of records in table
Dim iMaxRecords       ' Setting - Maximum number of records to be retreived from table
Dim iPageSize         ' Setting - Number of records to be viewed at one time
Dim iPageCount        ' The number of pages in the recordset.
Dim iRecordCount      ' The number of records in the recordset.
Dim iPage             ' The current page that we are on.
Dim iRecord           ' Counter used to iterate through the recordset.
Dim iStart            ' The record that we are starting on.
Dim iFinish           ' The record that we are finishing on.
Dim sRowColor         ' Row color to highlight every other record in results
Dim iPosition         ' Detail page
Dim sAction           ' ?????
Dim iRecordsUpdated   ' Detail Page - Records Affected
Dim iRecordsDeleted   ' Detail Page - Records Affected
Dim sSubTitle         ' Detail Page - Records Affected
Dim oFSO              ' Detail Page - File System Object - To verify image

' JScript Timer variables
Dim iTimeStart        ' This the start time - used to determine how long page takes to create.
Dim iTimeEnd          ' This the finish time - used to determine how long page takes to create.

%>
<!-- #INCLUDE file="DB2ASP_security.asp" -->
<%

' JScript - Start Timer
iTimeStart = GetTimeStamp()
sErr = ""
%>
<SCRIPT language="JScript" RUNAT=Server>
// This is a JScript timer function to measure how long it takes to render this page
// You can delete this functionality if you'd like.
// To Remove: Delete this function and all other lines commented with "JScript Timer" below
function GetTimeStamp() {
    var dt = new Date();
    return Date.UTC(dt.getYear(),dt.getMonth(),dt.getDate(), dt.getHours(),dt.getMinutes(),dt.getSeconds(),dt.getMilliseconds());
}
</SCRIPT>
<!-- #INCLUDE file="DB2ASP_conn.asp" -->
