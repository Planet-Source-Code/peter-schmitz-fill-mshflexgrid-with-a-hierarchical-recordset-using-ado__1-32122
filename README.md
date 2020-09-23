<div align="center">

## Fill MsHFlexGrid with a Hierarchical RecordSet \(Using ADO\)


</div>

### Description

Fills a Hierarchical Flexgrid based on a Hierarchical Recordset (a one-to-many relationship). Uses the Northwind database. Code is documented.
 
### More Info
 
A reference to MS ADO should be made (Project, References, check Microsoft ActiveX Data Objects Library)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Schmitz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-schmitz.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-schmitz-fill-mshflexgrid-with-a-hierarchical-recordset-using-ado__1-32122/archive/master.zip)

### API Declarations

```
Public RS As Recordset
Public CN As Connection
```


### Source Code

```
Dim SQL As String
  Set RS = New Recordset
  Set CN = New Connection
  Dim rsChild As Variant
  ' Define SQL String
  ' The statement between the first pair of brackets defines the
  ' Parent-recordset.
  ' The statement between the second pair of brackets defines the
  ' child-recordset. The WHERE clause contains a questionmark, which
  ' identifies this as a parameterised value.
  ' The RELATE statement defines which columns the recordsets connect with.
  ' In this case, PARAMETER 0 points back to the questionmark used earlier.
  ' Basically this is the equivalent of the JOIN .. ON statement in T-SQL.
  ' For more info about hierarchical recordset creations look here:
  ' http://support.microsoft.com/default.aspx?scid=kb;en-us;Q189657
  SQL = "SHAPE {SELECT FirstName, LastName, EmployeeID FROM employees} APPEND ({SELECT OrderID FROM orders WHERE EmployeeID = ?} AS Orders RELATE EmployeeID TO PARAMETER 0)"
  ' Open connection
  ' We use MSDataShape because of the hierarchical recordset.
  ' Change Servername to your own SQL-Server, and alter the login-ID / password
  CN.Open "Provider=MSDataShape;Driver={SQL Server};Server=RNT07;Database=NorthWind", "sa", ""
  RS.Open SQL, CN
  ' The following part can be used for debugging purposes
  ' It will spit the Recordset records into the Immediate Window (CTRL + G)
  '
  'While Not RS.EOF
  '   Debug.Print RS("FirstName"), RS("Lastname")
  '     rsChild = RS("Orders")
  '     While Not rsChild.EOF
  '       Debug.Print rsChild(0)
                    ' rsChild contains just one column.
                    ' If you'd have more columns
                    ' simply add ,rsChild(1) etc
  '       rsChild.MoveNext
  '     Wend
  '     RS.MoveNext
  'Wend
  Set MSflexGrid1.DataSource = RS
  ' Close Recordset object and destroy it
  RS.Close
  Set RS = Nothing
  ' Close Connection object and destroy iy
  CN.Close
  Set CN = Nothing
```

