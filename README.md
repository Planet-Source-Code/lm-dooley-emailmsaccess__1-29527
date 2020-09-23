<div align="center">

## EmailMSACCESS


</div>

### Description

Allows you to adress and write an MS Access report directly into an email message. (saves attaching pesky rtf files making it easier for the reader. The code attached is what we use on a weekly basis and can be used as a template with your tables/queries
 
### More Info
 
The user of this code will have to go in and set up using his/her database

The report length will be limited to the maximum lentgh of a strin variable. I've figured out how to go bigger than that but I don't need to at this time

setting the toggle incorrectly at the end of the sendobject command can let you send the email without seeing it first


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[LM Dooley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lm-dooley.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VBA MS Access
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lm-dooley-emailmsaccess__1-29527/archive/master.zip)





### Source Code

```

Private Sub Command_Email_Click()
Dim dbs As Database, rst As Recordset
Dim strcrlist1
  Dim strsql As String
   Dim stDocName As String
   Set dbs = CurrentDb
   Dim test, i, intspace
   intspace = "   "
strsql = "SELECT Main_data_table.Date_concurrence, "
strsql = strsql & "Main_data_table.ID, "
strsql = strsql & "Main_data_table.CR_number, "
strsql = strsql & "Main_data_table.Resp_dept, "
strsql = strsql & "Main_data_table.Resp_indiv, "
strsql = strsql & "Main_data_table.Description, "
strsql = strsql & "Main_data_table.Signif_level, "
strsql = strsql & "Main_data_table.Eval_level "
strsql = strsql & "FROM Main_data_table "
strsql = strsql & "WHERE (((Main_data_table.Date_concurrence) Between #" & [Forms]![Phil is Sick or on Vacation]![datea] & "# And #" & [Forms]![Phil is Sick or on Vacation]![dateb] & "#)) "
strsql = strsql & "ORDER BY Main_data_table.CR_number;"
  Set rst = dbs.OpenRecordset(strsql)
With rst
rst.MoveLast
Debug.Print rst.RecordCount
End With
test = rst.RecordCount
Me!Test1 = test
strcrlist = ""
With rst
rst.MoveFirst
For i = 1 To test
strcrlist = strcrlist & Chr$(127) & !CR_number & intspace & !Signif_level & intspace & !Resp_dept & intspace & intspace & !Resp_indiv & Chr$(13) & Chr$(10) & !Description & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
rst.MoveNext
Next i
strcrlist = "CR#" & intspace & "Sig Lvl" & intspace & "Resp Dept" & intspace & intspace & "Resp Indv/ Description" & Chr$(13) & Chr$(10) & strcrlist
End With
Me!test2 = strcrlist
  Dim list
  'stDocName = "Condition Reports created this week auto"
   strnsra = "Please review the following CR listing for possible NS and RA reportable implications and return the package to me. If any items need further looking into, I have all the associated information and will gladly share any with you."
 strnsra = strnsra & Chr$(13) & Chr$(10) & "This review may be done electronically by responding to this email" & Chr$(13) & Chr$(10)
  strcrlist = strnsra & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & strcrlist
  DoCmd.SendObject , , , "Michael Johnson", "Phillip Wood", , "CRs Generated This Week-NS and RA Review", strcrlist
End Sub
```

