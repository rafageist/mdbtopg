Attribute VB_Name = "mdbtopg"
Option Compare Database

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' MDB2Pg
'' Create SQL Script for PostgreSQL from MS Access Database
'' by Rafael Rodríguez Ramírez
'' 2/8/2009
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Fix name of tables and columns (clean "tildes" and spaces)

Function NameRepair(ByVal n As String) As String
    n = Replace(n, "á", "a")
    n = Replace(n, "é", "e")
    n = Replace(n, "í", "i")
    n = Replace(n, "ó", "o")
    n = Replace(n, "ú", "u")
    n = Replace(n, "ñ", "n")
    n = Replace(n, "Ñ", "N")
    n = Replace(n, "ü", "u")
    n = Replace(n, " ", "")
    If LCase(n) = "desc" Then n = "descrip"
    NameRepair = n
End Function

' Type convert from dbTypes to pgTypes

Function MapaType(ByVal tipo As Long) As String

    Select Case tipo
        Case dbBigInt: MapaType = "bigint"
        Case dbBinary: MapaType = "binary"
        Case dbBoolean: MapaType = "bool"
        Case dbByte: MapaType = "integer"
        Case dbChar: MapaType = "char"
        Case dbCurrency: MapaType = "float"
        Case dbDate: MapaType = "date"
        Case dbDecimal: MapaType = "float"
        Case dbDouble: MapaType = "float"
        Case dbFloat: MapaType = "float"
        Case dbGUID: MapaType = "guid"
        Case dbInteger: MapaType = "integer"
        Case dbLong: MapaType = "bigint"
        Case dbLongBinary: MapaType = "varchar"
        Case dbMemo: MapaType = "text"
        Case dbNumeric: MapaType = "numeric"
        Case dbSingle: MapaType = "float"
        Case dbText: MapaType = "varchar"
        Case dbTime: MapaType = "time"
        Case dbTimeStamp: MapaType = "timestamp"
        Case dbVarBinary: MapaType = "varbinary"
    End Select
    
End Function

'Create SQL Script
'@param path contain the file system path to script
Sub mdbtopg(path As String)
    
    Open path For Output As #1
    
    For i = 1 To CurrentDb.TableDefs.Count
        If Mid(CurrentDb.TableDefs(i - 1).Name, 1, 4) <> "MSys" Then
            Debug.Print "[INFO] Working with table " & CurrentDb.TableDefs(i - 1).Name
            p = i - 1

            Print #1, "CREATE TABLE " & NameRepair(CurrentDb.TableDefs(p).Name) & "("
            
            'primary key
            For j = 1 To CurrentDb.TableDefs(p).Indexes.Count
                If CurrentDb.TableDefs(p).Indexes(j - 1).Primary = True Then
                    Print #1, "  PRIMARY KEY (";
                    For k = 1 To CurrentDb.TableDefs(p).Indexes(j - 1).Fields.Count
                        Print #1, NameRepair(CurrentDb.TableDefs(p).Indexes(j - 1).Fields(k - 1).Name) & IIf(k < CurrentDb.TableDefs(p).Indexes(j - 1).Fields.Count, ",", "");
                    Next
                    Print #1, "),"
                    Exit For
                End If
            Next
            
            'fields
            For j = 1 To CurrentDb.TableDefs(p).Fields.Count
                Print #1, "  " & NameRepair(CurrentDb.TableDefs(p).Fields(j - 1).Name) & " " & MapaType(CurrentDb.TableDefs(p).Fields(j - 1).Type) & IIf(j < CurrentDb.TableDefs(p).Fields.Count, ",", "")
            Next
            
            Print #1, ");"
            Print #1, "--"
        End If
    Next
        
    'Foreign keys
    
    For i = 0 To CurrentDb.Relations.Count - 1
        Debug.Print "[INFO] Working with relation " & CurrentDb.Relations(i).Name
        Print #1, "ALTER TABLE " & NameRepair(CurrentDb.Relations(i).ForeignTable) & " ADD FOREIGN KEY (";
        For j = 1 To CurrentDb.Relations(i).Fields.Count
            Print #1, NameRepair(CurrentDb.Relations(i).Fields(j - 1).ForeignName) & IIf(j < CurrentDb.Relations(i).Fields.Count, ",", "");
        Next
        Print #1, ") REFERENCES " & NameRepair(CurrentDb.Relations(i).Table) & "(";
        For j = 1 To CurrentDb.Relations(i).Fields.Count
            Print #1, NameRepair(CurrentDb.Relations(i).Fields(j - 1).Name) & IIf(j < CurrentDb.Relations(i).Fields.Count, ",", "");
        Next
        Print #1, ") ON DELETE CASCADE ON UPDATE CASCADE;"
    Next
        
    Close
    
End Sub
