'****************************************************************
'*
'*                      COUNTU UDF Module
'*
'* Module which adds the following User Defined Functions COUNTU,
'* COUNTUIF, COUNTUIFS which count the number of items in a given
'* range excluding duplicates.
'*
'* Module requires a reference to the Microsoft Scripting Runtime
'* library.
'*
'* Version: 1.3
'* Created: 13-11-2014
'*
'****************************************************************

Option Explicit

'****************************************************************
'*
'* COUNTU
'*
'* Counts the number of cells in a range excluding any duplicates
'* found.
'*
'* Usage:
'*
'* =COUNTU(range)
'*
'****************************************************************
Public Function COUNTU(ParamArray ranges() As Variant)

    Dim Values As Dictionary
    Dim targetRange As Variant
    Dim trimmedRange As range
  
    Set Values = New Dictionary
    
    For Each targetRange In ranges

        Set trimmedRange = TrimRangeToTarget(targetRange, ActiveSheet.UsedRange)
    
        Dim cell As range
        
        For Each cell In trimmedRange
        
            If Not Values.Exists(cell.Value) Then
            
                If Not IsError(cell.Value) Then
    
                    If cell.Value <> "" Then
                
                        Values.Add cell.Value, 1
                    
                    End If
                
                End If
            
            End If
            
        Next cell
        
        Set trimmedRange = Nothing
    
    Next targetRange
    
    COUNTU = Values.Count

    Set Values = Nothing

End Function

'****************************************************************
'*
'* COUNTUIF
'*
'* Counts the number of cells in a range that meet the given
'* condition excluding any duplicates found in the
'* duplicate_range.
'*
'* Usage:
'*
'* =COUNTUIF(duplicate_range, criteria, criteria_range)
'*
'****************************************************************
Public Function COUNTUIF(CountRange As range, Criteria As Variant, CriteriaRange As range)
    
    COUNTUIF = COUNTUIFS(CountRange, CriteriaRange, Criteria)
    
End Function

'****************************************************************
'*
'* COUNTUIFS
'*
'* Counts the number of cells specified by a given set of
'* conditions or criteria excluding any duplicates found in the
'* duplicate_range.
'*
'* Usage:
'*
'* =COUNTUIFS(duplicate_range, criteria_range1, criteria1, …)
'*
'****************************************************************
Public Function COUNTUIFS(CountRange As range, ParamArray conditions() As Variant)

    If (UBound(conditions) + 1) Mod 2 <> 0 Then
    
        COUNTUIFS = CVErr(xlErrName)
    
        Exit Function
        
    End If
    
    Dim Values As Dictionary
    
    Dim i As Integer
    Dim passed As Boolean
    
    Dim CriteriaRange As range
    
    Dim Container() As Variant
    ReDim Container(UBound(conditions) / 2)
    
    For i = 0 To UBound(conditions) Step 2
    
        Set CriteriaRange = conditions(i)
    
        Container(i / 2) = CriteriaRange.Value
    
    Next i
    
    Set CriteriaRange = Nothing
    
    Dim CellRangeContainer() As Variant
    CellRangeContainer() = CountRange.Value

    Set Values = New Dictionary

    Dim x As Long

    For x = 1 To UBound(CellRangeContainer())
    
        If Not Values.Exists(CellRangeContainer(x, 1)) Then
        
            passed = True
                  
            For i = 0 To UBound(conditions) Step 2
                
                If IsError(Container(i / 2)(x, 1)) Then
                
                    passed = False
                    
                    Exit For
                    
                ElseIf Not compare(Container(i / 2)(x, 1), conditions(i + 1)) Then
                    
                    passed = False
                    
                    Exit For
                                    
                End If
            
            Next i
            
            If passed Then
            
                Values.Add CellRangeContainer(x, 1), 0
            
            End If
        
        End If
        
    Next x
    
    COUNTUIFS = Values.Count

    Set Values = Nothing
    
End Function

'****************************************************************
'*
'* Compare
'*
'* Enables the use of comparison operators for both strings and
'* numbers.
'*
'****************************************************************
Private Function compare(a As Variant, b As Variant) As Boolean

    Dim result As Boolean
    result = False
    
    If Not IsNumeric(a) Or a = Empty Then
        If Left(b, 2) = "<>" Then
            result = a <> Mid(b, 3)
        Else
            result = a Like b
        End If
    Else
        If IsNumeric(b) Then
                result = a = b
        Else
            If Left(b, 2) = ">=" Then
                result = a >= CDbl(Mid(b, 3))
            ElseIf Left(b, 2) = "<=" Then
                result = a <= CDbl(Mid(b, 3))
            ElseIf Left(b, 2) = "<>" Then
                result = a <> CDbl(Mid(b, 3))
            ElseIf Left(b, 1) = "<" Then
                result = a < CDbl(Mid(b, 2))
            ElseIf Left(b, 1) = ">" Then
                result = a > CDbl(Mid(b, 2))
            Else
                result = False
            End If
        End If
    End If
    
    compare = result

End Function

'****************************************************************
'*
'* TrimRangeToTarget
'*
'* Restricts the height and width of the targetRange to the height
'* and width of the containerRange and returns a resized range.
'*
'****************************************************************
Private Function TrimRangeToTarget(ByVal targetRange As range, ByVal containerRange As range) As range

    Dim tx0, ty0, tx1, ty1 As Long
    Dim ux1, uy1 As Long
    Dim x1, y1 As Long

    tx0 = targetRange.Column
    ty0 = targetRange.Row
    tx1 = targetRange.Columns.Count + tx0 - 1
    ty1 = targetRange.Rows.Count + ty0 - 1

    ux1 = containerRange.Columns.Count + containerRange.Column - 1
    uy1 = containerRange.Rows.Count + containerRange.Row - 1

    If tx1 > ux1 Then x1 = ux1 Else x1 = tx1
    If ty1 > uy1 Then y1 = uy1 Else y1 = ty1
      
    Set TrimRangeToTarget = ActiveSheet.range(Cells(ty0, tx0), Cells(y1, x1))

End Function

