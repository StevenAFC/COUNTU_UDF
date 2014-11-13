Attribute VB_Name = "COUNTU_UDF"
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
'* Version: 1.0
'* Created: 12-11-2014
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
Public Function COUNTU(ParamArray conditions() As Variant)

    Dim Values As Dictionary
    Dim target As Variant
    
    Dim TargetContainer() As Variant
    
    Set Values = New Dictionary
    
    For Each target In conditions
    
        TargetContainer = target.Value

        Dim i As Long
        
        For i = 1 To UBound(TargetContainer)
        
            If Not Values.Exists(TargetContainer(i, 1)) Then
            
                If Not IsError(TargetContainer(i, 1)) Then
    
                    If TargetContainer(i, 1) <> "" Then
                
                        Values.Add TargetContainer(i, 1), 1
                    
                    End If
                
                End If
            
            End If
            
        Next i
    
    Next target
    
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
Public Function COUNTUIF(CountRange As Range, Criteria As Variant, CriteriaRange As Range)
    
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
Public Function COUNTUIFS(CountRange As Range, ParamArray conditions() As Variant)

    If (UBound(conditions) + 1) Mod 2 <> 0 Then
    
        COUNTUIFS = CVErr(xlErrName)
    
        Exit Function
        
    End If
    
    Dim Values As Dictionary
    
    Dim i As Integer
    Dim passed As Boolean
    
    Dim CriteriaRange As Range
    
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
                    
                ElseIf Not Container(i / 2)(x, 1) = conditions(i + 1) Then
    
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
