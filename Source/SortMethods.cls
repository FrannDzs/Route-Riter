VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function BubbleSort(SortArray As Variant, ArrDim As Variant, _
                            SortType As Variant, Optional SortCol As Variant) As Variant
    
    Dim lLoop1 As Variant
    Dim lLoop2 As Variant
    Dim lLoop3 As Variant
    Dim lTemp As Variant

If ArrDim = 1 Then
    If SortType = "ASC" Then
        For lLoop1 = UBound(SortArray) To LBound(SortArray) Step -1
            For lLoop2 = LBound(SortArray) + 1 To lLoop1
                If SortArray(lLoop2 - 1) > SortArray(lLoop2) Then
                    lTemp = SortArray(lLoop2 - 1)
                    SortArray(lLoop2 - 1) = SortArray(lLoop2)
                    SortArray(lLoop2) = lTemp
                End If
            Next lLoop2
        Next lLoop1
    ElseIf SortType = "DSC" Then
        For lLoop1 = UBound(SortArray) To LBound(SortArray) Step -1
            For lLoop2 = LBound(SortArray) + 1 To lLoop1
                If SortArray(lLoop2 - 1) < SortArray(lLoop2) Then
                    lTemp = SortArray(lLoop2 - 1)
                    SortArray(lLoop2 - 1) = SortArray(lLoop2)
                    SortArray(lLoop2) = lTemp
                End If
            Next lLoop2
        Next lLoop1
    Else
        Exit Function
    End If
ElseIf ArrDim = 2 Then
    If SortType = "ASC" Then
        For lLoop1 = UBound(SortArray) To LBound(SortArray) Step -1
            For lLoop2 = LBound(SortArray) + 1 To lLoop1
                If SortArray(lLoop2 - 1, SortCol) > SortArray(lLoop2, SortCol) Then
                    For lLoop3 = 0 To UBound(SortArray, 2)
                        lTemp = SortArray(lLoop2 - 1, lLoop3)
                        SortArray(lLoop2 - 1, lLoop3) = SortArray(lLoop2, lLoop3)
                        SortArray(lLoop2, lLoop3) = lTemp
                    Next
                End If
            Next lLoop2
        Next lLoop1
    ElseIf SortType = "DSC" Then
        For lLoop1 = UBound(SortArray) To LBound(SortArray) Step -1
            For lLoop2 = LBound(SortArray) + 1 To lLoop1
                If SortArray(lLoop2 - 1, SortCol) < SortArray(lLoop2, SortCol) Then
                    For lLoop3 = 0 To UBound(SortArray, 2)
                        lTemp = SortArray(lLoop2 - 1, lLoop3)
                        SortArray(lLoop2 - 1, lLoop3) = SortArray(lLoop2, lLoop3)
                        SortArray(lLoop2, lLoop3) = lTemp
                    Next
                End If
            Next lLoop2
        Next lLoop1
    Else
        Exit Function
    End If
End If
       
    BubbleSort = SortArray
    
End Function

Public Function SelectionSort(SortArray As Variant, ArrDim As Variant, _
                            SortType As Variant, Optional SortCol As Variant) As Variant

    
Dim lLoop1 As Variant
Dim lLoop2 As Variant
Dim lLoop3 As Variant
Dim lMin As Variant
Dim lTemp As Variant
Dim lMax As Variant

If ArrDim = 1 Then
    If SortType = "ASC" Then
        For lLoop1 = LBound(SortArray) To UBound(SortArray) - 1
            lMin = lLoop1
            For lLoop2 = lLoop1 + 1 To UBound(SortArray)
                If SortArray(lLoop2) < SortArray(lMin) Then
                    lMin = lLoop2
                End If
            Next lLoop2
            lTemp = SortArray(lMin)
            SortArray(lMin) = SortArray(lLoop1)
            SortArray(lLoop1) = lTemp
        Next lLoop1
    ElseIf SortType = "DSC" Then
        For lLoop1 = LBound(SortArray) To UBound(SortArray) - 1
            lMax = lLoop1
            For lLoop2 = lLoop1 + 1 To UBound(SortArray)
                If SortArray(lLoop2) > SortArray(lMax) Then
                    lMax = lLoop2
                End If
            Next lLoop2
            lTemp = SortArray(lMax)
            SortArray(lMax) = SortArray(lLoop1)
            SortArray(lLoop1) = lTemp
        Next lLoop1
    Else
        Exit Function
    End If
ElseIf ArrDim = 2 Then
    If SortType = "ASC" Then
        For lLoop1 = LBound(SortArray) To UBound(SortArray) - 1
            lMin = lLoop1
            For lLoop2 = lLoop1 + 1 To UBound(SortArray)
                If SortArray(lLoop2, SortCol) < SortArray(lMin, SortCol) Then
                    lMin = lLoop2
                End If
            Next lLoop2
            For lLoop3 = 0 To UBound(SortArray, 2)
                lTemp = SortArray(lMin, lLoop3)
                SortArray(lMin, lLoop3) = SortArray(lLoop1, lLoop3)
                SortArray(lLoop1, lLoop3) = lTemp
            Next
        Next lLoop1
    ElseIf SortType = "DSC" Then
        For lLoop1 = LBound(SortArray) To UBound(SortArray) - 1
            lMax = lLoop1
            For lLoop2 = lLoop1 + 1 To UBound(SortArray)
                If SortArray(lLoop2, SortCol) > SortArray(lMax, SortCol) Then
                    lMax = lLoop2
                End If
            Next lLoop2
            For lLoop3 = 0 To UBound(SortArray, 2)
                lTemp = SortArray(lMax, lLoop3)
                SortArray(lMax, lLoop3) = SortArray(lLoop1, lLoop3)
                SortArray(lLoop1, lLoop3) = lTemp
            Next
        Next lLoop1
    Else
        Exit Function
    End If
End If

    SelectionSort = SortArray
    
End Function

Public Function QuickSort(SortArray As Variant, ArrDim As Variant, _
                            SortType As Variant, Optional SortCol As Variant) As Variant

    Dim Leftpoint, Rightpoint As Double
    Leftpoint = LBound(SortArray)
    Rightpoint = UBound(SortArray)
    Call Quick(Leftpoint, Rightpoint, SortArray, ArrDim, SortType, SortCol)
    QuickSort = SortArray
End Function

'Private Sub Quick(Leftpoint As Variant, Rightpoint As Variant, SortArray As Variant)
Private Function Quick(Leftpoint As Variant, Rightpoint As Variant, _
                        SortArray As Variant, ArrDim As Variant, _
                        SortType As Variant, Optional SortCol As Variant) As Variant

    Dim lLoop1 As Variant
    Dim Passedright, Passedleft As Double
    Dim Ref As Boolean
    Dim lTemp As Variant
    Passedleft = Leftpoint
    Passedright = Rightpoint
    Ref = False

If ArrDim = 1 Then
    If SortType = "ASC" Then
        Do Until Leftpoint = Rightpoint
            If SortArray(Rightpoint) < SortArray(Leftpoint) Then
                lTemp = SortArray(Rightpoint)
                SortArray(Rightpoint) = SortArray(Leftpoint)
                SortArray(Leftpoint) = lTemp
                If Ref = False Then
                    Ref = True
                Else
                    Ref = False
                End If
            End If
            If Ref = False Then
                Rightpoint = Rightpoint - 1
            Else
                Leftpoint = Leftpoint + 1
            End If
        Loop
        
        If Leftpoint - Passedleft > 1 Then
            Call Quick(Passedleft, Leftpoint - 1, SortArray, ArrDim, SortType, SortCol)
        End If
    
        If Passedright - Rightpoint > 1 Then
            Call Quick(Leftpoint + 1, Passedright, SortArray, ArrDim, SortType, SortCol)
        End If
        
    ElseIf SortType = "DSC" Then
    
        Do Until Leftpoint = Rightpoint
            If SortArray(Rightpoint) > SortArray(Leftpoint) Then
                lTemp = SortArray(Rightpoint)
                SortArray(Rightpoint) = SortArray(Leftpoint)
                SortArray(Leftpoint) = lTemp
                If Ref = False Then
                    Ref = True
                Else
                    Ref = False
                End If
            End If
            If Ref = False Then
                Rightpoint = Rightpoint - 1
            Else
                Leftpoint = Leftpoint + 1
            End If
        Loop
        
        If Leftpoint - Passedleft > 1 Then
            Call Quick(Passedleft, Leftpoint - 1, SortArray, ArrDim, SortType, SortCol)
        End If
    
        If Passedright - Rightpoint > 1 Then
            Call Quick(Leftpoint + 1, Passedright, SortArray, ArrDim, SortType, SortCol)
        End If
    
    Else
        Exit Function
    End If
ElseIf ArrDim = 2 Then
    If SortType = "ASC" Then
        Do Until Leftpoint = Rightpoint
            If SortArray(Rightpoint, SortCol) < SortArray(Leftpoint, SortCol) Then
                
                For lLoop1 = 0 To UBound(SortArray, 2)
                    lTemp = SortArray(Rightpoint, lLoop1)
                    SortArray(Rightpoint, lLoop1) = SortArray(Leftpoint, lLoop1)
                    SortArray(Leftpoint, lLoop1) = lTemp
                Next
                If Ref = False Then
                    Ref = True
                Else
                    Ref = False
                End If
            End If
            If Ref = False Then
                Rightpoint = Rightpoint - 1
            Else
                Leftpoint = Leftpoint + 1
            End If
        Loop
        
        If Leftpoint - Passedleft > 1 Then
            Call Quick(Passedleft, Leftpoint - 1, SortArray, ArrDim, SortType, SortCol)
        End If
    
        If Passedright - Rightpoint > 1 Then
            Call Quick(Leftpoint + 1, Passedright, SortArray, ArrDim, SortType, SortCol)
        End If
    
    ElseIf SortType = "DSC" Then
    
        Do Until Leftpoint = Rightpoint
            If SortArray(Rightpoint, SortCol) > SortArray(Leftpoint, SortCol) Then
                
                For lLoop1 = 0 To UBound(SortArray, 2)
                    lTemp = SortArray(Rightpoint, lLoop1)
                    SortArray(Rightpoint, lLoop1) = SortArray(Leftpoint, lLoop1)
                    SortArray(Leftpoint, lLoop1) = lTemp
                Next
                If Ref = False Then
                    Ref = True
                Else
                    Ref = False
                End If
            End If
            If Ref = False Then
                Rightpoint = Rightpoint - 1
            Else
                Leftpoint = Leftpoint + 1
            End If
        Loop
        
        If Leftpoint - Passedleft > 1 Then
            Call Quick(Passedleft, Leftpoint - 1, SortArray, ArrDim, SortType, SortCol)
        End If
    
        If Passedright - Rightpoint > 1 Then
            Call Quick(Leftpoint + 1, Passedright, SortArray, ArrDim, SortType, SortCol)
        End If
    
    Else
        Exit Function
    End If
End If

End Function

