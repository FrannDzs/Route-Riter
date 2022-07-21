Attribute VB_Name = "TileNames"
Private wt_ew_min
Private wt_ew_max
Private wt_ns_min
Private wt_ns_max
Private TileFileName As String, Direction() As Integer
Private Prefix As String, quad As Integer

Public Function TileName(wt_ew, wt_ns) As String
ReDim Direction(2, 2) As Integer

TileFileName = vbNullString
'world tile max and min
wt_ew_min = -16384
wt_ew_max = 16384
wt_ns_min = -16384
wt_ns_max = 16384

'for convience
ne = 1
se = 2
SW = 3
NW = 0

Direction(0, 0) = SW
Direction(0, 1) = NW
Direction(1, 1) = ne
Direction(1, 0) = se

TileFileName = vbNullString   'null string
Prefix = "-"

While update_position(wt_ew, wt_ns) >= 0
    If Prefix = "-" Then
            Prefix = "_"
        If quad = ne Then
                append_char = "4"
        ElseIf quad = se Then
                append_char = "8"
        ElseIf quad = SW Then
                append_char = "C"
        Else
                'quad = NW Then
                append_char = "0"
        End If
    Else
        Prefix = "-"
        If quad = ne Then
                add_value = 1
        ElseIf quad = se Then
                add_value = 2
        ElseIf quad = SW Then
                add_value = 3
        Else
                'quad = NW
                add_value = 0
        End If
        TileFileName = TileFileName + Hex(CDec("&H0" + append_char) + add_value)
    End If
Wend    'end while (update_position>0)
'display tilename
frmUtils.Text3(2) = Prefix + LCase(TileFileName)
End Function


Public Function TileName2(wt_ew, wt_ns, strTName As String) As String
ReDim Direction(2, 2) As Integer

TileFileName = vbNullString
'world tile max and min
wt_ew_min = -16384
wt_ew_max = 16384
wt_ns_min = -16384
wt_ns_max = 16384

'for convience
ne = 1
se = 2
SW = 3
NW = 0

Direction(0, 0) = SW
Direction(0, 1) = NW
Direction(1, 1) = ne
Direction(1, 0) = se

TileFileName = vbNullString   'null string
Prefix = "-"

While update_position(wt_ew, wt_ns) >= 0
    If Prefix = "-" Then
            Prefix = "_"
        If quad = ne Then
                append_char = "4"
        ElseIf quad = se Then
                append_char = "8"
        ElseIf quad = SW Then
                append_char = "C"
        Else
                'quad = NW Then
                append_char = "0"
        End If
    Else
        Prefix = "-"
        If quad = ne Then
                add_value = 1
        ElseIf quad = se Then
                add_value = 2
        ElseIf quad = SW Then
                add_value = 3
        Else
                'quad = NW
                add_value = 0
        End If
        TileFileName = TileFileName + Hex(CDec("&H0" + append_char) + add_value)
    End If
Wend    'end while (update_position>0)
'display tilename
strTName = Prefix + LCase(TileFileName)
End Function
Function update_position(wt_ew_tgt, wt_ns_tgt)
wt_ew_avg = (wt_ew_max + wt_ew_min) / 2
If wt_ew_tgt >= (wt_ew_max + wt_ew_min) / 2 Then
    'very left side
    idx1 = 1
    wt_ew_min = wt_ew_avg
Else
    idx1 = 0
    wt_ew_max = wt_ew_avg
End If

wt_ns_avg = Int((wt_ns_max + wt_ns_min) / 2)
If wt_ns_tgt >= wt_ns_avg Then
    idx2 = 1
    wt_ns_min = wt_ns_avg
Else
    idx2 = 0
    wt_ns_max = wt_ns_avg
End If

If Len(TileFileName) < 8 Then
    'crude!!!!
    update_position = Direction(idx1, idx2)
Else
    update_position = -1
End If
quad = update_position
End Function
