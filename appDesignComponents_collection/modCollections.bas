Attribute VB_Name = "modCollections"
'/'=====. about .===================================================,
'|       '-----'                                                    |
'|     > appDesingComponents Collection                             |
'|         ¤ v: 1.0                                                 |
'|         ¤ price: free for any kind of use                        |
'|     ------------------------------------------------------------ |
'|     > author(s):                                                 |
'|         ¤ ivan stimac, croatia                                   |
'|           mail: ivan.stimac@po.htnet.hr or flashboy01@gmail.com  |
'|         ¤                                                        |
'|           mail:                                                  |
'|         ¤                                                        |
'|           mail:                                                  |
'|     ------------------------------------------------------------ |
'|     > thanks:                                                    |
'|         ¤ Ariad Software - ascPaintEffects class, modFile        |
'|         ¤ Mark Gordon - power resize                             |
'|         ¤                                                        |
'|         ¤                                                        |
'|         ¤                                                        |
'|         ¤                                                        |
'|     ------------------------------------------------------------ |
'|     > please:                                                    |
'|         ¤ rate it                                                |
'|         ¤ report me bugs                                         |
'|         ¤ add your components there and share them with us       |
'|              if you do that, add also your name as author        |
'|                                                                 /
'|                                                                /
'|_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _,


Public Function replaceData(ByRef mColl As Collection, ByVal dataIndex As Integer, newData As Variant)
    Dim i As Integer
    Dim tmpColl As New Collection
    
    For i = 1 To mColl.Count
        If i <> dataIndex Then
            tmpColl.Add mColl.Item(i)
        Else
            tmpColl.Add newData
        End If
    Next i
    
    Do While mColl.Count > 0
        mColl.Remove 1
    Loop
    
    For i = 1 To tmpColl.Count
        mColl.Add tmpColl.Item(i)
    Next i
End Function


Public Function isInList(ByRef mColl As Collection, ByVal Data As Variant) As Boolean
    Dim i As Integer
    isInList = False
    If mColl.Count <= 0 Then Exit Function
    For i = 1 To mColl.Count
        If mColl.Item(i) = Data Then
            isInList = True
            Exit For
        End If
    Next i
End Function

Public Function getDataIndex(ByRef mColl As Collection, ByVal Data As Variant) As Integer
    Dim i As Integer
    getDataIndex = -1
    For i = 1 To mColl.Count
        If mColl.Item(i) = Data Then
            getDataIndex = i
            Exit For
        End If
    Next i
End Function

Public Function clearColl(ByRef mColl As Collection)
    If mColl.Count <= 0 Then Exit Function
    Do While mColl.Count > 0
        mColl.Remove 1
    Loop
End Function
