Option Explicit

' ============================================================
' 主入口 - 选择处理类型
' ============================================================
Sub 处理抵销分录()
    Dim choice As Integer
    choice = MsgBox("请选择处理类型：" & vbCrLf & vbCrLf & _
                    "是(Yes)  = 内部往来未抵销分录" & vbCrLf & _
                    "否(No)   = 内部交易未抵销分录", _
                    vbYesNoCancel + vbQuestion, "选择处理类型")
    If choice = vbYes Then
        Call 处理内部往来未抵销分录
    ElseIf choice = vbNo Then
        Call 处理内部交易未抵销分录
    End If
End Sub

' ============================================================
' 内部往来未抵销分录 - 主流程
' ============================================================
Sub 处理内部往来未抵销分录()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 确保K列有标题
    If ws.Cells(1, 11).Value = "" Then ws.Cells(1, 11).Value = "备注"

    Call 步骤1_借正贷负(ws)
    Call 步骤2_删除零值行(ws)
    Call 步骤345_处理所有组(ws, True)
    Call 步骤6_重新排序(ws)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "内部往来未抵销分录处理完成！", vbInformation
End Sub

' ============================================================
' 内部交易未抵销分录 - 主流程
' ============================================================
Sub 处理内部交易未抵销分录()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If ws.Cells(1, 11).Value = "" Then ws.Cells(1, 11).Value = "备注"

    Call 步骤1_借正贷负(ws)
    Call 步骤2_删除零值行(ws)
    Call 步骤345_处理所有组(ws, False)
    Call 步骤6_重新排序(ws)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "内部交易未抵销分录处理完成！", vbInformation
End Sub

' ============================================================
' 步骤1: 将G列贷方金额调整为相反数（借正贷负）
' ============================================================
Sub 步骤1_借正贷负(ws As Worksheet)
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim v As Variant
        v = ws.Cells(i, 7).Value
        If IsNumeric(v) And v <> "" Then
            ws.Cells(i, 7).Value = -CDbl(v)
        End If
    Next i
End Sub

' ============================================================
' 步骤2: 删除借贷金额均为零或空的行（小计行除外）
' ============================================================
Sub 步骤2_删除零值行(ws As Worksheet)
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = lastRow To 2 Step -1
        If Trim(ws.Cells(i, 2).Value) = "小计" Then GoTo NextRow
        Dim dv As Variant, cv As Variant
        dv = ws.Cells(i, 6).Value
        cv = ws.Cells(i, 7).Value
        Dim dZero As Boolean, cZero As Boolean
        dZero = (IsEmpty(dv) Or dv = "" Or CDbl(dv & "0") = 0)
        cZero = (IsEmpty(cv) Or cv = "" Or CDbl(cv & "0") = 0)
        If dZero And cZero Then ws.Rows(i).Delete
NextRow:
    Next i
End Sub

' ============================================================
' 步骤3-5: 从下往上找每个"小计"组，生成并合并补充分录
' isWangLai=True → 内部往来, False → 内部交易
' ============================================================
Sub 步骤345_处理所有组(ws As Worksheet, isWangLai As Boolean)
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    i = lastRow
    Do While i >= 2
        If Trim(ws.Cells(i, 2).Value) = "小计" Then
            Dim gStart As Long
            gStart = 找组起始行(ws, i)
            If isWangLai Then
                Call 生成补充分录组_往来(ws, gStart, i)
            Else
                Call 生成补充分录组_交易(ws, gStart, i)
            End If
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        End If
        i = i - 1
    Loop
End Sub

' ============================================================
' 找小计行对应组的起始行（上一个小计行的下一行，或第2行）
' ============================================================
Function 找组起始行(ws As Worksheet, subtotalRow As Long) As Long
    Dim i As Long
    For i = subtotalRow - 1 To 2 Step -1
        If Trim(ws.Cells(i, 2).Value) = "小计" Then
            找组起始行 = i + 1
            Exit Function
        End If
    Next i
    找组起始行 = 2
End Function

' ============================================================
' 内部往来: 为一个组生成补充分录（步骤3-5）
' ============================================================
Sub 生成补充分录组_往来(ws As Worksheet, gStart As Long, subtotalRow As Long)
    ' 收集组内不重复的未抵销原因
    Dim reasons() As String, rCount As Integer
    Call 收集原因(ws, gStart, subtotalRow - 1, reasons, rCount)
    If rCount = 0 Then Exit Sub

    ' 构建补充分录列表（每条 = Array(B,C,E,金额,I备注)）
    Dim entries() As Variant
    Dim eCount As Integer
    eCount = 0

    Dim k As Integer
    For k = 0 To rCount - 1
        Call 生成往来单条原因(ws, gStart, subtotalRow - 1, reasons(k), entries, eCount)
    Next k

    ' 步骤5：合并相同(B,C,E)且备注为空（有金额）的条目
    Dim merged() As Variant, mCount As Integer
    Call 合并分录(entries, eCount, merged, mCount)

    If mCount = 0 Then Exit Sub

    ' 插入行并填充
    ws.Rows(subtotalRow).Resize(mCount).Insert Shift:=xlShiftDown
    Call 填充补充分录行(ws, subtotalRow, merged, mCount)
End Sub

' ============================================================
' 内部交易: 为一个组生成补充分录（步骤3-5）
' ============================================================
Sub 生成补充分录组_交易(ws As Worksheet, gStart As Long, subtotalRow As Long)
    Dim reasons() As String, rCount As Integer
    Call 收集原因(ws, gStart, subtotalRow - 1, reasons, rCount)
    If rCount = 0 Then Exit Sub

    Dim entries() As Variant, eCount As Integer
    eCount = 0

    Dim k As Integer
    For k = 0 To rCount - 1
        Call 生成交易单条原因(ws, gStart, subtotalRow - 1, reasons(k), entries, eCount)
    Next k

    Dim merged() As Variant, mCount As Integer
    Call 合并分录(entries, eCount, merged, mCount)

    If mCount = 0 Then Exit Sub

    ws.Rows(subtotalRow).Resize(mCount).Insert Shift:=xlShiftDown
    Call 填充补充分录行(ws, subtotalRow, merged, mCount)
End Sub

' ============================================================
' 收集组内不重复的未抵销原因（J列）
' ============================================================
Sub 收集原因(ws As Worksheet, gStart As Long, gEnd As Long, _
             ByRef reasons() As String, ByRef rCount As Integer)
    rCount = 0
    ReDim reasons(0)
    Dim i As Long, r As String, j As Integer, found As Boolean
    For i = gStart To gEnd
        r = Trim(ws.Cells(i, 10).Value)
        If r = "" Then GoTo Skip
        found = False
        For j = 0 To rCount - 1
            If reasons(j) = r Then found = True: Exit For
        Next j
        If Not found Then
            ReDim Preserve reasons(rCount)
            reasons(rCount) = r
            rCount = rCount + 1
        End If
Skip:
    Next i
End Sub

' ============================================================
' 内部往来: 为单个原因生成补充分录条目
' ============================================================
Sub 生成往来单条原因(ws As Worksheet, gStart As Long, gEnd As Long, _
                     reason As String, ByRef entries() As Variant, ByRef eCount As Integer)

    ' 取出该原因的所有行号
    Dim mRows() As Long, mCount As Integer
    Call 找原因行(ws, gStart, gEnd, reason, mRows, mCount)
    If mCount = 0 Then Exit Sub

    Dim 本方 As String, 对方 As String, 新科目 As String, 备注 As String
    Dim amt As Double
    备注 = ""

    Select Case reason

        Case "待结转增值税无法抵销"
            ' 本方=原对方, 对方=原本方
            本方 = ws.Cells(mRows(0), 3).Value
            对方 = ws.Cells(mRows(0), 2).Value
            新科目 = "其他流动负债-合同负债待转销项税金"
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, 新科目, amt, "")

        Case "合同进度确认不一致"
            Dim 本方科目1 As Variant, 对方科目1 As Variant
            本方科目1 = Array("预付款项-常规业务", "应付账款-常规业务", "应付账款-供应链金融", _
                             "应付账款-待结算增值税", "其他应付款-其他常规业务", _
                             "一年内到期的非流动负债-长期应付款-其他", "长期应付款-其他常规业务")
            对方科目1 = Array("应收账款-常规业务", "应收账款-供应链金融", "合同资产-已完工未结算", _
                             "合同资产-质保金", "其他应收款-其他常规业务", _
                             "一年内到期的非流动资产-长期应收款-其他", "长期应收款-其他常规业务", _
                             "其他非流动资产-合同资产质保金", "合同负债-其他业务", _
                             "合同负债-预收工程款", "合同负债-已结算未完工")
            Call 处理分科目生成(ws, mRows, mCount, 本方科目1, 对方科目1, "已完工未结算", entries, eCount)

        Case "商品销售结算不一致"
            Dim 本方科目2 As Variant, 对方科目2 As Variant
            本方科目2 = Array("预付款项-常规业务", "应付账款-常规业务", "应付账款-供应链金融", _
                             "应付账款-待结算增值税", "其他应付款-其他常规业务", _
                             "一年内到期的非流动负债-长期应付款-其他", "长期应付款-其他常规业务")
            对方科目2 = Array("应收账款-常规业务", "应收账款-供应链金融", "其他应收款-其他常规业务", _
                             "合同负债-其他业务", "合同负债-预收工程款", "合同负债-已结算未完工")
            Call 处理分科目生成(ws, mRows, mCount, 本方科目2, 对方科目2, "已完工未结算", entries, eCount)

        Case "票据背书外部单位"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "应付票据-第三方", amt, "")

        Case "票据背书内部单位"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "应付票据-内部单位", amt, "")

        Case "票据贴现"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "短期借款-第三方", amt, "")

        Case "内部应收保理至外部非银行保理商"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "其他应付款-借款", amt, "")

        Case "内部应收保理至外部银行"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "短期借款-第三方", amt, "")

        Case "内部应收保理至内部单位"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "应付账款-内部单位", amt, "")

        Case "其他"
            ' 仅填充原因，不计算金额
            Call 添加分录(entries, eCount, reason, "", "", 0, "补充分录-其他")

        Case Else
            Call 添加分录(entries, eCount, reason, "", "", 0, "补充分录-超出说明文档列示原因或科目")

    End Select
End Sub

' ============================================================
' 内部交易: 为单个原因生成补充分录条目
' ============================================================
Sub 生成交易单条原因(ws As Worksheet, gStart As Long, gEnd As Long, _
                     reason As String, ByRef entries() As Variant, ByRef eCount As Integer)

    Dim mRows() As Long, mCount As Integer
    Call 找原因行(ws, gStart, gEnd, reason, mRows, mCount)
    If mCount = 0 Then Exit Sub

    Dim 本方 As String, 对方 As String, amt As Double

    Select Case reason

        Case "合同进度确认不一致"
            Dim 本方科目T As Variant, 对方科目T As Variant
            本方科目T = Array("主营业务成本-工程总分包", "主营业务成本-其他业务")
            对方科目T = Array("主营业务收入-工程总分包", "主营业务收入-其他业务", "主营业务收入-销售商品收入")
            Call 处理分科目生成(ws, mRows, mCount, 本方科目T, 对方科目T, "主营业务收入", entries, eCount)

        Case "PPP项目公司收入成本不一致"
            本方 = ws.Cells(mRows(0), 2).Value
            对方 = ws.Cells(mRows(0), 3).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "主营业务收入", amt, "")

        Case "利息收入/支出不一致"
            ' 本方=原对方, 对方=原本方
            本方 = ws.Cells(mRows(0), 3).Value
            对方 = ws.Cells(mRows(0), 2).Value
            amt = 计算补充金额(ws, mRows, mCount)
            Call 添加分录(entries, eCount, 本方, 对方, "财务费用-借款利息支出", amt, "")

        Case "其他"
            Call 添加分录(entries, eCount, reason, "", "", 0, "补充分录-其他")

        Case Else
            Call 添加分录(entries, eCount, reason, "", "", 0, "补充分录-超出说明文档列示原因或科目")

    End Select
End Sub

' ============================================================
' 按科目列表判断本方/对方，生成补充分录
' 支持同原因下多个(本方,对方)对 → 分别生成（步骤4第11点）
' ============================================================
Sub 处理分科目生成(ws As Worksheet, mRows() As Long, mCount As Integer, _
                   本方科目 As Variant, 对方科目 As Variant, 新科目 As String, _
                   ByRef entries() As Variant, ByRef eCount As Integer)

    ' 收集所有行对应的(本方,对方)映射
    Dim pairs() As String  ' "本方|对方"
    Dim pCount As Integer
    pCount = 0
    ReDim pairs(0)

    Dim pairBen() As String, pairDui() As String
    ReDim pairBen(0), pairDui(0)

    Dim i As Integer, j As Integer
    Dim rowNum As Long, subj As String
    Dim ben As String, dui As String, pkey As String, found As Boolean

    For i = 0 To mCount - 1
        rowNum = mRows(i)
        subj = Trim(ws.Cells(rowNum, 5).Value)
        If 在数组中(subj, 本方科目) Then
            ben = ws.Cells(rowNum, 2).Value
            dui = ws.Cells(rowNum, 3).Value
        ElseIf 在数组中(subj, 对方科目) Then
            ben = ws.Cells(rowNum, 3).Value
            dui = ws.Cells(rowNum, 2).Value
        Else
            ' 科目超出范围 → 直接添加占位条目
            Call 添加分录(entries, eCount, ws.Cells(rowNum, 2).Value, _
                         ws.Cells(rowNum, 3).Value, subj, 0, _
                         "补充分录-超出说明文档列示原因或科目")
            GoTo NextPairRow
        End If

        pkey = ben & "|" & dui
        found = False
        For j = 0 To pCount - 1
            If pairs(j) = pkey Then found = True: Exit For
        Next j
        If Not found Then
            ReDim Preserve pairs(pCount)
            ReDim Preserve pairBen(pCount)
            ReDim Preserve pairDui(pCount)
            pairs(pCount) = pkey
            pairBen(pCount) = ben
            pairDui(pCount) = dui
            pCount = pCount + 1
        End If
NextPairRow:
    Next i

    ' 对每个(本方,对方)对：若多对则分别生成，若单对则合并（步骤4第11点）
    For j = 0 To pCount - 1
        Dim subRows() As Long, sCount As Integer
        sCount = 0
        ReDim subRows(0)

        For i = 0 To mCount - 1
            rowNum = mRows(i)
            subj = Trim(ws.Cells(rowNum, 5).Value)
            Dim cb As String, cd As String
            If 在数组中(subj, 本方科目) Then
                cb = ws.Cells(rowNum, 2).Value
                cd = ws.Cells(rowNum, 3).Value
            ElseIf 在数组中(subj, 对方科目) Then
                cb = ws.Cells(rowNum, 3).Value
                cd = ws.Cells(rowNum, 2).Value
            Else
                GoTo NextSubRow
            End If
            If cb = pairBen(j) And cd = pairDui(j) Then
                ReDim Preserve subRows(sCount)
                subRows(sCount) = rowNum
                sCount = sCount + 1
            End If
NextSubRow:
        Next i

        If sCount > 0 Then
            Dim subAmt As Double
            subAmt = 计算补充金额(ws, subRows, sCount)
            Call 添加分录(entries, eCount, pairBen(j), pairDui(j), 新科目, subAmt, "")
        End If
    Next j
End Sub

' ============================================================
' 计算补充金额：-(sum(F借方) + sum(G贷方_after_step1))
' G已在步骤1取反，所以 G列当前值 = -原始贷方
' 补充金额 = -(sumDebit + sumCreditNegated)
' ============================================================
Function 计算补充金额(ws As Worksheet, mRows() As Long, mCount As Integer) As Double
    Dim i As Integer, sumD As Double, sumC As Double
    sumD = 0: sumC = 0
    For i = 0 To mCount - 1
        Dim dv As Variant, cv As Variant
        dv = ws.Cells(mRows(i), 6).Value
        cv = ws.Cells(mRows(i), 7).Value
        If IsNumeric(dv) And dv <> "" Then sumD = sumD + CDbl(dv)
        If IsNumeric(cv) And cv <> "" Then sumC = sumC + CDbl(cv)
    Next i
    计算补充金额 = -(sumD + sumC)
End Function

' ============================================================
' 找某原因在组内的所有行号
' ============================================================
Sub 找原因行(ws As Worksheet, gStart As Long, gEnd As Long, reason As String, _
             ByRef mRows() As Long, ByRef mCount As Integer)
    mCount = 0
    ReDim mRows(0)
    Dim i As Long
    For i = gStart To gEnd
        If Trim(ws.Cells(i, 10).Value) = reason Then
            ReDim Preserve mRows(mCount)
            mRows(mCount) = i
            mCount = mCount + 1
        End If
    Next i
End Sub

' ============================================================
' 添加一条补充分录到entries数组
' entries(n) = Array(B列, C列, E列, 金额, I列备注)
' ============================================================
Sub 添加分录(ByRef entries() As Variant, ByRef eCount As Integer, _
             colB As String, colC As String, colE As String, _
             金额 As Double, colI As String)
    ReDim Preserve entries(eCount)
    entries(eCount) = Array(colB, colC, colE, 金额, colI)
    eCount = eCount + 1
End Sub

' ============================================================
' 步骤5: 合并相同(B,C,E)且无备注的分录（对有备注的去重）
' ============================================================
Sub 合并分录(entries() As Variant, eCount As Integer, _
             ByRef merged() As Variant, ByRef mCount As Integer)
    mCount = 0
    If eCount = 0 Then Exit Sub
    ReDim merged(0)

    Dim keys() As String
    Dim kCount As Integer
    kCount = 0
    ReDim keys(0)

    Dim i As Integer, j As Integer
    Dim ea As Variant, key As String, found As Boolean

    ' 建立唯一key列表：有金额的(或备注为空的) → B|C|E（可合并）
    '                   有备注无金额的         → B|C|E|REM|备注（去重保唯一）
    For i = 0 To eCount - 1
        ea = entries(i)
        Dim hasAmt As Boolean
        hasAmt = (CDbl(ea(3)) <> 0 Or CStr(ea(4)) = "")
        If hasAmt Then
            key = CStr(ea(0)) & "|" & CStr(ea(1)) & "|" & CStr(ea(2))
        Else
            key = CStr(ea(0)) & "|" & CStr(ea(1)) & "|" & CStr(ea(2)) & "|REM|" & CStr(ea(4))
        End If

        found = False
        For j = 0 To kCount - 1
            If keys(j) = key Then found = True: Exit For
        Next j
        If Not found Then
            ReDim Preserve keys(kCount)
            keys(kCount) = key
            kCount = kCount + 1
        End If
    Next i

    ' 对每个key汇总金额
    For i = 0 To kCount - 1
        key = keys(i)
        Dim totalAmt As Double
        totalAmt = 0
        Dim b As String, c As String, e As String, rem As String
        b = "": c = "": e = "": rem = ""
        Dim firstFound As Boolean
        firstFound = False

        For j = 0 To eCount - 1
            ea = entries(j)
            Dim haA As Boolean
            haA = (CDbl(ea(3)) <> 0 Or CStr(ea(4)) = "")
            Dim ek As String
            If haA Then
                ek = CStr(ea(0)) & "|" & CStr(ea(1)) & "|" & CStr(ea(2))
            Else
                ek = CStr(ea(0)) & "|" & CStr(ea(1)) & "|" & CStr(ea(2)) & "|REM|" & CStr(ea(4))
            End If

            If ek = key Then
                If Not firstFound Then
                    b = CStr(ea(0)): c = CStr(ea(1)): e = CStr(ea(2)): rem = CStr(ea(4))
                    firstFound = True
                End If
                totalAmt = totalAmt + CDbl(ea(3))
            End If
        Next j

        If firstFound Then
            ReDim Preserve merged(mCount)
            merged(mCount) = Array(b, c, e, totalAmt, rem)
            mCount = mCount + 1
        End If
    Next i
End Sub

' ============================================================
' 将补充分录填充到已插入的空行中
' ============================================================
Sub 填充补充分录行(ws As Worksheet, startRow As Long, merged() As Variant, mCount As Integer)
    Dim i As Integer
    For i = 0 To mCount - 1
        Dim r As Long
        r = startRow + i
        Dim ea As Variant
        ea = merged(i)

        Dim amt As Double
        amt = CDbl(ea(3))

        ws.Cells(r, 2).Value = ea(0)     ' 本方单位 或 未抵销原因（对其他类）
        ws.Cells(r, 3).Value = ea(1)     ' 对方单位
        ws.Cells(r, 5).Value = ea(2)     ' 科目

        ' 借正贷负放置金额
        If amt > 0 Then
            ws.Cells(r, 6).Value = amt
            ws.Cells(r, 7).Value = ""
        ElseIf amt < 0 Then
            ws.Cells(r, 6).Value = ""
            ws.Cells(r, 7).Value = amt   ' 负数，与步骤1一致
        Else
            ws.Cells(r, 6).Value = ""
            ws.Cells(r, 7).Value = ""
        End If

        ws.Cells(r, 9).Value = ea(4)     ' I列备注
        ws.Cells(r, 11).Value = "补充分录"  ' K列标注
    Next i
End Sub

' ============================================================
' 步骤6: 重新排序序号（A列）
' 原始数据行（非小计、非补充分录）按组计数，每条完整分录一个序号
' ============================================================
Sub 步骤6_重新排序(ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim seqNum As Long
    seqNum = 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Range("A2:A" & lastRow).ClearContents

    ' 按组编号：每个小计区块内的第一条原始数据行分配序号
    Dim inGroup As Boolean
    inGroup = False

    For i = 2 To lastRow
        Dim bVal As String
        bVal = Trim(ws.Cells(i, 2).Value)
        Dim kVal As String
        kVal = Trim(ws.Cells(i, 11).Value)

        If bVal = "小计" Then
            inGroup = False  ' 遇到小计，重置，下一组重新计数
        ElseIf kVal = "补充分录" Or bVal = "" Then
            ' 补充分录行和空行不编号，也不影响inGroup状态
        Else
            ' 正常原始数据行
            If Not inGroup Then
                ws.Cells(i, 1).Value = seqNum
                seqNum = seqNum + 1
                inGroup = True
            End If
        End If
    Next i
End Sub

' ============================================================
' 工具函数：检查字符串是否在Variant数组中
' ============================================================
Function 在数组中(val As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) = Trim(val) Then
            在数组中 = True
            Exit Function
        End If
    Next i
    在数组中 = False
End Function
