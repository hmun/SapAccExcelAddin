Public Class TSAP_DocData

    Public aHdrRec As TDataRec
    Public aCurRec As TDataRec
    Public aData As TData
    Public aAmounts As Dictionary(Of String, TData)

    ' no static list anymore - field list is now read from function metadata
    'Private Hdr_Fields() As String = {"OBJ_TYPE", "OBJ_KEY", "OBJ_SYS", "BUS_ACT", "USERNAME", "HEADER_TXT", "COMP_CODE", "DOC_DATE", "PSTNG_DATE", "TRANS_DATE", "FISC_YEAR", "FIS_PERIOD", "DOC_TYPE", "REF_DOC_NO", "AC_DOC_NO", "OBJ_KEY_R", "REASON_REV", "COMPO_ACC", "REF_DOC_NO_LONG", "ACC_PRINCIPLE", "NEG_POSTNG", "OBJ_KEY_INV", "BILL_CATEGORY", "VATDATE", "INVOICE_REC_DATE", "ECS_ENV", "PARTIAL_REV", "DOC_STATUS", "TAX_CALC_DATE", "GLO_REF1_HD", "GLO_DAT1_HD", "GLO_REF2_HD", "GLO_DAT2_HD", "GLO_REF3_HD", "GLO_DAT3_HD", "GLO_REF4_HD", "GLO_DAT4_HD", "GLO_REF5_HD", "GLO_DAT5_HD", "GLO_BP1_HD", "GLO_BP2_HD", "EV_POSTNG_CTRL", "LEDGER_GROUP", "PLANNED_REV_DATE"}
    'Private Gla_Fields() As String = {"GL_ACCOUNT", "ITEM_TEXT", "STAT_CON", "LOG_PROC", "AC_DOC_NO", "REF_KEY_1", "REF_KEY_2", "REF_KEY_3", "ACCT_KEY", "ACCT_TYPE", "DOC_TYPE", "COMP_CODE", "BUS_AREA", "FUNC_AREA", "PLANT", "FIS_PERIOD", "FISC_YEAR", "PSTNG_DATE", "VALUE_DATE", "FM_AREA", "CUSTOMER", "CSHDIS_IND", "VENDOR_NO", "ALLOC_NMBR", "TAX_CODE", "TAXJURCODE", "EXT_OBJECT_ID", "BUS_SCENARIO", "COSTOBJECT", "COSTCENTER", "ACTTYPE", "PROFIT_CTR", "PART_PRCTR", "NETWORK", "WBS_ELEMENT", "ORDERID", "ORDER_ITNO", "ROUTING_NO", "ACTIVITY", "COND_TYPE", "COND_COUNT", "COND_ST_NO", "FUND", "FUNDS_CTR", "CMMT_ITEM", "CO_BUSPROC", "ASSET_NO", "SUB_NUMBER", "BILL_TYPE", "SALES_ORD", "S_ORD_ITEM", "DISTR_CHAN", "DIVISION", "SALESORG", "SALES_GRP", "SALES_OFF", "SOLD_TO", "DE_CRE_IND", "P_EL_PRCTR", "XMFRW", "QUANTITY", "BASE_UOM", "BASE_UOM_ISO", "INV_QTY", "INV_QTY_SU", "SALES_UNIT", "SALES_UNIT_ISO", "PO_PR_QNT", "PO_PR_UOM", "PO_PR_UOM_ISO", "ENTRY_QNT", "ENTRY_UOM", "ENTRY_UOM_ISO", "VOLUME", "VOLUMEUNIT", "VOLUMEUNIT_ISO", "GROSS_WT", "NET_WEIGHT", "UNIT_OF_WT", "UNIT_OF_WT_ISO", "ITEM_CAT", "MATERIAL", "MATL_TYPE", "MVT_IND", "REVAL_IND", "ORIG_GROUP", "ORIG_MAT", "SERIAL_NO", "PART_ACCT", "TR_PART_BA", "TRADE_ID", "VAL_AREA", "VAL_TYPE", "ASVAL_DATE", "PO_NUMBER", "PO_ITEM", "ITM_NUMBER", "COND_CATEGORY", "FUNC_AREA_LONG", "CMMT_ITEM_LONG", "GRANT_NBR", "CS_TRANS_T", "MEASURE", "SEGMENT", "PARTNER_SEGMENT", "RES_DOC", "RES_ITEM", "BILLING_PERIOD_START_DATE", "BILLING_PERIOD_END_DATE", "PPA_EX_IND", "FASTPAY", "PARTNER_GRANT_NBR", "BUDGET_PERIOD", "PARTNER_BUDGET_PERIOD", "PARTNER_FUND", "ITEMNO_TAX", "PAYMENT_TYPE", "EXPENSE_TYPE", "PROGRAM_PROFILE", "MATERIAL_LONG", "HOUSEBANKID", "HOUSEBANKACCTID", "PERSON_NO", "ACROBJ_TYPE", "ACROBJ_ID", "ACRSUBOBJ_ID", "ACRITEM_TYPE", "VALOBJTYPE", "VALOBJ_ID", "VALSUBOBJ_ID"}
    'Private Cus_Fields() As String = {"CUSTOMER", "GL_ACCOUNT", "REF_KEY_1", "REF_KEY_2", "REF_KEY_3", "COMP_CODE", "BUS_AREA", "PMNTTRMS", "BLINE_DATE", "DSCT_DAYS1", "DSCT_DAYS2", "NETTERMS", "DSCT_PCT1", "DSCT_PCT2", "PYMT_METH", "PMTMTHSUPL", "PAYMT_REF", "DUNN_KEY", "DUNN_BLOCK", "PMNT_BLOCK", "VAT_REG_NO", "ALLOC_NMBR", "ITEM_TEXT", "PARTNER_BK", "SCBANK_IND", "BUSINESSPLACE", "SECTIONCODE", "BRANCH", "PYMT_CUR", "PYMT_CUR_ISO", "PYMT_AMT", "C_CTR_AREA", "BANK_ID", "SUPCOUNTRY", "SUPCOUNTRY_ISO", "TAX_CODE", "TAXJURCODE", "TAX_DATE", "SP_GL_IND", "PARTNER_GUID", "ALT_PAYEE", "ALT_PAYEE_BANK", "DUNN_AREA", "CASE_GUID", "PROFIT_CTR", "FUND", "GRANT_NBR", "MEASURE", "HOUSEBANKACCTID", "RES_DOC", "RES_ITEM", "FUND_LONG", "DISPUTE_IF_TYPE", "BUDGET_PERIOD", "PAYS_PROV", "PAYS_TRAN", "SEPA_MANDATE_ID", "PART_BUSINESSPLACE", "REP_COUNTRY_EU"}
    'Private Ven_Fields() As String = {"VENDOR_NO", "GL_ACCOUNT", "REF_KEY_1", "REF_KEY_2", "REF_KEY_3", "COMP_CODE", "BUS_AREA", "PMNTTRMS", "BLINE_DATE", "DSCT_DAYS1", "DSCT_DAYS2", "NETTERMS", "DSCT_PCT1", "DSCT_PCT2", "PYMT_METH", "PMTMTHSUPL", "PMNT_BLOCK", "SCBANK_IND", "SUPCOUNTRY", "SUPCOUNTRY_ISO", "BLLSRV_IND", "ALLOC_NMBR", "ITEM_TEXT", "PO_SUB_NO", "PO_CHECKDG", "PO_REF_NO", "W_TAX_CODE", "BUSINESSPLACE", "SECTIONCODE", "INSTR1", "INSTR2", "INSTR3", "INSTR4", "BRANCH", "PYMT_CUR", "PYMT_AMT", "PYMT_CUR_ISO", "SP_GL_IND", "TAX_CODE", "TAX_DATE", "TAXJURCODE", "ALT_PAYEE", "ALT_PAYEE_BANK", "PARTNER_BK", "BANK_ID", "PARTNER_GUID", "PROFIT_CTR", "FUND", "GRANT_NBR", "MEASURE", "HOUSEBANKACCTID", "BUDGET_PERIOD", "PPA_EX_IND", "PART_BUSINESSPLACE", "PAYMT_REF"}
    'Private Tax_Fields() As String = {"GL_ACCOUNT", "COND_KEY", "ACCT_KEY", "TAX_CODE", "TAX_RATE", "TAX_DATE", "TAXJURCODE", "TAXJURCODE_DEEP", "TAXJURCODE_LEVEL", "ITEMNO_TAX", "DIRECT_TAX"}
    '   Private Amt_Fields() As String = {"CURR_TYPE", "CURRENCY", "CURRENCY_ISO", "AMT_DOCCUR", "EXCH_RATE", "EXCH_RATE_V", "AMT_BASE", "DISC_BASE", "DISC_AMT", "TAX_AMT"}
    '   removed currency fields as they are allready in the aCurRec
    'Private Amt_Fields() As String = {"CURR_TYPE", "AMT_DOCCUR", "EXCH_RATE", "EXCH_RATE_V", "AMT_BASE", "DISC_BASE", "DISC_AMT", "TAX_AMT"}
    'Private Cpd_Fields() As String = {"NAME", "NAME_2", "NAME_3", "NAME_4", "POSTL_CODE", "CITY", "COUNTRY", "COUNTRY_ISO", "STREET", "PO_BOX", "POBX_PCD", "POBK_CURAC", "BANK_ACCT", "BANK_NO", "BANK_CTRY", "BANK_CTRY_ISO", "TAX_NO_1", "TAX_NO_2", "TAX", "EQUAL_TAX", "REGION", "CTRL_KEY", "INSTR_KEY", "DME_IND", "LANGU_ISO", "IBAN", "SWIFT_CODE", "TAX_NO_3", "TAX_NO_4", "TITLE", "TAX_NO_5", "GLO_RE1_OT"}

    Private Hdr_Fields() As String = {}
    Private Gla_Fields() As String = {}
    Private Cus_Fields() As String = {}
    Private Ven_Fields() As String = {}
    Private Tax_Fields() As String = {}
    Private Amt_Fields() As String = {}
    Private Cpd_Fields() As String = {}

    Private aAccPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sHd As String = "DOCUMENTHEADER"
    Private Const sCp As String = "CUSTOMERCPD"
    Private Const sGL As String = "ACCOUNTGL"
    Private Const sCu As String = "ACCOUNTRECEIVABLE"
    Private Const sVe As String = "ACCOUNTPAYABLE"
    Private Const sTx As String = "ACCOUNTTAX"
    Private Const sAm As String = "CURRENCYAMOUNT"
    Private Const sPa As String = "CRITERIA"

    Public Sub New(ByRef pAccPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr, ByRef pSAPAcctngDocument As SAPAcctngDocument, pTest As Boolean)
        aAccPar = pAccPar
        aIntPar = pIntPar
        ' get Metadata
        pSAPAcctngDocument.getMeta(Hdr_Fields, Gla_Fields, Cus_Fields, Ven_Fields, Tax_Fields, Amt_Fields, Cpd_Fields, pTest)
    End Sub

    Public Function checkHeader() As Boolean
        checkHeader = If(Not (aHdrRec.aTDataRecCol.Contains("HD-COMP_CODE") Or aHdrRec.aTDataRecCol.Contains("GL+CU+VE+HD-COMP_CODE")) Or
                         Not aHdrRec.aTDataRecCol.Contains("HD-DOC_DATE") Or
                         Not aHdrRec.aTDataRecCol.Contains("HD-PSTNG_DATE") Or
                         Not aHdrRec.aTDataRecCol.Contains("HD-DOC_TYPE") Or
                         Not aHdrRec.aTDataRecCol.Contains("HD-DOC_DATE") Or
                         Not aCurRec.aTDataRecCol.Contains("A00-CURRENCY"), False, True)
    End Function

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec
        aCurRec = New TDataRec
        Dim aPostRec As New TDataRec
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec
        Dim aNewCurRec As New TDataRec
        aPostRec = pData.getPostingRecord()
        If IsNothing(aPostRec) Then
            log.Debug("fillHeader - " & "aPostRec Is Nothing -> Nothing To post, don't fill Header")
        fillHeader = False
            Exit Function
        End If
        For Each aKvb In aAccPar.getData()
            aTStrRec = aKvb.Value
            If valid_Hdr_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            ElseIf valid_Cur_Field(aTStrRec) Then
                aNewCurRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
        Next
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        For Each aTStrRec In aPostRec.aTDataRecCol
            If valid_Hdr_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            ElseIf valid_Cur_Field(aTStrRec) Then
                aNewCurRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            End If
        Next
        aHdrRec = aNewHdrRec
        aCurRec = aNewCurRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aAccTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        Dim isPA As Boolean
        aData = New TData(aIntPar)
        aAmounts = New Dictionary(Of String, TData)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            isPA = aTDataRec.getIsPa(aIntPar)
            Dim aAccType As String = aTDataRec.getAccType(aIntPar).ToUpper
            Dim aIntAccTStrRec As SAPCommon.TStrRec = aTDataRec.getAccTStrRec(aIntPar)
            Select Case aAccType
                Case "S", "G"
                    aAccTStrRec = New SAPCommon.TStrRec
                    aAccTStrRec.setValues(sGL, "GL_ACCOUNT", aIntAccTStrRec.Value, aIntAccTStrRec.Currency, aIntAccTStrRec.Format)
                    aData.addValue(CStr(aCnt), aAccTStrRec)
                    ' add the valid gl-account fields
                    For Each aTStrRec In aTDataRec.aTDataRecCol
                        If valid_Gla_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sGL)
                        ElseIf valid_Amt_Field(aTStrRec) Then
                            addAmountRecord(CStr(aCnt), aTStrRec)
                        ElseIf valid_Ext_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec)
                        End If
                        If valid_Pa_Field(aTStrRec) And isPA Then
                            aTStrRec.Fieldname = If(aIntPar.value(sPa, aTStrRec.Fieldname) <> "", aIntPar.value(sPa, aTStrRec.Fieldname), aTStrRec.Fieldname)
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sPa)
                        End If
                            If valid_Tax_Field(aTStrRec) Then 'TX information can be for TX and GL
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sTx)
                        End If
                    Next
                Case "D", "C"
                    aAccTStrRec = New SAPCommon.TStrRec
                    aAccTStrRec.setValues(sCu, "CUSTOMER", aIntAccTStrRec.Value, aIntAccTStrRec.Currency, aIntAccTStrRec.Format)
                    aData.addValue(CStr(aCnt), aAccTStrRec)
                    ' add the valid customer-account fields
                    For Each aTStrRec In aTDataRec.aTDataRecCol
                        If valid_Cus_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCu)
                        ElseIf valid_Cpd_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sCp)
                        ElseIf valid_Amt_Field(aTStrRec) Then
                            addAmountRecord(CStr(aCnt), aTStrRec)
                        ElseIf valid_Ext_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec)
                        End If
                    Next
                Case "K", "V"
                    aAccTStrRec = New SAPCommon.TStrRec
                    aAccTStrRec.setValues(sVe, "VENDOR_NO", aIntAccTStrRec.Value, aIntAccTStrRec.Currency, aIntAccTStrRec.Format)
                    aData.addValue(CStr(aCnt), aAccTStrRec)
                    ' add the valid vendor-account fields
                    For Each aTStrRec In aTDataRec.aTDataRecCol
                        If valid_Ven_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sVe)
                        ElseIf valid_Amt_Field(aTStrRec) Then
                            addAmountRecord(CStr(aCnt), aTStrRec)
                        ElseIf valid_Ext_Field(aTStrRec) Then
                            aData.addValue(CStr(aCnt), aTStrRec)
                        End If
                    Next
            End Select
            aCnt += 1
        Next
    End Function

    Private Function addAmountRecord(pDataKey As String, pTStrRec As SAPCommon.TStrRec) As Boolean
        addAmountRecord = False
        Dim aTData As TData
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCurrencyType As String
        aCurrencyType = Right(pTStrRec.Strucname, 2)
        If aAmounts.ContainsKey(pDataKey) Then
            aTData = aAmounts(pDataKey)
        Else
            aTData = New TData(aIntPar)
            aAmounts.Add(pDataKey, aTData)
        End If
        aTStrRec = New SAPCommon.TStrRec
        aTStrRec.setValues(sAm, "CURR_TYPE", aCurrencyType)
        aTData.addValue(aCurrencyType, aTStrRec)
        aTData.addValue(aCurrencyType, pTStrRec, pNewStrucname:=sAm)
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Hdr_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("HD", aStrucName) Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hdr_Fields)
        End If
    End Function

    Public Function valid_Cur_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cur_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("A00", aStrucName) Or isInArray("A10", aStrucName) Or isInArray("A20", aStrucName) Or isInArray("A30", aStrucName) Or isInArray("A40", aStrucName) Then
            valid_Cur_Field = If(pTStrRec.Fieldname = "CURRENCY", True, False)
        End If
    End Function

    Public Function valid_Gla_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Gla_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("GL", aStrucName) Then
            valid_Gla_Field = isInArray(pTStrRec.Fieldname, Gla_Fields)
        End If
    End Function

    Public Function valid_Pa_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Pa_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("PA", aStrucName) Then
            ' no validity check for CO-PA field names
            valid_Pa_Field = True
        End If
    End Function

    Public Function valid_Cus_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cus_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CU", aStrucName) Then
            valid_Cus_Field = isInArray(pTStrRec.Fieldname, Cus_Fields)
        End If
    End Function

    Public Function valid_Cpd_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Cpd_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("CP", aStrucName) Then
            valid_Cpd_Field = isInArray(pTStrRec.Fieldname, Cpd_Fields)
        End If
    End Function

    Public Function valid_Ven_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Ven_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("VE", aStrucName) Then
            valid_Ven_Field = isInArray(pTStrRec.Fieldname, Ven_Fields)
        End If
    End Function

    Public Function valid_Tax_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Tax_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("TX", aStrucName) Then
            valid_Tax_Field = isInArray(pTStrRec.Fieldname, Tax_Fields)
        End If
    End Function

    Public Function valid_Amt_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Amt_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        If isInArray("A00", aStrucName) Or isInArray("A10", aStrucName) Or isInArray("A20", aStrucName) Or isInArray("A30", aStrucName) Or isInArray("A40", aStrucName) Then
            valid_Amt_Field = isInArray(pTStrRec.Fieldname, Amt_Fields)
        End If
    End Function

    Public Function valid_Ext_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        Dim aValExtString As String = If(aIntPar.value("STR", "VALEXT") <> "", aIntPar.value("STR", "VALEXT"), "")
        valid_Ext_Field = False
        aStrucName = Split(aValExtString, ",")
        If isInArray(pTStrRec.Strucname, aStrucName) Then
            valid_Ext_Field = True
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getDocCurrency() As String
        Dim aTStrRec As SAPCommon.TStrRec = aCurRec.aTDataRecCol("A00-CURRENCY")
        If IsNothing(aTStrRec) Then
            getDocCurrency = ""
        Else
            getDocCurrency = aTStrRec.Value
        End If
    End Function

    Public Function getCurrency(pCurrencyType As String) As String
        Dim aTStrRec As SAPCommon.TStrRec = aCurRec.aTDataRecCol("A" & pCurrencyType & "-CURRENCY")
        If IsNothing(aTStrRec) Then
            getCurrency = ""
        Else
            getCurrency = aTStrRec.Value
        End If
    End Function

    Public Function getCompanyCode() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getCompanyCode = ""
        For Each aTStrRec In aHdrRec.aTDataRecCol
            If aTStrRec.Fieldname = "COMP_CODE" Then
                getCompanyCode = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("DBG", "DUMPHEADER") <> "", aIntPar.value("DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the DBG-DUMPHEADR Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
            ' dump the Currencies
            aFieldArray = {}
            aValueArray = {}
            For Each aTStrRec In aCurRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(3, 1), aDWS.Cells(3, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(4, 1), aDWS.Cells(4, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("DBG", "DUMPDATA") <> "", aIntPar.value("DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapAccAddIn.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aKvB_Am As KeyValuePair(Of String, TDataRec)
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec
            Dim aDataRec_Am As New TDataRec
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB In aData.aTDataDic
                aDataRec = aKvB.Value
                Dim aFieldArray() As String = {}
                Dim aValueArray() As String = {}
                For Each aTStrRec In aDataRec.aTDataRecCol
                    Array.Resize(aFieldArray, aFieldArray.Length + 1)
                    aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                    Array.Resize(aValueArray, aValueArray.Length + 1)
                    aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                Next
                aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                aRange.Value = aFieldArray
                aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                aRange.Value = aValueArray
                i += 2
                ' dump the amount data
                aData_Am = aAmounts(aKvB.Key)
                For Each aKvB_Am In aData_Am.aTDataDic
                    aDataRec_Am = aKvB_Am.Value
                    aFieldArray = {}
                    aValueArray = {}
                    For Each aTStrRec In aDataRec_Am.aTDataRecCol
                        Array.Resize(aFieldArray, aFieldArray.Length + 1)
                        aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                        Array.Resize(aValueArray, aValueArray.Length + 1)
                        aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                    Next
                    aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                    aRange.Value = aFieldArray
                    aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                    aRange.Value = aValueArray
                    i += 2
                Next
            Next
        End If
    End Sub

End Class
