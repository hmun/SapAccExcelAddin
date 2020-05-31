Public Class BInv
    Public aBInv As Dictionary(Of String, BInvRec) = New Dictionary(Of String, BInvRec)

    Public Sub addBInv(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                        pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                        pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pPO_NUMBER As String, pPO_ITEM As String,
                        pITEM_TEXT As String, pITEM_AMOUNT As String, pQUANTITY As String, pPO_UNIT As String, pGL_ACCOUNT As String,
                        pCOSTCENTER As String, pNETWORK As String, pACTIVITY As String, pWBS_ELEM As String, pASSET_NO As String,
                        pSUB_NUMBER As String, pORDERID As String)
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            aBInvRec.setValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT, pDIFF_INV, pINVOICE_DOC_ITEM, "", pPO_NUMBER, pPO_ITEM, pITEM_TEXT, pITEM_AMOUNT, pQUANTITY, pPO_UNIT, pGL_ACCOUNT, pCOSTCENTER, pNETWORK, pACTIVITY, pWBS_ELEM, pASSET_NO, pSUB_NUMBER, pORDERID)
        Else
            aBInvRec = New BInvRec
            aBInvRec.setValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT, pDIFF_INV, pINVOICE_DOC_ITEM, "", pPO_NUMBER, pPO_ITEM, pITEM_TEXT, pITEM_AMOUNT, pQUANTITY, pPO_UNIT, pGL_ACCOUNT, pCOSTCENTER, pNETWORK, pACTIVITY, pWBS_ELEM, pASSET_NO, pSUB_NUMBER, pORDERID)
            aBInv.Add(aKey, aBInvRec)
        End If
    End Sub

    Public Function getITEM_TEXT(pINV_DOC_NO As String, pFISC_YEAR As String, pINVOICE_DOC_ITEM As String) As String
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            getITEM_TEXT = aBInvRec.aITEM_TEXT.Value
        Else
            getITEM_TEXT = ""
        End If
    End Function

    Public Function getPO_NUMBER(pINV_DOC_NO As String, pFISC_YEAR As String, pINVOICE_DOC_ITEM As String) As String
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            getPO_NUMBER = aBInvRec.aPO_NUMBER.Value
        Else
            getPO_NUMBER = ""
        End If
    End Function

    Public Function getPO_ITEM(pINV_DOC_NO As String, pFISC_YEAR As String, pINVOICE_DOC_ITEM As String) As String
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            getPO_ITEM = aBInvRec.aPO_ITEM.Value
        Else
            getPO_ITEM = ""
        End If
    End Function

    Public Sub addBInvAcc(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                        pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                        pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pSERIAL_NO As String,
                        pITEM_TEXT As String, pPO_NUMBER As String, pPO_ITEM As String,
                        pITEM_AMOUNT As String, pQUANTITY As String, pPO_UNIT As String, pGL_ACCOUNT As String,
                        pCOSTCENTER As String, pNETWORK As String, pACTIVITY As String, pWBS_ELEM As String, pASSET_NO As String,
                        pSUB_NUMBER As String, pORDERID As String)
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM & "-" & pSERIAL_NO
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            aBInvRec.setAccValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT,
                                  pDIFF_INV, pINVOICE_DOC_ITEM, pSERIAL_NO, pITEM_TEXT, pPO_NUMBER, pPO_ITEM, pITEM_AMOUNT, pQUANTITY, pPO_UNIT, pGL_ACCOUNT, pCOSTCENTER, pNETWORK,
                                  pACTIVITY, pWBS_ELEM, pASSET_NO, pSUB_NUMBER, pORDERID)
        Else
            aBInvRec = New BInvRec
            aBInvRec.setAccValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT,
                                  pDIFF_INV, pINVOICE_DOC_ITEM, pSERIAL_NO, pITEM_TEXT, pPO_NUMBER, pPO_ITEM, pITEM_AMOUNT, pQUANTITY, pPO_UNIT, pGL_ACCOUNT, pCOSTCENTER, pNETWORK,
                                  pACTIVITY, pWBS_ELEM, pASSET_NO, pSUB_NUMBER, pORDERID)
            aBInv.Add(aKey, aBInvRec)
        End If
    End Sub

    Public Sub addBInvItem(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                        pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                        pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pPO_NUMBER As String, pPO_ITEM As String,
                        pITEM_TEXT As String)
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            aBInvRec.setItemValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT, pDIFF_INV, pINVOICE_DOC_ITEM, pPO_NUMBER, pPO_ITEM, pITEM_TEXT)
        Else
            aBInvRec = New BInvRec
            aBInvRec.setItemValues(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pDOC_TYPE, pDOC_DATE, pPSTNG_DATE, pREF_DOC_NO, pCOMP_CODE, pCURRENCY, pEXCH_RATE, pHEADER_TXT, pDIFF_INV, pINVOICE_DOC_ITEM, pPO_NUMBER, pPO_ITEM, pITEM_TEXT)
            aBInv.Add(aKey, aBInvRec)
        End If
    End Sub

    Public Sub addBInvMessage(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pINVOICE_DOC_ITEM As String)
        Dim aBInvRec As BInvRec
        Dim aKey As String
        aKey = pINV_DOC_NO & "-" & pFISC_YEAR & "-" & pINVOICE_DOC_ITEM
        If aBInv.ContainsKey(aKey) Then
            aBInvRec = aBInv(aKey)
            aBInvRec.setMessage(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pINVOICE_DOC_ITEM)
        Else
            aBInvRec = New BInvRec
            aBInvRec.setMessage(pRET_MESSAGE, pINV_DOC_NO, pFISC_YEAR, pINVOICE_DOC_ITEM)
            aBInv.Add(aKey, aBInvRec)
        End If
    End Sub

End Class

Public Class BInvRec
    Public aRET_MESSAGE As SAPCommon.TField
    Public aINV_DOC_NO As SAPCommon.TField
    Public aFISC_YEAR As SAPCommon.TField
    Public aDOC_TYPE As SAPCommon.TField
    Public aDOC_DATE As SAPCommon.TField
    Public aPSTNG_DATE As SAPCommon.TField
    Public aREF_DOC_NO As SAPCommon.TField
    Public aCOMP_CODE As SAPCommon.TField
    Public aCURRENCY As SAPCommon.TField
    Public aEXCH_RATE As SAPCommon.TField
    Public aHEADER_TXT As SAPCommon.TField
    Public aDIFF_INV As SAPCommon.TField
    Public aINVOICE_DOC_ITEM As SAPCommon.TField
    Public aSERIAL_NO As SAPCommon.TField
    Public aPO_NUMBER As SAPCommon.TField
    Public aPO_ITEM As SAPCommon.TField
    Public aITEM_TEXT As SAPCommon.TField
    Public aITEM_AMOUNT As SAPCommon.TField
    Public aQUANTITY As SAPCommon.TField
    Public aPO_UNIT As SAPCommon.TField
    Public aGL_ACCOUNT As SAPCommon.TField
    Public aCOSTCENTER As SAPCommon.TField
    Public aNETWORK As SAPCommon.TField
    Public aACTIVITY As SAPCommon.TField
    Public aWBS_ELEM As SAPCommon.TField
    Public aASSET_NO As SAPCommon.TField
    Public aSUB_NUMBER As SAPCommon.TField
    Public aORDERID As SAPCommon.TField

    Public Sub New()
        aRET_MESSAGE = New SAPCommon.TField
        aINV_DOC_NO = New SAPCommon.TField
        aFISC_YEAR = New SAPCommon.TField
        aDOC_TYPE = New SAPCommon.TField
        aDOC_DATE = New SAPCommon.TField
        aPSTNG_DATE = New SAPCommon.TField
        aREF_DOC_NO = New SAPCommon.TField
        aCOMP_CODE = New SAPCommon.TField
        aCURRENCY = New SAPCommon.TField
        aEXCH_RATE = New SAPCommon.TField
        aHEADER_TXT = New SAPCommon.TField
        aDIFF_INV = New SAPCommon.TField
        aINVOICE_DOC_ITEM = New SAPCommon.TField
        aSERIAL_NO = New SAPCommon.TField
        aPO_NUMBER = New SAPCommon.TField
        aPO_ITEM = New SAPCommon.TField
        aITEM_TEXT = New SAPCommon.TField
        aITEM_AMOUNT = New SAPCommon.TField
        aQUANTITY = New SAPCommon.TField
        aPO_UNIT = New SAPCommon.TField
        aGL_ACCOUNT = New SAPCommon.TField
        aCOSTCENTER = New SAPCommon.TField
        aNETWORK = New SAPCommon.TField
        aACTIVITY = New SAPCommon.TField
        aWBS_ELEM = New SAPCommon.TField
        aASSET_NO = New SAPCommon.TField
        aSUB_NUMBER = New SAPCommon.TField
        aORDERID = New SAPCommon.TField
    End Sub

    Public Sub setValues(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                          pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                          pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pSERIAL_NO As String, pPO_NUMBER As String, pPO_ITEM As String,
                          pITEM_TEXT As String, pITEM_AMOUNT As String, pQUANTITY As String, pPO_UNIT As String, pGL_ACCOUNT As String,
                          pCOSTCENTER As String, pNETWORK As String, pACTIVITY As String, pWBS_ELEM As String, pASSET_NO As String,
                          pSUB_NUMBER As String, pORDERID As String)
        aRET_MESSAGE = New SAPCommon.TField("RET_MESSAGE", CStr(pRET_MESSAGE))
        aINV_DOC_NO = New SAPCommon.TField("INV_DOC_NO", CStr(pINV_DOC_NO))
        aFISC_YEAR = New SAPCommon.TField("FISC_YEAR", CStr(pFISC_YEAR))
        aDOC_TYPE = New SAPCommon.TField("DOC_TYPE", CStr(pDOC_TYPE))
        aDOC_DATE = New SAPCommon.TField("DOC_DATE", CStr(pDOC_DATE))
        aPSTNG_DATE = New SAPCommon.TField("PSTNG_DATE", CStr(pPSTNG_DATE))
        aREF_DOC_NO = New SAPCommon.TField("REF_DOC_NO", CStr(pREF_DOC_NO))
        aCOMP_CODE = New SAPCommon.TField("COMP_CODE", CStr(pCOMP_CODE))
        aCURRENCY = New SAPCommon.TField("CURRENCY", CStr(pCURRENCY))
        aEXCH_RATE = New SAPCommon.TField("EXCH_RATE", CStr(pEXCH_RATE))
        aHEADER_TXT = New SAPCommon.TField("HEADER_TXT", CStr(pHEADER_TXT))
        aDIFF_INV = New SAPCommon.TField("DIFF_INV", CStr(pDIFF_INV))
        aINVOICE_DOC_ITEM = New SAPCommon.TField("INVOICE_DOC_ITEM", CStr(pINVOICE_DOC_ITEM))
        aSERIAL_NO = New SAPCommon.TField("SERIAL_NO", CStr(pSERIAL_NO))
        aPO_NUMBER = New SAPCommon.TField("PO_NUMBER", CStr(pPO_NUMBER))
        aPO_ITEM = New SAPCommon.TField("PO_ITEM", CStr(pPO_ITEM))
        aITEM_TEXT = New SAPCommon.TField("ITEM_TEXT", CStr(pITEM_TEXT))
        aITEM_AMOUNT = New SAPCommon.TField("ITEM_AMOUNT", CStr(pITEM_AMOUNT))
        aQUANTITY = New SAPCommon.TField("QUANTITY", CStr(pQUANTITY))
        aPO_UNIT = New SAPCommon.TField("PO_UNIT", CStr(pPO_UNIT))
        aGL_ACCOUNT = New SAPCommon.TField("GL_ACCOUNT", CStr(pGL_ACCOUNT))
        aCOSTCENTER = New SAPCommon.TField("COSTCENTER", CStr(pCOSTCENTER))
        aNETWORK = New SAPCommon.TField("NETWORK", CStr(pNETWORK))
        aACTIVITY = New SAPCommon.TField("ACTIVITY", CStr(pACTIVITY))
        aWBS_ELEM = New SAPCommon.TField("WBS_ELEM", CStr(pWBS_ELEM))
        aASSET_NO = New SAPCommon.TField("ASSET_NO", CStr(pASSET_NO))
        aSUB_NUMBER = New SAPCommon.TField("SUB_NUMBER", CStr(pSUB_NUMBER))
        aORDERID = New SAPCommon.TField("ORDERID", CStr(pORDERID))
    End Sub

    Public Sub setItemValues(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                          pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                          pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pPO_NUMBER As String, pPO_ITEM As String,
                          pITEM_TEXT As String)
        aRET_MESSAGE = New SAPCommon.TField("RET_MESSAGE", CStr(pRET_MESSAGE))
        aINV_DOC_NO = New SAPCommon.TField("INV_DOC_NO", CStr(pINV_DOC_NO))
        aFISC_YEAR = New SAPCommon.TField("FISC_YEAR", CStr(pFISC_YEAR))
        aDOC_TYPE = New SAPCommon.TField("DOC_TYPE", CStr(pDOC_TYPE))
        aDOC_DATE = New SAPCommon.TField("DOC_DATE", CStr(pDOC_DATE))
        aPSTNG_DATE = New SAPCommon.TField("PSTNG_DATE", CStr(pPSTNG_DATE))
        aREF_DOC_NO = New SAPCommon.TField("REF_DOC_NO", CStr(pREF_DOC_NO))
        aCOMP_CODE = New SAPCommon.TField("COMP_CODE", CStr(pCOMP_CODE))
        aCURRENCY = New SAPCommon.TField("CURRENCY", CStr(pCURRENCY))
        aEXCH_RATE = New SAPCommon.TField("EXCH_RATE", CStr(pEXCH_RATE))
        aHEADER_TXT = New SAPCommon.TField("HEADER_TXT", CStr(pHEADER_TXT))
        aDIFF_INV = New SAPCommon.TField("DIFF_INV", CStr(pDIFF_INV))

        aINVOICE_DOC_ITEM = New SAPCommon.TField("INVOICE_DOC_ITEM", CStr(pINVOICE_DOC_ITEM))
        aPO_NUMBER = New SAPCommon.TField("PO_NUMBER", CStr(pPO_NUMBER))
        aPO_ITEM = New SAPCommon.TField("PO_ITEM", CStr(pPO_ITEM))
        aITEM_TEXT = New SAPCommon.TField("ITEM_TEXT", CStr(pITEM_TEXT))
    End Sub

    Public Sub setAccValues(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pDOC_TYPE As String, pDOC_DATE As String,
                          pPSTNG_DATE As String, pREF_DOC_NO As String, pCOMP_CODE As String, pCURRENCY As String, pEXCH_RATE As String,
                          pHEADER_TXT As String, pDIFF_INV As String, pINVOICE_DOC_ITEM As String, pSERIAL_NO As String,
                          pITEM_TEXT As String, pPO_NUMBER As String, pPO_ITEM As String,
                          pITEM_AMOUNT As String, pQUANTITY As String, pPO_UNIT As String, pGL_ACCOUNT As String,
                          pCOSTCENTER As String, pNETWORK As String, pACTIVITY As String, pWBS_ELEM As String, pASSET_NO As String,
                          pSUB_NUMBER As String, pORDERID As String)
        aRET_MESSAGE = New SAPCommon.TField("RET_MESSAGE", CStr(pRET_MESSAGE))
        aINV_DOC_NO = New SAPCommon.TField("INV_DOC_NO", CStr(pINV_DOC_NO))
        aFISC_YEAR = New SAPCommon.TField("FISC_YEAR", CStr(pFISC_YEAR))
        aDOC_TYPE = New SAPCommon.TField("DOC_TYPE", CStr(pDOC_TYPE))
        aDOC_DATE = New SAPCommon.TField("DOC_DATE", CStr(pDOC_DATE))
        aPSTNG_DATE = New SAPCommon.TField("PSTNG_DATE", CStr(pPSTNG_DATE))
        aREF_DOC_NO = New SAPCommon.TField("REF_DOC_NO", CStr(pREF_DOC_NO))
        aCOMP_CODE = New SAPCommon.TField("COMP_CODE", CStr(pCOMP_CODE))
        aCURRENCY = New SAPCommon.TField("CURRENCY", CStr(pCURRENCY))
        aEXCH_RATE = New SAPCommon.TField("EXCH_RATE", CStr(pEXCH_RATE))
        aHEADER_TXT = New SAPCommon.TField("HEADER_TXT", CStr(pHEADER_TXT))
        aDIFF_INV = New SAPCommon.TField("DIFF_INV", CStr(pDIFF_INV))

        aINVOICE_DOC_ITEM = New SAPCommon.TField("INVOICE_DOC_ITEM", CStr(pINVOICE_DOC_ITEM))
        aSERIAL_NO = New SAPCommon.TField("SERIAL_NO", CStr(pSERIAL_NO))
        aITEM_TEXT = New SAPCommon.TField("ITEM_TEXT", CStr(pITEM_TEXT))
        aPO_NUMBER = New SAPCommon.TField("PO_NUMBER", CStr(pPO_NUMBER))
        aPO_ITEM = New SAPCommon.TField("PO_ITEM", CStr(pPO_ITEM))
        aITEM_AMOUNT = New SAPCommon.TField("ITEM_AMOUNT", CStr(pITEM_AMOUNT))
        aQUANTITY = New SAPCommon.TField("QUANTITY", CStr(pQUANTITY))
        aPO_UNIT = New SAPCommon.TField("PO_UNIT", CStr(pPO_UNIT))
        aGL_ACCOUNT = New SAPCommon.TField("GL_ACCOUNT", CStr(pGL_ACCOUNT))
        aCOSTCENTER = New SAPCommon.TField("COSTCENTER", CStr(pCOSTCENTER))
        aNETWORK = New SAPCommon.TField("NETWORK", CStr(pNETWORK))
        aACTIVITY = New SAPCommon.TField("ACTIVITY", CStr(pACTIVITY))
        aWBS_ELEM = New SAPCommon.TField("WBS_ELEM", CStr(pWBS_ELEM))
        aASSET_NO = New SAPCommon.TField("ASSET_NO", CStr(pASSET_NO))
        aSUB_NUMBER = New SAPCommon.TField("SUB_NUMBER", CStr(pSUB_NUMBER))
        aORDERID = New SAPCommon.TField("ORDERID", CStr(pORDERID))
    End Sub

    Public Sub setMessage(pRET_MESSAGE As String, pINV_DOC_NO As String, pFISC_YEAR As String, pINVOICE_DOC_ITEM As String)
        aRET_MESSAGE = New SAPCommon.TField("RET_MESSAGE", CStr(pRET_MESSAGE))
        aINV_DOC_NO = New SAPCommon.TField("INV_DOC_NO", CStr(pINV_DOC_NO))
        aFISC_YEAR = New SAPCommon.TField("FISC_YEAR", CStr(pFISC_YEAR))
        aINVOICE_DOC_ITEM = New SAPCommon.TField("INVOICE_DOC_ITEM", CStr(pINVOICE_DOC_ITEM))
    End Sub

    Public Function getKey() As String
        Dim aKey As String
        aKey = aINV_DOC_NO.Value & "-" & aFISC_YEAR.Value & "-" & aINVOICE_DOC_ITEM.Value & "-" & aSERIAL_NO.Value
        getKey = aKey
    End Function

    Public Function getKeyR() As String
        Dim aKey As String
        aKey = aINV_DOC_NO.Value & "-" & aFISC_YEAR.Value & "-" & aINVOICE_DOC_ITEM.Value & "-" & aSERIAL_NO.Value
        getKeyR = aKey
    End Function

    Public Function toStringValue() As Object
        Dim aArray(26) As String
        aArray(0) = aRET_MESSAGE.Value
        aArray(1) = aINV_DOC_NO.Value
        aArray(2) = aFISC_YEAR.Value
        aArray(3) = aDOC_TYPE.Value
        aArray(4) = aDOC_DATE.Value
        aArray(5) = aPSTNG_DATE.Value
        aArray(6) = aREF_DOC_NO.Value
        aArray(7) = aCOMP_CODE.Value
        aArray(8) = aCURRENCY.Value
        aArray(9) = aEXCH_RATE.Value
        aArray(10) = aHEADER_TXT.Value
        aArray(11) = aDIFF_INV.Value
        aArray(12) = aINVOICE_DOC_ITEM.Value
        aArray(13) = aSERIAL_NO.Value
        aArray(14) = aPO_NUMBER.Value
        aArray(15) = aPO_ITEM.Value
        aArray(16) = aITEM_TEXT.Value
        aArray(17) = aITEM_AMOUNT.Value
        aArray(18) = aQUANTITY.Value
        aArray(19) = aPO_UNIT.Value
        aArray(20) = aGL_ACCOUNT.Value
        aArray(21) = aCOSTCENTER.Value
        aArray(22) = aNETWORK.Value
        aArray(23) = aACTIVITY.Value
        If aWBS_ELEM.Value = "00000000" Then
            aArray(24) = ""
        Else
            aArray(24) = aWBS_ELEM.Value
        End If
        '    aArray(24) = aASSET_NO.Value
        '    aArray(25) = aSUB_NUMBER.Value
        aArray(25) = aORDERID.Value
        toStringValue = aArray
    End Function
End Class
