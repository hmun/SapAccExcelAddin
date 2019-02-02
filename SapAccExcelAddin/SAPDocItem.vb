' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPDocItem

    Public ACCTYPE As String
    Public NEWKO As String
    Public Betrag As Double
    Public MWSKZ As String
    Public SGTXT As String
    Public AUFNR As String
    Public MATNR As String
    Public WERKS As String
    Public KOSTL As String
    Public LIFNR As String
    Public PA As String
    Public VKORG As String
    Public VTWEG As String
    Public SPART As String
    Public KNDNR As String
    Public KTGRM As String
    Public PRCTR As String
    Public PMNTTRMS As String
    Public BLINE_DATE As String
    Public ALLOC_NMBR As String
    Public BETR2 As Double
    Public CURRTYP2 As String
    Public WAERS2 As String
    Public BETR3 As Double
    Public CURRTYP3 As String
    Public WAERS3 As String
    Public BETR4 As Double
    Public CURRTYP4 As String
    Public WAERS4 As String
    Public WBS As String
    Public SEGMENT As String
    Public FUNC_AREA As String
    Public TRADE_ID As String
    Public BUS_AREA As String
    Public BEWAR As String
    Public NETWORK As String
    Public ACTIVITY As String
    Public COMP_CODE As String
    Public PARTNER_SEGMENT As String
    Public PART_PRCTR As String
    Public ZZETXT As String
    Public ZZHFMC1 As String
    Public ZZHFMC3 As String
    Public MTART As String
    Public REF_KEY_3 As String
    Public PMNT_BLOCK As String
    Public SP_GL_IND As String
    Public TXJCD As String
    Public ZZDIM06 As String
    Public ZZDIM07 As String
    Public BUPLA As String
    Public SALES_ORD As String
    Public S_ORD_ITEM As String
    Public PARTNER_BK As String

    Public Function create(pACCTYPE As String, pNEWKO As String, pBetrag As Double, pMWSKZ As String, pSGTXT As String,
                       pAUFNR As String, pMATNR As String, pWERKS As String, pKOSTL As String,
                       pLIFNR As String,
                       pPA As String, pVKORG As String, pVTWEG As String, pSPART As String,
                       pKNDNR As String, pKTGRM As String, pPRCTR As String,
                       pPMNTTRMS As String, pBLINE_DATE As String, pALLOC_NMBR As String,
                       pBETR2 As Double, pCURRTYP2 As String, pWAERS2 As String,
                       pBETR3 As Double, pCURRTYP3 As String, pWAERS3 As String,
                       pBETR4 As Double, pCURRTYP4 As String, pWAERS4 As String,
                       pWBS As String, pSEGMENT As String, pFUNC_AREA As String,
                       pTRADE_ID As String, pBUS_AREA As String, pBEWAR As String,
                       pNETWORK As String, pACTIVITY As String, pCOMP_CODE As String, pPARTNER_SEGMENT As String,
                       pPART_PRCTR As String, pZZETXT As String, pZZHFMC1 As String, pZZHFMC3 As String, pMTART As String,
                       pREF_KEY_3 As String, pPMNT_BLOCK As String, pSP_GL_IND As String,
                       pTXJCD As String, pZZDIM06 As String, pZZDIM07 As String, pBUPLA As String,
                       pSALES_ORD As String, pS_ORD_ITEM As String, pPARTNER_BK As String) As SAPDocItem
        Dim aSAPDocItem As New SAPDocItem
        aSAPDocItem.ACCTYPE = pACCTYPE
        aSAPDocItem.NEWKO = pNEWKO
        aSAPDocItem.Betrag = pBetrag
        aSAPDocItem.MWSKZ = pMWSKZ
        aSAPDocItem.SGTXT = pSGTXT
        aSAPDocItem.AUFNR = pAUFNR
        aSAPDocItem.MATNR = pMATNR
        aSAPDocItem.WERKS = pWERKS
        aSAPDocItem.KOSTL = pKOSTL
        aSAPDocItem.LIFNR = pLIFNR
        aSAPDocItem.PA = pPA
        aSAPDocItem.VKORG = pVKORG
        aSAPDocItem.VTWEG = pVTWEG
        aSAPDocItem.SPART = pSPART
        aSAPDocItem.KNDNR = pKNDNR
        aSAPDocItem.KTGRM = pKTGRM
        aSAPDocItem.PRCTR = pPRCTR
        aSAPDocItem.MWSKZ = pMWSKZ
        aSAPDocItem.PMNTTRMS = pPMNTTRMS
        aSAPDocItem.BLINE_DATE = pBLINE_DATE
        aSAPDocItem.ALLOC_NMBR = pALLOC_NMBR
        aSAPDocItem.BETR2 = pBETR2
        aSAPDocItem.CURRTYP2 = pCURRTYP2
        aSAPDocItem.WAERS2 = pWAERS2
        aSAPDocItem.BETR3 = pBETR3
        aSAPDocItem.CURRTYP3 = pCURRTYP3
        aSAPDocItem.WAERS3 = pWAERS3
        aSAPDocItem.BETR4 = pBETR4
        aSAPDocItem.CURRTYP4 = pCURRTYP4
        aSAPDocItem.WAERS4 = pWAERS4
        aSAPDocItem.WBS = pWBS
        aSAPDocItem.SEGMENT = pSEGMENT
        aSAPDocItem.FUNC_AREA = pFUNC_AREA
        aSAPDocItem.TRADE_ID = pTRADE_ID
        aSAPDocItem.BUS_AREA = pBUS_AREA
        aSAPDocItem.BEWAR = pBEWAR
        aSAPDocItem.NETWORK = pNETWORK
        aSAPDocItem.ACTIVITY = pACTIVITY
        aSAPDocItem.COMP_CODE = pCOMP_CODE
        aSAPDocItem.PARTNER_SEGMENT = pPARTNER_SEGMENT
        aSAPDocItem.PART_PRCTR = pPART_PRCTR
        aSAPDocItem.ZZETXT = pZZETXT
        aSAPDocItem.ZZHFMC1 = pZZHFMC1
        aSAPDocItem.ZZHFMC3 = pZZHFMC3
        aSAPDocItem.MTART = pMTART
        aSAPDocItem.REF_KEY_3 = pREF_KEY_3
        aSAPDocItem.PMNT_BLOCK = pPMNT_BLOCK
        aSAPDocItem.SP_GL_IND = pSP_GL_IND
        aSAPDocItem.TXJCD = pTXJCD
        aSAPDocItem.ZZDIM06 = pZZDIM06
        aSAPDocItem.ZZDIM07 = pZZDIM07
        aSAPDocItem.BUPLA = pBUPLA
        aSAPDocItem.SALES_ORD = pSALES_ORD
        aSAPDocItem.S_ORD_ITEM = pS_ORD_ITEM
        aSAPDocItem.PARTNER_BK = pPARTNER_BK

        create = aSAPDocItem
    End Function

End Class
