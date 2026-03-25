Option Strict Off
Option Explicit On
Module AIG_KGJT_mdlVariables
    '//  SAP MANAGE UI API 6.5 SDK Sample
    '//****************************************************************************
    '//
    '//  File:      mdlVariables.bas
    '//
    '//  Copyright (c) SAP MANAGE
    '//
    '// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
    '// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
    '// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
    '// PARTICULAR PURPOSE.
    '//
    '//****************************************************************************

    Public SBO_Application As SAPbouiCOM.Application

    '// to handle form operations
    Public oForm As SAPbouiCOM.Form

    Public oMatrix As SAPbouiCOM.Matrix
    Public oColumns As SAPbouiCOM.Columns
    Public oColumn As SAPbouiCOM.Column
    Public oItem As SAPbouiCOM.Item
    Public oItems As SAPbouiCOM.Items
    Public oCells As SAPbouiCOM.Cells
    Public oCell As SAPbouiCOM.Cell
    Public oComboBox As SAPbouiCOM.ComboBox
    Public oCheckBox As SAPbouiCOM.CheckBox
    Public oButton As SAPbouiCOM.Button
    Public oEditText As SAPbouiCOM.EditText
    Public oOptionBtn As SAPbouiCOM.OptionBtn
    Public oNewItem As SAPbouiCOM.Item
    Public oStatic As SAPbouiCOM.StaticText

    Public sMenuUID As String
    Public bRunning As Boolean
    Public DocNum As String
    Public CodeGodkendTimer As String
    Public strOrdreType As String
    Public ConvertCode As String
    Public HuskLastSecound As Long = 0
    'Til KGJT
    'Public bFDate As Boolean = False
    'Public bTDate As Boolean = False
    'Public dateSelected As String
    'Public bLoggedOn As Boolean = False
    'Public bolCreateOrderView As Boolean = False
    'Public bolChangeUser As Boolean = False
    'Public isTerminalUser As Boolean = False
    'Public isAdministrator As Boolean = False
    'Public bAutoUd As Boolean = False
    'Public bStampOneShot As Boolean = False
    'Public dStampLogOnTime As Date
    'Public bCallFromOO As Boolean = False
    'Public bBrugTools As Boolean = True
    'Public LonRapport As String
    'Public DBPass As String
    'Public printname As String
    'Public printcopy As String
    Public rs As SAPbobsCOM.Recordset
    Public rs1 As SAPbobsCOM.Recordset
    Public rs2 As SAPbobsCOM.Recordset
    Public rs3 As SAPbobsCOM.Recordset
    Public rs4 As SAPbobsCOM.Recordset

    'Public SboGuiApi As SAPbouiCOM.SboGuiApi

    'Public FormID As Object

End Module