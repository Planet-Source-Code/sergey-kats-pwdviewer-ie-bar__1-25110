VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordViewBus.clsCrypto
' Cryptography class, encapsulates MS CryptoAPI
'
'------------------------------------------------------------------------------
' This class was taken from the VBPJ Jan 2001 article by Dan Appleman.
' Thanks goes out to the author.
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

Private Declare Function CryptAcquireContext Lib "Advapi32.dll" Alias "CryptAcquireContextA" ( _
                            hProv As Long, _
                            ByVal pszContainer As String, _
                            ByVal pszProvider As String, _
                            ByVal dwProvType As Long, _
                            ByVal dwFlags As Long _
                            ) As Long
' Provider type
Private Const PROV_RSA_FULL = 1
' Provider name
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
' Flags
Private Const CRYPT_NEWKEYSET = &H8&

Private Declare Function CryptCreateHash Lib "Advapi32.dll" ( _
                            ByVal hProv As Long, _
                            ByVal algid As Long, _
                            ByVal hkey As Long, _
                            ByVal dwFlags As Long, _
                            phHash As Long) As Long

Private Declare Function CryptHashDataString Lib "Advapi32.dll" Alias "CryptHashData" ( _
                            ByVal hHash As Long, _
                            ByVal bData As String, _
                            ByVal dwDataLen As Long, _
                            ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam Lib "Advapi32.dll" ( _
                            ByVal hHash As Long, _
                            ByVal dwParam As Long, _
                            pbData As Long, _
                            dwDataLength As Long, _
                            ByVal dwFlags As Long _
                            ) As Long
    
Private Declare Function CryptGetHashParamString Lib "Advapi32.dll" Alias "CryptGetHashParam" ( _
                            ByVal hHash As Long, _
                            ByVal dwParam As Long, _
                            ByVal pbData As String, _
                            dwDataLength As Long, _
                            ByVal dwFlags As Long _
                            ) As Long
' Parameter types
Private Const HP_HASHVAL = 2
Private Const HP_HASHSIZE = 4

Private Declare Function CryptDeriveKey Lib "Advapi32.dll" ( _
                            ByVal hProv As Long, _
                            ByVal algid As Long, _
                            ByVal hBaseData As Long, _
                            ByVal dwFlags As Long, _
                            hkey As Long _
                            ) As Long

' Algorithms
Private Const CALG_RC4 = &H6801&
Private Const CALG_MD5 = &H8003&
' Flags
Const CRYPT_EXPORTABLE = &H1&

Private Declare Function CryptEncrypt Lib "Advapi32.dll" ( _
                            ByVal hkey As Long, _
                            ByVal hHash As Long, _
                            ByVal Final As Long, _
                            ByVal dwFlags As Long, _
                            ByVal pData As String, _
                            dwDataLength As Long, _
                            ByVal dwBufLength As Long _
                            ) As Long
   
Private Declare Function CryptDecrypt Lib "Advapi32.dll" ( _
                            ByVal hkey As Long, _
                            ByVal hHash As Long, _
                            ByVal Final As Long, _
                            ByVal dwFlags As Long, _
                            ByVal pData As String, _
                            dwDataLength As Long _
                            ) As Long
' Errors from Encrypt/Decrypt
Private Const NTE_BAD_LENGTH = &H80090004

Private Declare Function CryptDestroyKey Lib "Advapi32.dll" ( _
                            ByVal hkey As Long _
                            ) As Long

Private Declare Function CryptDestroyHash Lib "Advapi32.dll" ( _
                            ByVal hHash As Long _
                            ) As Long

Private Declare Function CryptReleaseContext Lib "Advapi32.dll" ( _
                            ByVal hProv As Long, _
                            ByVal dwFlags As Long _
                            ) As Long


Public Function DoEncrypt(ByVal sData As String, ByVal sKey As String) As String
    
    Dim hProv As Long
    Dim hHash As Long
    Dim hkey As Long
    
    ' Initialize Crypto APIs
    hProv = GetCryptContext
    ' Generate MD5 hash from the key string
    hHash = CreateHash(hProv, sKey)
    ' Create RC4 key from hash
    hkey = GetCryptKey(hProv, hHash)
    ' Encrypt the data
    DoEncrypt = Encrypt(hkey, sData)
    
    ' Release key and hash
    CryptDestroyKey hkey
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
End Function

Public Function DoDecrypt(ByVal sData As String, ByVal sKey As String) As String
    
    Dim hProv As Long
    Dim hHash As Long
    Dim hkey As Long
    
    ' Initialize Crypto APIs
    hProv = GetCryptContext
    ' Generate MD5 hash from the key string
    hHash = CreateHash(hProv, sKey)
    ' Create RC4 key from hash
    hkey = GetCryptKey(hProv, hHash)
    ' Decrypt the data
    DoDecrypt = Decrypt(hkey, sData)
    
    ' Release key and hash
    CryptDestroyKey hkey
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
End Function

Public Function DoHash(ByVal sKey As String) As String
    
    Dim hProv As Long
    Dim hHash As Long
    
    ' Initialize Crypto APIs
    hProv = GetCryptContext
    ' Generate MD5 hash from the key string
    hHash = CreateHash(hProv, sKey)
    ' Generate string representation of the hash
    DoHash = GetHashString(hHash)
    
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0
    
End Function


Private Function GetCryptContext() As Long
    
    Dim lProv As Long
    Dim sContainer As String
    
    sContainer = "PWDVIEWER"
    
    If 0 = CryptAcquireContext(lProv, sContainer, MS_DEF_PROV, PROV_RSA_FULL, 0) Then
        ' Try again with new key
        If 0 = CryptAcquireContext(lProv, sContainer, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_NEWKEYSET) Then
            ' Error
        Else
            GetCryptContext = lProv
        End If
    Else
        GetCryptContext = lProv
    End If
    
End Function

Private Function CreateHash(ByVal hProv As Long, ByVal sToHash As String) As Long
   
    Dim res As Long
    Dim hHash As Long
    
    If 0 = CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash) Then
        ' Error
        Exit Function
    End If
    If 0 = CryptHashDataString(hHash, sToHash, Len(sToHash), 0) Then
        ' Error
        Exit Function
    End If
    CreateHash = hHash
    
End Function

Private Function GetHashString(ByVal hHash As Long) As String
    
    Dim lLen As Long
    Dim lPtrLen As Long
    Dim sHash As String
    
    lPtrLen = 4
    
    If 0 = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, lPtrLen, 0) Then
        ' Error
    Else
        sHash = String$(lLen, 0)
        If 0 = CryptGetHashParamString(hHash, HP_HASHVAL, sHash, lLen, 0) Then
            ' Error
        Else
            GetHashString = Left$(sHash, lLen)
        End If
    End If
    
End Function

Private Function GetCryptKey(ByVal hProv As Long, ByVal hHash As Long) As Long
    
    Dim hkey As Long
    
    If 0 = CryptDeriveKey(hProv, CALG_RC4, hHash, CRYPT_EXPORTABLE, hkey) Then
        ' Error
    Else
        GetCryptKey = hkey
    End If
    
End Function

Private Function Encrypt(ByVal hkey As Long, ByVal sData As String) As String
    
    Dim lLen As Long
    Dim lBufLen As Long
    Dim lRet As Long
    
    lLen = Len(sData)
    lBufLen = lLen
    
    lRet = CryptEncrypt(hkey, 0, 1, 0, sData, lLen, lBufLen)
    If lRet = 0 Then
        If Err.LastDllError = NTE_BAD_LENGTH Then
            lLen = Len(sData)
            sData = sData & String$(lBufLen - lLen, 0)
            If 0 = CryptEncrypt(hkey, 0, 1, 0, sData, lLen, lBufLen) Then
                ' Error
            Else
                Encrypt = sData
            End If
        End If
    Else
        Encrypt = sData
    End If
    
End Function

Private Function Decrypt(ByVal hkey As Long, ByVal sData As String) As String
    
    Dim lLen As Long
    Dim lRet As Long
    
    lLen = Len(sData)
    
    lRet = CryptDecrypt(hkey, 0, 1, 0, sData, lLen)
    If lRet = 0 Then
        If Err.LastDllError = NTE_BAD_LENGTH Then
            sData = sData & String$(lLen - Len(sData), 0)
            lLen = Len(sData)
            If 0 = CryptDecrypt(hkey, 0, 1, 0, sData, lLen) Then
                ' Error
            Else
                Decrypt = sData
            End If
        End If
    Else
        Decrypt = sData
    End If
    
End Function
