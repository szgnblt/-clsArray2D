VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr_Data As Variant
Public Property Get SatirBaslangic() As Variant

Call ArrayiKontrolEt("SatirBaslangic()")
SatirBaslangic = LBound(arr_Data, 1)
    
End Property
Public Property Get SatirBitis() As Variant

Call ArrayiKontrolEt("SatirBitis()")
SatirBitis = UBound(arr_Data, 1)
    
End Property
Public Property Get SatirSayisi() As Variant

SatirSayisi = UBound(arr_Data, 1) - LBound(arr_Data, 1) + 1

End Property
Public Property Get SutunBaslangic() As Variant

Call ArrayiKontrolEt("SutunBaslangic()")
SutunBaslangic = LBound(arr_Data, 2)

End Property
Public Property Get SutunBitis() As Variant

Call ArrayiKontrolEt("SutunBitis()")
SutunBitis = UBound(arr_Data, 2)

End Property
Public Property Get SutunSayisi() As Variant

Call ArrayiKontrolEt("SutunBitis()")
SutunSayisi = UBound(arr_Data, 2) - LBound(arr_Data, 2) + 1

End Property
Public Property Get Data() As Variant

Data = arr_Data
    
End Property

Public Property Let Data(ByRef Data As Variant)
   
If IsEmpty(Data) = True Or IsArray(Data) = False Then
    ReDim arr_Data(1 To 1, 1 To 1)
    arr_Data(1, 1) = Data
ElseIf IsArray(Data) = True Then
    If BoyutlariKontrolEt(Data) <> 2 Then
        MsgBox "Veri Array Formatinda Degil", _
        vbInformation, "Sayin " & Environ("UserName")
    End If
    arr_Data = Data
End If
    
End Property
Public Property Let Value(ByVal satir As Long, ByVal sutun As Long, ByVal Data As Variant)

arr_Data(satir, sutun) = Data

End Property
Public Property Get Value(ByVal satir As Long, ByVal sutun As Long) As Variant

Value = arr_Data(satir, sutun)

End Property
Public Sub AlandanYukle(ByVal rg As Range _
                    , Optional basligiKaldir As Boolean = False _
                    , Optional basligiSatirSayisi As Long = 1 _
                    , Optional SoldanBasligiKaldir As Boolean = False _
                    , Optional soldanBaslikSayisi As Long = 1)


If basligiKaldir = True Or SoldanBasligiKaldir = True Then

    Dim SatirSayisi As Long, rowOffset As Long
    SatirSayisi = IIf(basligiKaldir = True, rg.rows.count - _
    basligiSatirSayisi, rg.rows.count)
    rowOffset = IIf(basligiKaldir = True, basligiSatirSayisi, 0)
    
    Dim SutunSayisi As Long, columnOffset As Long
    SutunSayisi = IIf(SoldanBasligiKaldir = True, rg.columns.count - _
    soldanBaslikSayisi, rg.columns.count)
    columnOffset = IIf(SoldanBasligiKaldir = True, soldanBaslikSayisi, 0)
    
    Set rg = rg.Offset(rowOffset, columnOffset).Resize(SatirSayisi, SutunSayisi)
    
End If

arr_Data = rg.Value

End Sub
Public Function TekrarEdenleriKaldir(ByVal sadeceSecilenler As Boolean, _
                                    ParamArray SutunIndeksi()) As Variant

'Dictionary ile Calisir. Uyumluluk icin Dictionary arastirmasi yapiniz.

Dim VeriArray()
Dim ret()
Dim indexes()
Dim ids()
Dim i As Long
Dim r As Long
Dim c As Long
Dim seper As String
Dim dict As Object
Dim sutunIndexArr As Variant
Dim arrSon As Variant

If UBound(SutunIndeksi) <= LBound(SutunIndeksi) And _
UBound(SutunIndeksi) <= 0 Then
    ReDim sutunIndexArr(0 To UBound(arr_Data, 2) - 1)
    For i = LBound(sutunIndexArr) To UBound(sutunIndexArr)
        sutunIndexArr(i) = i + 1
    Next i
Else
    sutunIndexArr = SutunIndeksi
End If
Set dict = CreateObject("Scripting.Dictionary")

VeriArray = arr_Data: seper = ChrW(-1)
ReDim ids(LBound(sutunIndexArr) To UBound(sutunIndexArr))

For r = UBound(VeriArray) To LBound(VeriArray) Step -1
    For c = LBound(sutunIndexArr) To UBound(sutunIndexArr)
        ids(c) = VeriArray(r, sutunIndexArr(c))
    Next
    dict(Join$(ids, seper)) = r
Next
indexes = dict.items()
Sort1DArraY indexes

If sadeceSecilenler Then
    ReDim ret(LBound(VeriArray) To LBound(VeriArray) + dict.count - 1, _
    LBound(sutunIndexArr) + 1 To UBound(sutunIndexArr) + 1)
Else
    ReDim ret(LBound(VeriArray) To LBound(VeriArray) + dict.count - 1, _
    LBound(VeriArray, 2) To UBound(VeriArray, 2))
End If

For c = IIf(sadeceSecilenler = True, LBound(ret, 2), _
LBound(VeriArray, 2)) To IIf(sadeceSecilenler = True, _
UBound(ret, 2), UBound(VeriArray, 2))

    For r = LBound(ret) To UBound(ret)
        If sadeceSecilenler Then
            ret(r, c) = VeriArray(indexes(r - 1), sutunIndexArr(c - 1))
        Else
            ret(r, c) = VeriArray(indexes(r - 1), c)
        End If
    Next
Next

arr_Data = ret

End Function
Private Sub Sort1DArraY(ByRef arr As Variant)

Dim i As Long
Dim j As Long
Dim temp As Variant

For i = LBound(arr) To UBound(arr)
    For j = i + 1 To UBound(arr)
        If arr(i) > arr(j) Then
            temp = arr(j)
            arr(j) = arr(i)
            arr(i) = temp
        End If
    Next j
Next i

End Sub
Public Sub ArraydenYukle(ByVal arr As Variant)

arr_Data = arr

End Sub
Public Sub ArraydenKismiYukle(ByVal arr As Variant, _
                              Optional satir As Long = 1, _
                              Optional sutun As Long = 1, _
                              Optional SatirSayisi As Long = 1, _
                              Optional SutunSayisi As Long = 1)
                              
Dim i As Long
Dim j As Long
Dim counterRow As Long
Dim counterColumn As Long
Dim boyut As Long

boyut = BoyutlariKontrolEt(arr)
If boyut > 2 Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If sutun > UBound(arr, 2) Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satir < LBound(arr, 1) Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satir + SatirSayisi > UBound(arr, 1) Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirSayisi, _
SutunBaslangic To SutunSayisi)

counterRow = 1: counterColumn = 1
For i = LBound(arr, 2) To UBound(arr, 2)
    If i >= sutun And i + SutunSayisi - 1 Then
        For j = LBound(arr, 1) To UBound(arr, 1)
            If j >= satir And j + SatirSayisi - 1 Then
                temp(counterRow, counterColumn) = arr(j, i)
                If counterRow = UBound(temp, 1) Then counterRow = 1: Exit For
                counterRow = counterRow + 1
            End If
        Next j
        If counterColumn = UBound(temp, 2) Then Exit For
        counterColumn = counterColumn + 1
    End If
Next i

arr_Data = temp

End Sub
Public Function SutunlariSatirlardanAl(ByVal satir As Long, _
                                        Optional ByVal SatirSayisi As Long = 1) As Variant
    
If satir - 1 + SatirSayisi > SatirBitis Then
    MsgBox "Array icin Satir Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newColumns As Variant
ReDim newColumns(1 To SutunBitis, 1 To SatirSayisi)

Dim lastRow As Long: lastRow = ((satir - 1) + SatirSayisi)
Dim currentRow As Long: currentRow = SatirBaslangic

Dim i As Long, j As Long
For i = satir To lastRow
    For j = 1 To SutunBitis
        newColumns(j, currentRow) = arr_Data(i, j)
    Next j
    currentRow = currentRow + 1
Next i

SutunlariSatirlardanAl = newColumns

End Function
Public Function SatirlariSutundanAl(ByVal sutun As Long, _
                                    Optional ByVal SutunSayisi As Long = 1) As Variant

Dim i As Long
Dim j As Long

If sutun - 1 + SutunSayisi > SutunBitis Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newRows As Variant
ReDim newRows(SutunBaslangic To SutunSayisi, _
SatirBaslangic To SatirBitis)

Dim lastColumn As Long: lastColumn = ((sutun - 1) + SutunSayisi)
Dim currentColumn As Long: currentColumn = SutunBaslangic


For j = sutun To lastColumn
    For i = SatirBaslangic To SatirBitis
        newRows(currentColumn, i) = arr_Data(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

SatirlariSutundanAl = newRows

End Function
Public Function SatirlariAl(ByVal satir As Long, _
                            Optional ByVal SatirSayisi As Long = 1) As Variant


Dim newRow As Variant
ReDim newRow(SatirBaslangic To SatirSayisi, _
SutunBaslangic To SutunBitis)

Dim lastRow As Long: lastRow = ((satir - 1) + SatirSayisi)
Dim currentRow As Long: currentRow = SatirBaslangic

Dim i As Long, j As Long
For i = satir To lastRow
    For j = SutunBaslangic To SutunBitis
        newRow(currentRow, j) = arr_Data(i, j)
    Next j
    currentRow = currentRow + 1
Next i

SatirlariAl = newRow

End Function
Public Function SutunlariAl(ByVal sutun As Long, _
                            Optional ByVal SutunSayisi As Long = 1) As Variant

If sutun - 1 + SutunSayisi > SutunBitis Then
    MsgBox "Array icin Sutun Sayisi Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim newColumn As Variant
ReDim newColumn(SatirBaslangic To SatirBitis, SutunBaslangic To SutunSayisi)

Dim lastColumn As Long: lastColumn = ((sutun - 1) + SutunSayisi)
Dim currentColumn As Long: currentColumn = SutunBaslangic

Dim i As Long, j As Long
For j = sutun To lastColumn
    For i = SatirBaslangic To SatirBitis
        newColumn(i, currentColumn) = arr_Data(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

SutunlariAl = newColumn

End Function
Public Sub VeriyiAlanaYazdir(ByVal rg As Range, _
                            Optional ByVal transpose As Boolean = False)
    
If transpose = True Then
    rg.Resize(SutunBitis, SatirBitis).Value = _
    SatirlariSutundanAl(1, SutunSayisi)
Else
    rg.Resize(SatirBitis, SutunBitis).Value = arr_Data
End If
    
End Sub
Public Sub SatirlariAlanaYazdir(ByVal alanaYaz As Range _
                                , ByVal satir As Long _
                                , Optional SatirSayisi As Long = 1 _
                                , Optional transpose As Boolean = False)

If transpose = True Then
    alanaYaz.Resize(SutunBitis, SatirSayisi).Value = _
    SutunlariSatirlardanAl(satir, SatirSayisi)
Else
    alanaYaz.Resize(SatirSayisi, SutunBitis).Value = _
    SatirlariAl(satir, SatirSayisi)
End If
    
End Sub
Public Sub SutunlariAlanaYazdir(ByVal alanaYaz As Range _
                                , ByVal sutun As Long _
                                , Optional SutunSayisi As Long = 1 _
                                , Optional transpose As Boolean = False)
    
If transpose = True Then
    alanaYaz.Resize(SutunSayisi, SatirBitis).Value = _
    SatirlariSutundanAl(sutun, SutunSayisi)
Else
    alanaYaz.Resize(SatirBitis, SutunSayisi).Value = _
    SutunlariAl(sutun, SutunSayisi)
End If
    
End Sub
Private Function ArrayKopyala(ByRef arrSource As Variant _
                            , ByRef arrDestination As Variant)
    
Dim i As Long, j As Long
For i = LBound(arrDestination, 1) To UBound(arrDestination, 1)
    For j = LBound(arrDestination, 2) To UBound(arrDestination, 2)
        arrDestination(i, j) = arrSource(i, j)
    Next j
Next i
    
End Function
Private Sub SatirKopyala(ByRef rowData As Variant _
                            , ByVal currentRow As Long _
                            , ByVal newRow As Long)
    
    
Dim i As Long
For i = LBound(arr_Data, 2) To UBound(arr_Data, 2)
    rowData(newRow, i) = arr_Data(currentRow, i)
Next i
    
End Sub
Private Function BoyutlariKontrolEt(ByRef arr As Variant) As Long

If IsArray(arr) = False Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

On Error GoTo errPlace

Dim dimensions As Long
Dim count As Long: count = 0

Do
    count = count + 1
    dimensions = UBound(arr, count)
Loop While True

errPlace:
BoyutlariKontrolEt = count - 1

End Function
Public Sub BosSutunEkle(ByVal sutunEkle As Long, _
                Optional ByVal SutunSayisi As Long = 1)
    
If sutunEkle > SutunBitis + 1 Then
    sutunEkle = SutunBitis + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis + SutunSayisi)

Dim i As Long, j As Long
For j = SutunBaslangic To sutunEkle - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j
    
For j = sutunEkle To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j + SutunSayisi) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp
        
End Sub
Public Sub SutunEkleme(ByVal sutunEkle As Long, ByRef yeniSutunlar As Variant)
    
If IsEmpty(yeniSutunlar) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If BoyutlariKontrolEt(yeniSutunlar) <> 2 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If UBound(yeniSutunlar, 1) < SatirBitis Then
    MsgBox "Yeni Array Formati Dogru Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunEkle < 1 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunEkle > SutunBitis + 1 Then
    sutunEkle = SutunBitis + 1
End If

Dim newColumnCount As Long
newColumnCount = UBound(yeniSutunlar, 2)

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis + newColumnCount)
Dim i As Long, j As Long, currentColumn As Long

currentColumn = sutunEkle
For j = SutunBaslangic To newColumnCount
    For i = SatirBaslangic To SatirBitis
        temp(i, currentColumn) = yeniSutunlar(i, j)
    Next i
    currentColumn = currentColumn + 1
Next j

For j = SutunBaslangic To sutunEkle - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j
    
For j = sutunEkle To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j + newColumnCount) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp
        
End Sub
Public Sub BosSatirEkle(ByVal satirEkle As Long, Optional ByVal SatirSayisi As Long = 1)
    
If satirEkle > SatirBitis + 1 Then
    satirEkle = SatirBitis + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis + _
SatirSayisi, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long
For i = SatirBaslangic To satirEkle - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i
    
For i = satirEkle To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i + SatirSayisi, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp
        
End Sub
Public Sub SatirEkleme(ByVal satirEkle As Long, ByRef yeniSatirlar As Variant)
    
If IsEmpty(yeniSatirlar) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If BoyutlariKontrolEt(yeniSatirlar) <> 2 Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If UBound(yeniSatirlar, 2) < SutunBitis Then
    MsgBox "Veride Yeterli Satir Yok", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If satirEkle < 1 Then
    MsgBox "Satir Sayisi Birden Buyuk Olmalidir", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If satirEkle > SatirBitis + 1 Then
    satirEkle = SatirBitis + 1
End If

Dim newRowCount As Long
newRowCount = UBound(yeniSatirlar, 1)

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis + _
newRowCount, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long, currentRow As Long
currentRow = satirEkle
For i = SatirBaslangic To newRowCount
    For j = SutunBaslangic To SutunBitis
        temp(currentRow, j) = yeniSatirlar(i, j)
    Next j
    currentRow = currentRow + 1
Next i

For i = SatirBaslangic To satirEkle - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i
    
For i = satirEkle To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i + newRowCount, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp
        
End Sub
Public Sub SutunlariKaldir(ByVal sutunuKaldir As Long, _
                            Optional ByVal SutunSayisi As Long = 1)
    
If sutunuKaldir < 1 Or sutunuKaldir > SutunBitis Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If sutunuKaldir = 1 And SutunSayisi >= SutunBitis Then
    MsgBox "Veri Sutun Sayisi Yeterli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

If sutunuKaldir - 1 + SutunSayisi > SutunBitis Then
    SutunSayisi = SutunBitis - sutunuKaldir + 1
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis, _
SutunBaslangic To SutunBitis - SutunSayisi)

Dim i As Long, j As Long
For j = SutunBaslangic To sutunuKaldir - 1
    For i = SatirBaslangic To SatirBitis
        temp(i, j) = arr_Data(i, j)
    Next i
Next j

For j = sutunuKaldir + SutunSayisi To SutunBitis
    For i = SatirBaslangic To SatirBitis
        temp(i, j - SutunSayisi) = arr_Data(i, j)
    Next i
Next j

arr_Data = temp

End Sub
Public Sub SatirlariSil(ByVal satiriKaldir As Long, _
                        Optional ByVal SatirSayisi As Long = 1)
    
If satiriKaldir < 1 Or satiriKaldir > SatirBitis Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If satiriKaldir - 1 + SatirSayisi > SatirBitis Then
    MsgBox "Veri Satir Sayisi Yeterli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim temp As Variant
ReDim temp(SatirBaslangic To SatirBitis - _
SatirSayisi, SutunBaslangic To SutunBitis)

Dim i As Long, j As Long
For i = SatirBaslangic To satiriKaldir - 1
    For j = SutunBaslangic To SutunBitis
        temp(i, j) = arr_Data(i, j)
    Next j
Next i

For i = satiriKaldir + SatirSayisi To SatirBitis
    For j = SutunBaslangic To SutunBitis
        temp(i - SatirSayisi, j) = arr_Data(i, j)
    Next j
Next i

arr_Data = temp

End Sub
Public Function ElementIndeks(ByVal elementAra As Variant _
                        , Optional ByVal sutun As Long = 1 _
                        , Optional ByVal baslangicPozisyonu As Long = 1 _
                        , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                        , Optional ByVal bosluklariKirp As Boolean = False) As Long
    
If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

ElementIndeks = -1

Dim i As Long
For i = baslangicPozisyonu To SatirBitis
    
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        ElementIndeks = i
        Exit For
    End If
    
Next i
        
End Function
Public Function ElementSonIndeks(ByVal elementAra As Variant _
                        , Optional ByVal sutun As Long = 1 _
                        , Optional ByVal baslangicPozisyonu As Long = 1 _
                        , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                        , Optional ByVal bosluklariKirp As Boolean = False) As Long
    
If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

ElementSonIndeks = -1

Dim i As Long
For i = SatirBitis To baslangicPozisyonu Step -1
    
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        ElementSonIndeks = i
        Exit For
    End If
    
Next i
        
End Function
Public Function CokEFiltrele(ByVal sadeceSecilenler As Boolean, _
                            ByVal karsilastirmaTuru As VbCompareMethod, _
                            ParamArray Kriter() As Variant) As Variant

Dim i As Long
Dim j As Long
Dim q As Long
Dim counter As Long
Dim upperLimit As Long
Dim matchFound As Boolean
Dim numCompare As Boolean
Dim dateCompare As Boolean
Dim varCompare As Variant
Dim tmpVar As Variant
Dim arr As Variant
Dim arrCevap As Variant
Dim arrSon As Variant

arr = arr_Data

If sadeceSecilenler Then
    ReDim arrCevap(1 To UBound(arr, 1), 1 To UBound(Kriter) / 2)
Else
    ReDim arrCevap(1 To UBound(arr, 1), 1 To UBound(arr, 2))
End If

counter = 1
For j = LBound(arr, 1) To UBound(arr, 1)
    matchFound = True
    For i = LBound(Kriter) To UBound(Kriter)
        If i Mod 2 = 0 Then
            If Kriter(i + 1) Like ">*" Or Kriter(i + 1) Like "<*" Then
                tmpVar = Replace(Replace(Kriter(i + 1), ">", ""), "<", "")
                If IsDate(tmpVar) Then dateCompare = True Else: dateCompare = False
                If isNumeric(tmpVar) Then numCompare = True Else: numCompare = False
                If dateCompare Then varCompare = CDate(tmpVar)
                If numCompare And Not dateCompare Then varCompare = CDbl(tmpVar)
                
                If LCase(TypeName(arr(j, Kriter(i)))) = "string" Then matchFound = False: Exit For
                If Kriter(i + 1) Like ">*" Then
                    If Not arr(j, Kriter(i)) > varCompare Then matchFound = False: Exit For
                ElseIf Kriter(i + 1) Like "<*" Then
                    If Not arr(j, Kriter(i)) < varCompare Then matchFound = False: Exit For
                End If
            Else
                If LCase(TypeName(arr(j, Kriter(i)))) <> "string" Then matchFound = False: Exit For
                If InStr(1, Kriter(i + 1), Chr(42)) > 0 Then
                    tmpVar = Replace(Kriter(i + 1), Chr(42), "")
                    
                    If Right(Kriter(i + 1), 1) = Chr(42) And _
                    Left(Kriter(i + 1), 1) = Chr(42) Then
                        If InStr(1, arr(j, Kriter(i)), tmpVar, _
                        karsilastirmaTuru) = 0 Then matchFound = False: Exit For
                    
                    ElseIf Right(Kriter(i + 1), 1) = Chr(42) Then
                        If karsilastirmaTuru = vbTextCompare Then
                            If LCase(Left(arr(j, Kriter(i)), Len(tmpVar))) <> _
                            LCase(tmpVar) Then matchFound = False: Exit For
                        ElseIf karsilastirmaTuru = vbBinaryCompare Then
                            If Left(arr(j, Kriter(i)), Len(tmpVar)) <> tmpVar _
                            Then matchFound = False: Exit For
                        End If
                        
                    ElseIf Left(Kriter(i + 1), 1) = Chr(42) Then
                        If karsilastirmaTuru = vbTextCompare Then
                            If LCase(Right(arr(j, Kriter(i)), Len(tmpVar))) <> _
                            LCase(tmpVar) Then matchFound = False: Exit For
                        ElseIf karsilastirmaTuru = vbBinaryCompare Then
                            If Right(arr(j, Kriter(i)), Len(tmpVar)) <> _
                            tmpVar Then matchFound = False: Exit For
                        End If
                        
                    End If
                Else
                
                    If Not StrComp(arr(j, Kriter(i)), _
                    Kriter(i + 1), karsilastirmaTuru) = 0 Then
                        matchFound = False: Exit For
                    End If
                    
                End If
            End If
        End If
    Next i
    If matchFound Then
        If Not sadeceSecilenler Then
            For q = LBound(arr, 2) To UBound(arr, 2)
                arrCevap(counter, q) = arr(j, q)
            Next q
        Else
            For q = LBound(Kriter) To UBound(Kriter)
                If q Mod 2 = 1 Then
                    arrCevap(counter, Int(q / 2) + 1) = _
                    arr(j, Kriter(q - 1))
                End If
            Next q
        End If
        counter = counter + 1
    End If
Next j

upperLimit = UBound(arrCevap, 2): i = 0: j = 0
ReDim arrSon(1 To counter - 1, 1 To upperLimit)

For i = LBound(arrSon, 2) To UBound(arrSon, 2)
    For j = LBound(arrSon, 1) To UBound(arrSon, 1)
        arrSon(j, i) = arrCevap(j, i)
    Next j
Next i

CokEFiltrele = arrSon

End Function
Public Function Filtrele(ByVal elementAra As Variant _
                    , Optional ByVal sutun As Long = 1 _
                    , Optional ByVal baslangicPozisyonu As Long = 1 _
                    , Optional ByVal karsilastirmaTuru As VbCompareMethod = vbTextCompare _
                    , Optional ByVal bosluklariKirp As Boolean = False) As Variant

If sutun > SutunBitis Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If
If baslangicPozisyonu > SatirBitis Then
    MsgBox "Baslangic Pozisyonu Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim filteredData As Variant
ReDim filteredData(1 To UBound(arr_Data, 1), 1 To UBound(arr_Data, 2))

Dim count As Long: count = 1
    
Dim i As Long, newRow As Long, filterEmpty As Boolean
newRow = 1
filterEmpty = True
For i = baslangicPozisyonu To SatirBitis
    If StrComp(IIf(bosluklariKirp = True, Trim(arr_Data(i, sutun)), _
    arr_Data(i, sutun)), elementAra, karsilastirmaTuru) = 0 Then
        SatirKopyala filteredData, i, newRow
        newRow = newRow + 1
        filterEmpty = False
    End If
Next i

If filterEmpty = True Then
    Filtrele = Empty
Else
    Dim finalData As Variant
    ReDim finalData(1 To newRow - 1, 1 To UBound(filteredData, 2))
    ArrayKopyala filteredData, finalData
End If

Filtrele = finalData

End Function
Public Function MaxMin(ByVal sutun As Long, MaxBul As Boolean, SayisalBul As Boolean) As Variant

If sutun > SutunBitis Or sutun < 1 Then
    MsgBox "Arama Elementi Gecerli Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

Dim lnMax As Variant
Dim lnMin As Variant
Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
If SayisalBul Then
    lnMin = findNumericorNotValues(arr, sutun, True): lnMax = 0
Else
    lnMin = findNumericorNotValues(arr, sutun, False): lnMax = ""
End If

For j = LBound(arr_Data, 2) To UBound(arr_Data, 2)
    If j = sutun Then
        For i = LBound(arr_Data, 1) To UBound(arr_Data, 1)
            If SayisalBul Then
                If Not IsEmpty(arr_Data(i, j)) Then
                    If isNumeric(arr_Data(i, j)) Then
                        lnMin = IIf(arr_Data(i, j) > lnMin, lnMin, arr_Data(i, j))
                        lnMax = IIf(arr_Data(i, j) > lnMax, arr_Data(i, j), lnMax)
                    End If
                End If
            Else
                If Not isNumeric(arr_Data(i, j)) Then
                    lnMin = IIf(arr_Data(i, j) > lnMin, lnMin, arr_Data(i, j))
                    lnMax = IIf(arr_Data(i, j) > lnMax, arr_Data(i, j), lnMax)
                End If
            End If
        Next i
        Exit For
    End If
Next j

If MaxBul Then MaxMin = lnMax Else: MaxMin = lnMin

End Function
Private Function findNumericorNotValues(arr As Variant, _
                                        sutun As Long, isNumerical As Boolean)

Dim i As Long

For i = LBound(arr, 1) To UBound(arr, 1)
    If isNumerical And isNumeric(arr(i, _
    sutun)) Then findNumericorNotValues = _
    CDbl(arr(i, sutun)): Exit For
    If Not isNumerical And Not isNumeric(arr(i, _
    sutun)) Then findNumericorNotValues = _
    CStr(arr(i, sutun)): Exit For
Next i

End Function
Public Sub TerseCevir()

Dim rows As Long, columns As Long
rows = SatirBitis
columns = SutunBitis

Dim reversedData As Variant
ReDim reversedData(1 To rows, 1 To columns)

Dim i As Long, j As Long, row As Long
row = 1
For i = rows To 1 Step -1
    For j = 1 To columns
        reversedData(row, j) = arr_Data(i, j)
    Next j
    row = row + 1
Next i

arr_Data = reversedData

End Sub
Public Function CokETopla(ToplanacakSutun As Long, _
                        ParamArray Kriter() As Variant) As Variant

Dim genelToplam As Double
Dim r As Long
Dim c As Long
Dim conditionsMet As Boolean
Dim arr As Variant

arr = Me.Data: genelToplam = 0
If ToplanacakSutun < LBound(arr, 2) Or _
ToplanacakSutun > UBound(arr, 2) Then Err.Raise vbError + 9

For r = LBound(arr, 1) To UBound(arr, 1)
    conditionsMet = True
    For c = LBound(Kriter) To UBound(Kriter) - 1
        If c Mod 2 = 0 Then
            If arr(r, Kriter(c)) <> Kriter(c + 1) Then conditionsMet = False: Exit For
        End If
    Next c
    If conditionsMet And isNumeric(arr(r, ToplanacakSutun)) Then genelToplam = _
    genelToplam + arr(r, ToplanacakSutun)
Next r
CokETopla = genelToplam

End Function
Public Function CokESay(ParamArray Kriter() As Variant) As Long

Dim genelToplam As Double
Dim r As Long
Dim c As Long
Dim conditionsMet As Boolean
Dim arr As Variant

arr = Me.Data: genelToplam = 0
For r = LBound(arr, 1) To UBound(arr, 1)
    conditionsMet = True
    For c = LBound(Kriter) To UBound(Kriter) - 1
        If c Mod 2 = 0 Then
            If arr(r, Kriter(c)) <> Kriter(c + 1) Then conditionsMet = False: Exit For
        End If
    Next c
    If conditionsMet Then genelToplam = genelToplam + 1
Next r
CokESay = genelToplam

End Function
Public Function Topla(ToplanacakSutun As Long) As Variant

Dim i As Long
Dim j As Long
Dim genelToplam As Double
Dim arr As Variant

arr = arr_Data: genelToplam = 0
For j = LBound(arr, 2) To UBound(arr, 2)
    If j = ToplanacakSutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If isNumeric(arr(i, j)) Then
                genelToplam = genelToplam + arr(i, j)
            End If
        Next i
    End If
Next j
Topla = CDbl(genelToplam)

End Function
Public Function Say(SayilacakSutun As Long, _
                    Optional bosluklarHaric As Boolean = False) As Variant

Dim i As Long
Dim j As Long
Dim genelToplam As Double
Dim arr As Variant

arr = arr_Data: genelToplam = 0
For j = LBound(arr, 2) To UBound(arr, 2)
    If j = SayilacakSutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If bosluklarHaric Then
                If arr(i, j) <> "" Then
                    genelToplam = genelToplam + 1
                End If
            Else
                genelToplam = genelToplam + 1
            End If
        Next i
    End If
Next j
Say = genelToplam

End Function
Public Sub Kirp(ByVal sutun As Long, Optional SayilarDahil As Boolean = False)

Dim i As Long
Dim j As Long
Dim arr As Variant

arr = arr_Data
For j = 1 To UBound(arr, 2)
    If j = sutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If isNumeric(arr(i, j)) And Not SayilarDahil Then GoTo skipTrim
            arr(i, j) = Trim(arr(i, j))
skipTrim:
        Next i
    End If
Next j
arr_Data = arr

End Sub
Public Function TekBoyutluyaCevir(ByVal sutun As Long) As Variant
 
Dim i As Long
Dim j As Long
Dim tmpVar As String
Dim arr As Variant
Dim arr1D As Variant

arr = arr_Data
ReDim arr1D(1 To UBound(arr, 1))

For j = LBound(arr, 2) To UBound(arr, 2)
    If j = sutun Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr1D(i) = arr(i, j)
        Next i
    End If
Next j
TekBoyutluyaCevir = arr1D

End Function
Public Sub TurkceKarakterleriCevir(BosluklariKalsin As Boolean, ParamArray Kriter() As Variant)

Dim i As Long
Dim j As Long
Dim tmpVar As String
Dim arr As Variant

arr = arr_Data
For j = LBound(Kriter) To UBound(Kriter)
    If j + 1 <= UBound(arr, 2) Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            If Not isNumeric(arr(i, j + 1)) Then
                tmpVar = IIf(BosluklariKalsin = True, arr(i, j + 1), Trim(arr(i, j + 1)))
                arr(i, j + 1) = TurkcetoLatin(tmpVar)
            End If
        Next i
    End If
Next j
arr_Data = arr

End Sub
Private Function TurkcetoLatin(ByVal strInput As String) As String

Dim i As Long
Dim arrTurkce() As Variant
Dim arrEnglish() As Variant

arrTurkce = Array(Chr(253), Chr(221), Chr(240), Chr(208), Chr(252), Chr(220), _
Chr(254), Chr(222), Chr(246), Chr(214), Chr(231), Chr(199))
arrEnglish = Array("i", "I", "g", "G", "u", "U", "s", "S", "o", "O", "c", "C")

For i = LBound(arrTurkce) To UBound(arrTurkce)
    If InStr(1, strInput, arrTurkce(i), vbBinaryCompare) > 0 Then
        strInput = Replace(strInput, arrTurkce(i), arrEnglish(i))
    End If
Next i
TurkcetoLatin = strInput

End Function
Public Sub Sirala(ByVal sutun As Long)

QuickSort SortArray:=arr_Data, lngColumn:=sutun

End Sub
Private Sub QuickSort(ByRef SortArray As Variant, Optional lngMin As Long = -1, _
                       Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    
On Error Resume Next

Dim i As Long
Dim j As Long
Dim varMid As Variant
Dim arrRowTemp As Variant
Dim lngColTemp As Long

If IsEmpty(SortArray) Then
    Exit Sub
End If
If InStr(TypeName(SortArray), "()") < 1 Then
    Exit Sub
End If
If lngMin = -1 Then
    lngMin = LBound(SortArray, 1)
End If
If lngMax = -1 Then
    lngMax = UBound(SortArray, 1)
End If
If lngMin >= lngMax Then
    Exit Sub
End If

i = lngMin
j = lngMax

varMid = Empty
varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

If IsObject(varMid) Then
    i = lngMax
    j = lngMin
ElseIf IsEmpty(varMid) Then
    i = lngMax
    j = lngMin
ElseIf IsNull(varMid) Then
    i = lngMax
    j = lngMin
ElseIf varMid = "" Then
    i = lngMax
    j = lngMin
ElseIf VarType(varMid) = vbError Then
    i = lngMax
    j = lngMin
ElseIf VarType(varMid) > 17 Then
    i = lngMax
    j = lngMin
End If

While i <= j
    While SortArray(i, lngColumn) < varMid And i < lngMax
        i = i + 1
    Wend
    While varMid < SortArray(j, lngColumn) And j > lngMin
        j = j - 1
    Wend

    If i <= j Then
        ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
        For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
            arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
            SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
            SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
        Next lngColTemp
        Erase arrRowTemp

        i = i + 1
        j = j - 1
    End If
Wend

If (lngMin < j) Then Call QuickSort(SortArray, lngMin, j, lngColumn)
If (i < lngMax) Then Call QuickSort(SortArray, i, lngMax, lngColumn)

End Sub
Public Sub ArrayiKontrolEt(ByVal procedure As String)

If IsEmpty(arr_Data) Then
    MsgBox "Veri Array Formatinda Degil", _
    vbInformation, "Sayin " & Environ("UserName")
End If

End Sub