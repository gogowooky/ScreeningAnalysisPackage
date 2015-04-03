Attribute VB_Name = "T1"
Rem EXPORT ON
 
Rem ******************************************************************************************************
Rem ******************************************************************************************************
Rem
Rem
Rem ScreeningAnalysisPackage.xlsm
Rem
Rem Copyright (c) 2015  Shinichiro Egashira
Rem
Rem
Rem
Rem 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を
Rem 取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。
Rem これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または
Rem 販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
Rem
Rem
Rem 上記の著作権表示および本許諾表示を､ソフトウェアのすべての複製または重要な部分に記載するものとします｡
Rem
Rem ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。
Rem ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、
Rem それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、
Rem ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の
Rem 請求、損害、その他の義務について何らの責任も負わないものとします。
Rem
Rem
Rem
Rem ScreeningAnalysisPackage.xlsm
Rem
Rem Copyright (c) 2015 Shinichiro Egashira
Rem
Rem
Rem Permission is hereby granted, free of charge, to any person obtaining a copy
Rem of this software and associated documentation files (the "Software"), to deal
Rem in the Software without restriction, including without limitation the rights
Rem to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
Rem copies of the Software, and to permit persons to whom the Software is
Rem furnished to do so, subject to the following conditions:
Rem
Rem The above copyright notice and this permission notice shall be included in all
Rem copies or substantial portions of the Software.
Rem
Rem THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
Rem IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
Rem FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
Rem AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
Rem LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
Rem OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
Rem SOFTWARE.
Rem
Rem ******************************************************************************************************
Rem ******************************************************************************************************

Rem EXPORT OFF
   
Rem ****************************************************************************************************************************************************************
Rem 固定参照値
Rem ****************************************************************************************************************************************************************

' Template内の名前付き領域の名前
Public Const PLATE_TYPE = "PLATE_TYPE"
Public Const PLATE_READER = "PLATE_READER"
Public Const PLATE_FORMAT = "PLATE_FORMAT"
'Const ASSAY_NAME = "ASSAY_NAME"
'Const ASSAY_DATE = "ASSAY_DATE"
'Const ASSAY_TIME = "ASSAY_TIME"
Const PLATE_WELL_POSITION = "WELL_POS"
Const PLATE_WELL_ROLE = "WELL_ROLE"
Const PLATE_COMPOUND_CONC = "CPD_CONC"
Public Const LABEL_TABLE = "TABLE"

Rem ********************************************************************************
Rem [名前] T1.SYSTEM( ... )
Rem
Rem [用途] Excel解析パッケージの情報を得る。
Rem ********************************************************************************
Public Function SYSTEM(Optional param As String = "")
	Application.Volatile
	Select Case param
		Case "title":          SYSTEM = "Screening Analysis Package for Excel"
		Case "version":        SYSTEM = "ver. 1.0.1"
		Case "update":         SYSTEM = "2015/04/03 12:21"
		Case "affiliation":    SYSTEM = "Drug Discovery Initiative (DDI)"
		Case "affiliation2":   SYSTEM = "The University of Tokyo"
		Case "affiliation3":   SYSTEM = "DDI"
		Case "homepage":       SYSTEM = "http://www.ddi.u-tokyo.ac.jp/wp/application/"
		Case "address":        SYSTEM = "Yakugaku #504B, 7-3-1 Hongo, Bunkyo-ku, Tokyo, 113-0033, JAPAN"
		Case "phone":          SYSTEM = "+81-3-5841-1960"
		Case "fax":            SYSTEM = "+81-3-5841-1959"
		Case "author":         SYSTEM = "Shinichiro Egashira"
		Case "copyright":      SYSTEM = "Copyright (c) 2015 " & SYSTEM("author")
		Case "mail":           SYSTEM = "ddiinfo@mol.f.u-tokyo.ac.jp"
		Case "original":       SYSTEM = "https://github.com/gogowooky/ScreeningAnalysisPackage"		
		Case "support_reader":     SYSTEM = T1M.SYSTEM_SUPPORT_PLATE_READER
		Case "support_plate_type": SYSTEM = T1M.SYSTEM_SUPPORT_PLATE_TYPE
		Case "today":          SYSTEM = DATE_ID(Date)
		Case "now":            SYSTEM = TIME_ID(Now)
		Case "excelver":       SYSTEM = Application.Version
		Case "pc":             SYSTEM = Array("Mac","Windows")(-CInt(CBool(InStr(Application.OperatingSystem,"Windows"))))
		Case "filename":       SYSTEM = Application.Caller.Parent.Parent.Name
		Case "path":           SYSTEM = Application.Caller.Parent.Parent.path
		Case "filepath":       SYSTEM = Application.Caller.Parent.Parent.path & Application.PathSeparator & Application.Caller.Parent.Parent.Name
		Case "parentdir":      SYSTEM = Mid( SYSTEM("path"), InStrRev( SYSTEM("path"), Application.PathSeparator) + 1)
		Case Else:             SYSTEM = SYSTEM("title") & " " & SYSTEM("version")
	End Select
End Function

Rem EXPORT ON

Rem ********************************************************************************
Rem 汎用関数
Rem ********************************************************************************

Rem
Rem [名前] T1.FORMAT_DATE
Rem [用途] 日付文字列(dat)を"yyyy/mm/dd"形式に変換する
Rem
Public Function DATE_ID( dat As Variant ) As String
   DATE_ID = Strings.Format(DateValue(dat), "yyyy/mm/dd")
End Function

Rem
Rem [名前] T1.FORMAT_TIME
Rem [用途] 時間文字列(tim)をhh:nn 形式に変換する
Rem
Public Function TIME_ID(tim As Variant) As String
	TIME_ID = Strings.Format(TimeValue(tim), "hh:nn")
End Function

Rem
Rem [名前] T1.CSV_AND
Rem [用途] csv1とcsv2のANDを返す
Rem
Public Function CSV_AND(csv1 As String, csv2 As String) As String
	Dim csv As String:  csv1 = "," & csv1 & ","
	Dim item As Variant
	For Each item In T1.CSV2ARY(csv2)
		If 0 < InStr(csv1, "," & CStr(item) & ",") Then csv = csv & "," & CStr(item)
	Next
	CSV_AND = Mid(csv, 2)
End Function

Rem
Rem [名前] T1.CSV_OR
Rem [用途] csv1とcsv2のORを返す
Rem
Public Function CSV_OR(csv1 As String, csv2 As String) As String
	Dim csv As String: csv = "," & csv1 & ","
	Dim item As Variant
	For Each item In T1.CSV2ARY(csv2)
		If item <> "" And InStr(csv, "," & CStr(item) & ",") = 0 Then csv1 = csv1 & "," & CStr(item)
	Next
	CSV_OR = csv1
End Function

Rem
Rem [名前] T1.CSV_SUB
Rem [用途] csv1とcsv2の差分を返す
Rem
Public Function CSV_SUB(csv1 As String, csv2 As String) As String
	Dim item As Variant
	Dim csv As String: csv2 = "," & csv2 & ","
	For Each item In T1.CSV2ARY(csv1)
		If item <> "" And InStr(csv2, "," & CStr(item) & ",") = 0 Then csv = csv & "," & CStr(item)
	Next
	CSV_SUB = Mid(csv, 2)
End Function

Rem
Rem [名前] T1.CSV2ARY
Rem [用途] CSV文字列を配列にして返す。
Rem
Public Function CSV2ARY(csvstr As String, Optional num As Integer = -1) As Variant
	If num = -1 Then
		CSV2ARY = Split(csvstr, ",")
	Else
		CSV2ARY = Split(csvstr, ",")(num)
	End If
End Function

Rem
Rem [名前] T1.V2LOOKUP
Rem [用途] 選択範囲の先頭２列をkeyに行を選択、行中の値を取得する
Rem
Public Function V2LOOKUP(col1key As String, col2key As String, rng As Range, colnum As Integer) As Variant
  Dim itm As Variant
	For Each itm In Application.Intersect(rng.Parent.UsedRange, rng).Rows
		If itm.Cells(1, 1).Value = col1key And itm.Cells(1, 2).Value = col2key Then
			V2LOOKUP = rng.Parent.Cells(rng.row + itm.row - 1, rng.Column + colnum).Value
			Exit Function
		End If
	Next
	V2LOOKUP = CVErr(xlErrNA)
End Function

Rem
Rem [名前] T1.VLOOKUP2
Rem [用途] VLOOKUP関数とほぼ同じだが、カラム選択に負の値を用いることが出来る。
Rem
Public Function VLOOKUP2(colkey As String, rng As Range, colnum As Integer) As Variant
  Dim itm As Variant
	For Each itm In Application.Intersect(rng.Parent.UsedRange, rng).Rows
		If itm.Cells(1, 1).Value = colkey Then
			VLOOKUP2 = rng.Parent.Cells(rng.row + itm.row - 1, rng.Column + colnum).Value
			Exit Function
		End If
	Next
	VLOOKUP2 = CVErr(xlErrNA)
End Function

Rem
Rem [名前] T1.VHLOOKUP
Rem [用途] 選択範囲の一行目・一列目をrowkey,colkeyで各々検索しクロスするセルの値を得る。
Rem
Public Function VHLOOKUP(rowkey As String, colkey As String, rng As Range) As Variant
  Dim rw As Variant
  Dim cl As Variant
	For Each rw In rng.Columns(1).Rows
		If rw.Value = rowkey Then
			For Each cl In Application.Intersect(rng.Parent.UsedRange, rng).Rows(1).Columns
				If cl.Value = colkey Then
					VHLOOKUP = rng.Cells(rw.row - rng.row + 1, cl.Column - rng.Column + 1).Value
					Exit Function
				End If
			Next
		End If
	Next
	VHLOOKUP = CVErr(xlErrNA)
End Function

Rem
Rem [名前] T1.FIND_ROW
Rem [用途] 選択範囲(rng)の第一列目に文字列(str1,str2,str3)を含む行の行番号を返す
Rem
Public Function FIND_ROW(rng As Range, str1 As String, Optional str2 As String = "", Optional str3 As String = "") As Variant
  On Error Resume Next
  Dim val As String: Dim rw As Variant
  FIND_ROW = 0
  For Each rw In rng.Rows
    val = rw.Columns(1).Value
    If 0 < InStr(val, str1) And 0 < InStr(val, str2) And 0 < InStr(val, str3) Then
			FIND_ROW = rw.row
			Exit Function
		End If
    If 10000 < rw.row Then Exit Function
  Next
End Function

Rem
Rem [名前] T1.RC2WELL
Rem [用途] 行(rw), 列(cl)と表示形式(param)を指定しwell pos文字列を返す。
Rem
Public Function RC2WELL(rw As Integer, cl As Integer, param As String) As Variant
  On Error GoTo ERR_RC2WELL
  RC2WELL = RESOURCE.GetWellpos(rw, cl, param)
  Exit Function
	
ERR_RC2WELL: ' 汎用
  Dim r As Variant: r = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF")
	Dim clstr As String: clstr = CStr(cl)
	Select Case param
		Case "rc":     RC2WELL = Array(rw, cl)
		Case "RC":     RC2WELL = "R" & CStr(rw) & "C" & clstr
		Case "pos":    RC2WELL = CStr(r(rw - 1) & clstr)
		Case "ROW":    RC2WELL = CStr(r(rw - 1))
		Case "row":    RC2WELL = rw
		Case "col":    RC2WELL = cl
		Case "COLUMN": RC2WELL = clstr
		Case "pos0":   RC2WELL = CStr(r(rw - 1)) & Right("0" & clstr, 2)
		Case "pos00":  RC2WELL = Right("0" & r(rw - 1), 2) & Right("0" & clstr, 2)
		Case Else:     RC2WELL = CVErr(xlErrRef)
	End Select
End Function


Rem
Rem [名前] T1.ROWTOOL
Rem [用途] ROWを並べるための便利関数。　コピーして値で貼り付けることを想定。
Rem
Public Function ROWTOOL() As String
  Dim rw As Integer
	For rw = 0 To Application.WorksheetFunction.MIN(32, Application.Caller.row - 1)
		If InStr(1, Application.Caller.Offset(-rw, 0).Formula, "T1.ROWTOOL", 1) = 0 Then Exit For
	Next
	ROWTOOL = T1.RC2WELL(rw, 1, "ROW")
End Function

Rem
Rem [名前] T1.POSTOOL
Rem [用途] well pos文字列を並べるための便利関数。　コピーして値で貼り付けることを想定。
Rem
Public Function POSTOOL(Optional param As String = "pos") As String
  Dim rw As Integer
  Dim cl As Integer
	For rw = 0 To Application.WorksheetFunction.MIN(32, Application.Caller.row - 1)
		If InStr(1, Application.Caller.Offset(-rw, 0).Formula, "=T1.POSTOOL", 1) = 0 Then Exit For
	Next
	For cl = 0 To Application.WorksheetFunction.MIN(48, Application.Caller.row - 1)
		If InStr(1, Application.Caller.Offset(0, -cl).Formula, "T1.POSTOOL", 1) = 0 Then Exit For
	Next
	POSTOOL = T1.RC2WELL(rw, cl, param)
End Function

Rem
Rem [名前] T1.PLATETOOL
Rem [用途] plateを描画させるための便利関数。　コピーして値で貼り付けることを想定。
Rem
Public Function PLATETOOL(Optional param As String = "pos") As String
  Dim rw As Integer
  Dim cl As Integer
	For rw = 1 To Application.WorksheetFunction.MIN(33, Application.Caller.row)
		If InStr(1, Application.Caller.Offset(-rw, 0).Formula, "=T1.PLATETOOL", 1) = 0 Then Exit For
	Next
	For cl = 1 To Application.WorksheetFunction.MIN(49, Application.Caller.Column)
		If InStr(1, Application.Caller.Offset(0, -cl).Formula, "=T1.PLATETOOL", 1) = 0 Then Exit For
	Next
   
	If rw = 1 Then
		If cl = 1 Then
			PLATETOOL = ""
		Else
			PLATETOOL = T1.RC2WELL(rw - 1, cl - 1, "COLUMN")
		End If
	ElseIf cl = 1 Then
		PLATETOOL = T1.RC2WELL(rw - 1, cl - 1, "ROW")
	Else
		PLATETOOL = T1.RC2WELL(rw - 1, cl - 1, "pos")
	End If

	With Application.Caller
		If rw = 1 Or cl = 1 Then
			.Borders(xlEdgeLeft).LineStyle = xlNone
			.Borders(xlEdgeTop).LineStyle = xlNone
			.Borders(xlEdgeBottom).LineStyle = xlNone
			.Borders(xlEdgeRight).LineStyle = xlNone
		Else
			.Borders(xlEdgeLeft).ColorIndex = 0
			.Borders(xlEdgeTop).ColorIndex = 0
			.Borders(xlEdgeBottom).ColorIndex = 0
			.Borders(xlEdgeRight).ColorIndex = 0
		End If
	End With
End Function

Rem
Rem [名前] T1.SELECT_WELLS
Rem [用途] 指定値(key)を指定領域(rng)で探索し、WELLPOSのCSV文字列として返す
Rem
Public Function SELECT_WELLS(rng As Range, comp As String, key As Variant) As String
	On Error Resume Next
	Application.Volatile
  SELECT_WELLS = ""
  Dim csv As String: csv = ""
  Dim r As Variant
	Dim flag As boolean
	For Each r In rng
		Select Case comp
			Case "like":   flag = 0 < InStr(r.Value, key)
			Case "match":  flag = r.Value = key
			Case "above?": flag = r.Value > key
			Case "equal?": flag = r.Value = key
			Case "below?": flag = r.Value < key
		End Select
		If flag Then csv = csv & T1.RC2WELL(r.row - rng.row + 1, r.Column - rng.Column + 1, "pos") & ","
	Next
	SELECT_WELLS = Left(csv, Len(csv) - 1)
End Function

Rem
Rem [名前] T1.NTH_VALUE
Rem [用途] 領域(rng)のnum番目の値を取得する
Rem
Public Function NTH_VALUE(rng As Range, num As Integer) As Variant
  Dim r As Variant
	For Each r In rng
		num = num - 1
		If num = 0 Then
			NTH_VALUE = r.Value
			Exit Function
		End If
	Next
End Function

Rem
Rem [名前] T1.NTH_ADDRESS
Rem [用途] 領域(rng)のnum番目のAddressを取得する
Rem
Public Function NTH_ADDRESS(rng As Range, Optional num As Integer = -1) As String
  Dim r As Variant
  If num < 0 Then
    NTH_ADDRESS = rng.Address
  Else
    For Each r In rng
      If num = 1 Then
        NTH_ADDRESS = r.Address
        Exit Function
      Else
        num = num - 1
      End If
    Next
  End If
End Function

Rem
Rem [名前] T1.統計関数
Rem
Public Function AVERAGE(rng As Range) As Variant
	AVERAGE = WorksheetFunction.AVERAGE(rng)
End Function

Public Function STDEV(rng As Range) As Variant
	STDEV = WorksheetFunction.STDEV(rng)
End Function

Public Function STDERR(rng As Range) As Variant
  STDERR = T1.STDEV(rng) / T1.COUNT(rng)
End Function

Public Function COUNT(rng As Range) As Variant
	COUNT = WorksheetFunction.COUNT(rng)
End Function

Public Function MAX(rng As Range) As Variant
	MAX = WorksheetFunction.MAX(rng)
End Function

Public Function MIN(rng As Range) As Variant
	MIN = WorksheetFunction.MIN(rng)
End Function

Public Function CV(rng As Range) As Variant
	CV = T1.STDEV(rng) / T1.AVERAGE(rng)
End Function

Public Function CVP(rng As Range) As Variant
	CVP = T1.CV(rng) * 100
End Function

Public Function RANK(val1, rng As Range) As Variant
	RANK = WorksheetFunction.RANK(val1, rng)
End Function

Public Function ZVALUE(val1, rng As Range) As Variant
  ZVALUE = (val1 - T1.AVERAGE(rng)) / T1.STDEV(rng)
End Function

Public Function SB_RATIO(rng1 As Range, rng2 As Range) As Variant
	SB_RATIO = T1.AVERAGE(rng2) / T1.AVERAGE(rng1)
End Function

Public Function TC_RATIO(rng1 As Range, rng2 As Range) As Variant
	TC_RATIO = T1.SB_RATIO(rng1, rng2)
End Function

Public Function SN_RATIO(rng1 As Range, rng2 As Range) As Variant
	SN_RATIO = T1.AVERAGE(rng2) / T1.STDEV(rng1)
End Function

Public Function DIFF(rng1 As Range, rng2 As Range) As Variant
	DIFF = T1.AVERAGE(rng2) - T1.AVERAGE(rng1)
End Function

Public Function ZPRIME(rng1 As Range, rng2 As Range) As Variant
	ZPRIME = 1 - 3 * (T1.STDEV(rng1) + T1.STDEV(rng2)) / Abs(T1.AVERAGE(rng2) - T1.AVERAGE(rng1))
End Function

Public Function PERCENTAGE(val1, rng1 As Range, rng2 As Range) As Variant
	PERCENTAGE = 100 * (val1 - T1.AVERAGE(rng1)) / (T1.AVERAGE(rng2) - T1.AVERAGE(rng1))
End Function

Public Function INHIBITION(val1, rng1 As Range, rng2 As Range) As Variant
	INHIBITION = 100 - 100 * (val1 - T1.AVERAGE(rng1)) / (T1.AVERAGE(rng2) - T1.AVERAGE(rng1))
End Function

Public Function RD_HALF(dose As Range, response As Range) As Variant
	Application.Volatile
	Dim i As Integer
	Dim y1 As Double
	Dim y2 As Double
	Dim length As Integer: length = dose.COUNT - 1

	RD_HALF = CVErr(xlErrNA)
	For i = 1 To length
		y1 = response.item(i).Value
		y2 = response.item(i + 1).Value
		If (y1 - 50) * (y2 - 50) < 0 Then
			RD_HALF = dose.item(i).Value + (dose.item(i + 1).Value - dose.item(i).Value) * _
				(50 - response.item(i).Value) / _
				(response.item(i + 1).Value - response.item(i).Value)
		ElseIf y1 = 50 Then
			RD_HALF = dose.item(i)
		ElseIf y2 = 50 Then
			RD_HALF = dose.item(i + 1)
		End If
	Next
End Function

Rem EXPORT OFF



Rem ********************************************************************************
Rem [名前] T1.WELL( ... )
Rem
Rem [用途] 各WELLの情報を得る。
Rem ********************************************************************************
Rem
Rem  WELL( "A1", ... ) : wellpos指定
Rem  WELL( "", ... )   : wellpos指定しないと、同関数を記述した領域全体をplateと見なした際の記述位置がwellposになる。
Rem
Rem  WELL( *, "cpdid" )    : アッセイ結果シートあれば化合物IDが取得できる。
Rem  WELL( *, "role" )     : WELL_ROLE値  WELL( *, "WELL_ROLE", "val" ) と同値
Rem  WELL( *, "conc" )     : CPD_CONC値   WELL( *, "CPD_CONC", "val" ) と同値
Rem  WELL( *, "roleconc" ) : "CPD1@0.1" のような 合成値
Rem  WELL( *, "rc" )    : well位置を配列で返す。　INDEX( T1.WELL(*,"rc"), 0 ) で取得できる。
Rem  WELL( *, "RC" )    : well位置を返す R1C1形式
Rem  WELL( *, "pos" )   : well位置を返す A1形式
Rem  WELL( *, "pos0" )  : well位置を返す A01形式
Rem  WELL( *, "pos00" ) : well位置を返す 0A01形式
Rem  WELL( *, "ROW" )   : 行を示すアルファベット
Rem  WELL( *, "COLUMN" ): 列を示す数値
Rem  WELL( *, (labelname), "val" )     : 名前領域(labelname)から値を得る。
Rem  WELL( *, (labelname), "adr" )     : 名前領域(labelname)からアドレス文字を得る。
Rem  WELL( *, (labelname), "above?", criteria ) : 名前領域(labelname)の値がcriteriaより大きい場合 hit を返す。
Rem  WELL( *, (labelname), "below?", criteria ) : 名前領域(labelname)の値がcriteriaより小さい場合 hit を返す。
Rem  WELL( *, (labelname), "equal?", criteria ) : 名前領域(labelname)の値がcriteriaと同じ値の場合 hit を返す。
Rem  WELL( *, (labelname), "rank", role )   : 名前領域(labelname)の値をrole内で比較したときの順位を返す。
Rem  WELL( *, (labelname), "zvalue", role ) : 名前領域(labelname)の値をrole内で比較したときのzvalue( well値/領域sd )を返す。
Rem  WELL( *, (labelname), "pcnt", role1, role2 )  : 名前領域(labelname)の値をrole1/role2をmin/maxとして求めた相対値を返す。
Rem  WELL( *, (labelname), "inhp", role1, role2 )  : 名前領域(labelname)の値をrole1/role2をmin/maxとして求めた相対値の補数値を返す。
Rem

Public Function well(wellpos As String, labelname As String, Optional func As String = "", Optional ref1 As Variant = Null, Optional ref2 As Variant = Null)
  On Error Resume Next
	Application.Volatile
   
  well = CVErr(xlErrRef)

	Dim ary As Variant
	If wellpos = "" Then
		ary = RESOURCE.GetHere(Application.Caller.Parent.Name, Application.Caller, "rc")
	Else
		ary = RESOURCE.GetRC(wellpos)
	End If

	Dim rw As Integer: rw = ary(0)
	Dim cl As Integer: cl = ary(1)
   
	' 機能計算
	If func = "" Then
		Select Case labelname                     ' WELL( wellpos/rc, labelname )
			Case "cpdid": well = RESOURCE.GetCpdID(Application.Caller.Parent.Name, T1.RC2WELL(rw, cl, "pos"))
			Case "role":  well = Range(PLATE_WELL_ROLE).Cells(rw, cl).Value
			Case "conc":  well = Range(PLATE_COMPOUND_CONC).Cells(rw, cl).Value
			Case "roleconc":
				cnc = Range(PLATE_COMPOUND_CONC).Cells(rw, cl).Value
				well = Range(PLATE_WELL_ROLE).Cells(rw, cl).Value
				If cnc = "0" Or cnc = "" Then well = well & "@" & Range(PLATE_COMPOUND_CONC).Cells(rw, cl).Value
			Case Else:    well = T1.RC2WELL(rw, cl, labelname) ' rc, RC, pos, pos0, pos00
		End Select

	Else
		Dim val0 As Variant
		val0 = Range(labelname).Cells(rw, cl).Value

		Select Case TypeName(ref1)
			Case "Null":                       ' WELL( wellpos/rc, labelname, func )
				Select Case func
					Case "val":  well = val0
					Case "adr":  well = Range(CStr(labelname)).Cells(rw + 1, cl + 1).Address
					Case "pcnt": well = T1.well(wellpos, labelname, "pcnt", "MIN", "MAX")
					Case "inhp": well = T1.well(wellpos, labelname, "inhp", "MIN", "MAX")
				End Select
            
			Case "Integer", "Double", "Range": ' WELL( wellpos/rc, labelname, func, 数値(ref1) )
				Dim criteria As Double
            
				If TypeName(ref1) = "Range" Then
					criteria = ref1.Value
				Else
					criteria = ref1
				End If
				well = ""
				Select Case func
					Case "above?": If val0 > criteria Then well = "hit"
					Case "below?": If val0 < criteria Then well = "hit"
					Case "equal?": If val0 = criteria Then well = "hit"
				End Select
            
			Case "String":
				Select Case TypeName(ref2)
					Case "Null":                 ' WELL( wellpos/rc, labelname, func, 文字(ref1) )
						Dim role As String
						role = Str(ref1)
						Select Case func
							Case "rank":   well = T1.RANK(val0, Range(T1.role(role, labelname, "adr")))
							Case "zvalue": well = val0 / T1.role(role, labelname, "sd")
							Case "ormalize": well = val0 - T1.role(role, labelname, "avr")
						End Select

					Case "String":               ' WELL( wellpos/rc, labelname, func, 文字(ref1), 文字(ref2) )
						Dim val1 As Double
						Dim val2 As Double
						val1 = T1.role(CStr(ref1), labelname, "avr")
						val2 = T1.role(CStr(ref2), labelname, "avr")
						Select Case func
							Case "pcnt":   well = 100 * (val0 - val1) / (val2 - val1)
							Case "inhp":   well = 100 - 100 * (val0 - val1) / (val2 - val1)
						End Select
				End Select
         
			Case Else: well = TypeName(ref1)
         
		End Select
	End If
   
End Function



Rem ********************************************************************************
Rem [名前] T1.LABEL( ... )
Rem
Rem [用途] データラベルの情報を得る。
Rem ********************************************************************************
Rem
Rem  LABEL( (labelname), ... ) : 特定labelの情報を得る。
Rem  LABEL( "", ... )      : 登録label全体の情報を得る。
Rem
Rem  LABEL( "", "plate" ) : plate用の全ラベル値をCSVで得る。
Rem  LABEL( "", "well" )  : well用の全ラベル値をCSVで得る。
Rem  LABEL( "", "table" ) : table用の全ラベル値をCSVで得る。
Rem  LABEL( "", "all" )   : 全ラベル値をCSVで得る。
Rem
Rem  LABEL( (labelname) )              : 特定labelの値を得る。
Rem  LABEL( (labelname), "val" )       : 特定labelの値を得る。
Rem  LABEL( (labelname), "exist" )     : 特定labelの存在を得る。
Rem  LABEL( (labelname), "adr" )       : 特定labelのアドレス文字を得る。
Rem  LABEL( (labelname), "count" )     : 特定labelのcell個数を得る。
Rem  LABEL( (labelname), "rows" )      : 特定labelの行数を得る。
Rem  LABEL( (labelname), "columns" )   : 特定labelの列数を得る。
Rem  LABEL( (labelname), "type" )      : 特定labelのtype(plate,well,table)を得る。
Rem  LABEL( (labelname), "zprime", ("MIN", "MAX") ) : 特定labelのzprime値を得る。 defaultの参照値は"MIN","MAX"の各roleから得る。
Rem  LABEL( (labelname), "sb",     ("MIN", "MAX") ) : 特定labelのsb(tc)値を得る。 defaultの参照値は"MIN","MAX"の各roleから得る。
Rem  LABEL( (labelname), "sn",     ("MIN", "MAX") ) : 特定labelのsn値を得る。 defaultの参照値は"MIN","MAX"の各roleから得る。
Rem

Public Function LABEL(labelname As String, Optional func As String = "", _
                                           Optional role1 As String = "MIN", _
                                           Optional role2 As String = "MAX") As Variant
  LABEL = CVErr(xlErrRef)
   On Error Resume Next
   Application.Volatile
   Dim nam As Variant

   If labelname = "" Then         ' LABEL( "", func )
      Dim csv As String: csv = ""
      For Each nam In Application.Caller.Parent.names
         If (func = "plate" And nam.RefersToRange.COUNT = 1) Or _
            (func = "well" And nam.RefersToRange.COUNT = T1.PLATE("", "type")) Or _
            (func = "table" And nam.Name = ("Template!" & LABEL_TABLE)) Or _
            (func = "all") Then
            pos = InStrRev(nam.Name, "!")
            csv = csv & Mid(nam.Name, pos + 1) & ","
         End If
      Next
      LABEL = Left(csv, Len(csv) - 1)

   Else
      Select Case func
         Case "exist?":
            LABEL = False
            For Each nam In Application.Caller.Parent.names
               If nam.Name = T1.PLATE() & "!" & labelname Then LABEL = True: Exit Function
            Next
         Case "val", "":  LABEL = Range(labelname).Value
         Case "adr":      LABEL = Range(labelname).Address
         Case "count":    LABEL = Range(labelname).COUNT
         Case "rows":     LABEL = Range(labelname).Rows
         Case "columns":  LABEL = Range(labelname).Columns
         Case "type":
            Select Case Range(labelname).COUNT
            Case 1:                    LABEL = "plate"
            Case T1.PLATE("", "type"): LABEL = "well"
            Case Else:                 LABEL = "table"
            End Select
         Case "zprime", "sb", "tc", "sn":
            avr1 = T1.role(role1, labelname, "avr")
            avr2 = T1.role(role2, labelname, "avr")
            sd1 = T1.role(role1, labelname, "sd")
            sd2 = T1.role(role2, labelname, "sd")
            Select Case func
               Case "zprime":   LABEL = 1 - 3 * (sd1 + sd2) / Abs(avr1 - avr2)
               Case "sb", "tc": LABEL = avr2 / avr1
               Case "sn":       LABEL = avr2 / sd1
            End Select
         Case Else:         Set LABEL = Range(labelname)
      End Select
   End If

End Function

Rem ********************************************************************************
Rem [名前] T1.ASSAY( ... )
Rem
Rem [用途] アッセイに関する各情報を得る。
Rem ********************************************************************************
Rem
Rem  ASSAY( "plates" )   : アッセイプレート情報を得る(CSV)
Rem
Public Function ASSAY(func)
  On Error Resume Next
  Dim csv As String
  Dim cl As Variant
  
  Select Case func
    Case "plates":
      For Each cl In Sheets(T1M.ASSAY_SUMMARY_SHEET_NAME).UsedRange.Columns(2).Rows
        If 1 < cl.row And cl.Value <> "" Then csv = csv & cl.Value & ","
      Next
      ASSAY = Left(csv, Len(csv) - 1)
  
  End Select
End Function



Rem ********************************************************************************
Rem [名前] T1.ROLE( ... )
Rem
Rem [用途] プレートマップに基づき各情報を得る。
Rem ********************************************************************************
Rem
Rem  ROLE( "", "roles" ) : 登録ROLEの情報を得る
Rem  ROLE( "", "concs" ) : 登録CONCの情報を得る
Rem
Rem  ROLE( "CPD10", ... )     : 特定ROLEの情報を得る(特定化合物)
Rem  ROLE( "CPD*", ... )      : 同上(全化合物)
Rem  ROLE( "CPD10@10", ... )  : 同上(特定濃度の特定化合物)
Rem
Rem  ROLE( *, "well" )
Rem  ROLE( *, "cpdid" )
Rem  ROLE( *, "cpdconc" )
Rem  ROLE( *, "cpdvol" )
Rem
Rem  ROLE( *, (labelname), "avr" )   : average, stdev, stderr, count, max, min, cv, cvp,
Rem  ROLE( *, (labelname), "adr" )
Rem  ROLE( *, (labelname), "adr", 数値 )
Rem  ROLE( *, (labelname), "val", 数値 )
Rem

Public Function role(rolename As String, labelname As String, Optional func As String = "", Optional param As Integer = 0)
  role = CVErr(xlErrRef)
  On Error Resume Next
  Application.Volatile
  
  With Application.Caller.Parent
    If rolename = "" Then
      Select Case labelname
        Case "roles": role = TSUKUBA_UTIL.EnumrateValues(.Range(PLATE_WELL_ROLE))
        Case "concs": role = TSUKUBA_UTIL.EnumrateValues(.Range(PLATE_COMPOUND_CONC))
      End Select
    Else
      If func = "" Then
        Select Case labelname
          Case "well":    role = RESOURCE.GetRoleWell(.Name, rolename)
          Case "cpdid":   role = RESOURCE.GetCpdID(.Name, CStr(T1.CSV2ARY(RESOURCE.GetRoleWell(.Name, rolename))(0)))
          Case "cpdvol":  role = RESOURCE.GetCpdVol(.Name, RESOURCE.GetRoleWell(.Name, rolename))
          Case "cpdconc": role = RESOURCE.GetCpdConc(.Name, RESOURCE.GetRoleWell(.Name, rolename))
        End Select
      Else
        'Dim rng As Range: Set rng = RESOURCE.GetRoleRange(Application.Caller.Parent.Name, labelname, rolename)
        Dim adr As String: adr = RESOURCE.GetRoleAddress(Application.Caller.Parent.Name, labelname, rolename)
        Select Case func
          Case "avr":   role = T1.AVERAGE(.Range(adr))
          Case "+2sd":  role = T1.AVERAGE(.Range(adr)) + 2 * T1.STDEV(.Range(adr))
          Case "-2sd":  role = T1.AVERAGE(.Range(adr)) - 2 * T1.STDEV(.Range(adr))
          Case "+3sd":  role = T1.AVERAGE(.Range(adr)) + 3 * T1.STDEV(.Range(adr))
          Case "-3sd":  role = T1.AVERAGE(.Range(adr)) - 3 * T1.STDEV(.Range(adr))
          Case "+4sd":  role = T1.AVERAGE(.Range(adr)) + 4 * T1.STDEV(.Range(adr))
          Case "-4sd":  role = T1.AVERAGE(.Range(adr)) - 4 * T1.STDEV(.Range(adr))
          Case "sd":    role = T1.STDEV(.Range(adr))
          Case "se":    role = T1.STDERR(.Range(adr))
          Case "count": role = T1.COUNT(.Range(adr))
          Case "max":   role = T1.MAX(.Range(adr))
          Case "min":   role = T1.MIN(.Range(adr))
          Case "cv":    role = T1.CV(.Range(adr))
          Case "cvp":   role = T1.CVP(.Range(adr))
          Case "val":   role = T1.NTH_VALUE(.Range(adr), param)
          Case "adr":
            If param = 0 Then
              role = T1.NTH_ADDRESS(.Range(adr))
            Else
              role = T1.NTH_ADDRESS(.Range(adr), param)
            End If
        End Select
      End If
    End If
  End With
   
End Function


Rem ********************************************************************************
Rem [名前] T1.PLATE( ... )
Rem
Rem [用途] プレート情報を得る。
Rem ********************************************************************************
Rem
Rem  PLATE() : 関数記述シート名(platename)を得る
Rem
Rem  PLATE( "AR0105535", ... ) : 特定PLATEの情報を得る
Rem  PLATE( "", ... )          : 関数記述シートPLATEの情報を得る
Rem
Rem  PLATE( *, "type" )   : plate type(24,96,384,1536)
Rem  PLATE( *, "reader" ) : plate reader(PHERASTER,FDSS,ENSPIRE,HTFC,FREE)
Rem  PLATE( *, "format" ) : plate format (PRIMARY,CONFIRMATION,DOSE_RESPONSE,FREE)
Rem  PLATE( *, "name" )   : plate name
Rem  PLATE( * )           : plate name
Rem  PLATE( *, "rawdatasheet" ) : plateに対応するrawdataのsheet名
Rem  PLATE( *, "rawdatafile" )  : plateに対応するrawdataのファイル名
Rem  PLATE( *, "labels" )       : plateに登録されている全label名
Rem  PLATE( *, "platelabels" )  : plateに登録されている全plate label名
Rem  PLATE( *, "welllabels" )   : plateに登録されている全well label名
Rem  PLATE( *, "tablelabels" )  : plateに登録されている全table label名
Rem
Rem  PLATE( *, (labelname) )
Rem  PLATE( *, (labelname), (well) )
Rem

Public Function PLATE(Optional platename As String = "", Optional func As String = "name", _
											Optional param As Variant = Null)
	PLATE = ""
	On Error Resume Next
	Application.Volatile
   
	Dim sht As String
	If platename = "" Then platename = Application.Caller.Parent.Name
	sht = "'" & platename & "'!"
	Dim nam As Variant
	Dim rw As Variant
   
	Select Case func
		Case "type":         PLATE = Range(sht & PLATE_TYPE).Value
		Case "reader":       PLATE = Range(sht & PLATE_READER).Value
		Case "format":       PLATE = Range(sht & PLATE_FORMAT).Value
		Case "name":         PLATE = platename
		Case "labels", "platelabels", "welllabels", "tablelabel":
			Dim csv As String: csv = ""
			For Each nam In Sheets(platename).names
				If (func = "platelabels" And nam.RefersToRange.COUNT = 1) Or _
					 (func = "welllabels" And nam.RefersToRange.COUNT = T1.PLATE(platename, "type")) Or _
					 (func = "tablelabel" And nam.Name = (platename & "!" & LABEL_TABLE)) Or _
					 (func = "labels") Then
					pos = InStrRev(nam.Name, "!")
					csv = csv & Mid(nam.Name, pos + 1) & ","
				End If
			Next
			PLATE = Left(csv, Len(csv) - 1)
         
		Case "rawdatasheet": PLATE = "(raw)" & platename
		Case "rawdatafile":
			With Sheets(T1M.ASSAY_SUMMARY_SHEET_NAME)
				For Each rw In .UsedRange.Rows
					If rw.Cells(1, 2).Value = platename Then
						PLATE = rw.Cells(1, 1).Value: Exit Function
					End If
				Next
			End With
		Case Else:
			Select Case TypeName(param)
				Case "Null":   PLATE = Sheets(platename).Range(func).Value
				Case "String":
					Select Case param
						Case "adr": PLATE = Sheets(platename).Range(func).Address
						Case Else: PLATE = Sheets(platename).Range(func).Cells(T1.well(CStr(param), "rc")(0), T1.well(CStr(param), "rc")(1)).Value
					End Select
			End Select
	End Select
End Function





Rem ********************************************************************************
Rem [名前] T1.TABLE( ... )
Rem
Rem [用途] ユーザーテーブル情報を得る。
Rem ********************************************************************************
Rem
Rem  TABLE( "name" )
Rem  TABLE( "items" )
Rem  TABLE( "records" )
Rem  TABLE( (items), (record_num) )
Rem

Public Function TABLE(func As String, Optional param As Integer = 0)
	TABLE = CVErr(xlErrRef)
	On Error Resume Next
	Application.Volatile
	Dim cl As Variant
   
	Select Case func
		Case "name": TABLE = LABEL_TABLE
		Case "items":
			Dim csv As String
			For Each cl In Range("TABLE").Rows(1).Columns
				If cl.Value <> "" Then csv = csv & cl.Value & ","
			Next
			TABLE = Left(csv, Len(csv) - 1)
		Case "records": TABLE = Range("TABLE").Rows.COUNT - 1
		Case Else:
			For Each cl In Range("TABLE").Rows(1).Columns
				If cl.Value = func Then
					TABLE = Range("TABLE").Cells(param + 1, cl.Column - 1).Value
					Exit Function
				End If
			Next
	End Select
End Function




Rem ****************************************************************************************************************************************************************
Rem 生データ参照
Rem ****************************************************************************************************************************************************************

Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem インターフェイス関数
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Rem
Rem  - 必須項目
Rem  PLATEREADER_INFO( "date" )
Rem  PLATEREADER_INFO( "time" )
Rem  PLATEREADER_INFO( "assay" )
Rem

Public Function PLATEREADER_INFO(param As String) As String
	Select Case T1.PLATE("", "reader")
		Case "PHERASTER":  PLATEREADER_INFO = T1.PHERASTER_INFO(param)
		Case "FDSS":       PLATEREADER_INFO = T1.FDSS_INFO(param)
		Case "ENSPIRE":    PLATEREADER_INFO = T1.ENSPIRE_INFO(param)
		Case "EZREADER":   PLATEREADER_INFO = T1.EZREADER_INFO(param)
		Case "HTFC":       PLATEREADER_INFO = T1.HTFC_INFO(param)
		Case "FREE":       PLATEREADER_INFO = T1.FREE_INFO(param)
	End Select
End Function


Rem
Rem  - 必須項目
Rem  PLATEREADER_VALUE( (well), ... ) : well位置指定してデータを得る
Rem  PLATEREADER_VALUE( "", ... )     : well位置は記述位置依存
Rem
Rem  PLATEREADER_INFO( *, (id) ) : id で区別されるデータを得る
Rem

Public Function PLATEREADER_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	Select Case T1.PLATE("", "reader")
		Case "PHERASTER":  PLATEREADER_VALUE = T1.PHERASTER_VALUE(wellpos, id, param1, param2, param3)
		Case "FDSS":       PLATEREADER_VALUE = T1.FDSS_VALUE(wellpos, id, param1, param2, param3)
		Case "ENSPIRE":    PLATEREADER_VALUE = T1.ENSPIRE_VALUE(wellpos, id, param1, param2, param3)
		Case "EZREADER":   PLATEREADER_VALUE = T1.EZREADER_VALUE(wellpos, id, param1, param2, param3)
		Case "HTFC":       PLATEREADER_VALUE = T1.HTFC_VALUE(wellpos, id, param1, param2, param3)
		Case "FREE":       PLATEREADER_VALUE = T1.FREE_VALUE(wellpos, id)
	End Select
End Function

Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem FREE
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function FREE_INFO(adr As String) As String
	On Error GoTo FREE_INFO_ERR
	Application.Volatile

	FREE_INFO = Sheets(T1.PLATE("", "rawdatasheet")).Range(adr).Value

FREE_INFO_ERR:
	FREE_INFO = CVErr(xlErrRef)
End Function

Public Function FREE_VALUE(wellpos As String, id As String) As Variant
	Application.Volatile
	On Error GoTo FREE_VALUE_ERR
   
	wellpos = T1.well(wellpos, "pos0")
	Select Case id
		Case "pos": FREE_VALUE = wellpos
		Case Else: FREE_VALUE = ""
	End Select
   
FREE_VALUE_ERR:
	FREE_VALUE = CVErr(xlErrRef)
End Function


Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem EZREADER
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function EZREADER_INFO(param As String) As String
	On Error GoTo EZREADER_INFO_ERR
	Application.Volatile

	With Sheets(T1.PLATE("", "rawdatasheet"))
		Dim dat As Variant
		dat = Split(.Range("Q2").Value, " ")
		Select Case param
			Case "date":   EZREADER_INFO = T1.DATE_ID(dat(0))
			Case "time":   EZREADER_INFO = T1.TIME_ID(dat(1))
			Case "assay":  EZREADER_INFO = .Range("T2").Value
			Case "chipid":   EZREADER_INFO = .Range("T2").Value
			Case "filepath": EZREADER_INFO = .Range("O2").Value
		End Select
		Exit Function
	End With

EZREADER_INFO_ERR:
	EZREADER_INFO = CVErr(xlErrRef)
End Function

Public Function EZREADER_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	Application.Volatile
	On Error GoTo EZREADER_VALUE_ERR
  Dim i As Integer: Dim j As Integer

	wellpos = T1.well(wellpos, "pos0")

	With Sheets(T1.PLATE("", "rawdatasheet"))
		For i = 1 To .UsedRange.Columns.COUNT
			If 0 < InStr(.Cells(1, i).Value, id) Then
				For j = 1 To .UsedRange.Rows.COUNT
					If 0 < InStr(.Cells(j, 2).Value, wellpos) Then
						EZREADER_VALUE = .Cells(j, i).Value
						Exit Function
					End If
				Next j
			End If
		Next i
	End With
EZREADER_VALUE_ERR:
	EZREADER_VALUE = CVErr(xlErrRef)
End Function


Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem ENSPIRE
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Function ENSPIRE_INFO(param As String) As String
   On Error GoTo ENSPIRE_INFO_ERR
   Application.Volatile

   With Sheets(T1.PLATE("", "rawdatasheet"))
      Dim dateTime As String
      Dim pos As Integer
        Dim i As Integer
      dateTime = .Cells(32, 5).Value
      pos = InStr(dateTime, " ")
      Select Case param
         Case "date":     ENSPIRE_INFO = T1.DATE_ID(Left(CStr(dateTime), pos))
         Case "time":     ENSPIRE_INFO = T1.TIME_ID(Mid(CStr(dateTime), pos + 1, 100))
         Case "assay":    ENSPIRE_INFO = Replace(.Cells(36, 5).Value, "Testname: ", "")
         Case "testname": ENSPIRE_INFO = Replace(.Cells(36, 5).Value, "Testname: ", "")
         Case Else
            For i = 1 To 300
               If 0 < InStr(.Cells(i, 1).Value, nam) Then
                  ENSPIRE_INFO = .Cells(i, 5)
               End If
            Next i
      End Select
      Exit Function
   End With
   
ENSPIRE_INFO_ERR:
   ENSPIRE_INFO = CVErr(xlErrRef)
End Function

Public Function ENSPIRE_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	On Error GoTo ENSPIRE_VALUE_ERR
	Application.Volatile

  Dim arr As Variant
	Dim rw As Integer
  Dim cl As Integer
	arr = T1.well(wellpos, "rc")
	rw = arr(0) - 1
	cl = arr(1) - 1

	With Sheets(T1.PLATE("", "rawdatasheet"))
		For i = 1 To .UsedRange.Rows.COUNT
			If 0 < InStr(.Cells(i, 1).Value, id) Then
				For j = 1 To 10
					If 0 < Len(.Cells(i + j, 4).Value) Then
						ENSPIRE_VALUE = .Cells(i + j + rw + 1, cl + 2).Value
						Exit Function
					End If
				Next j
			End If
		Next i
	End With
ENSPIRE_VALUE_ERR:
	ENSPIRE_VALUE = CVErr(xlErrRef)
End Function


Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem HTFC
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function HTFC_INFO(param As String) As String
	On Error GoTo HTFC_INFO_ERR
	Application.Volatile

	With Sheets(T1.PLATE("", "rawdatasheet"))
		Select Case param
			Case "date":       Dim dat As Variant
				dat = Split(Replace(.Range("G1").Value, "Export Date: ", ""), "/")
				HTFC_INFO = T1.DATE_ID(dat(2) & "/" & dat(0) & "/" & dat(1))
			Case "time":       HTFC_INFO = T1.TIME_ID(Replace(.Range("H1").Value, "Export Time: ", ""))
			Case "assay":   HTFC_INFO = Replace(.Range("B1").Value, "Analysis: ", "")

			Case "experiment": HTFC_INFO = Replace(.Range("A1").Value, "Experiment: ", "")
			Case "name":       HTFC_INFO = Replace(.Range("A1").Value, "Experiment: ", "")
			Case "analysis":   HTFC_INFO = Replace(.Range("B1").Value, "Analysis: ", "")
			Case "user":       HTFC_INFO = Replace(.Range("C1").Value, "User: ", "")
			Case "plate":      HTFC_INFO = Replace(.Range("D1").Value, "Plate: ", "")
			Case "platetype":  HTFC_INFO = Replace(.Range("E1").Value, "Plate Type: ", "")
			Case "plateorder": HTFC_INFO = Replace(.Range("F1").Value, "Plate Order: ", "")
		End Select
	End With
   
	Exit Function
HTFC_INFO_ERR:
	HTFC_INFO = CVErr(xlErrRef)
End Function


Public Function HTFC_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	Application.Volatile
	On Error GoTo HTFC_VALUE_ERR

  Dim i As Integer
	Dim rw As Integer
	Dim cl As Integer
	Dim arr As Variant
	arr = T1.well(wellpos, "rc")
	rw = arr(0) - 1
	cl = arr(1) - 1

	With Sheets(T1.PLATE("", "rawdatasheet"))
		For i = 1 To .UsedRange.Rows.COUNT
			If 0 < InStr(.Cells(i, 1).Value, id) Then
				HTFC_VALUE = .Cells(i + rw + 2, cl + 2).Value
				Exit Function
			End If
		Next i
	End With
HTFC_VALUE_ERR:
	HTFC_VALUE = CVErr(xlErrRef)
End Function



Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem PHERASTER
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function PHERASTER_INFO(param As String) As String
	On Error GoTo PHERASTER_INFO_ERR
	Application.Volatile

	With Sheets(T1.PLATE("", "rawdatasheet"))
		Dim dateTime As Variant
		dateTime = Split(.Cells(2, 1).Value, " ")

		Select Case param
			Case "date":     PHERASTER_INFO = T1.DATE_ID(CStr(dateTime(1)))
			Case "time":     PHERASTER_INFO = T1.TIME_ID(CStr(dateTime(4)))
			Case "assay":    PHERASTER_INFO = Replace(.Cells(1, 1).Value, "Testname: ", "")
			Case "testname": PHERASTER_INFO = Replace(.Cells(1, 1).Value, "Testname: ", "")
		End Select
      
	End With

	Exit Function
PHERASTER_INFO_ERR:
	PHERASTER_INFO = CVErr(xlErrRef)
End Function


Public Function PHERASTER_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	On Error GoTo PHERASTER_VALUE_ERR
	Application.Volatile

	Dim rw As Integer
  Dim cl As Integer
  Dim arr As Variant
	arr = T1.well(wellpos, "rc")
	rw = arr(0) - 1
	cl = arr(1) - 1

	With Sheets(T1.PLATE("", "rawdatasheet"))
		For i = 1 To .UsedRange.Rows.COUNT
			If 0 < InStr(.Cells(i, 1).Value, id) Then
				For j = 1 To 10
					If 0 < Len(.Cells(i + j, 4).Value) Then
						PHERASTER_VALUE = .Cells(i + j + rw, cl + 1).Value
						Exit Function
					End If
				Next j
			End If
		Next i
	End With

PHERASTER_VALUE_ERR:
	PHERASTER_VALUE = CVErr(xlErrRef)
End Function


Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Rem FDSS
Rem ::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Function FDSS_INFO(param As String) As String
	On Error GoTo FDSS_INFO_ERR
	Application.Volatile

	Dim rstname As String
	rstname = T1.PLATE("", "rawdatasheet")

	Select Case param
		Case "date":         FDSS_INFO = T1.DATE_ID(FDSS_INFO_sub(rstname, "B", "Date : "))
		Case "time":         FDSS_INFO = T1.TIME_ID(Strings.Format(FDSS_INFO_sub(rstname, "B", "Time : "), "hh:mm:ss"))
		Case "assay":        FDSS_INFO = T1.FDSS_INFO_sub(rstname, "C", "Protocol Name")
		Case "protocolname": FDSS_INFO = T1.FDSS_INFO_sub(rstname, "C", "Protocol Name")
		Case "name":         FDSS_INFO = T1.FDSS_INFO_sub(rstname, "C", "Protocol Name")
		Case "datafilename": FDSS_INFO = T1.FDSS_INFO_sub(rstname, "B", "Data File Name : ")
		Case "sensitivity":  FDSS_INFO = T1.FDSS_INFO_sub(rstname, "C", "Sensitivity")
	End Select

	Exit Function
FDSS_INFO_ERR:
	FDSS_INFO = CVErr(xlErrRef)
End Function

Private Function FDSS_INFO_sub(stname As String, row As String, key As String) As String
  Dim i As Integer
	With Sheets(stname)
		If row = "B" Then
			For i = 1 To .UsedRange.Rows.COUNT
				If 0 < InStr(.Cells(i, 2).Value, key) Then
					FDSS_INFO_sub = .Cells(i, 3).Value
					Exit Function
				End If
			Next i
		ElseIf row = "C" Then
			For i = 1 To .UsedRange.Rows.COUNT
				If 0 < InStr(.Cells(i, 3).Value, key) Then
					FDSS_INFO_sub = .Cells(i, 4).Value
					Exit Function
				End If
			Next i
		End If
	End With
End Function

Public Function FDSS_VALUE(wellpos As String, id As String, Optional param1 As Variant = Null, Optional param2 As Variant = Null, Optional param3 As Variant = Null) As Variant
	On Error GoTo FDSS_VALUE_ERR
	Application.Volatile
   
	Dim arr As Variant
	Dim wellrow As Double
	Dim wellcol As Double
	arr = T1.well(wellpos, "rc")
	wellrow = arr(0) - 1
	wellcol = arr(1) - 1

	Dim timerow As Double
	With Worksheets(T1.PLATE("", "rawdatasheet"))
		For timerow = 1 To 50
			If InStr(.Cells(timerow, 2).Value, "No.") Then Exit For
		Next timerow
	End With
   
	If IsNull(param2) Then
		Dim tp As Double
		tp = CDbl(param1) * 1000
		FDSS_VALUE = T1.FDSS_VALUE_1tp(wellrow, wellcol, id, timerow, tp)

	Else
		Dim tp1 As Double
		Dim tp2 As Double
		Dim func As String
		tp1 = CDbl(param1) * 1000
		tp2 = CDbl(param2) * 1000
		func = CStr(param3)
		FDSS_VALUE = T1.FDSS_VALUE_2tp(wellrow, wellcol, id, timerow, tp1, tp2, func)
	End If
	Exit Function
   
FDSS_VALUE_ERR:
	FDSS_VALUE = CVErr(xlErrRef)
End Function


Private Function FDSS_VALUE_1tp(wellrow As Double, wellcol As Double, reftype As String, timerow As Double, timepoint As Double) As Variant
	On Error GoTo FDSS_VALUE_1tp_ERR
	Application.Volatile

	Dim timecol As Double
	Dim rownum As Double
	Dim typeoffset As Double

	With Worksheets(T1.PLATE("", "rawdatasheet"))
		For timecol = 5 To .UsedRange.Columns.COUNT
			If timepoint <= .Cells(timerow, timecol).Value Then Exit For
		Next timecol
		For rownum = 1 To 10
			If .Cells(timerow + 1, 4).Value = .Cells(timerow + rownum + 1, 4).Value Then Exit For
		Next rownum
		If reftype = "" Then
			typeoffset = 0
		Else
			For typeoffset = 0 To 10
				If InStr(.Cells(timerow + 1 + typeoffset, 4).Value, reftype) Then Exit For
			Next typeoffset
		End If

		FDSS_VALUE_1tp = .Cells(timerow + 1 + typeoffset + (wellrow * 24 + wellcol) * rownum, timecol).Value
	End With

	Exit Function
FDSS_VALUE_1tp_ERR:
	FDSS_VALUE_1tp = CVErr(xlErrRef)
End Function


Private Function FDSS_VALUE_2tp(wellrow As Double, wellcol As Double, reftype As String, timerow As Double, timepoint As Double, timepoint2 As Double, func As String) As Variant
	On Error GoTo FDSS_VALUE_2tp_ERR
	Application.Volatile
	Dim tim As Long
	Dim timecol As Double
	Dim timecol1 As Double
	Dim timecol2 As Double
	Dim rownum As Double
	Dim typeoffset As Double
	Dim datarow As Double
	Dim rawdatasht As String

	With Sheets(T1.PLATE("", "rawdatasheet"))
		For timecol1 = 5 To .UsedRange.Columns.COUNT
			If timepoint <= .Cells(timerow, timecol1).Value Then Exit For
		Next timecol1
		For timecol2 = 5 To .UsedRange.Columns.COUNT
			If timepoint2 <= .Cells(timerow, timecol2).Value Then Exit For
		Next timecol2
		If timecol2 < timecol1 Then
			timecol = timecol1
			timecol2 = timecol1
			timecol1 = timecol
		End If

		For rownum = 1 To 10
			If .Cells(timerow + 1, 4).Value = .Cells(timerow + rownum + 1, 4).Value Then Exit For
		Next rownum
      
		If reftype = "" Then
			typeoffset = 0
		Else
			For typeoffset = 0 To 10
				If InStr(.Cells(timerow + 1 + typeoffset, 4).Value, reftype) Then Exit For
			Next typeoffset
		End If

		datarow = timerow + 1 + typeoffset + (wellrow * 24 + wellcol) * rownum

		If func = "diff" Then
			FDSS_VALUE_2tp = .Cells(datarow, timecol2).Value - .Cells(datarow, timecol1).Value
		Else
			rawdatasht = "'(raw)" & Application.Caller.Parent.Name & "'!"
			Set DataRange = .Range(.Cells(datarow, timecol1), .Cells(datarow, timecol2))
			Set TimeRange = .Range(.Cells(timerow, timecol1), .Cells(timerow, timecol2))
			Select Case func
				Case "adr":        FDSS_VALUE_2tp = rawdatasht & DataRange.Address
				Case "timeadr":    FDSS_VALUE_2tp = rawdatasht & TimeRange.Addressrange
				Case "avr":        FDSS_VALUE_2tp = T1.AVERAGE(DataRange)
				Case "sd":         FDSS_VALUE_2tp = T1.STDEV(DataRange)
				Case "se":         FDSS_VALUE_2tp = T1.STDEV(DataRange) / T1.COUNT(DataRange)
				Case "count":      FDSS_VALUE_2tp = T1.COUNT(DataRange)
				Case "max":        FDSS_VALUE_2tp = T1.MAX(DataRange)
				Case "min":        FDSS_VALUE_2tp = T1.MIN(DataRange)
				Case "extent":     FDSS_VALUE_2tp = T1.MAX(DataRange) - T1.MIN(DataRange)
				Case "slope":      FDSS_VALUE_2tp = WorksheetFunction.Slope(DataRange, TimeRange)
				Case "intercept":  FDSS_VALUE_2tp = WorksheetFunction.Intercept(DataRange, TimeRange)
				Case "rsq":        FDSS_VALUE_2tp = WorksheetFunction.RSq(DataRange, TimeRange)
			End Select
		End If
		Exit Function
	End With

FDSS_VALUE_2tp_ERR:
	FDSS_VALUE_2tp = CVErr(xlErrRef)
End Function






















