Option Explicit

		' ------------------------------------------------------------
		' 代入
		' ------------------------------------------------------------
		Public Sub Substitution(ByRef distination, ByRef source)
			If IsObject(source) Then
				Set distination = source
			Else
				distination = source
			End If
		End Sub

		Public Function IIF(ByVal expr, ByVal TrueResult, ByVal FalseResult)
			If expr Then
				Substitution IIF, TrueResult
			Else
				Substitution IIF, FalseResult
			End If
		End Function

		Public Function IsString(ByVal var)
			If Not (IsObject(var) Or IsNull(var)) Then
				If IsNumeric(var) Then
					IsString = ((var+1)-1) <> var
				Else
					IsString = True
				End If
			End If
		End Function

		' ------------------------------------------------------------
		' YYYYMMDD編集
		' ------------------------------------------------------------
		Function YYYYMMDD(dt)
			Dim y, m, d
			y = "0" & Year(dt)
			m = "0" & Month(dt)
			d = "0" & Day(dt)
			YYYYMMDD = Right(y,4) & Right(m,2) & Right(d,2)
		End Function

		' ------------------------------------------------------------
		' HHNNSS編集
		' ------------------------------------------------------------
		Function HHNNSS(dt)
			Dim h, n, s
			h = "0" & hour(dt)
			n = "0" & Minute(dt)
			s = "0" & Second(dt)
			HHNNSS = Right(h,2) & Right(n,2) & Right(s,2)
		End Function

		' ------------------------------------------------------------
		' YYYYMMDDHHMMSS
		' ------------------------------------------------------------
		Function YMDHMS(dt)
			YMDHMS = YYYYMMDD(dt) & HHNNSS(dt)
		End Function

		' ------------------------------------------------------------
		' "YYYY/MM/DD HH:MM:SS"
		' ------------------------------------------------------------
		Function ToDTS(dt)
			Dim src, d, t
			src = YYYYMMDD(dt)
			d = Mid(src, 1, 4) & "/" & Mid(src, 5, 2) & "/" & Mid(src, 7, 2)
			src = HHNNSS(dt)
			t = Mid(src, 9, 2) & ":" & Mid(src, 11, 2) & ":" & Mid(src, 13, 2)
			ToDTS = d & " " & t
		End Function

		' ----------------------------------------
		' 10桁の数字文字列
		' ----------------------------------------
		Function Numeric10(num)
			Numeric10 = Right("0000000000" & CStr(num), 10)
		End Function

		' ------------------------------------------------------------
		' 文書編集
		' ------------------------------------------------------------
		Function MessageFormat(format, arr)
			Dim dic, par
			Set dic = CreateObject("Scripting.Dictionary")
			Set par = CreateObject("Scripting.Dictionary")

			' パラメータ
			For idx=LBound(arr) To UBound(arr)
				par.Add Cstr(idx+1), arr(idx)
			Next

			' 正規表現で書式ブロックを解析
			Dim regEx, Matches, text
			Set regEx = New RegExp
			regEx.Pattern = "\$\{([0-9]+)[:]?([0-9]*)\}"
			regEx.IgnoreCase = True
			regEx.Global = True
			text = sun(format)
			Set Matches = regEx.Execute(text)

			' 書式ブロックをパラメータで置換
			Dim RetStr, pos, Match, SubMatch, fmt, idx, key
			RetStr = ""
			pos = 1
			For Each Match in Matches
				' テキストブロックを追加
				RetStr = RetStr & unsun(Mid(text, pos, Match.FirstIndex+1 - pos))
				pos = Match.FirstIndex + Match.Length + 1
				' 書式に従ってパラメータをフォーマット
				key = Match.Value
				If Not dic.Exists(key) Then
					Set fmt = new formatter
					fmt.Init Match, par
					dic.Add key, fmt
				Else
					Set fmt = dic.Item(key)
				End IF
				' フォーマット済みパラメータを追加
				RetStr = RetStr & fmt.value
			Next
			' 残ったテキストブロックを追加
			RetStr = RetStr & unsun(Mid(text, pos))
			MessageFormat = RetStr
		End Function

' --------------------------------------------------------------------------------

		'Const fmt = "<${1} ${2:2} ${3}\n\t${1}\${3}\\${2:4} ${4>"
		'MsgBox(fmt & vbLf & MessageFormat(fmt, Array("aaa","bbb","ccc")))
		Private Function sun(str)
			sun = str
			sun = Replace(sun, "&", "&amp;")
			sun = Replace(sun, "<", "&lt;")
			sun = Replace(sun, ">", "&gt;")
			sun = Replace(sun, "\\", "<ESC>")
			sun = Replace(sun, "\n", "<RET>")
			sun = Replace(sun, "\t", "<TAB>")
			sun = Replace(sun, "\$", "<DOL>")
		End Function

		Private Function unsun(str)
			unsun = str
			unsun = Replace(unsun, "<DOL>", "$")
			unsun = Replace(unsun, "<TAB>", vbTab)
			unsun = Replace(unsun, "<RET>", vbCrLf)
			unsun = Replace(unsun, "<ESC>", "\")
			unsun = Replace(unsun, "&gt;", ">")
			unsun = Replace(unsun, "&lt;", "<")
			unsun = Replace(unsun, "&amp;", "&")
		End Function

' --------------------------------------------------------------------------------

	Class formatter
		Dim key
		Dim paramnum
		Dim width
		Dim value

		Sub Init(Match, params)
			key = Match.Value
			paramnum = CInt("0" & Match.SubMatches(0))
			width = CInt("0" & Match.SubMatches(1))
			Dim prmkey
			prmkey = CStr(paramnum)
			If params.Exists(prmkey) Then
				' 対応バラメータあり
				value = params.Item(prmkey)
				if width>LenB(value) then
					value = Space(width-LenB(value)) & value
				End If
			Else
				' 対応バラメータなし
				value = Space(width)
			End IF
		End Sub
	End Class

' --------------------------------------------------------------------------------

		' ------------------------------------------------------------
		' ファイルに UTF-8 エンコーディングで保存します
		' (元ネタ) http://d.hatena.ne.jp/replication/20091117/1258418243
		' ------------------------------------------------------------
		Sub SaveAsUTF8(filename, text)
			On Error Resume Next

			' ADODB.Streamのモード
			Dim adTypeBinary : adTypeBinary = 1
			Dim adTypeText : adTypeText = 2
			Dim adSaveCreateOverWrite : adSaveCreateOverWrite = 2

			' ADODB.Streamを作成
			Dim pre : Set pre = CreateObject("ADODB.Stream")
			' 最初はテキストモードでUTF-8で書き込む
			pre.Type = adTypeText
			pre.Charset = "UTF-8"
			pre.Open()
			pre.WriteText(text)
			' バイナリモードにするためにPositionを一度0に戻す
			' Readするためにはバイナリタイプでないといけない
			pre.Position = 0
			pre.Type = adTypeBinary
			' Positionを3にしてから読み込むことで最初の3バイトをスキップする
			' つまりBOMをスキップします
			pre.Position = 3
			Dim bin : bin = pre.Read()
			pre.Close()

			' 読み込んだバイナリデータをバイナリデータとしてファイルに出力する
			' ここは一般的な書き方なので説明を省略
			Dim stm : Set stm = CreateObject("ADODB.Stream")
			stm.Type = adTypeBinary
			stm.Open()
			stm.Write(bin)
			stm.SaveToFile filename, adSaveCreateOverWrite ' force overwrite
			stm.Close()

		End Sub
